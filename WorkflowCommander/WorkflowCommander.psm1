﻿#########################################################################################
# WorkflowCommander, copyrighted by Joel Wiesmann, 2016
# 
# Warm welcome to my code, whatever wisdom you try to find here.
#
# This file is part of WorkflowCommander.
# See http://www.binpress.com/license/view/l/9b201d0301d19b7bd87a3c7c6ae34bcd for full license details.
#
# THIS SOFTWARE IS PROVIDED BY THE AUTHOR ``AS IS'' AND ANY EXPRESS OR IMPLIED WARRANTIES, 
# INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
# A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY DIRECT,
# INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
# LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS;
# OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
# CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY
# WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#########################################################################################

$VerbosePreference = 'Continue'

# Default values for failure / empty / OK result
$WFCFAILURE = 'FAIL'
$WFCEMPTY   = 'EMPTY'
$WFCOK      = 'OK'

#########################################################################################
# Published functions
#########################################################################################
function Search-aeObject {
  <#
      .SYNOPSIS
      Equals the AE "search for object".

      .DESCRIPTION
      search-aeObject returns the same information as known from the GUI. The output can be piped to
      export-aeObject or to another search-aeObject.

      .PARAMETER aeConnection
      WorkflowCommander Connection Object (new-aeConnection).

      .PARAMETER name
      Name of the object to search. Supports wildcards like the GUI search. Defaults to "*"

      .PARAMETER path
      Path where the object is stored. By default "/". 

      .PARAMETER objType
      Limits the object types that should be searched for. By default, all object types are searched.

      .PARAMETER text
      Text to search for in process, archivekeys etc..

      .PARAMETER textType
      Limit the places to search for the text specified with -text.

      .PARAMETER noRecursiveSearch
      Switch to disable recursive search feature. This limits the search to the specified -path.

      .PARAMETER dateSearch
      Enable search based on date/time constraints. Choose between create/modify/used.

      .PARAMETER fromDateTime
      From date / time. Input as [datetime] or 'YYYY-MM-DD HH:MM:SS'

      .PARAMETER toDateTime
      To date / time. Input as [datetime] or 'YYYY-MM-DD HH:MM:SS'

      .EXAMPLE
      search-aeObject -ae $ae -name "*JOB*" -path "/PRODUCTION" -objType JOBS
      Searches for objects that name matches "*JOB*" and are stored below /PRODUCTION and are of type JOBS.

      .LINK
      http://workflowcommander.blogspot.com

      .OUTPUTS
      Array of findings.
  #>
  Param(
    [Parameter(mandatory,HelpMessage='AE connection object returned by new-aeConnection.')]
    [Alias('ae')]
    [WFC.Core.WFCConnection]$aeConnection,
    [Parameter(ValueFromPipeline,ValueFromPipelineByPropertyName)]
    [string]$name       = '*',
    [Parameter(ValueFromPipelineByPropertyName)]
    [string]$path       = $null,
    [Parameter(ValueFromPipelineByPropertyName)]
    [Alias('type')]
    [ValidateSet('JOBS',
        'JOBP','CALE','CALL','CITC',
        'CLNT','CODE','CONN','CPIT',
        'DASH','DOCU','EVNT','FILTER',
        'FOLD','HOST','HOSTG','HSTA',
        'JOBF','JOBG','JOBI','JOBQ',
        'JSCH','LOGIN','PERIOD','PRPT',
        'QUEUE','SCRI','SERV','STORE',
        'SYNC','TZ','USER','USERG',
        'VARA','XSL','Executeable'
    )]
    [string[]]$objType  = $null,
    [string]$text       = $null,
    [ValidateSet('archive','process','title','documentation','varakey','varavalue','varaall')]  
    [string[]]$textType = $null,
    [switch]$NonRecursiveSearch,
    [ValidateSet('noConstraint', 'Created', 'Modified', 'Used')]
    [string]$dateSearch = 'noConstraint',
    [datetime]$fromDateTime = [DateTime]::Today,
    [datetime]$toDateTime = [datetime]::Now,
    [switch]$searchForUsage
  )

  begin {
    Write-Debug -Message '** Search-aeObject start'
    $resultSet = @()
  }
  
  process {
    $subResultSet = @()
    $search = [com.uc4.communication.requests.SearchObject]::new()
    
    ###################
    # Search for object or for usage?
    ###################
    $search.setSearchUseOfObjects($searchForUsage)

    ###################
    # Objecttype filter
    ###################
    # Depending on whether we want to filter for object types or not we select all or only specific object types.
    if ($objType -eq $null) {
      $search.selectAllObjectTypes() 
    }
    else {
      $search.unselectAllObjectTypes()

      # This activates the object type search selection. To support new object types, just add the new type 
      # in the parameter validation. i.E. JOBS => $search.setTypeJOBS($true)
      foreach ($typeFilter in $objType) {
        $search.('setType' + $typeFilter)($true)
      }
    }
  
    ###################
    # Date filter, if any.
    ###################
    if ($dateSearch -ne 'noContraint') {
      $fromDateTimeFilter = [com.uc4.api.Datetime]::new($fromDateTime.ToString('yyyy-MM-dd HH:mm:ss'))
      $toDateTimeFilter   = [com.uc4.api.Datetime]::new($toDateTime.ToString('yyyy-MM-dd HH:mm:ss'))
    
      switch ($dateSearch) {
        'Modified' { $search.setDateSelectionModified($fromDateTimeFilter, $toDateTimeFilter) }
        'Created'  { $search.setDateSelectionCreated($fromDateTimeFilter, $toDateTimeFilter) }
        'Used'     { $search.setDateSelectionUsed($fromDateTimeFilter, $toDateTimeFilter) }
      }
    }

    ###################
    # Text search
    ###################
    # Unavailable when NOT searching for usage
    if (! $searchForUsage) {
      # If we do text-based search, we either search all fields or only specific ones. We handle this the same way
      # as we support object types. By default we search in all fields but we can limit the search. Also useful -
      # we combine here the search-in-VARA logic.
      if ($text -ne $null) {
        # Search in all fields and VARAs. At least I am doing this usually because of common lazyness, however to prevent 
        # performance impact in proper scripts or large environments this should not be done.
        if ($textType -eq $null) {
          $searchArchive       = $true
          $searchProcess       = $true
          $searchTitle         = $true
          $searchDocumentation = $true
          $searchVARAKey       = $true
          $searchVARAValue     = $true
          $searchVARAAllCols   = $true
        }
        else {
          foreach ($textFilter in $textType) {
            switch ($textFilter) {
              'archive'       { $searchArchive       = $true }
              'process'       { $searchProcess       = $true }
              'title'         { $searchTitle         = $true }
              'documentation' { $searchDocumentation = $true }
              'varakey'       { $searchVARAKey       = $true }
              'varavalue'     { $searchVARAValue     = $true }
              'varaall'       { $searchVARAAllCols   = $true }
              default { 
                # We need to abort here as the risk is that all objects will be returned.
                write-warning -Message ('! Textfilter ' + $textFilter + ' not known / supported. Supported: ' +
                'archive, process, title, documentation, varakey, varavalue, varaall') 
                return
              }
            }
          }
        }
          
        # By default unallocated variables will be $false so we need no further logic.
        $search.setTextSearch($text, $searchProcess, $searchDocumentation, $searchTitle, $searchArchive)
        $search.setSearchInVariable($text, $searchVARAKey, $searchVARAValue, $searchVARAAllCols, [com.uc4.communication.requests.SearchObject+VariableDataType]::CHARACTER)
      }
    }
  
    ###################
    # Name & path filter
    ###################
    $search.setname($name)

    # This is a bit special. When searching for objects with folder information, we must encode the client name as top-level
    # folder. So searching for /PROD is going to be /1000/PROD if the search is executed on client 1000.
    if ($path -notmatch '^/') {
      $path = ('/' + $path)
    }
    $search.setSearchLocation(([string]$aeConnection.client + $path), (! $NonRecursiveSearch))

    ###################
    # Start search and gather results
    ###################
    try {
      $aeConnection.sendRequest($search)
    }
    catch {
      Write-Warning -message ('! Failed to query the AE: ' + $_.Exception.Message)
      $resultSet += (_new-aeEmptySearchResult -name "$name" -result $WFCFAILURE)
      return
    }
    
    Write-Verbose -Message ('* Found ' + $search.size() + ' results. ' + $search.getMessageBox())

    try {
      [java.util.Iterator]$iterator = $search.resultIterator()
      for ($result = 0; $result -lt $search.size(); $result++) {
        # The top level folder of the object equals to the client number. We don't want this to not confuse import-aeObject or other
        # functions. Instead we encode the source client information in an extra "client" field.
        [com.uc4.api.SearchResultItem]$item = $iterator.next()
        $subResultSet += $item
      }
    }
    catch {
      Write-Warning -Message "! Issue with gathering result items."
    }
    
    # This feature can be very handy if comparing a list of objects (i.E. coming from a text file) with what is actually available on a system.
    # Of course it could be misleading as well - i.E. if the namefilter is set to "*" and the datefilter says all created objects of today. That
    # Would result in a result with objectname "*" with no further information, but this should not be the common use-case for using this switch.
    if ($subResultSet.length -eq 0) {
      $subResultSet += (_new-aeEmptySearchResult -name "$name" -result $WFCEMPTY)
    }
    
    $resultSet += $subResultSet
  }
  
  end {
    Write-Debug -Message '** Search-aeObject end'
    return ($resultSet | Sort-Object -property Path)
  }
}

function Export-aeObject {
  <#
      .SYNOPSIS
      Export a object to an XML file.

      .DESCRIPTION
      This is the export-as-xml functionality. It is done on a per-object basis so every object results in a separate file.
      You can pipe the output of this function to import-aeObject or pipe the output of search-aeObject into this function.
      Exported XMLs contain the original AE path (this is a WorkflowCommander feature). This can be considered when the
      object is being imported.

      .PARAMETER aeConnection
      WorkflowCommander AE connection object.

      .PARAMETER name
      Name of the AE object to export.

      .PARAMETER file
      Filename or directory name to where the object should be exported. If a directory is specified, the XML file will be named like the object.

      .EXAMPLE
      export-aeObject -ae $ae -name MYJOB.WIN01 -file C:\temp
      Exports MYJOB.WIN01 to c:\temp\MYJOB.WIN01.xml

      export-aeObject -ae $ae -name MYJOB.WIN01 -file C:\temp\output.xml
      Exports MYJOB.WIN01 to c:\temp\output.xml.xml

      .LINK
      http://workflowcommander.blogspot.com

      .OUTPUTS
      List of successfully exported objects and path to their XML files.
  #>
  Param(
    [Parameter(mandatory,HelpMessage='AE connection object returned by new-aeConnection.')]
    [Alias('ae')]
    [WFC.Core.WFCConnection]$aeConnection,
    [Parameter(ValueFromPipeline,HelpMessage='Name of object to export.',ValueFromPipelineByPropertyName,Mandatory)]
    [string[]]$name,
    [Parameter(Mandatory,HelpMessage='File or directory to export the object to.')]
    [Alias('directory')]
    [IO.FileInfo]$file
  )
  
  begin {
    Write-Debug -Message '** Export-aeObject start'
    $resultSet = @()
    $startExportDateTime = ''
  }
  
  process {
    # As export-aeobject is often fed by search-aeobject output, it might be that there was no result. So we should only set the
    # starttime if we really at least received one pipe entry.
    if (! $startExportDateTime) {
      $startExportDateTime = [datetime]::Now
    }
  
    # Before we export we need to identify whether the object exists and whether the type is exporteable
    $objectInfo = search-aeObject -aeConnection $aeConnection -name "$name"
  
    # There are limitations when it comes to 
    if ($objectInfo.type -eq 'USER' -or $objectInfo.type -eq 'FOLD' -or $objectInfo.result -eq 'UNDEF' -or $objectinfo.result -eq 'EMPTY') {
      Write-Verbose -Message ('* ' + $name + ' not exported.')
      $resultSet += _new-aeObjectExportObject -name "$name" -type $objectInfo.type -result $WFCEMPTY 
      return
    }

    Write-Debug -Message ('** Starting export of ' + $name)
      
    # Determine XMl file name based on object name.
    if ($file.Attributes.HasFlag([IO.FileAttributes]::Directory) -and $file.Attributes -ne -1) {
      $outputFilename = ($file.fullname + '\' + $name + '.xml')
    }
    else {
      $outputFilename = $file
    }

    if (Get-ChildItem -path $outputFilename -ErrorAction SilentlyContinue) {
      Write-Warning -Message ('! File ' + $outputFilename + ' does already exist and will be overwritten.')
    }
            
    try {  
      $aeExportRequest = [com.uc4.communication.requests.ExportObject]::new($name, [java.io.file]::new($outputFilename))
      $aeConnection.sendRequest($aeExportRequest)
    }
    catch {
      Write-Warning -Message ('! AE export failure for ' + $name + ' this should not happen! Please report exception: ' + $_.Exception.Message)
      $resultSet += _new-aeObjectExportObject -name "$name" -type $objectInfo.type -result $WFCFAILURE
      return
    }
    
    # The above often returns no error in case that invalid objects have been exported.
    $msg = $aeExportRequest.getMessageBox()
    if($msg) {
      Write-Warning -Message ('! AE export failure for ' + $name + ': ' + $msg)
      $resultSet += _new-aeObjectExportObject -name "$name" -type $objectInfo.type -result $WFCFAILURE
      return
    }      
    # We can encode the (folder-)path where the object has been exported from in case that the export has not been done from a V11+ system.
    # This information is encoded in the uc-name element.
    [xml]$xmlObjectInfo = get-content -path $outputFilename
    $attribute = $xmlObjectInfo.CreateAttribute('WorkflowCommander','aeObjectPath','devnull')
    $path = $objectInfo.path
    $attribute.value = $path
    $null = $xmlObjectInfo.'uc-export'.SetAttributeNode($attribute)
    $xmlObjectInfo.OuterXml | Out-File -Encoding ASCII -FilePath $outputFilename

    # When exporting UC_* VARAs, the path will be empty.
    if (! $path) {
      $path = '-'
    }

    # Finally return the information on the exported object      
    $resultSet += _new-aeObjectExportObject -name "$name" -type $objectInfo.type -path "$path" -file $outputFilename -result $WFCOK 
    Write-Verbose -Message ('* Object ' + $name + ' has been exported successfully.')
  }
  
  end {
    if ($startExportDateTime) {
      Write-Verbose -Message ('* Exported ' + @($resultSet | Where-Object { $_.File -ne $null }).length + ' objects. Duration: ' + ([datetime]::Now - $startExportDateTime).toString())
    }
    else {
      Write-Warning -Message ('! Nothing exported. Likely that the input was null (like empty search-aeObject).')
    }
    
    Write-Debug -Message '** Export-aeObject end'
    
    return ($resultSet | Sort-Object -Property Name)
  }
}

function Import-aeObject {
  <#
      .SYNOPSIS
      Import a single XML or a directory containing multiple XML files.

      .DESCRIPTION
      This equals the import - XML file functionality from the GUI. Plus you can choose where to import the file to
      or, if not specified, it will consider the path information encoded in the XML by export-aeObject.

      .PARAMETER aeConnection
      WorkflowCommander AE connection object.

      .PARAMETER file
      File or directory to import. If a directory has been specified, all XML files within will be considered.

      .PARAMETER path
      AE path to import the object(s) to. Like /PRODUCTION

      .PARAMETER noOverwrite
      This will prevent that existing objects will be overwritten.

      .EXAMPLE
      import-aeObject -ae $ae -file C:\temp\demo.xml -path /PRODUCTION
      Imports the c:\temp\demo.xml object to /PRODUCTION

      import-aeObject -ae $ae -file C:\temp\ 
      Imports all XML files stored within C:\temp and uses the path information encoded in the XML files.

      .OUTPUTS
      Outputs the total amount of successfully imported objects.
  #>
  [CmdletBinding(SupportsShouldProcess)]
  param(
    [Parameter(mandatory,HelpMessage='AE connection object returned by new-aeConnection.')]
    [Alias('ae')]
    [WFC.Core.WFCConnection]$aeConnection,
    [Parameter(ValueFromPipeline,HelpMessage='File or directory to import XMLs from.',ValueFromPipelineByPropertyName,Mandatory)]
    [AllowNull()]
    [Alias('directory')]
    [io.fileinfo[]]$file,
    [Parameter(ValueFromPipelineByPropertyName)]
    [string]$path = $null,
    [switch]$noOverwrite
  )

  begin {
    Write-Debug -Message '** Import-aeObject start'
    $resultSet = @()
    $startImportDateTime = $null
  }

  process {
    if (! $startImportDateTime) {
      $startImportDateTime = [datetime]::Now
    }
  
    # This might happen if we pipe from a export-aeObject without getting rid of EMPTY results
    if (! $file) {
      Write-Debug -Message '** Skipping import as no file is defined.'
      return
    }
  
    # If a directory has been specified, we process the single XML files within
    if ($file.attributes.HasFlag([IO.FileAttributes]::Directory)) {
      Write-Verbose -Message '* Directory has been specified, loading all XMLs within directory.'
      $file = Get-ChildItem -Path ($file.fullname + '\*.xml')
    }
    
    if ($file.Length -eq 0) {
      Write-Warning -Message '! No XML files to import.'
      return
    }
    
    foreach ($xmlFile in $file) {
      Write-Debug -Message ('Processing ' + $xmlFile)
      # The path defines, in what folder structure the object should be stored into. This information can be inputted with 3 methods:
      # 1. as -path parameter to import-aeObject
      # 2. encoded into the XML object as WorkflowCommander:aeObjectPath attribute to uc-name XML element
      try {
        # If the parameter has not been specified, we must load the XML file
        Write-Debug -Message ('Reading base information from ' + $xmlFile)
        [xml]$xmlData = get-content -Path $xmlFile

        # Get name, type and aeObjectPath
        $identifiedType = $xmlData.'uc-export'.FirstChild.LocalName
        $identifiedName = (select-xml -xml $xmlData -xpath '/uc-export/*/@name').Node.'#text'
        $identifiedPath = (select-xml -xml $xmlData -XPath '//*[@WorkflowCommander:aeObjectPath]' -Namespace @{ 'WorkflowCommander' = 'devnull' }).Node.aeObjectPath
      }
      catch {
        Write-Warning -Message ('! Issues loading XML file ' + $xmlFile + ' to extract AE folder destination. Unsafe import will be skipped.')
        $resultSet += _new-aeImportResult -name '-' -file $xmlFile -result $WFCFAILURE
      }

      if (! $identifiedPath) {       
        Write-Debug -Message ('** Object will be imported to parametrized destination folder ' + $path)
        $identifiedPath = $path
      }
      
      # If identifiedPath is empty, we can not safely import the object.
      if (! $identifiedPath) {
        $resultSet += _new-aeImportResult -name $identifiedName -file $xmlFile -type $identifiedType -result $WFCFAILURE
        Write-Warning -Message ('! AE path for object import of ' + $xmlFile.fullname + ' could not be identified.')
        return
      }

      Write-Verbose -Message ('* Importing ' + $xmlFile.fullname + ' to ' + $identifiedPath)

      try {
        # Make sure the destination folder path exists. This will create the necessary subfolder structure and return the iFolder
        # destination object.
        Write-Debug -Message ('** Making sure that destination path exists: ' + $identifiedPath)
        $targetFolder = new-aeFolder -aeConnection $aeConnection -path $identifiedPath

        if ($pscmdlet.ShouldProcess('Import ' + $xmlFile.fullname + ' to ' + $identifiedPath) ) {
          $importRequest = [com.uc4.communication.requests.ImportObject]::new(
            [java.io.file]$xmlfile.fullname,
            $targetFolder,
            (! $noOverwrite),
            $true
          )
      
          $aeConnection.sendRequest($importRequest)
          $aeMsg = $importRequest.getImportMessages()
        }
      }
      catch {
        $resultSet += _new-aeImportResult -name $identifiedName -path $identifiedPath -file $xmlFile -type $identifiedType -result $WFCFAILURE
        Write-Warning -Message ('! Import of ' + $xmlFile.fullname + ' failed! AE says: ' + $aeMsg)
        return
      }

      if ($aeMsg -match 'U04005758') {
        Write-Warning -Message ('! File ' + $xmlFile.fullname + ' has not been imported because object already exists.')
        $resultSet += _new-aeImportResult -name $identifiedName -path $identifiedPath -file $xmlFile -type $identifiedType -result $WFCEMPTY
      }
      else {
        $resultSet += _new-aeImportResult -name $identifiedName -path $identifiedPath -file $xmlFile -type $identifiedType -result $WFCOK
      }
    }
  }
      
  end {
    if ($startImportDateTime) {
      Write-Verbose -Message ('* Imported ' + $resultSet.length + ' objects. Duration: ' + ([datetime]::Now - $startImportDateTime).toString())
    }
    else {
      Write-Warning -Message ('! Nothing imported as file/directory was not existing or empty.')
    }

    Write-Debug -Message '** Import-aeObject end'
    return ($resultSet | Sort-Object -Property Name)
  }
}

function New-aeFolder  {
  <#
      .SYNOPSIS
      Create folder(s) on AE system. Works recursively.

      .DESCRIPTION
      Creates folders. Missing folders will be automatically created.

      .PARAMETER aeConnection
      WorkflowCommander AE Connection object.

      .PARAMETER path
      Path to create.

      .EXAMPLE
      new-aeFolder -ae $ae -path /PRODUCTION/SYSTEMA/FOLDER1
      Creates /PRODUCTION/SYSTEMA/FOLDER1. 

      .LINK
      http://workflowcommander.blogspot.com

      .OUTPUTS
      IFolder AE API Object.
  #>
  [CmdletBinding(SupportsShouldProcess)]
  param (
    [Parameter(Mandatory,HelpMessage='AE connection object returned by new-aeConnection.')]
    [Alias('ae')]
    [WFC.Core.WFCConnection]$aeConnection,
    [Parameter(Mandatory,HelpMessage='Path to create on AE server.')]
    [string[]]$path
  )
 
  begin {
    Write-Debug -Message '** new-aeFolder start'
    $resultSet = @()
  }

  process { 
    # TODO: replace logic with get-aeFolder
    $folderBrowser = [com.uc4.communication.requests.FolderTree]::new()

    try {
      $aeConnection.sendRequest($folderBrowser)
    }
    catch {
      throw('! Could not send request to AE. Please check exception: ' + $_.Exception.Message)
    }
  
    # Remove leading and ending slash for appropriate foreach loop
    $path = $path -replace '^/', '' -replace '/$', ''
    $absPath = '/'
  
    foreach ($subFolder in $path.split('/')) {
      $folderName = ($absPath + $subFolder)
      if (! $folderBrowser.getFolder($folderName)) {
        write-verbose -Message ('* Creating folder structure ' + $folderName)
        try {
          if ($PSCmdlet.ShouldProcess('Folder ' + $folderName + ' does not exist and must be created.')) {
            $createFolderRequest = [com.uc4.communication.requests.CreateObject]::new(
              [com.uc4.api.UC4ObjectName]::new($subFolder), 
              [com.uc4.api.Template]::FOLD, 
              $folderBrowser.getFolder($absPath)
            )
            $aeConnection.sendRequest($createFolderRequest)
      
            # Reload folderBrowser
            $folderBrowser = [com.uc4.communication.requests.FolderTree]::new()
            $aeConnection.sendRequest($folderBrowser)
          }
        }
        catch {
          throw ('! Issue creating folder.')
        }
      }
      $absPath = ($absPath + $subFolder + '/')
    }
    
    $resultSet += $folderBrowser.getFolder($absPath)
  }
  
  end {
    Write-Debug -Message '** new-aeFolder end'
    return ($resultSet | Sort-Object Path)
  }
}

function Get-aeStatistic  {
<#
      .SYNOPSIS
      Get statistic of a single object.

      .DESCRIPTION
      Equals the rightclick-statistic on a single object functionality.

      .PARAMETER aeConnection
      WorkflowCommander AE Connection object.

      .PARAMETER name
      Name of object of the statistic to show.

      .PARAMETER amount
      Amount of statistic entries to show. Default is the users setting.

      .EXAMPLE
      Get-aeStatistic -ae $ae -name JOBF.002
      Get all available statistic entries (depending on user settings and available statistic entries).

      Get-aeStatistic -ae $ae -name JOBF.002 -amount 1
      Get latest statistic entry only.

      .LINK
      http://workflowcommander.blogspot.com

      .OUTPUTS
      StatisticSearchItem items
  #>
  param (
    [Parameter(mandatory,HelpMessage='AE connection object returned by new-aeConnection.')]
    [Alias('ae')]
    [object]$aeConnection,
    [Parameter(mandatory,HelpMessage='Objectname',ValueFromPipelineByPropertyName,ValueFromPipeline)]
    [string]$name,
    [int]$amount = 1
  )

  begin {
    Write-Debug -Message '** Get-aeStatistic start'
    $resultSet = @()
  }

  process {
    try {
      $getStatistic = [com.uc4.communication.requests.ObjectStatistics]::new([com.uc4.api.UC4ObjectName]::new($name),$amount)
      $aeConnection.sendRequest($getStatistic)
    }
    catch {
      Write-Warning -message ('! Failed to query the AE: ' + $_.Exception.Message)
      $resultSet += (_new-aeEmptyStatisticResult -name "$name" -result $WFCFAILURE)
      return
    }

    write-verbose -message ('* Getting ' + $getStatistic.size() + ' statistic entries')
    if ($getStatistic.size() -eq 0) {
      $resultSet += (_new-aeEmptyStatisticResult -name "$name" -result $WFCEMPTY)
    }
    else {
      $statisticIterator = $getSTatistic.resultIterator()
      while ($statisticIterator.hasNext()) {
        $resultSet += $statisticIterator.next()
      }
    }
  }
  
  end {
    Write-Debug -Message '** Get-aeStatistic end'
    return $resultSet
  }
}


function Search-aeStatistic  {
  param (
    [Parameter(mandatory,HelpMessage='AE connection object returned by new-aeConnection.')]
    [Alias('ae')]
    [object]$aeConnection,
    [Parameter(mandatory,HelpMessage='Objectname',ValueFromPipelineByPropertyName,ValueFromPipeline)]
    [string[]]$name
  )

  $searchStatistic = [com.uc4.communication.requests.GenericStatistics]::new()
  $searchStatistic.setObjectName('JOBF.002')
  $searchStatistic.selectAllTypes()
  $aeConnection.sendRequest($searchStatistic)
  $searchStatistic.size()
  $statisticentry = $searchStatistic.resultIterator().next()
  $statisticentry
}
