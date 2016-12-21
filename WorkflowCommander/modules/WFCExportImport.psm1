#########################################################################################
# WorkflowCommander, copyrighted by Joel Wiesmann, 2016
# <  THIS CODE IS EXPERIMENTAL  >
#
# Get newest tipps and tricks on my blog:
# http://workflowcommander.wordpress.com
# ... or get in touch with me directly (joel.wiesmann <at> gmail <dot> com)
# 
# Read the LICENSE.txt provided with this software. GNU GPLv3
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
    http://workflowcommander.wordpress.com

    .OUTPUTS
    List of successfully exported objects and path to their XML files.
#>
[cmdletbinding(SupportsShouldProcess)]
Param(
  [Parameter(mandatory,HelpMessage='AE connection object returned by new-aeConnection.')]
  [Alias('ae')]
  [Object]$aeConnection,
  [Parameter(ValueFromPipeline,HelpMessage='Name of object to export.',ValueFromPipelineByPropertyName,Mandatory)]
  [string[]]$name,
  [Parameter(Mandatory,HelpMessage='File or directory to export the object to.')]
  [Alias('directory')]
  [IO.FileInfo]$file
)
  
begin {
  Write-Debug -Message '** Export-aeObject start'
  $resultSet = @()
  $startExportDateTime = [datetime]::Now
}
  
process {
  # Before we export we need to identify whether the object exists and whether the type is exporteable
  $objectInfo = search-aeObject -aeConnection $aeConnection -name "$name"
  
  # Some objects are not exportable. Same applies if object does not exist.
  if (@('USER','FOLD').contains($objectInfo.type)) {
    Write-Warning -Message ('! ' + $name + ' not exported because it is of type: ' + $objectInfo.type)
    $resultSet += New-WFCObjectExportObject -name "$name" -type $objectInfo.type -result EMPTY
    return
  }
    
  # If the object was not identifieable, we do not try to export 
  if (@('FAIL','EMPTY').contains($objectInfo.result)) {
    Write-Verbose -Message ('* ' + $name + ' not exported because was not identifieable (result was ' + $objectInfo.result + ')')
    $resultSet += New-WFCObjectExportObject -name "$name" -type $objectInfo.type -result EMPTY
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


  if ($pscmdlet.ShouldProcess('Export ' + $name + ' to ' + $outputFilename + ' with path encoding: ' + $objectInfo.path) ) {
    try {  
      $aeExportRequest = [com.uc4.communication.requests.ExportObject]::new($name, [java.io.file]::new($outputFilename))
      $aeConnection.sendRequest($aeExportRequest)
    }
    catch {
      Write-Warning -Message ('! AE export failure for ' + $name + ' this should not happen! Please report exception: ' + $aeExportRequest.getAllMessageBoxes())
      $resultSet += New-WFCObjectExportObject -name "$name" -type $objectInfo.type -result FAIL
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

    # Finally return the information on the exported object      
    $resultSet += New-WFCObjectExportObject -name "$name" -type $objectInfo.type -path "$path" -file $outputFilename -result OK
    Write-Verbose -Message ('* Object ' + $name + ' has been exported successfully.')
    }
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
    [Object]$aeConnection,
    [Parameter(ValueFromPipeline,HelpMessage='File or directory to import XMLs from.',ValueFromPipelineByPropertyName,Mandatory)]
    [AllowNull()]
    [Alias('directory')]
    [io.fileinfo]$file,
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
        Write-Warning -Message ('! Issues loading XML file ' + $xmlFile + ' to extract basic information and folder destination. Unsafe import will be skipped.')
        $resultSet += New-WFCImportResult -name '' -file $xmlFile -result FAIL
        return
      }

      if (! $identifiedPath) {       
        Write-Debug -Message ('** Object will be imported to parametrized destination folder ' + $path)
        $identifiedPath = $path
      }
      
      # If identifiedPath is empty, we can not safely import the object.
      if (! $identifiedPath) {
        $resultSet += New-WFCImportResult -name "$identifiedName" -file $xmlFile -type $identifiedType -result FAIL
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
        Write-Warning -Message ('! Import of ' + $xmlFile.fullname + ' failed! AE says: ' + $importRequest.getAllMessageBoxes()  + ' ' + $_.Exception.GetType().FullName + ' ' + $_.Exception.Message)
        $resultSet += New-WFCImportResult -name $identifiedName -path $identifiedPath -file $xmlFile -type $identifiedType -result FAIL
        return
      }

      if ($aeMsg -match 'U04005758') {
        Write-Warning -Message ('! File ' + $xmlFile.fullname + ' has not been imported because object already exists.')
        $result = 'EMPTY'
      }
      else {
        # This will make sure that the import took place at the expected location even if the object was overwritten at the origin location
        $mvResult = Move-aeObject -aeConnection $aeConnection -name $identifiedName -path $identifiedPath
        
        if ($mvResult -eq 'FAIL') {
          $result = 'FAIL'
        }
        else {
          $result = 'OK'
        }
      }
      
      $resultSet += New-WFCImportResult -name $identifiedName -path $identifiedPath -file $xmlFile -type $identifiedType -result $result
    }
  }
      
  end {
    Write-Verbose -Message ('* Imported ' + $resultSet.length + ' objects. Duration: ' + ([datetime]::Now - $startImportDateTime).toString())
    Write-Debug -Message '** Import-aeObject end'
    return ($resultSet | Sort-Object -Property Name)
  }
}