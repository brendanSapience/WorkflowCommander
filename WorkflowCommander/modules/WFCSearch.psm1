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

function search-aeObject {
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

      .PARAMETER type
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
      search-aeObject -ae $ae -name "*JOB*" -path "/PRODUCTION" -type JOBS
      Searches for objects that name matches "*JOB*" and are stored below /PRODUCTION and are of type JOBS.

      .LINK
      http://workflowcommander.wordpress.com

      .OUTPUTS
      Array of findings.
  #>
  Param(
    [Parameter(Mandatory,HelpMessage='AE connection object returned by new-aeConnection.')]
    [Alias('ae')]
    [Object]$aeConnection,
    [Parameter(ValueFromPipeline,ValueFromPipelineByPropertyName)]
    [string]$name = '*',
    [Parameter(ValueFromPipelineByPropertyName)]
    [string]$path = $null,
    [Parameter(ValueFromPipelineByPropertyName)]
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
    [string[]]$type = $null,
    [string]$text = $null,
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
    $search = [com.uc4.communication.requests.SearchObject]::new()
    
    ###################
    # Search for object or for usage
    ###################
    $search.setSearchUseOfObjects($searchForUsage)

    ###################
    # Objecttype filter
    ###################
    # Depending on whether we want to filter for object types or not we select all or only specific object types.
    if ($type -eq $null) {
      $search.selectAllObjectTypes() 
    }
    else {
      $search.unselectAllObjectTypes()

      # This activates the object type search selection. To support new object types, just add the new type 
      # in the parameter validation. i.E. JOBS => $search.setTypeJOBS($true)
      foreach ($typeFilter in $type) {
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

    # Finally after all the parametrization of the search, send it to the AE and receive the results.
    try {
      $aeConnection.sendRequest($search)
    }
    catch {
      Write-Warning -message ('! Failed to query the AE: ' + $search.getAllMessageBoxes() + ' ' + $_.Exception.GetType().FullName + ' ' + $_.Exception.Message)
      $resultSet += New-WFCEmptySearchResult -name "$name" -result FAIL
      return
    }
    
    Write-Verbose -Message ('* Found ' + $search.size() + ' results. ')
    
    if ($search.size() -eq 0) {
      $resultSet += New-WFCEmptySearchResult -name "$name" -result EMPTY
    }
    else {
      $searchIterator = $search.resultIterator()
      while ($searchIterator.hasNext()) {
        $resultSet += $searchIterator.next()
      }
    }
  }
  
  end {
    Write-Debug -Message '** Search-aeObject end'
    return ($resultSet | Sort-Object -property Path)
  }
}
