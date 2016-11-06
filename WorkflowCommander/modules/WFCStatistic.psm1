#########################################################################################
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

function Get-aeStatistic  {
  <#
      .SYNOPSIS
      Get statistic of a single object.

      .DESCRIPTION
      Equals the rightclick-statistic on a single object functionality.

      .PARAMETER aeConnection
      WorkflowCommander AE Connection object.

      .PARAMETER name
      Name of object of the statistic to show. This must not contain any wildcards - you might pipe from search-aeObject if you require this.

      .PARAMETER amount
      Amount of statistic entries to show. Default is the users setting (which is the default). Order is always newest to oldest entry. 

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
      $getStatistic = [com.uc4.communication.requests.ObjectStatistics]::new([com.uc4.api.UC4ObjectName]::new($name), $amount)
      $aeConnection.sendRequest($getStatistic)
    }
    catch {
      Write-Warning -message ('! Failed to query the AE: ' + $_)
      $resultSet += (New-WFCEmptyStatisticResult -name "$name" -result FAIL)
      return
    }

    write-verbose -message ('* Query resulted in ' + $getStatistic.size() + ' statistic entries')
    if ($getStatistic.size() -eq 0) {
      $resultSet += (New-WFCEmptyStatisticResult -name "$name" -result EMPTY)
    }
    else {
      $statisticIterator = $getStatistic.resultIterator()
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
  <#
      .SYNOPSIS
      Search for statistic entries.

      .DESCRIPTION
      Equals the period search in the AWA GUI.

      .PARAMETER aeConnection
      WorkflowCommander AE Connection object.

      .PARAMETER name
      Name of object of the statistic to show. Wildcards are possible.

      .EXAMPLE
      search-aeStatistic -ae $ae -name JOBF.002 -xxxxxx
      x.

      .LINK
      http://workflowcommander.blogspot.com

      .OUTPUTS
      StatisticSearchItem items
  #>
  param (
    [Parameter(mandatory,HelpMessage='AE connection object returned by new-aeConnection.')]
    [Alias('ae')]
    [object]$aeConnection,
    [Parameter(ValueFromPipelineByPropertyName,ValueFromPipeline)]
    [string]$name = $null,
    [string]$alias = $null,
    [string]$archiveKey1 = $null,
    [string]$archiveKey2 = $null,
    [switch]$archiveKeyAND,
    [string]$dstHost = $null,
    [string]$srcHost = $null,
    [string]$status = $null,
    [Parameter(ValueFromPipelineByPropertyName)]
    [int]$runid = $null,
    [Parameter(ValueFromPipelineByPropertyName)]
    [int]$topRunid = $null,
    [string]$queue = $null,
    [ValidateSet('noConstraint', 'Activation', 'Start', 'End')]
    [string]$dateSearch = 'noConstraint',
    [datetime]$fromDateTime = [DateTime]::Today,
    [datetime]$toDateTime = [datetime]::Now,
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
    [string[]]$type = $null
  )

  begin {
    Write-Debug -Message '** Search-aeStatistic start'
    $resultSet = @()
  }

  process {
    $searchStatistic = [com.uc4.communication.requests.GenericStatistics]::new()
  
    $searchStatistic.setObjectName($name)
    $searchStatistic.setAlias($alias)
    $searchStatistic.setArchiveKey1($archiveKey1)
    $searchStatistic.setArchiveKey2($archiveKey2)
    $searchStatistic.setArchiveKeyAndRelation($archiveKeyAND)
 
    switch($dateSearch) {
      'noConstraint' { $searchStatistic.setDateSelectionNone() }
      'activation' { $searchStatistic.setDateSelectionActivation() }
      'start' { $searchStatistic.setDateSelectionStart() }
      'end' { $searchStatistic.setDateSelectionEnd() }
    }

    $searchStatistic.setFromDate([com.uc4.api.Datetime]::new($fromDateTime.ToString('yyyy-MM-dd HH:mm:ss')))
    $searchStatistic.setToDate([com.uc4.api.Datetime]::new($toDateTime.ToString('yyyy-MM-dd HH:mm:ss')))
    
    $searchStatistic.setDestinationHost($dstHost)
    $searchStatistic.setSourceHost($srcHost)
  
    # TODO: Status is numeric require mapping
    $searchStatistic.setStatus($status)
    
    $searchStatistic.setRunID($runid)  
    $searchStatistic.setTopRunID($topRunid)
  
    ###################
    # Objecttype filter
    ###################
    # Depending on whether we want to filter for object types or not we select all or only specific object types.
    if ($type -eq $null) {
      $searchStatistic.selectAllTypes()
    }
    else {
      $searchStatistic.unselectAllTypes()

      # This activates the object type search selection. To support new object types, just add the new type 
      # in the parameter validation. i.E. JOBS => $search.setTypeJOBS($true)
      foreach ($typeFilter in $type) {
        $searchStatistic.('setType' + $typeFilter)($true)
      }
    }
    
    # Missing parametrization
    #$searchStatistic.setPlatform*($true)
  
    $searchStatistic.setQueue($queue)
    
    try {
      $aeConnection.sendRequest($searchStatistic)
    }
    catch {
      $resultSet += New-WFCEmptyStatisticResult -name "$name" -result FAIL
      Write-Warning ('! Failed to query the AE: ' + $searchStatistic.getAllMessageBoxes())
      return
    }

    $aeMsg = $searchStatistic.getAllMessageBoxes()
    if ($aeMsg -ne '') {
      Write-Warning -Message ('! Querying the AE failed: ' + $aeMsg)
      $resultSet += New-WFCEmptyStatisticResult -name "$name" -result FAIL
      return
    }

    write-verbose -message ('* Query resulted in ' + $searchStatistic.size() + ' statistic entries')
    if ($searchStatistic.size() -eq 0) {
      $resultSet += (New-WFCEmptyStatisticResult -name "$name" -result EMPTY)
    }
    else {
      $iterator = $searchStatistic.resultIterator()
      while($iterator.hasNext()) {
        $resultSet += $iterator.next()
      }
    }  
  }
  
  end {
    Write-Debug -Message '** Search-aeStatistic end'
    return $resultSet
  }
  
}
