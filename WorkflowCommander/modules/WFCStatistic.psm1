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
      $resultSet += (New-WFCEmptyStatisticResult -name "$name" -result $WFCFAILURE)
      return
    }

    write-verbose -message ('* Getting ' + $getStatistic.size() + ' statistic entries')
    if ($getStatistic.size() -eq 0) {
      $resultSet += (New-WFCEmptyStatisticResult -name "$name" -result $WFCEMPTY)
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
