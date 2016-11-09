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

# Profile directory defines where to search for profile files. You can set this to any folder you like.
# Make sure you have an ending \ !!!
$global:WFCPROFILES = ([Environment]::GetFolderPath("MyDocuments") + '\')

# Mapping table for search that involves status
$global:WFCSTATUS = @{
  'ENDED_SKIPPED_CONDITIONS' = '1933'
  'RET_0' = '0000'
  'RET_1' = '0001'
  'RET_2' = '0002'
  'ANY_ABEND' = '1800-1899'
  'ANY_OK' = '1900-1999'
  'ANY_SKIPPED' = '1920,1922,1930,1931,1933,1940,1941,1942'
  'ENDED_CANCEL' = '1850,1851'
  'ENDED_EMPTY' = '1910,1912'
  'ENDED_ESCALATED' = '1856'
  'ENDED_INACTIVE_MANUAL' = '1922'
  'ENDED_INACTIV' = '1919,1920,1921,1922,1925'
  'ENDED_NOT_OK' = '1800'
  'ENDED_NOT_OK_SYNC' = '1801'
  'ENDED_OK' = '1900'
  'ENDED_OK_OR_EMPTY' = '1900,1910,1912'
  'ENDED_OK_OR_INACTIV' = '1900,1919,1920,1921,1922,1925'
  'ENDED_SKIPPED' = '1930,1931,1933'
  'ENDED_SKIPPED_SYNC' = '1931'
  'ENDED_TIMEOUT' = '1940,1941,1942'
  'ENDED_TRUNCATE' = '1911'
  'ENDED_UNDEFINED' = '1815'
  'ENDED_VANISHED' = '1810'
  'FAULT_ALREADY_RUNNING' = '1822'
  'FAULT_NO_HOST' = '1821'
  'FAULT_OTHER' = '1820'
  'RET_3' = '0003'
  'RET_4' = '0004'
  'RET_5' = '0005'
  'ANY_ABEND_EXCEPT_FAULT' = '1800-1819,1823-1899'
  'ANY_EXCEPT_FAULT' = '1800-1819,1823-1999'
  'ANY_ACTIVE' = '1300-1799'
  'ANY_BLOCKED' = '1560,1562'
  'ANY_BLOCKED_OR_STOPPED' = '1560-1564'
  'ANY_OK_OR_UNBLOCKED' = '1900-1999,1899'
  'ANY_STOPPED' = '1561,1563-1564'
  'ANY_WAITING' = '1301,1682-1700,1709-1710'
  'ENDED_OK_OR_UNBLOCKED' = '1900,1899'
  'WAITING_AGENT' = '1685,1688-1689,1694,1696'
  'WAITING_AGENT_OR_AGENTGROUP' = '1685-1689,1694,1696'
  'WAITING_AGENTGROUP' = '1686-1687'
  'WAITING_EXTERNAL' = '1690'
  'WAITING_GROUP' = '1710'
  'WAITING_QUEUE' = '1684'
  'WAITING_SYNC' = '1697'
  'ANY_RUNNING' = '1550,1541,1542,1545,1546,1551-1564,1566-1576,1578-1583,1590-1593,1682,1685,1686,1701'
}

function New-aeConnection {
  Param(
    [Parameter(ParameterSetName = "profile",HelpMessage = "Name of profile to load.",Mandatory)]
    [string]$profile,
    [Parameter(ParameterSetName = "new",HelpMessage = "AE client to login into.",Mandatory)]
    [int]$client,
    [Parameter(ParameterSetName = "new",HelpMessage = "AE server / IP.",Mandatory)]
    [string]$server,
    [Parameter(ParameterSetName = "new",HelpMessage = "AE user to login with.",Mandatory)]
    [string]$username,
    [Parameter(ParameterSetName = "new")]
    [String]$department = $null,
    [Parameter(ParameterSetName = "new")]
    [string]$saveAsProfile = $null,
    [Parameter(ParameterSetName = "new")]
    [int]$port = 2217,
    [Parameter(ParameterSetName = "new",Mandatory)]
    [SecureString]$password
  )

  begin {
    Write-Debug -message "** New-aeConnection start."
  }

  process {
    Write-Verbose -Message "* Initiating connection..."

    if ($profile) {
      try {
        Write-Debug -Message ("** Creating connection using profile: " + $WFCPROFILES + $profile  + ".xml")
        $WFCConnection = [WFC.Core.WFCConnection]::new($WFCPROFILES + $profile  + '.xml')
      }
      catch {
        throw ($_)
      }
    }
    else {
      Write-Debug -message "** Creating connection using adhoc profile / credentials."
      $WFCConnection = [WFC.Core.WFCConnection]::new($username, $department, $password, $server, $client, $port)
    }

    if ($saveAsProfile) {
      try {
        $profileFilename = $WFCConnection.saveProfile($WFCPROFILES + $saveAsProfile + '.xml')
        Write-Verbose -message "* Profile stored: " + $profileFilename.FullName
      }
      catch {
        throw
      }
    }

    if ($WFCConnection.message) {
      write-warning -message $WFCConnection.message
      return
    }

    Write-Warning -message "**************************************************************************************"
    Write-Warning -message "* WorkflowCommander is a copyrighted work by Joel Wiesmann (joel.wiesmann@gmail.com) *"
    Write-Warning -message "* >  Do only continue if you read the disclaimer, manual & licensing information.  < *"
    Write-Warning -message "**************************************************************************************"
    Write-Warning -message "* Connected:"
    Write-Warning -message ("* User       " + $WFCConnection.username)
    Write-Warning -message ("* Department " + $WFCConnection.department)
    Write-Warning -message ("* Systemname " + $WFCConnection.systemname)
    Write-Warning -message ("* Client     " + $WFCConnection.client)
    Write-Warning -message ("* License    '" + $WFCConnection.licensedTo + "' expires in: " + $WFCConnection.expiryDays + " days.")

    if ($WFCConnection.expiryDays -lt 30) {
      Write-Warning -message ("! Your license is about to expire. Like the product? Get a license!")
      Write-Warning -message ("! joel.wiesmann@gmail.com / XING / workflowcommander.blogspot.com")
    }

    return $WFCConnection
  }

  end {
    Write-Debug -Message ("** New-aeConnection ended.")
  }
}
