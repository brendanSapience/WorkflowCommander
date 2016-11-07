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
$WFCPROFILES = ([Environment]::GetFolderPath("MyDocuments") + '\')

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
