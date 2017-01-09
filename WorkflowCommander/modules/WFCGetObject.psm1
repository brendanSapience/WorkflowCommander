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

# This is an internal cmdlet for the moment. Please do not use, it will change.
function get-aeObject() {
  Param(
    [Parameter(mandatory,HelpMessage='AE connection object returned by new-aeConnection.')]
    [Alias('ae')]
    [Object]$aeConnection,
    [Parameter(ValueFromPipeline,HelpMessage='Name of object to get.',ValueFromPipelineByPropertyName,Mandatory)]
    [string]$name
  )
  
  begin {
    Write-Debug -Message '** Get-aeObject start'
    $resultSet = @()
  }
  
  process {
    $objectRequest = [com.uc4.communication.requests.OpenObject]::new($name, $true, $true)
    try {
      $aeConnection.sendRequest($objectRequest)
    }
    catch {
      Write-Warning -Message ('! Failed to query the AE: ' + $objectRequest.getAllMessageBoxes() + ' ' + $_.Exception.GetType().FullName + ' ' + $_.Exception.Message)
      return $null
    }
    $resultSet += $objectRequest
  }

  end {
    Write-Debug -Message '** Get-aeObject end'
    return $resultSet
  }
}