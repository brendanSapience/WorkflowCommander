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
            try {
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
            catch {
              Write-Warning -message ('! Failed to create folder: ' + $folderBrowser.getAllMessageBoxes())
              return
            }
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

