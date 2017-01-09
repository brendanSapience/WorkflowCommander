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

function new-aeFolder  {
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
      http://workflowcommander.wordpress.com

      .OUTPUTS
      IFolder AE API Object.
  #>
  [CmdletBinding(SupportsShouldProcess)]
  param (
    [Parameter(Mandatory,HelpMessage='AE connection object returned by new-aeConnection.')]
    [Alias('ae')]
    [Object]$aeConnection,
    [Parameter(Mandatory,HelpMessage='Path to create on AE server.',ValueFromPipeline,ValueFromPipelineByPropertyName)]
    [string]$path
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
      return
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

function get-aeFolder {
  <#
      .SYNOPSIS
      Receive a iFolder object representing a folder. 

      .DESCRIPTION
      This is mostly interesting for internal usage. For checking availability or sub-folder structures use search-aeObject.

      .PARAMETER aeConnection
      WorkflowCommander AE Connection object.

      .PARAMETER path
      Path to get. If not existing, it will return $null.

      .EXAMPLE
      Get-aeFolder -ae $ae -path /PRODUCTION/SYSTEMA/FOLDER1
      Receive IFolder object.

      .LINK
      http://workflowcommander.wordpress.com

      .OUTPUTS
      IFolder.
  #>
  param (
    [Parameter(Mandatory,HelpMessage='AE connection object returned by new-aeConnection.')]
    [Alias('ae')]
    [Object]$aeConnection,
    [Parameter(Mandatory,HelpMessage='Name of the path to get.',ValueFromPipeline,ValueFromPipelineByPropertyName)]
    [string]$path
  )
  
  begin {
    Write-Debug -Message '** Get-aeFolder start'
    $resultSet = @()
    
    $folderBrowser = [com.uc4.communication.requests.FolderTree]::new()
    try {
      $aeConnection.sendRequest($folderBrowser)
    }
    catch {
      throw('! Could not send request to AE. Please check exception: ' + $_.Exception.Message)
      return
    }
  }

  process {
    $resultSet += $folderBrowser.getFolder($path)
  }

  end {
    Write-Debug -Message '** Get-aeFolder end'
    return $resultSet
  }

}
function move-aeObject  {
  <#
      .SYNOPSIS
      Move object to folder.

      .DESCRIPTION
      Move an object to a new folder. This is not for renaming an object.

      .PARAMETER aeConnection
      WorkflowCommander AE Connection object.

      .PARAMETER path
      Path to move object to. If path does not exist, it will be created.

      .EXAMPLE
      Move-aeObject -ae $ae -name OBJECTNAME -path /PRODUCTION/SYSTEMA/FOLDER1
      Creates /PRODUCTION/SYSTEMA/FOLDER1 (if necessary) and moves object.

      .LINK
      http://workflowcommander.wordpress.com

      .OUTPUTS
      Search result
  #>
  [CmdletBinding(SupportsShouldProcess)]
  param (
    [Parameter(Mandatory,HelpMessage='AE connection object returned by new-aeConnection.')]
    [Alias('ae')]
    [Object]$aeConnection,
    [Parameter(Mandatory,HelpMessage='Name of object to move.',ValueFromPipeline,ValueFromPipelineByPropertyName)]
    [string]$name,
    [Parameter(Mandatory,HelpMessage='Path to move object to.',ValueFromPipelineByPropertyName)]
    [string]$path
  )
 
  begin {
    Write-Debug -Message '** Move-aeObject start'
    $resultSet = @()
  }

  process {
    # Get object to relocate
    $object = Search-aeObject -aeConnection $aeConnection -name $name
    
    if ($object.result -ne 'OK') {
      Write-Warning ('! Object to move could not be found: ' + $name)
      return
    }
    
    # If object is already at the correct place, stop here
    if ($object.path -eq $path) {
      Write-Verbose -Message ('* ' + $name + ' is already at target destination.')
      $resultSet += $object
      return
    }
    
    # Create destination folder if necessary. If already available the command won't harm the system.
    $dstPath = New-aeFolder -aeConnection $aeConnection -path $path
    $srcPath = Get-aeFolder -aeConnection $aeConnection -path $object.path
        
    try {
      # This does not really make sense like this.
      $folderListRequest = [com.uc4.communication.requests.FolderList]::new($srcPath)
      $aeConnection.sendRequest($folderListRequest)
      $objectFolderList = $folderListRequest.findByName($name)
      
      $moveRequest = [com.uc4.communication.requests.MoveObject]::new($objectFolderList, $srcPath, $dstPath)
      $aeConnection.sendRequest($moveRequest)
    }
    catch {
      Write-Warning -message ('! Failed to move object to folder ' + $_)
      return
    }
    
    # Check whether object has been moved properly
    $object = Search-aeObject -aeConnection $aeConnection -name $name
    if ($object.path -ne $path) {
      Write-Warning -Message ('! Moving ' + $name + ' to ' + $path + ' failed.')
      $object.result = 'FAIL'
    }

    $resultSet += $object
  }

  end {
    Write-Debug -Message '** Move-aeObject ended'
    return $resultSet
  }
}