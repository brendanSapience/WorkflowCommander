#########################################################################################
# WorkflowCommander, copyrighted by Joel Wiesmann, 2017
# <  THIS CODE IS EXPERIMENTAL  >
#
# https://github.com/JoelWiesmann/WorkflowCommander
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
#
# Added by Brendan Sapience - Jan 16 2017
#########################################################################################

function create-aeObject {
  <#
      .SYNOPSIS
      Equals the AE "search for object".

      .DESCRIPTION
      create-aeObject creates AE Objects based on a Template, Name, Folder and Title

      .PARAMETER aeConnection
      WorkflowCommander Connection Object (new-aeConnection).

      .PARAMETER name
      Name of the object to search. Supports wildcards like the GUI search. Defaults to "*"

      .PARAMETER path
      Path where the object is to be stored. By default "/". 

      .PARAMETER template
      Type of object to create

      .PARAMETER title
      Title of the object to be created // only works for FOLD Objects!

      .EXAMPLE
      create-aeObject -ae $ae -name "MY.JOB" -path "/PRODUCTION" -template "JOBS_WINDOWS" -title "My Object Title"

      .LINK
      

      .OUTPUTS
      
  #>
   Param(
    [Parameter(Mandatory,HelpMessage='AE connection object returned by new-aeConnection.')]
    [Alias('ae')]
    [Object]$aeConnection,
    [Parameter(ValueFromPipeline,ValueFromPipelineByPropertyName)]
    [string]$name = '',
    [Parameter(ValueFromPipelineByPropertyName)]
    [string]$path = $null,
    [Parameter(ValueFromPipelineByPropertyName)]
    [ValidateSet('CALE','CALL','CODE','CPIT','DOCU','EVNT_CONS','EVNT_FILE','EVNT_TIME','HOSTG','JOBF','JOBG','JOBI','JOBP','JOBP_IF','JOBP_FOREACH',
    'JOBQ_PS','JOBQ_R3','JOBS_BS2000','JOBS_GCOS8','JOBS_JMX','JOBS_MPE','JOBS_NSK','JOBS_OA','JOBS_MVS','JOBS_OS400','JOBS_PS','JOBS_R3','JOBS_SIEBEL','JOBS_SQL',
    'JOBS_UNIX','JOBS_VMS','JOBS_WINDOWS','JSCH','LOGIN','SCRI','STORE','SYNC','TZ','USER','USRG','VARA','FOLD','QUEUE','DASH','JOBS.SAP','PERIOD','SLO'
    )]
    [string]$template = $null,
    [string]$title = $null
    
  )

  begin {
    Write-Debug -Message '** create-aeObject start'
    $resultSet = @()
  }
  
  process {

  # converting the template name passed to a Template object
  $templateO = [com.uc4.api.Template]::getTemplateFor($template)

  if (!$templateO){
    Write-Error "Error, Template $template does not appear to be a valid Automic Template."
    return
  }

  # converting the path passed to a IFolder object
  $folderO = get-aeFolder -ae $ae -path $path
  
  if (!$folderO){
    Write-Error "Error, Folder $path does not appear to be a valid or existing path."
    return
  }

  # creating a request for object creation
  $req = [com.uc4.communication.requests.CreateObject]::new($name,$templateO,$folderO)

  # setting the title, if any. this is only for FOLD objects..
  $req.setTitle($title)

  # Submit the request to AE..
    try {
      $ae.sendRequest($req)
    }
    catch {
      Write-Warning -message ('! Failed to query the AE: ' + $req.getAllMessageBoxes() + ' ' + $_.Exception.GetType().FullName + ' ' + $_.Exception.Message)
      return
    }

    # if the request has a non-NULL message box, then something went wrong.. showing the error message in that case.
    if($req.getMessageBox()){
        $MessageBoxObj = $req.getMessageBox()
        Write-Warning "$MessageBoxObj"
        return
    }
  }
  
  
  end {
    Write-Debug -Message '** Create-aeObject end'
    return ($resultSet | Sort-Object -property Path)
  }
}
