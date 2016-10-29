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
$VerbosePreference = 'Continue'

#########################################################################################
# Internal functions
#########################################################################################
function _new-aeObjectExportObject {
  param(
    [Parameter(Mandatory,HelpMessage='AE object name')]
    [string]$name,
    [string]$type,
    [Parameter(Mandatory,HelpMessage='XML filename of exported AE object')]
    [string]$path,
    [IO.FileInfo]$file = $null
  )

  $Object = New-Object -TypeName PSObject
  $Object.PsObject.TypeNames.Insert(0, 'WFC.PS.AEObjectExport')
  
  Add-Member -InputObject $Object -NotePropertyMembers @{
    'Name' = $name;
    'Type' = $type;
    'Path' = $path;
    'File' = $file
  }
    
  return $Object
}
