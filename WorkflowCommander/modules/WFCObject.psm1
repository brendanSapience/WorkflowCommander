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

# Returned after an object export request. This object type is also used for -showNoExport
function New-WFCObjectExportObject {
  param(
    [Parameter(Mandatory,HelpMessage='AE object name')]
    [string]$name,
    [string]$type,
    [string]$path,
    [IO.FileInfo]$file = $null,
    [ValidateSet('OK','EMPTY','FAIL')]
    [Parameter(Mandatory,HelpMessage='Result')]
    [string]$result
  )

  $Object = New-Object -TypeName PSObject
  $Object.PsObject.TypeNames.Insert(0, 'WFC.PS.AEObjectExport')
  
  Add-Member -InputObject $Object -NotePropertyMembers @{
    'Name' = $name;
    'Type' = $type;
    'Path' = $path;
    'File' = $file;
    'Result' = $result;
  }
    
  return $Object
}

# For empty com.uc4.api.StatisticSearchItem
function New-WFCEmptyStatisticResult {
  param(
    [AllowEmptyString()]
    [Parameter(Mandatory,HelpMessage='AE object name')]
    [string]$name,
    [ValidateSet('OK','EMPTY','FAIL')]
    [Parameter(Mandatory,HelpMessage='Result')]
    [string]$result
  )

  $Object = New-Object -TypeName PSObject
  $Object.PsObject.TypeNames.Insert(0, 'WFC.PS.AEStatisticItem')
  
  Add-Member -InputObject $Object -NotePropertyMembers @{
    'Name' = $name;
    'Result' = $result;
  }
    
  return $Object
}

# As we cannot create an empty com.uc4.api.SearchResultItem we create an own type.
function New-WFCEmptySearchResult {
  param(
    [Parameter(Mandatory,HelpMessage='AE object name')]
    [string]$name,
    [ValidateSet('EMPTY','FAIL')]
    [Parameter(Mandatory,HelpMessage='Result')]
    [string]$result
  )
  
  $emptyResult = New-Object -TypeName PSObject
  $emptyResult.PsObject.TypeNames.Insert(0, 'WFC.PS.AEEmptySearchResult')
  Add-Member -InputObject $emptyResult -NotePropertyMembers @{
    'Name' = $name; 
    'Type' = '';
    'Path' = '';
    'Title' = '';
    'Result' = $result 
  }
  return $emptyResult  
}

# To identify whether an import was successful or not, we return all imported objects. This type can also be
# used to identify a failed import with an UNDEF type.
function New-WFCImportResult {
  param(
    [AllowEmptyString()]
    [Parameter(Mandatory,HelpMessage='AE object name')]
    [string]$name,
    [string]$type,
    [string]$path,
    [Parameter(Mandatory,HelpMessage='File to import XML.')]
    [string]$file,
    [ValidateSet('OK','EMPTY','FAIL')]
    [Parameter(Mandatory,HelpMessage='Result')]
    [string]$result
  )
  
  $importResult = New-Object -TypeName PSObject
  $importResult.PsObject.TypeNames.Insert(0, 'WFC.PS.AEObjectImport')
  Add-Member -InputObject $importResult -NotePropertyMembers @{
    'Name' = $name; 
    'Type' = $type; 
    'Path' = $path; 
    'File' = $file;
    'Result' = $result;
  }
  return $importResult  
}