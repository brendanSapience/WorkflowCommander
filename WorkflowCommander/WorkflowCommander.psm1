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
#########################################################################################

# This is about the only setting that you might want to adapt to your environment.
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

function new-aeConnection {
  <#
      .SYNOPSIS
      Create connection object to an Automation Engine instance.

      .DESCRIPTION
      WorkflowCommander allows you to connect to multiple AE instances at the same time. To identify these, 
      new-aeConnection will return a connection object. This connection object must  be specified to the
      commandlets so the commandlets know on what AE they should work.

      .PARAMETER profile
      Load previously saved connection profile. This will directly connect you with the AE server.

      .PARAMETER client
      AE client number to connect to.

      .PARAMETER server
      Servername / IP address of your AE system.

      .PARAMETER username
      Username to use for AE login.

      .PARAMETER password
      Password to use for AE login.

      .PARAMETER department
      Optionally the department to use.

      .PARAMETER port
      AE port if not standard.

      .PARAMETER saveAsProfile
      Saves the connection information to a XML file that can be loaded using -profile.

      .EXAMPLE
      $ae = new-aeConnection -client 1000 -server 127.0.0.1 -username admin -password myLuckyPwd123
      Connects to client 1000 on 127.0.0.1.

      .EXAMPLE
      $ae = new-aeConnection -client 1000 -server 127.0.0.1 -username admin -password myLuckyPwd123 -saveAsProfile client1000
      Same connection as above and saves the connection data to a profile named "client1000".

      .EXAMPLE
      $ae = new-aeConnection -profile client1000
      Loads the profile "client1000".

      .LINK
      http://automationfreak.blogspot.com

      .OUTPUTS
      WorkflowCommander Connection Object.
  #>
  param (
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
    [SecureString]$password,
    [Parameter(ParameterSetName = "listonly",Mandatory)]
    [Switch]$listProfiles
  )

  Write-Warning '========================================================'  
  Write-Warning '!   THIS SOFTWARE IS EXPERIMENTAL - READ THE LICENSE   !'
  Write-Warning '! Have fun, but be careful and know what you are doing !'
  Write-Warning '========================================================'
  Write-Warning '* Use -verbose switch for extended output or set $ErrorActionPreference to continue.'
  Write-Warning '* Use -whatif switch to simulate importing & exporting objects.'
  Write-Warning '> For more information, visit https://github.com/JoelWiesmann/WorkflowCommander'
  write-warning '> Send feedback to joel.wiesmann@gmail.com'
  
  # TODO: This is just a beta implementation. Do not use in your scripts as it should return an object. 
  if ($listProfiles) {
    foreach ($xmlFile in (Get-ChildItem -Path ($global:WFCPROFILES + '*.xml'))) {
      Write-Debug ('* Analyzing ' + $xmlFile)
      [xml]$potentialProfile = get-content -path $xmlFile
      if ($potentialProfile.'WFCProfile') {
        write-verbose  ($xmlFile.name -replace '.xml','')
      }
      else {
        Write-Verbose -Message ('* ' + $xmlFile + ' is no WFC profile')
      }
    }
    return  
  }
  
  # If a profile has been specified - load it or die trying.
  # The profile contains all necessary settings to connect to an AE system. Also it might contain user preferences / default settings.
  if ($profile) {
    try { 
      $aeProfile = Get-aeProfile -profileName $profile 
    }
    catch {
      throw($_.exception)
      return
    }
  }
  else {
    $aeProfile = New-aeProfile -username $username -department $department -server $server -port $port -client $client
  }
  
  # Open up the AE connection and try to login. The most common exception that could happen here
  Write-Verbose -Message ('* Connecting to Automation Engine (' + $aeProfile.workflowCommander.server + ':' + $aeProfile.workflowCommander.port + '), please wait...')

  try {
    $aeConnection = [com.uc4.communication.Connection]::Open($aeProfile.WorkflowCommander.server, $aeProfile.WorkflowCommander.port)
    $aelogin = $aeConnection.login(
      $aeProfile.WorkflowCommander.client,
      $aeProfile.WorkflowCommander.username,
      $aeProfile.WorkflowCommander.department,
      $aeProfile.WorkflowCommander.password,
      'E'
    )
  }
  catch {
    throw ('! Connection to the Automation Engine could not be established. Errormessage: ' + $_.exception)
    return
  }

  if (! $aelogin.isLoginSuccessful()) {
    throw ('! Login was not successful: ' + $aelogin.getMessageBox())
    return
  }
  
  Write-Verbose -Message '* Connection successfully established'
  $systemInfo = $aeConnection.getSessionInfo()
  Write-Verbose -Message ('** Username:   ' + $systemInfo.getUserName())
  Write-Verbose -Message ('** Department: ' + $systemInfo.getDepartment())
  Write-Verbose -Message ('** System:     ' + $systemInfo.getSystemName()) 
  Write-Verbose -Message ('** Version:    ' + $systeminfo.getServerVersion())
  write-verbose -Message ('** Client:     ' + $systemInfo.getClient())

  # If the login was successful and it was requested to store the profile, do so now.
  if ($saveAsProfile) {
    Save-aeProfile -profile $aeProfile -profileName $saveAsProfile
  }

  # Remove the password as it would be cleartext in memory..
  $aeProfile.WorkflowCommander.Password = 'Nope :).'

  return New-aeConnectionObject -aeConnection $aeConnection -aeLogin $aelogin -aeProfile $aeProfile -aeSystemInfo $systemInfo
}

function new-aeConnectionObject {
  param (
    [Parameter(mandatory,HelpMessage='AE connection object returned by new-aeConnection.')]
    [Alias('ae')]
    [Object]$aeConnection,
    [Parameter(mandatory,HelpMessage='AE login object.')]
    [Object]$aeLogin,
    [Parameter(mandatory,HelpMessage='AE profile object.')]
    [Object]$aeProfile,
    [Parameter(mandatory,HelpMessage='AE system info sobject.')]
    [Object]$aeSystemInfo
  )
  
  $Object = New-Object -TypeName PSObject
  $Object.PsObject.TypeNames.Insert(0, 'WFC.Core.WFCConnection')
  
  Add-Member -InputObject $Object -NotePropertyMembers @{
    'aeConnection' = $aeConnection;
    'aeLogin'      = $aeLogin;
    'aeProfile'    = $aeProfile;
    'aeSystemInfo' = $aeSystemInfo;
  }
  
  # Add a method to wrap the requests
  Add-Member -InputObject $Object -MemberType ScriptMethod -Name "SendRequest" -Value {
    param([Object]$request)
    $this.'aeConnection'.sendRequestAndWait($request)
  }
  
  return $Object
}

function new-aeProfile {
  param(
    [Parameter(mandatory,HelpMessage='Username to login into AE system')]
    [string]$username,
    [string]$department = $null,
    [Parameter(mandatory,HelpMessage='Password to login into AE system')]
    [securestring]$password,
    [Parameter(mandatory,HelpMessage='Server / IP to AE system')]
    [string]$server,
    [Parameter(mandatory,HelpMessage='TCP port of AE system')]
    [string]$port,
    [Parameter(mandatory,HelpMessage='Client number to login')]
    [int]$client  
  )
  
  # This reverts the securestring. 
  $undoEncryption = new-object System.Net.NetworkCredential('', $password)
  Remove-Variable password
  $password = $undoEncryption.Password

  # Create settings directory if not already existing
  $null = new-item -ItemType Directory -force -path $global:WFCPROFILES 
  
  # We create a basic profile XML and then input the variables as string values.
  [xml]$profileXML = '<WorkflowCommander><version/><username/><department/><password/><server/><port/><client/></WorkflowCommander>'
  # This might get useful when we introduce new features and the profile must be updated.
  $profileXML.WorkflowCommander.version    = '1' 
  $profileXML.WorkflowCommander.username   = [string]$username
  $profileXML.WorkflowCommander.department = [string]$department
  $profileXML.WorkflowCommander.password   = [string]$password
  $profileXML.WorkflowCommander.server     = [string]$server
  $profileXML.WorkflowCommander.port       = [string]$port
  $profileXML.WorkflowCommander.client     = [string]$client

  return $profileXML
}

function save-aeProfile {
  param(
    [Parameter(mandatory,HelpMessage='XML object that contains profile information.')]
    [xml]$profile,
    [Parameter(mandatory,HelpMessage='Name of the profile. Will be part of output XML file.')]
    [string]$profileName
  )

  $profileAbsoluteName = ($global:WFCPROFILES + $profileName + '.xml') 

  try { 
    # Convert the password to a securestring object so the password won't be stored unencrypted
    $profile.WorkflowCommander.password = [string]($profile.WorkflowCommander.password | ConvertTo-SecureString  -AsPlainText -Force | ConvertFrom-SecureString)
    $null = $profile.outerxml | Out-File -FilePath $profileAbsoluteName -ErrorAction Stop 
  }
  catch {
    Write-Warning -Message ('! Profile could not be written. Check if ' + $global:WFCPROFILES + ' exists and is writeable.')
  }
  finally {
    Write-Verbose -Message ('* Profile saved to ' + $profileAbsoluteName)
  }
}

function get-aeProfile {
  param(
    [Parameter(mandatory,HelpMessage='Name of profile to load.')]
    [string]$profileName
  )

  $profileFile = ($global:WFCPROFILES + $profileName + '.xml')
  Write-Verbose -Message ('* Loading profile ' + $profileFile)
  
  # Try to read the profile XML file (if available). There is no check yet whether the XML file is really an "aeProfile". 
  try {
    [xml]$profileXML = Get-Content -Path $profileFile -ErrorAction Stop
   
    # Revert the password encryption. 
    $undoEncryption = new-object System.Net.NetworkCredential('', ($profileXML.WorkflowCommander.password | ConvertTo-SecureString))
    $profileXML.WorkflowCommander.password = $undoEncryption.Password
    
  }
  catch {
    throw ('Could not load profile named ' + $profileName + '. Please check whether a profile XML exists in ' + $global:WFCPROFILES)
    return $null
  }

  return $profileXML
}