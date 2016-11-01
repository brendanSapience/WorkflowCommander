
@{
  RootModule = 'WorkflowCommander.psm1'
  ModuleVersion = '1.0.0'
#  CompatiblePSEditions = @()
  GUID = '82697cb4-cb69-4ed7-a218-c08e3e788ccd'
  Author = 'Joel Wiesmann (joel.wiesmann@gmail.com)'
  CompanyName = 'WorkflowCommander'
  Copyright = '(c) 2016 Joel Wiesmann. All rights reserved.'
  Description = 'Access the AE from PowerShell'
  PowerShellVersion = '3.0'
# PowerShellHostName = ''
# PowerShellHostVersion = ''
# DotNetFrameworkVersion = ''
# CLRVersion = ''
# ProcessorArchitecture = ''
  RequiredModules = @()
  RequiredAssemblies = @(
    'lib\Automic.dll'
  )
  ScriptsToProcess = @(
    'WorkflowCommander.enum.ps1'
  )
  TypesToProcess = @(
    'WorkflowCommander.types.ps1xml'
  )
  
  FormatsToProcess = @(
    'WorkflowCommander.format.ps1xml'
  )
  
  NestedModules = @(
    'lib\WorkflowCommanderCore.dll',
    'WorkflowCommander.internal.psm1'
  )

  FunctionsToExport = @(
   'New-aeFolder',
   'Search-aeObject',
   'Export-aeObject',
   'Import-aeObject',
   'Search-aeStatistic'
  )
  CmdletsToExport = @(
    'New-aeConnection'
  )
  VariablesToExport = '*'
  AliasesToExport = @()
# DscResourcesToExport = @()
# ModuleList = @()
# FileList = @()
PrivateData = @{
    PSData = @{
        # Tags = @()
        # LicenseUri = ''
        # ProjectUri = ''
        # IconUri = ''
        # ReleaseNotes = ''
    } 
}
# DefaultCommandPrefix = ''
}

