@{
  RootModule = 'WorkflowCommanderVision.psm1'
  ModuleVersion = '1.0'
# CompatiblePSEditions = @()
  GUID = '60f86de4-af2f-4393-b9e7-a09cc55efcf2'
  Author = 'Joel Wiesmann (joel.wiesmann@gmail.com)'
  CompanyName = 'WorkflowCommander'
  Copyright = '(c) 2017 Joel Wiesmann. All rights reserved.'
  Description = 'Convert AE workflows to Visio'
  PowerShellVersion = '3.0'
# PowerShellHostName = ''
# PowerShellHostVersion = ''
# DotNetFrameworkVersion = ''
# CLRVersion = ''
# ProcessorArchitecture = ''
  RequiredModules = @( 'WorkflowCommander' )
  RequiredAssemblies = @(  )
  ScriptsToProcess = @('data\visioEnum.ps1')
  TypesToProcess = @()
  FormatsToProcess = @()
  
  NestedModules = @()
  FunctionsToExport = @('convert-aeWorkflowToVisio', 'get-aeVisionData')
  VariablesToExport = @()
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