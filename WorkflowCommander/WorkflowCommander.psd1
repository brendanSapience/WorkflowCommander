
@{
  RootModule = 'WorkflowCommander.psm1'
  ModuleVersion = '1.0.0'
#  CompatiblePSEditions = @()
  GUID = '82697cb4-cb69-4ed7-a218-c08e3e788ccd'
  Author = 'Joel Wiesmann (joel.wiesmann@gmail.com)'
  CompanyName = 'WorkflowCommander'
  Copyright = '(c) 2017 Joel Wiesmann. All rights reserved.'
  Description = 'Access the AE from PowerShell'
  PowerShellVersion = '3.0'
# PowerShellHostName = ''
# PowerShellHostVersion = ''
# DotNetFrameworkVersion = ''
# CLRVersion = ''
# ProcessorArchitecture = ''
#  RequiredModules = @(  )
  RequiredAssemblies = @(
    'lib\Automic.dll'
  )
#  ScriptsToProcess = @(  )
  TypesToProcess = @(
    'types\WFC.type.ps1xml',
    'types\WFCSearch.type.ps1xml',
    'types\WFCStatistic.type.ps1xml'
  )
  
  FormatsToProcess = @(
    'formats\WFC.format.ps1xml',
    'formats\WFCSearch.format.ps1xml',
    'formats\WFCStatistic.format.ps1xml',
    'formats\WFCExportImport.format.ps1xml'
  )
  
  NestedModules = @(
    'modules\WFCGetObject.psm1',
    'modules\WFCSearch.psm1',
    'modules\WFCStatistic.psm1',
    'modules\WFCExportImport.psm1',
    'modules\WFCObjectFolderStructure.psm1',
    'modules\WFCObject.psm1'
  )

  FunctionsToExport = '*'
  VariablesToExport = @('WFCPROFILES', 'WFCSTATUS')
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

