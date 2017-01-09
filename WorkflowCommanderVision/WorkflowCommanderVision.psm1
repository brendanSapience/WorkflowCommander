#########################################################################################
# WorkflowCommanderVision, copyrighted by Joel Wiesmann, 2016
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


#################################################################################################################################################################
# INTERNAL: Process stencil fields and set values
#################################################################################################################################################################
function set-shapeData() {
  param(
    [Parameter(Mandatory)]
    [Object]$taskData,
    [Parameter(Mandatory)]
    [Object]$shape
  )

  # Does the shape have any data fields?
  try {
    $dataFieldCount = $shape.rowcount(243)
  }
  catch { 
    return
  }

  # If our taskdata contains a field i.e. "demo" and we have a field named "demo" in the visio stencil, 
  # this field will receive the value of the according taskdata.
  if ($dataFieldCount -gt 0) {
      for ($fieldCount = 0; $fieldCount -lt $dataFieldCount; $fieldCount++) {
        $dataRow = $shape.CellsSRC($visSectionProp, $fieldCount, $visCustPropsValue).name
        $name = $shape.Cells($dataRow).RowNameU
        try {
          $shape.Cells($dataRow).Formula = [string]('"' + ($taskData.$name -replace '"', '') + '"')
        }
        catch {
          Write-Warning -Message ('! Cannot set field named ' + $name)
        }
      }
  }

  # Recursively call data assignment function if there are any sub-shapes..
  $subShapeCount = $shape.Shapes.count
  if ($subShapeCount -ne 0) {
      for ($shapeNum = 1; $shapeNum -le $subShapeCount; $shapeNum++) {
          set-shapeData  -taskData $taskData -shape $shape.Shapes.item($shapeNum)
      }
  }
}

#################################################################################################################################################################
# INTERNAL: Convert data structure to Visio
#################################################################################################################################################################
function convert-dataToWorkflow() {
  param(
    [Parameter(Mandatory)]
    [Object]$page,
    [Parameter(Mandatory)]
    [Object]$workflow,
    [Parameter(Mandatory)]
    [Object]$mastershapes,
    [Parameter(Mandatory)]
    [Float]$xSpacing,
    [Parameter(Mandatory)]
    [Float]$ySpacing
  )

  # Layout options making look the workflow like a workflow.
  $page.PageSheet.CellsSRC($visSectionObject, $visRowPageLayout, $visPLORouteStyle).FormulaForceU     = '6'
  $page.PageSheet.CellsSRC($visSectionObject, $visRowPageLayout, $visPLOLineAdjustFrom).FormulaForceU = '1'
  $page.PageSheet.CellsSRC($visSectionObject, $visRowPageLayout, $visPLOLineAdjustTo).FormulaForceU   = '2'
  $page.PageSheet.CellsSRC($visSectionObject, $visRowPageLayout, $visPLOLineRouteExt).FormulaForceU   = '1'

  $shapes = @{}
  # Iterate through taskdata. Each record will result in one "box"
  write-verbose -message '* Drawing tasks...' 
  foreach ($task in $workflow.'taskData') {
    Write-Debug -Message '** Dropping task'
    # Verify if stencil is known. If not, drop the default stencil.
    if (! $mastershapes.([String]$task.type)) {
      $obj = $mastershapes.'default'
    }
    else { 
      $obj = $mastershapes.([String]$task.type) 
    }
    
    # Drop stencil onto page, use the x/y coordinates we get from the export.
    $shapes.($task.lnr) = $page.Drop($obj, ([int]$task.x * $xSpacing), $task.y * $ySpacing)
    Write-Verbose -Message ('* Dropping shape for task lnr ' + $task.lnr + ' (' + $task.name + ')')
    set-shapeData -taskData $task -shape $shapes.($task.lnr)
  }
  
  # Now the same for relation. Each relation record is one connector between 2 task-shapes.
  Write-Verbose -Message '* Drawing relations'
  foreach ($relation in $workflow.'relData') {
    if (-not $relation.prelnr -or -not $relation.lnr) { continue }
    # AutoConnect won't return a shape so we set a unique name to find the connector
    Write-Verbose -Message ('* Linking ' + $relation.lnr + ' to predecessor ' + $relation.prelnr)
    $connectorShape       = $mastershapes.'connector'
    $connectorName        = ('con_' + $relation.lnr + ':' + $relation.prelnr)
    $connectorShape.nameU = $connectorName

    try {
      $shapes.($relation.lnr).AutoConnect($shapes.($relation.prelnr), 0, $connectorShape)
    }
    catch {
      write-warning -message('! Linking lnr: ' + $relation.lnr + ' with prelnr: ' + $relation.prelnr + ' failed.') 
    }
  }
  
  # Finally resize to fit contents so it looks pretty.
  $page.ResizeToFitContents()
}


#################################################################################################################################################################
# INTERNAL: Save / saveas / export current page
#################################################################################################################################################################
function save-drawing() {
  param(
    [Parameter(Mandatory)]
    [Object]$file,
    [Parameter(Mandatory)]
    [Object]$document,
    [Parameter(Mandatory)]
    [Object]$page
  )

  # If we edited an existing VSD, we will use save(). For saving new VSDs, we use safeas() and for other file formats we have to export().
  # This is basically such a simple thing but the various possibilities make it complex.
  if ($file.Extension -match '.vsd') {
    if ($file.Exists) {
      if ($document.save() -eq 0) { 
        Write-Verbose -message ('* Document successfully saved to ' + $file.FullName)
      }
      else { 
        Write-Warning -message ('! Saving ' + $file.FullName + ' failed!')
      }
    }
    else {
      if ($document.saveas($file.FullName) -eq 0) { 
        Write-Verbose -message ('* Document successfully saved as ' + $file.FullName)
      }
      else { 
        Write-Warning -Message ('! Saving ' + $file.FullName + ' failed!') 
      }
    }
  }
  else {
    if ($page.export($file.FullName) -ne 0) { 
      Write-Verbose -message ('* Document successfully exported to ' + $file.FullName)
    }
    else { 
      Write-Warning -message ('! Exporting to ' + $file.FullName + ' failed!')
    }
  }
}

#################################################################################################################################################################
# Cmdlet for converting ae Workflows to Visio
#################################################################################################################################################################

function get-aeVisionData() {
  <#
      .SYNOPSIS
      Convert either SQL data dumps to visionData array // get data directly from AE to produce visionData.

      .DESCRIPTION
      WFC::Vision has a very own data format it expects to draw the workflow. This format is being created by this
      cmdlet either out of a task & relation file or based on AE data.

      .PARAMETER name
      If data should be gathered form the AE - name of the workflow.

      .PARAMETER aeConnection
      If data should be gathered form the AE - WFC::Core connection to get the object from.

      .PARAMETER csvTaskDataFile
      CSV backend file containing the task definition. See https://workflowcommander.wordpress.com. 

      .PARAMETER csvRelDataFile
      CSV backend file containing the task definition. See manual.

      .EXAMPLE
      get-aeVisionData -ae $ae -name JOELS_WORKFLOW
      Get the workflow JOELS_WORKFLOW in Vision-compatible format. Can be piped to convert-aeWorkflowToVisio.

      .NOTES
      See Philipp Elmer's article on Workflow Vision https://www.philippelmer.com/gastbeitrag-mit-workflow-vision/
  #>

  param(
    [Parameter(ParameterSetName='fromFile')]
    [Parameter(ParameterSetName='WFC',ValueFromPipeline,ValueFromPipelineByPropertyName,HelpMessage='Workflow object name to load and convert',Mandatory)]
    [string]$name,
    # WorkflowCommander support
    [Alias('ae')]
    [Parameter(ParameterSetName='WFC',Mandatory)]
    [Object]$aeConnection,
    # This works like the original "Workflow Visio(n)"
    [Parameter(ParameterSetName='fromFile',Mandatory)]
    [IO.FileInfo]$csvTaskDataFile,
    [Parameter(ParameterSetName='fromFile',Mandatory)]
    [IO.FileInfo]$csvRelDataFile
  )
  
  Begin {
    # Multidimensional table ($visionData) holds all data that is required to draw. Because we create an interim data array we can 
    # very easily mutate data before we convert it.
    # The format is:
    # $workflows.<workflowname>
    #                          .'taskdata' = Hash array with minimum keys lnr, x, y, type
    #                          .'reldata'  = Hash array with minimum keys prelnr, lnr
    $workflows = @{}
  }

  Process {
    # Prepare data structure
    $workflows.$name = @{}
    $taskData = @()
    $relData = @()

    #################################################################################################################################################################
    # Load data from export file
    #################################################################################################################################################################
    if ($PSCmdlet.ParameterSetName -eq 'fromFile') {
      Write-Verbose -Message '* Loading data from file...'
      # Load / verify the workflow data 
      try {
        $taskData = ConvertFrom-Csv (Get-Content -Path $csvTaskDataFile) -Delimiter ';'
        $relData  = ConvertFrom-Csv (Get-Content -Path $csvRelDataFile)  -Delimiter ';'
      }
      catch {
        Write-Warning -Message ('! Issues loading data (' + $_.Exception.Message + ').')
        return
      }
    }
    
    #################################################################################################################################################################
    # Get data from Automation Engine
    #################################################################################################################################################################
    if ($PSCmdlet.ParameterSetName -eq 'WFC') {
      Write-Verbose -Message ('* Getting ' + $name + ' workflow data...')
      $aeObject = Get-aeObject -aeConnection $aeConnection -name $name

      if ($aeObject.getType() -ne 'JOBP') {
        Write-Warning -Message ('! ' + $name + ' is not existing or not of type JOBP!')
        $workflows.Remove($name)
        return
      }

      $tasks = $aeObject.getUC4Object().taskIterator()
      while ($tasks.hasnext()) {
        Write-Debug -Message '* And another task...'
        $task = $tasks.next()
        
        # For some strange reason this seems not to work. Might be an API issue to be analysed later..
        # if ($task.isInactive()) { $inact = 0 } else { $inact = 1 }
        
        $taskData += @{
          'lnr'         = $task.getLnr()
          'name'        = $task.getTaskName()
          'description' = $task.getTaskTitle()
          'x'           = $task.getX()
          'y'           = $task.getY()
          'data'        = ''
          'agent'       = $task.getHostName()
          'type'        = $task.getType()
        }

        $dependencies = $task.dependencies().iterator()
        while ($dependencies.hasNext()) {
          $dependency = $dependencies.next()
          $relData += @{
            'lnr'    = $task.getLnr()
            'preLnr' = $dependency.getTask().getLnr()
          }
        }
      }
    }

    # Mirror the Y-axis in the task specification
    $maxYvalue = ($taskData.GetEnumerator() | Sort-Object -Property Y -Descending | Select-Object -First 1).Y
    foreach ($task in $taskData) {
      $task.Y = ($task.Y - $maxYvalue) * -1
    }

    # Assign data and proceed with next...
    $workflows.$name.'taskData' = $taskData
    $workflows.$name.'relData'  = $relData
  }

  End {
    return $workflows
  }
}
function convert-aeWorkflowToVisio() {
  <#
      .SYNOPSIS
      Convert an AE Vision Data Array to a Visio representation.

      .DESCRIPTION
      This cmdlet allows you to export an AE workflow to either Visio VSD or any supported export format (JPG, PNG, etc.).

      .PARAMETER visionData
      Array that contains the visionData. Might also be piped from get-aeVisionData

      .PARAMETER file
      Where to write the output to. This can be a folder or file. If folder is specified, the file will be named like the workflow, so if
      multiple workflows are converted, each one will get it's own file.

      .PARAMETER extension
      When working with an output folder, the extension defines the ending of the files. All supported Visio formats can be specified here.

      .PARAMETER stencilFile
      Stencil file to use. You can do a stencil on your own - just name the shapes like the AE types (JOBP, JOBF... + a default)

      .PARAMETER xSpacing
      Multiplier for spacing between shapes on x axis. The higher the more space inbetween shapes.

      .PARAMETER ySpacing
      Multiplier for spacing between shapes on y axis. The higher the more space inbetween shapes.

      .PARAMETER visioVisible
      If set to $true, you can watch Visio building the workflow.

      .EXAMPLE
      search-aeObject -ae $ae -type jobp | get-aeVisionData -ae $ae | convert-aeWorkflowToVisio -file c:\temp\demo.vsd
      Gets all workflows on $ae using WFC::Core connection $ae and charts them into c:\temp\dmo.vsd.

      .NOTES
      See Philipp Elmer's article on Workflow Vision https://www.philippelmer.com/gastbeitrag-mit-workflow-vision/
  #>

  param(
    [Alias('folder')]
    [Parameter(Mandatory,HelpMessage='Output file or folder')]
    [IO.FileInfo]$file,
    [Parameter(Mandatory,ValueFromPipeline,HelpMessage='Array containing data to draw')]
    [Hashtable[]]$visionData,
    [String]$extension = 'vsd',
    [IO.FileInfo]$stencilFile = (([IO.FileInfo](Get-Module WorkflowCommanderVision).path).Directory.ToString() + '\data\default.vss'),
    [double]$xSpacing = 2.5,
    [double]$ySpacing = 1.5,
    [bool]$visioVisible = $false
  )

  #################################################################################################################################################################
  # Initialize Visio and preload stencils
  #################################################################################################################################################################
  try {
    $application = New-Object -ComObject Visio.Application
    $application.visible = $visioVisible
    $documents = $application.Documents
    
    # Load the mastershapes from the stencil
    $mastershapes = @{}
    $objectStencil = $application.Documents.Add($stencilFile)
    foreach ($stencil in $Objectstencil.Masters) {
      $mastershapes.($stencil.name) = $stencil
    }
  
    if (! $mastershapes.'default') {
      Write-Warning -Message ('! ' + $stencilFile + ' does not contain a "default" mastershape.')
    }
  
    if (! $mastershapes.'connector') {
      Write-Warning -Message ('! ' + $stencilFile + ' does not contain a "connector" mastershape.')
    }
  }
  catch {
    Write-Warning -Message ('! Could not initialize Visio (' + $_.Exception.Message + ').')
    exit 1
  }
  
  # Outputfile might be a folder or a file. In case that it's a folder we will export one file per workflow
  if ((Get-Item -path $file -ErrorAction SilentlyContinue) -is [System.IO.DirectoryInfo]) {
    $outputToFolder = $file
  }
  
  #################################################################################################################################################################
  # Visualize
  #################################################################################################################################################################
  # There is a lot of Visio document/page management ongoing here. This is because of the various possibilities to save the files.
  # Visio must be managed differently whether you're exporting to other formats or create additional drawings within the same 
  # VSD document.
  foreach ($workflow in $visionData.keys) {
    Write-Verbose -Message ('* Processing ' + $workflow)

    # In case we output to a folder, we need to determine the final filename
    if ($outputToFolder) {
      $file = ($outputToFolder.FullName + '/' + $workflow + '.' + $extension)
    }
  
    if ($file.extension -eq '.vsd' -or $extension -match 'vsd') {
      $visioFormat = $true
    }
  
    # If output format is vsd AND the output file is already existing, the documentation might be added to the existing tab 
    # or the existing page gets replaced.
    $file.refresh()
    if ($visioFormat -and $file.Exists) {
      try { $document = $documents.Open($file.FullName) }
      catch {
        write-warning -Message ('! Could not open visio file for writing. It is likely that the file is already opened. Please kill all open visio processes and try again.')
        $application.quit()
        exit 1
      }
    }
    else {
      # Create new document and delete the empty default page.
      $document = $documents.Add('') 
      $removeFirst = $true
    }
  
    # If our Visio document already contains a tab named as the workflow we're going to create, we need to remove the tab and recreate the content.
    # The problem is, that if there is only one tab at all in the document, the removal of the tab will automatically create a new page. So we first
    # detect if there is something to delete and if so, we will finally delete the page after we created the new one.
    for ($pageNumber = 1; $pageNumber -le $document.Pages.count; $pageNumber++) {
      if ($document.pages.item($pageNumber).NameU -eq [string]$workflow) {
        $del = $pageNumber
      }
    }

    $page = $document.Pages.Add()
    
    if ($del) {
      write-verbose -message '* Deleting existing page with same pagename for recreation.'
      $document.pages.item($del).delete(1)
    }
    $page.name = [string]$workflow
  
    # Fill page with content
    convert-dataToWorkflow -page $page -mastershapes $mastershapes -workflow $visionData.$workflow -xSpacing $xspacing -ySpacing $ySpacing
  
    if ($removeFirst) {
      $document.pages.item(1).delete(1) 
    }
     
    save-drawing -document $document -page $page -file $file
    $document.saved = $true
    $document.close()
  }
  
  #################################################################################################################################################################
  # Finalize / shutdown
  #################################################################################################################################################################  
  # Properly close Visio or there will be hanging processes.
  $objectStencil.saved = $true
  $application.quit()
}