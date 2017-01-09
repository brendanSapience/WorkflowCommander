# Installation
Precondition:
* one or more AE 10+ systems you can connect to from your workstation
* a Powershell v3+ installation
* for Workflow visualization you will need a Visio installation

To install the modules:
- Download / git clone this repository (you basically need the WorkflowCommander* folder contents).
- Download the Automic AE GUI Client for AE V12 (for AE V10 you must download the AE V11.2 client).
- Search for an "Automic.dll" file and copy it into the WorkflowCommander/lib directory
- Start a PowerShell session and move (or copy) the WorkflowCommander directories to the Powershell module folder:
```powershell
move-item c:\temp\WorkflowCommander* $HOME\Documents\WindowsPowerShell\Modules
```

To verify installation, this command should return the module information. If nothing comes back, check out google for Powershell Module installation.
```powershell
Get-Module -name WorkflowCommander
Get-Module -name WorkflowCommanderVision
``` 

# WorkflowCommander (WFC::CORE)
## Some very basic things to begin
Commands usually return result objects. Those contain a property named "result" that may contain the following string:
* *EMPTY* (i.e. object not found, no statistic entry found)
* *OK* (i.e. import was successful, object was found)
* *FAIL* (i.e. export failed, any kind of connection issues)

Now enjoy the examples. They will help you to get into this useful set of cmdlets.

## Connecting and disconnecting
```powershell
$ae = new-aeConnection -server 127.0.0.1 -port 2217 -client 2000 -user ADMIN -department HR -saveAsProfile client2k
```
Next time you can connect using the profile:
```powershell
$ae = new-aeConnection -profile client2k
```
To disconnect, close your powershell session or remove the connection variable.
```powershell
remove-variable ae
```

##Search for objects

```powershell
search-aeObject -ae $ae -name MYOBJECT
```

Search for all JOBF objects within /DEMO
```powershell
search-aeObject -ae $ae -path /DEMO -type JOBF
```
Search for objects having a title containing "Import":
```powershell
search-aeObject -ae $ae -textType title -text "Import"
```

Search for objects named OBJECTA and OBJECTB:
```powershell
@("OBJECTA","OBJECTB") | search-aeObject -ae $ae
```

As you can pipe directly, you could also get those object names from a file:
```powershell
get-content obj.txt | search-aeObject -ae $ae
```

...or a CSV. The headers must equal to the parameter. Notice that you have a heading named "path" search-aeObject will consider this as well.
```powershell
Get-Content .\test.csv | ConvertFrom-Csv -Delimiter ";" | search-aeObject -ae $ae
```

Check whether all objects in /DEMO on clientA are available on clientB:
```powershell
$clientA = new-aeConnection -profile clientA
$clientB = new-aeConnection -profile clientB
search-aeObject -ae $clientA -path /DEMO | search-aeObject -ae $clientB
```

##Statistics

Show latest statistic entry of an object named MYOBJ:
```powershell
get-aeStatistic -ae $ae -name MYOBJ
```

Get last ten statistic entries (or maximum amount of found entries)
```powershell
get-aeStatistic -ae $ae -name MYJOB -amount 10
```

Search for all objects below /DEMO and show last statistic entry:
```powershell
search-aeObject -ae $ae -path /DEMO | get-aeStatistic -ae $ae
```

##Periods search
To be done...

##Object export
WFC exports one object into one file - if you export ten objects, you will get 10 XML files. This is handy in cases where you do source control on these single XML files. The export command will return what object has been exported with which path information. The object export will encode the object's path information into the XML file.
Export a single object by it's name
```powershell
export-aeObject -ae $ae -name MYOBJ -file c:\temp\out.xml
```

If the -file specifies a folder, the name of the XML file equals the object name. This allows you to export multiple objects at once. The list of objects might be gathered by a search:
```powershell
search-aeObject -ae $ae -path /DEMO | export-aeObject -ae $ae -file c:\temp
```

If you have a file that contains object names:
```powershell
get-content -path myfile.txt | export-aeObject -ae $ae -file c:\temp
```

##Object (XML) import
The object import works best together with XML files that have been exported with WFC. You can either import a single file or a batch of files by specifying a folder. WFC Object imports always forces the placement of objects to the destination folder. So no links will be created - instead you will always find the object at the expected location.
Import a single XML file, force it to be imported to /DEMO
```powershell
import-aeObject -ae $ae -file c:\temp\obj.xml -path /DEMO
```

Import a folder full of XML files. It is automatically determined where the objects must be imported to (same folder they have been exported from):
```powershell
import-aeObject -ae $ae -file c:\temp\
```

##Move objects
Right now this cmdlet will only allow you to move an object to another folder. No wildcards or other features are available right now.
Move object to folder (independant whether it's already there or not)
```powershell
move-aeObject -ae $ae -name MYOBJ -path /NEW/PLACE
```

##Create AE folders
To create an AE folder, simply give the folder structure you want to create. It will create all subfolders.
Create folder with subfolders:
```powershell
new-aeFolder -ae $ae -path /MY/FOLDER/AND/SUB/FOLDERS
```

# WorkflowCommanderVision (WFC::VISION)

Convert a workflow to Visio
```powershell
$ae = new-aeConnection -profile myAEConnectionProfile
get-aeVisionData -ae $ae -name JOELS_WORKFLOW | convert-aeWorkflowToVisio -file c:\temp\demo.vsd
```

Wow. That was simple. So why not having JPGs of all workflows below an AE-folder? 
```powershell
search-aeObject -ae $ae -path /FOLDER -type JOBP | get-aeVisionData -ae $ae | convert-aeWorkflowToVisio -file c:\temp\ -extension jpg
```

You can also use the Workflow Visio(n) dumps you might know from [Philipp Elmers Blog](http://philippelmer.com/automicblog/)

That's it for the beginning folks. Feel free to send feedback to joel.wiesmann at gmail.com or add me on XING.

