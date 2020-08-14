# RegAsmUtilityHelper
Helper working with COM utilities and VB.  Updating the Utility and retesting the assemblies so that they can be access from the VB code can be a huge pain so I wrote the Utility that automates a lot of the process and make it easier to iterate while doing development or testing.

## Default Paths

* **esignUtilityPath** - `~\source\repos\Document E-Signing\ESign.Utility\bin\Debug`

## Overrides 

* **esignUtilityPath** - full path to Esign.Utility.dll. It is important to pass the full path and not use relative pathing here *i.e. `C:Users\Connor\source\repos\Document E-Signing\ESign.Utility\bin\Debug`
* **verbose** - prevents the script from clearing the console after execution

## After Execution
This script only loads the assemblies into memory, it does not update them in the VB context.  Once Excel has opened follow the following steps.
1. Click on the `Developer` tab in excel and select `Visual Basic`
2. This should open the VB editor.  Select `Tools` and `References`
3. You should now see a window with all of the loaded assemblies.  If you have run this before you will have `Json.NET....` and `Esign_Utility`. If you would like update one of the existing utilities unselect it from the list.
4. Click `Browse...` and navigate to your Esign.Utility build directory and select the `.tlb` assembly file.
5. Once you have loaded all the assemblies that you need click `Ok` and you are ready to execute the `VB` code.

## Other Notes
* This script needs to run as an admin to update the assemblies in memory.  
* Also note that this script will close out of Excel and you should make sure you save any changes before running it. 
* Newtonsoft needs to be loaded for the Utility to work.  If you have not loaded it before you need to execute the newtonsoft option which will load the appropriate assemblies.  Then be sure to load it in the VB UI
