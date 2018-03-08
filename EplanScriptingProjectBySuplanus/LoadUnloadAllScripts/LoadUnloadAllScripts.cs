using System.IO;
using System.Windows.Forms;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.Scripting;

public class RegisterScriptMenu

{
	[Start]
	[DeclareMenu()]
	public void MenuFunction()
	{
		Eplan.EplApi.Gui.Menu aMenu = new Eplan.EplApi.Gui.Menu();
		aMenu.AddMenuItem(	"Load all scripts",
							"LoadScripts",
							"Loads all standard scripts",
							System.UInt32.Parse("35226"),   // Menu ID of Utilities / Scripts / To Run
                            System.Int32.Parse("0"), 	// ID of the following menu. (0 for default)
							false,
							true);
		Eplan.EplApi.Gui.Menu bMenu = new Eplan.EplApi.Gui.Menu();
		bMenu.AddMenuItem(	"Unload all scripts",
							"UnloadScripts",
							"Unoads all standard scripts",
							System.UInt32.Parse("35228"), 	//Menu ID of Utilities / Scripts / Unload
                            System.Int32.Parse("1"), 	// ID of the following menu. (0 for default)
							true,
							false);														
	}


	[DeclareAction("LoadScripts")]  // Action for "Load all scripts" 	//Script loads all scripts of the selected folder
    public void LoadScriptsProject()
	{
	string path = PathMap.SubstitutePath("$(MD_SCRIPTS)");
	string message = "";
	int i = 0;
	
	System.IO.FileInfo[] fi = new System.IO.DirectoryInfo(path).GetFiles("*.cs", SearchOption.AllDirectories); 
	System.IO.FileInfo[] fi1 = new System.IO.DirectoryInfo(path).GetFiles("*.vb", SearchOption.AllDirectories);	
		foreach(var fileInfo in fi)
		{
			CommandLineInterpreter oCLA = new CommandLineInterpreter();
			ActionCallingContext aca = new ActionCallingContext();
			aca.AddParameter("ScriptFile", fileInfo.FullName);
			oCLA.Execute("RegisterScript", aca);
			
			if (i > 0)
			{
			message += "\n";
			}	
			message += fileInfo.Name;
			i++;
		}
		foreach(var fileInfo in fi1)
		{
			CommandLineInterpreter oCLA = new CommandLineInterpreter();
			ActionCallingContext aca = new ActionCallingContext();
			aca.AddParameter("ScriptFile", fileInfo.FullName);
			oCLA.Execute("RegisterScript", aca);
			
			if (i > 0)
			{
			message += "\n";
			}	
			message += fileInfo.Name;
			i++;
		}
		MessageBox.Show("The following scripts were loaded successfully:" + "\n" + message, "Loaded scripts", MessageBoxButtons.OK, MessageBoxIcon.Information);
	}
	
	[DeclareAction("UnloadScripts")]    // Action for "Unload all scripts"
	public void UnloadScriptsProject()
	{	
	string path1 = PathMap.SubstitutePath("$(MD_SCRIPTS)");	
	
	System.IO.FileInfo[] fi = new System.IO.DirectoryInfo(path1).GetFiles("*.cs", SearchOption.AllDirectories); 
	System.IO.FileInfo[] fi1 = new System.IO.DirectoryInfo(path1).GetFiles("*.vb", SearchOption.AllDirectories);
	
		foreach(var fileInfo in fi)
		{
			CommandLineInterpreter oCLA = new CommandLineInterpreter();
			ActionCallingContext aca = new ActionCallingContext();
			aca.AddParameter("ScriptFile", fileInfo.FullName);
			oCLA.Execute("UnRegisterScript", aca);	
		}
		foreach(var fileInfo in fi1)
		{
			CommandLineInterpreter oCLA = new CommandLineInterpreter();
			ActionCallingContext aca = new ActionCallingContext();
			aca.AddParameter("ScriptFile", fileInfo.FullName);
			oCLA.Execute("UnRegisterScript", aca);	
		}
		MessageBox.Show("All scripts were unloaded", "Unloaded scripts", MessageBoxButtons.OK, MessageBoxIcon.Information);
	}	
}

