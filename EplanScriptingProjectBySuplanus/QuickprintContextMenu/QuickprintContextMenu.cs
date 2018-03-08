//
//  Dieses Skript fügt einen Menüpunkt "Seite drucken" im Kontextmenü des 
//  grafischen Editors hinzu. Die Seite wird über den vom Benutzer als
//  Standard definierten Drucker ausgegeben.
//
//  Christian Klasen  
//


//
// This script adds a menu item "Print page" in the context menu of the
// graphical editor. The page is about the user as
// Default defined printer output.
//
// Justin Walz
//


using System.Windows.Forms;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Scripting;

public class QuickprintContextMenu
{
    [DeclareRegister]
    public void Register()
    {
        MessageBox.Show("Script loaded.");

        return;
    }

    [DeclareUnregister]
    public void UnRegister()
    {
        MessageBox.Show("Unload the script.");

        return;
    }


    [DeclareAction("MenuAction")]
    public void ActionFunction()
    {
        CommandLineInterpreter oCLI = new CommandLineInterpreter();
        ActionCallingContext acc = new ActionCallingContext();

        string strPages = string.Empty;

        acc.AddParameter("TYPE", "PAGES");

        oCLI.Execute("selectionset", acc);

        acc.GetParameter("PAGES", ref strPages);
    
        acc.AddParameter("PAGENAME", strPages);

        oCLI.Execute("print", acc);

        //MessageBox.Show("Page printed");
        return;
    }

    [DeclareMenu]
    public void MenuFunction()
    {
        Eplan.EplApi.Gui.ContextMenu oMenu =
            new Eplan.EplApi.Gui.ContextMenu();

        Eplan.EplApi.Gui.ContextMenuLocation oLocation =
            new Eplan.EplApi.Gui.ContextMenuLocation(
                "Editor",
                "Ged"
                );

        oMenu.AddMenuItem(
            oLocation,
            "Print Page",
            "MenuAction",
            true,
            false
            );

        return;
    }

}




