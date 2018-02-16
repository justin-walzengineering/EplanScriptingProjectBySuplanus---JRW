using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Scripting;
using System.Windows.Forms;

// Goal:
// This will create event handlers for updating Plugs
// Connection data is not immediately automatically updated after changes.
// When the user runs the print command or PDF export command the GeneratePlugs sub is ran, which updates all connection types. 

// Load script in Eplan using [Utilities]>[Scripts]>[Load]
// Then choose the file from the file location. 
// The file will be a .cs extension. 

public class AutomaticGeneratingPlugs
{
    [DeclareEventHandler("onActionStart.String.PrnPrintDialogShow")]
    public void ehPrint()
    {
        GeneratePlugs();
    }

    [DeclareEventHandler("onActionStart.String.XPdfExportAction")]
    public void ehPdf()
    {
        GeneratePlugs();
    }

    private static void GeneratePlugs()
    {
        ActionCallingContext acc = new ActionCallingContext();
        acc.AddParameter("TYPE", "Plugs");
        new CommandLineInterpreter().Execute("generate",acc);
//        MessageBox.Show("Connection update completed!"); // This can be used for testing to monitor that the private sub ran correctly. 
    }
}