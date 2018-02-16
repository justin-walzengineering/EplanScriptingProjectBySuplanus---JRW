using System.IO;
using System.Windows.Forms;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Scripting;

public class OpenProjectAndSetPartsDb
{
    [DeclareAction("OpenProjectAndSetPartsDb")]
    public void OpenProjectAndSetPartsDbVoid(string PROJECT,string DATABASE)
    {
        if (File.Exists(DATABASE))
        {
            Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();
            oSettings.SetStringSetting("USER.PartsManagementGui.Database", DATABASE, 0);
            MessageBox.Show("Eingestellte Datenbank:\n" + DATABASE, "OpenProjectAndSetPartsDb", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        else
        {
            MessageBox.Show("Datenbank nicht gefanden:\n" + DATABASE + "\n\n Es wurde keine Änderung an den Einstellungen vorgenommen.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        if (File.Exists(PROJECT))
        {
            ActionCallingContext accProjectOpen = new ActionCallingContext();
            accProjectOpen.AddParameter("Project", PROJECT);
            new CommandLineInterpreter().Execute("ProjectOpen", accProjectOpen);
        }
        else
        {
            MessageBox.Show("Projekt nicht gefanden:\n" + PROJECT, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        return;

    }

}