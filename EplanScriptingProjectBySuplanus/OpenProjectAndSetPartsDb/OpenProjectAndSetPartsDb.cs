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
            MessageBox.Show("Set database:\n" + DATABASE, "OpenProjectAndSetPartsDb", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        else
        {
            MessageBox.Show("Database not found:\n" + DATABASE + "\n\n There was no change to the settings.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            MessageBox.Show("Project not found:\n" + PROJECT, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        return;

    }

}