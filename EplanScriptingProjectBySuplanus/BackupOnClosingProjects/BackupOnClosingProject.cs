using System.Windows.Forms;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.Scripting;

// Goal:
// This will create event handler in which every time the project is closed the user will be asked if a backup of project should be created

// Load script in Eplan using [Utilities]>[Scripts]>[Load]
// Then choose the file from the file location. 
// The file will be a .cs extension. 

public class BackupOnClosingProject
{
	[DeclareEventHandler("onActionStart.String.XPrjActionProjectClose")]
	public void Function()
	{
		string strProjectname = PathMap.SubstitutePath("$(PROJECTNAME)");
		string strFullProjectname = PathMap.SubstitutePath("$(P)");
		string strDestination = strFullProjectname;
		
		DialogResult Result = MessageBox.Show(
            "Should a backup for the project\n'"
			+ strProjectname +
            "'\nbe generated? ",
            "Data backup",
			MessageBoxButtons.YesNo,
			MessageBoxIcon.Question
			);

		if (Result == DialogResult.Yes)

		  {

				string myTime = System.DateTime.Now.ToString("yyyy.MM.dd");
				string hour = System.DateTime.Now.Hour.ToString();
				string minute = System.DateTime.Now.Minute.ToString();

				Progress progress = new Progress("SimpleProgress");
				progress.BeginPart(100, "");
				progress.SetAllowCancel(true);

				if (!progress.Canceled())
				{
					progress.BeginPart(33, "Backup");
					ActionCallingContext backupContext = new ActionCallingContext();
					backupContext.AddParameter("BACKUPMEDIA", "DISK");
					backupContext.AddParameter("BACKUPMETHOD", "BACKUP");
					backupContext.AddParameter("COMPRESSPRJ", "0");
					backupContext.AddParameter("INCLEXTDOCS", "1");
					backupContext.AddParameter("BACKUPAMOUNT", "BACKUPAMOUNT_ALL");
					backupContext.AddParameter("INCLIMAGES", "1");
					backupContext.AddParameter("LogMsgActionDone", "true");
					backupContext.AddParameter("DESTINATIONPATH", strDestination);
					backupContext.AddParameter("PROJECTNAME", strFullProjectname);
					backupContext.AddParameter("TYPE", "PROJECT");
					backupContext.AddParameter("ARCHIVENAME", strProjectname + "_" + myTime + "_" + hour + "." + minute + ".");
					new CommandLineInterpreter().Execute("backup", backupContext);
					progress.EndPart();
					
				}
				progress.EndPart(true);
			}
			
			return;
		}
	}