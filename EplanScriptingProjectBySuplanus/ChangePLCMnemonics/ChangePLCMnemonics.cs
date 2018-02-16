﻿using System.Windows.Forms;

// © by NAIROLF
// SPS-Mnemonik (E/A - I/Q) in "Adressen / Zuordnungslisten" einfach umschalten
//#############################################################################################
// ChangeLog:
// --------------------------------------------------------------------------------------------
// 2012-12-27   V1.0    NAIROLF Ersterstellung
//#############################################################################################

// Edited by
// Justin R. Walz
// 02/16/2018

// Goal:
// This will create an action that will find and replace PLC IO data.
// This can be used when going from PLC's whose IO format is "E" and "A" to "I" and "Q", or vice versa.
// I could not get this action to run. 

// Load script in Eplan using [Utilities]>[Scripts]>[Load]
// Then choose the file from the file location. 
// The file will be a .cs extension. 


using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Scripting;

public class NAIROLF_ContextMenu_ChangePLCMnemonics
{
    public static string sSourceText = string.Empty;           


    [DeclareAction("NAIROLF_ChangePLCMnemonics")]
    public void ChangePLCMnemonik(string DestMnemonik)
    {
        // Check target mnemonic
        if (DestMnemonik == string.Empty)
        {
            // No target mnemonic defined
            return;
        }

        // Empty clipboard
        System.Windows.Forms.Clipboard.Clear();

        // Fill clipboard
        CommandLineInterpreter oCLI = new CommandLineInterpreter();
        oCLI.Execute("GfDlgMgrActionIGfWind /function:Copy");

        MessageBox.Show("Test 1 completed!"); // This can be used for testing to monitor that the private sub ran correctly. 

        // Try mnemonic exchange
        #region change mnemonik

        

        if (System.Windows.Forms.Clipboard.ContainsText())
        {
            sSourceText = System.Windows.Forms.Clipboard.GetText();
            if (sSourceText != string.Empty)
            {                
                switch (DestMnemonik)
                {
                    case "IQ":
                        //sSourceText = sSourceText.Replace("PED", "PID").Replace("PAD", "PQD").Replace("PEW", "PIW").Replace("PAW", "PQW").Replace("ED", "ID").Replace("AD", "QD").Replace("EW", "IW").Replace("AW", "QW").Replace("EB", "IB").Replace("AB", "QB").Replace("E", "I").Replace("A", "Q");
                        sSourceText = sSourceText.Replace("E", "I").Replace("A", "Q");
                        break;
                    case "EA":
                        //sSourceText = sSourceText.Replace("PID", "PED").Replace("PQD", "PAD").Replace("PIW", "PEW").Replace("PQW", "PAW").Replace("ID", "ED").Replace("QD", "AD").Replace("IW", "EW").Replace("QW", "AW").Replace("IB", "EB").Replace("QB", "AB").Replace("I", "E").Replace("Q", "A");
                        sSourceText = sSourceText.Replace("I", "E").Replace("Q", "A");
                        break;
                    default:
                        return;
                }                

                try
                {
                    System.Windows.Forms.Clipboard.SetText(sSourceText);
                    oCLI.Execute("GfDlgMgrActionIGfWind /function:Paste");
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error...", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        #endregion
    }

    [DeclareMenu()]
    public void CreateContextMenus()
    {

        Eplan.EplApi.Gui.ContextMenuLocation oCTXLoc = new Eplan.EplApi.Gui.ContextMenuLocation();
        Eplan.EplApi.Gui.ContextMenu oCTXMenu = new Eplan.EplApi.Gui.ContextMenu();

        #region 1st menu-entry
        try
        {           
            oCTXLoc.DialogName = "XPlcIoDataDlg";
            oCTXLoc.ContextMenuName = "1024";
            oCTXMenu.AddMenuItem(oCTXLoc, "[SPS-Exchange mnemonics]: E /A -> I/Q", "NAIROLF_ChangePLCMnemonics /DestMnemonik:IQ", false, false);
        }
        catch (System.Exception ex)
        {
            MessageBox.Show(ex.Message, "Error...", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        #endregion

        #region 2nd menu-entry
        try
        {            
            oCTXLoc.DialogName = "XPlcIoDataDlg";
            oCTXLoc.ContextMenuName = "1024";
            oCTXMenu.AddMenuItem(oCTXLoc, "[SPS-Mnemonik tauschen]: I/Q -> E/A", "NAIROLF_ChangePLCMnemonics /DestMnemonik:EA", false, false);
        }
        catch (System.Exception ex)
        {
            MessageBox.Show(ex.Message, "Error...", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        #endregion
    }      
}