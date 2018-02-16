// ConnectionPointDesignationReverse.cs
//
// Erweitert das Kontextmenü von 'ConnectionTerms',
// im Dialog 'Eigenschaften (Schaltzeichen): Allgemeines Betriebsmittel',
// um den Menüpunkt 'Reihenfolge drehen'.
// Es wird die Eingabe im Feld 'ConnectionTerms' automatisch gedreht.
//
// Copyright by Frank Schöneck, 2015
// letzte Änderung:
// V1.0.0, 04.03.2015, Frank Schöneck, Projektbeginn
//
// für Eplan Electric P8, ab V2.3


// ConnectionPointDesignationReverse.cs
//
// extends the context menu of 'port names',
// in the dialog 'Properties (schematic): General Resources',
// to the menu item 'turn order'.
// The input in the field 'Connection designations' is automatically rotated.
//
// Copyright by Frank Schöneck, 2015
// last change:
// V1.0.1, 02/16/2018, Justin Walz, converted to english, added comments
//
// for Eplan Electric P8, from V2.3

using System;
using System.Windows.Forms;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.Scripting;

public class ConnectionPointDesignationReverse
{

	[DeclareMenu]
	public void ProjectCopyContextMenu()
	{
		// Context - Menu item
		string menuText = getMenuText();
		Eplan.EplApi.Gui.ContextMenu oContextMenu = new Eplan.EplApi.Gui.ContextMenu();
		Eplan.EplApi.Gui.ContextMenuLocation oContextMenuLocation = new Eplan.EplApi.Gui.ContextMenuLocation("XDTDataDialog", "4006");
		oContextMenu.AddMenuItem(oContextMenuLocation, menuText, "ConnectionPointDesignationReverse", true, false);
	}

	[DeclareAction("ConnectionPointDesignationReverse")]
	public void Action()
	{
		try
		{
			string sSourceText = string.Empty;
			string sReturnText = string.Empty;
			string EplanCRLF = "¶";

            // Empty clipboard
            System.Windows.Forms.Clipboard.Clear();

            // Fill clipboard

            CommandLineInterpreter oCLI = new CommandLineInterpreter();
			oCLI.Execute("GfDlgMgrActionIGfWind /function:SelectAll"); // Select all

            oCLI.Execute("GfDlgMgrActionIGfWind /function:Copy"); // Copy

            if (System.Windows.Forms.Clipboard.ContainsText())
			{
				sSourceText = System.Windows.Forms.Clipboard.GetText();
				if (sSourceText != string.Empty)
				{
					string[] sConnectionTerms = sSourceText.Split(new string[] { EplanCRLF }, StringSplitOptions.None);

					if (sConnectionTerms.Length > 2) // More than 2 connection names
                    {
						Decider eDecision = new Decider();
						EnumDecisionReturn eAnswer = eDecision.Decide(EnumDecisionType.eYesNoDecision,
                            "Should the connection designations be turned in pairs? ",
                            "Turn the order",
							EnumDecisionReturn.eYES,
							EnumDecisionReturn.eYES,
							"ConnectionPointDesignationReverse",
							true,
							EnumDecisionIcon.eQUESTION);

						if (eAnswer == EnumDecisionReturn.eYES)
						{
                            // Rebuild string

                            for (int i = 0; i < sConnectionTerms.Length; i = i + 2)
							{
								sReturnText += sConnectionTerms[i + 1] + EplanCRLF + sConnectionTerms[i] + EplanCRLF;
							}
						}
						else
						{
                            // Turn the string array

                            Array.Reverse(sConnectionTerms);


                            // Rebuild string

                            foreach (string sConnection in sConnectionTerms)
							{
								sReturnText += sConnection + EplanCRLF;
							}
						}
					}
                    else // Only 2 connection names
                    {
                        // Turn the string array

                        Array.Reverse(sConnectionTerms);

                        // Rebuild string
                        foreach (string sConnection in sConnectionTerms)
						{
							sReturnText += sConnection + EplanCRLF;
						}
					}

                    // Remove last character again

                    sReturnText = sReturnText.Substring(0, sReturnText.Length - 1);

                    // Insert clipboard

                    System.Windows.Forms.Clipboard.SetText(sReturnText);
					oCLI.Execute("GfDlgMgrActionIGfWind /function:SelectAll"); // Select all
                    oCLI.Execute("GfDlgMgrActionIGfWindDelete"); // Clear
					oCLI.Execute("GfDlgMgrActionIGfWind /function:Paste"); // Insert
                }
			}
		}
		catch (System.Exception ex)
		{
			MessageBox.Show(ex.Message, "Turn order, error", MessageBoxButtons.OK, MessageBoxIcon.Error);
		}
		return;
	}

	// Returns the menu item text in the GUI language if available.
	private string getMenuText()
	{
		MultiLangString muLangMenuText = new MultiLangString();
		muLangMenuText.SetAsString(
			"en_US@rotate order;"
			);

		ISOCode guiLanguage = new Languages().GuiLanguage;
		return muLangMenuText.GetString((ISOCode.Language)guiLanguage.GetNumber());
	}

}
