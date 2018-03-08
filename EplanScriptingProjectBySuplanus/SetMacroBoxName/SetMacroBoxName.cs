﻿// V1:      nairolf         05.12.2010
// V2:      Johann Weiher   07.12.2010
// V2.1:    Johann Weiher   07.12.2010
// V2.2:    nairolf         10.05.2011 
// V3.0:	nairolf			16.05.2011
// V3.1:	Straight-Potter	2012
// V3.2:	FrankS			17.05.2013
// Makrokastenname aus Seitenstruktur setzen

// V1: nairolf 05.12.2010
// V2: Johann Weiher 07.12.2010
// V2.1: Johann Weiher 07.12.2010
// V2.2: nairolf 10.05.2011
// V3.0: nairolf 16.05.2011
// V3.1: Straight Potter 2012
// V3.2: FrankS 17.05.2013
// set the macro box name from the page structure

using System.Windows.Forms;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Scripting;

public class SetMacroBoxName
{

	[DeclareAction("SetMacroBoxName")]
	public bool MyMacroBoxName(string EXTENSION, string GETNAMEFROMLEVEL)
	{
        //parameter description:
        //----------------------
        //EXTENSION			...	macroname extionsion (e.g. '.ems')

        // GETNAMEFROMLEVEL ... get macro name form level code (eg 1 = Functional assignment, 2 = Attachment, 3 = Location, 4 = Location, 5 = Document type, 6 = User defined, P = Page name)
        try
        {
			string sPages = string.Empty;
			ActionCallingContext oCTX1 = new ActionCallingContext();
			CommandLineInterpreter oCLI1 = new CommandLineInterpreter();
			oCTX1.AddParameter("TYPE", "PAGES");
			oCLI1.Execute("selectionset", oCTX1);
			oCTX1.GetParameter("PAGES", ref sPages);
			string[] sarrPages = sPages.Split(';');

			if (sarrPages.Length > 1)
			{
				MessageBox.Show("More than one page marked. \nAction not possible......", "Note...");
				return false;
			}

			#region get macroname
			string sPageName = sarrPages[0];
            // ensure unique level codes:
            // Functional assignment -> $
            // Site ->%
            sPageName = sPageName.Replace("==", "$").Replace("++", "%");
			//get location from pagename
			string sMacroBoxName = string.Empty;

			//add needed / wanted structures to macroname
			#region generate macroname
			string[] sNeededLevels = GETNAMEFROMLEVEL.Split('|');
			foreach (string sLevel in sNeededLevels)
			{
				switch (sLevel)
				{
					case "1":
						if (sMacroBoxName.EndsWith(@"\"))
						{
							sMacroBoxName += ExtractLevelName(sPageName, "$");
						}
						else
						{
							sMacroBoxName += "\\" + ExtractLevelName(sPageName, "$");
						}
						break;
					case "2":
						if (sMacroBoxName.EndsWith(@"\"))
						{
							sMacroBoxName += ExtractLevelName(sPageName, "=");
						}
						else
						{
							sMacroBoxName += "\\" + ExtractLevelName(sPageName, "=");
						}
						break;
					case "3":
						if (sMacroBoxName.EndsWith(@"\"))
						{
							sMacroBoxName = sMacroBoxName + ExtractLevelName(sPageName, "%");
						}
						else
						{
							sMacroBoxName = sMacroBoxName + "\\" + ExtractLevelName(sPageName, "%");
						}
						break;
					case "4":
						if (sMacroBoxName.EndsWith(@"\"))
						{
							sMacroBoxName = sMacroBoxName + ExtractLevelName(sPageName, "+");
						}
						else
						{
							sMacroBoxName = sMacroBoxName + "\\" + ExtractLevelName(sPageName, "+");
						}
						break;
					case "5":
						if (sMacroBoxName.EndsWith(@"\"))
						{
							sMacroBoxName = sMacroBoxName + ExtractLevelName(sPageName, "&");
						}
						else
						{
							sMacroBoxName = sMacroBoxName + "\\" + ExtractLevelName(sPageName, "&");
						}
						break;
					case "6":
						if (sMacroBoxName.EndsWith(@"\"))
						{
							sMacroBoxName = sMacroBoxName + ExtractLevelName(sPageName, "#");
						}
						else
						{
							sMacroBoxName = sMacroBoxName + "\\" + ExtractLevelName(sPageName, "#");
						}
						break;
					case "P": //Page name
						if (sMacroBoxName.EndsWith(@"\"))
						{
							sMacroBoxName = sMacroBoxName + ExtractLevelName(sPageName, "/");
						}
						else
						{
							sMacroBoxName = sMacroBoxName + "\\" + ExtractLevelName(sPageName, "/");
						}
						break;
					default:
						break;
				}
			}
			#endregion

			//Clean-up macroname            
			if (sMacroBoxName.EndsWith(@"\"))
			{
				sMacroBoxName = sMacroBoxName.Remove(sMacroBoxName.Length - 1, 1);
			}
			if (sMacroBoxName.StartsWith(@"\"))
			{
				sMacroBoxName = sMacroBoxName.Substring(1);
			}

			if (sMacroBoxName == string.Empty)
			{
				MessageBox.Show("No macro name could be determined...", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}

			sMacroBoxName = sMacroBoxName.Replace(".", @"\") + EXTENSION;
			#endregion

			//set macrobox: macroname
			string quote = "\"";
			CommandLineInterpreter oCLI2 = new CommandLineInterpreter();
			oCLI2.Execute("XEsSetPropertyAction /PropertyId:23001 /PropertyIndex:0 /PropertyValue:" + quote + sMacroBoxName + quote);

			return true;
		}
		catch (System.Exception ex)
		{
			MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			return false;
		}
	}
	//========================================================================================================================================		
	private string ExtractLevelName(string sPage, string sLevel)
	{
		string sLevelName = string.Empty;

		if (sPage.Contains(sLevel))
		{
            //check existing level codes (remove all text of following level code)
            #region Functional assignment

            if (sLevel == "$")
			{
				if (sPage.Contains("="))
				{
					sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("=") - sPage.IndexOf(sLevel));
					goto DONE;
				}
				if (sPage.Contains("%"))
				{
					sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("%") - sPage.IndexOf(sLevel));
					goto DONE;
				}
				if (sPage.Contains("+"))
				{
					sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("+") - sPage.IndexOf(sLevel));
					goto DONE;
				}
				if (sPage.Contains("&"))
				{
					//check if document type is at beginning
					if (sPage.StartsWith("&"))
					{
						//no extracting needed, when document type at beginning						
					}
					else
					{
						sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("&") - sPage.IndexOf(sLevel));
						goto DONE;
					}
				}
				if (sPage.Contains("#"))
				{
					//check if user-defined is at beginning
					if (sPage.StartsWith("#"))
					{
						//no extracting needed, when user-defined at beginning
					}
					else
					{
						sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("#") - sPage.IndexOf(sLevel));
						goto DONE;
					}
				}
				//no further structure identifier existing
				sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("/") - sPage.IndexOf(sLevel));
			}
            #endregion

            #region Investment
            if (sLevel == "=")
			{
				if (sPage.Contains("%"))
				{
					sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("%") - sPage.IndexOf(sLevel));
					goto DONE;
				}
				if (sPage.Contains("+"))
				{
					sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("+") - sPage.IndexOf(sLevel));
					goto DONE;
				}
				if (sPage.Contains("&"))
				{
					//check if document type is at beginning
					if (sPage.StartsWith("&"))
					{
						//no extracting needed, when document type at beginning						
					}
					else
					{
						sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("&") - sPage.IndexOf(sLevel));
						goto DONE;
					}
				}
				if (sPage.Contains("#"))
				{
					//check if user-defined is at beginning
					if (sPage.StartsWith("#"))
					{
						//no extracting needed, when user-defined at beginning
					}
					else
					{
						sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("#") - sPage.IndexOf(sLevel));
						goto DONE;
					}
				}

				//no further structure identifier existing
				sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("/") - sPage.IndexOf(sLevel));
			}
            #endregion

            #region Site
            if (sLevel == "%")
			{
				if (sPage.Contains("+"))
				{
					sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("+") - sPage.IndexOf(sLevel));
					goto DONE;
				}
				if (sPage.Contains("&"))
				{
					//check if document type is at beginning
					if (sPage.StartsWith("&"))
					{
						//no extracting needed, when document type at beginning						
					}
					else
					{
						sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("&") - sPage.IndexOf(sLevel));
						goto DONE;
					}
				}
				if (sPage.Contains("#"))
				{
					//check if user-defined is at beginning
					if (sPage.StartsWith("#"))
					{
						//no extracting needed, when user-defined at beginning
					}
					else
					{
						sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("#") - sPage.IndexOf(sLevel));
						goto DONE;
					}
				}

				//no further structure identifier existing
				sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("/") - sPage.IndexOf(sLevel));
			}
            #endregion

            #region Installation
            if (sLevel == "+")
			{
				if (sPage.Contains("&"))
				{
					//check if document type is at beginning
					if (sPage.StartsWith("&"))
					{
						//no extracting needed, when document type at beginning						
					}
					else
					{
						sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("&") - sPage.IndexOf(sLevel));
						goto DONE;
					}
				}

				if (sPage.Contains("#"))
				{
					//check if user-defined is at beginning
					if (sPage.StartsWith("#"))
					{
						//no extracting needed, when user-defined at beginning
					}
					else
					{
						sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("#") - sPage.IndexOf(sLevel));
						goto DONE;
					}
				}

				//no further structure identifier existing
				sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("/") - sPage.IndexOf(sLevel));
			}
            #endregion

            #region Document Type

            if (sLevel == "&")
			{
				//check if document type is at beginning
				if (sPage.StartsWith("&"))
				{
					//check further existing structures						
					if (sPage.Contains("$"))
					{
						sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("$") - sPage.IndexOf(sLevel));
						goto DONE;
					}
					if (sPage.Contains("="))
					{
						sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("=") - sPage.IndexOf(sLevel));
						goto DONE;
					}
					if (sPage.Contains("%"))
					{
						sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("%") - sPage.IndexOf(sLevel));
						goto DONE;
					}
					if (sPage.Contains("+"))
					{
						sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("+") - sPage.IndexOf(sLevel));
						goto DONE;
					}
					if (sPage.Contains("#"))
					{
						sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("#") - sPage.IndexOf(sLevel));
						goto DONE;
					}

					//no further structure identifier existing
					sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("/") - sPage.IndexOf(sLevel));
					goto DONE;
				}
				else
				{
					//document type NOT at beginning
					if (sPage.Contains("#"))
					{
						if (sPage.StartsWith("#"))
						{
							//no extracting needed, when user-defined at beginning
						}
						else
						{
							sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("#") - sPage.IndexOf(sLevel));
							goto DONE;
						}
					}

					//no further structure identifier existing
					sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("/") - sPage.IndexOf(sLevel));
				}

			}
            #endregion

            #region Custom
            if (sLevel == "#")
			{
				//check if user defined is at beginning
				if (sPage.StartsWith("#"))
				{
					//check further existing structures						
					if (sPage.Contains("$"))
					{
						sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("$") - sPage.IndexOf(sLevel));
						goto DONE;
					}
					if (sPage.Contains("="))
					{
						sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("=") - sPage.IndexOf(sLevel));
						goto DONE;
					}
					if (sPage.Contains("%"))
					{
						sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("%") - sPage.IndexOf(sLevel));
						goto DONE;
					}
					if (sPage.Contains("+"))
					{
						sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("+") - sPage.IndexOf(sLevel));
						goto DONE;
					}
					if (sPage.Contains("&"))
					{
						sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("&") - sPage.IndexOf(sLevel));
						goto DONE;
					}

					//no further structure identifier existing
					sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("/") - sPage.IndexOf(sLevel));
					goto DONE;
				}
				else
				{
					//document typ NOT at beginning
					if (sPage.Contains("#"))
					{
						sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("/") - sPage.IndexOf(sLevel));
						goto DONE;
					}
				}
			}
            #endregion

            #region Pagename

            if (sLevel == "/")
			{
				//nur den Seitenname
				sLevelName = sPage.Substring(sPage.IndexOf("/"), sPage.Length - sPage.IndexOf("/"));
				goto DONE;
			}
			#endregion
		}
	DONE:
		//remove page no		
		if (sLevelName == string.Empty && sPage.Contains(sLevel))
		{
			sLevelName = sPage.Substring(sPage.IndexOf(sLevel), sPage.IndexOf("/") - sPage.IndexOf(sLevel));
		}

		if (sLevelName != string.Empty)
		{
			sLevelName = sLevelName.Replace(sLevel, string.Empty);
		}


		return sLevelName;
	}
}