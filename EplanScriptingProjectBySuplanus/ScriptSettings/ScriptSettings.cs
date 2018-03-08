/* ================================================================================================
 * ScriptSettings
 * ================================================================================================
 * ibKastl GmbH
 * Dr.-Seitz-Str. 13a
 * 82418 Murnau
 * 08841 6286833
 * info@ibkastl.de
 * http://www.ibkastl.de
 * ================================================================================================
 * Version
 * 1.2
 * - Toolbar Support
 * ================================================================================================
*/
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.Scripting;

namespace ibKastl.Scripts.Global.ScriptSettings
{
  class ScriptSettings
  {
    readonly string SettingsFile = Path.Combine(PathMap.SubstitutePath("$(MD_SCRIPTS)"),
           "ScriptSettings", "SettingsFile.xml");

    private readonly string ScriptSettingsApplication = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
        @"ibKastl\ScriptSettings\ScriptSettings.exe");

    [DeclareAction("ScriptSettings")]
    public void Action(string name, string parameter, out string setting)
    {
      // Run ScriptSettings
      if (RunScriptSettings(name, parameter, out setting)) return;

      // Check
      if (!File.Exists(SettingsFile))
      {
        MessageBox.Show("Datei für Script-Einstellungen nicht gefunden." +
            Environment.NewLine + SettingsFile +
            Environment.NewLine + Environment.NewLine + "Der Vorgang wurde abgebrochen.",
            "Warnung - ScriptSettings", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        setting = null;
        return;
      }
      try
      {
        bool isOk = !string.IsNullOrEmpty(name) && string.IsNullOrEmpty(parameter);

        // Get value
        XmlDocument xmlDocument = new XmlDocument();
        xmlDocument.Load(SettingsFile);
        string url;

        // CompanyName
        if (parameter != null && parameter.Equals("COMPANYNAME"))
        {
          url = "/DataContext/CompanyName";
        }
        // OK
        else if (isOk)
        {
          url = string.Format("/DataContext/Scripts/ScriptObject[@Name='{0}']", name);
          XmlNodeList rankListSchemaName = xmlDocument.SelectNodes(url);
          if (rankListSchemaName != null && rankListSchemaName.Count > 0)
          {
            setting = true.ToString();
            return;
          }
          else
          {
            setting = null;
            return;
          }
        }
        // Parameter
        else
        {
          url = string.Format(
                        "/DataContext/Scripts/ScriptObject[@Name='{0}']/Settings/Setting[@Parameter='{1}']/Value",
                        name, parameter);
        }

        setting = GetValueFromXml(xmlDocument, url);

        if (string.IsNullOrEmpty(setting))
        {
          MessageBox.Show("Einstellungen nicht gefunden." + Environment.NewLine +
              "Scriptname: " + name + Environment.NewLine +
              "Parameter: " + parameter + Environment.NewLine +
              "Der Vorgang wurde abgebrochen.",
              "Warnung - ScriptSettings", MessageBoxButtons.OK, MessageBoxIcon.Warning);
          setting = null;
        }
      }
      catch (Exception exception)
      {
        MessageBox.Show(exception.Message, "Fehler - ScriptSettings", MessageBoxButtons.OK,
            MessageBoxIcon.Error);
        setting = null;
        return;
      }



    }

    private bool RunScriptSettings(string name, string parameter, out string setting)
    {
      try
      {
        if (name.ToUpper().Equals("NULL") && parameter.ToUpper().Equals("NULL"))
        {
          if (File.Exists(ScriptSettingsApplication))
          {
            string quote = "\"";
            if (ScriptSettingsApplication != null)
            {
              Process.Start(new ProcessStartInfo
              {
                Arguments = "/file:" + quote + SettingsFile + quote,
                // ReSharper disable once AssignNullToNotNullAttribute
                WorkingDirectory = Path.GetDirectoryName(ScriptSettingsApplication),
                FileName = Path.GetFileName(ScriptSettingsApplication)
              });
            }
            setting = null;
            return true;
          }
          else
          {
            MessageBox.Show("Programm 'ScriptSettings' nicht gefunden." +
                            Environment.NewLine +
                            ScriptSettingsApplication +
                            Environment.NewLine +
                            "Der Vorgang wurde abgebrochen.",
                "Warnung - ScriptSettings", MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            setting = null;
            return true;
          }
        }
        else
        {
          setting = null;
          return false;
        }
      }
      catch (Exception exception)
      {
        MessageBox.Show(exception.Message, "Fehler", MessageBoxButtons.OK,
            MessageBoxIcon.Error);
        setting = null;
        return true;
      }
    }

    private static string GetValueFromXml(XmlDocument xmlDocument, string url)
    {
      XmlNodeList rankListSchemaName = xmlDocument.SelectNodes(url);
      if (rankListSchemaName != null && rankListSchemaName.Count > 0)
      {
        return rankListSchemaName[0].InnerText;
      }
      else
      {
        return null;
      }
    }
  }


  public class ToolbarCreator
  {
    private bool _createSeparateToolBar;

    private string _toolbarName;
    private Eplan.EplApi.Gui.Toolbar _toolbar = new Eplan.EplApi.Gui.Toolbar();
    private List<IGrouping<string, ScriptObject>> _groups;

    [DeclareAction("ToolbarCreator")]
    public void Function(string type)
    {
      // Read
      string filename = Path.Combine(PathMap.SubstitutePath("$(MD_SCRIPTS)"),
        "ScriptSettings", "SettingsFile.xml");

      XmlSerializer serializer = new XmlSerializer(typeof(ScriptSettingsXmlObject),
        new XmlRootAttribute("DataContext"));
      StreamReader reader = new StreamReader(filename);

      var scriptSettingsXmlObject = (ScriptSettingsXmlObject)serializer.Deserialize(reader);
      _groups = scriptSettingsXmlObject.Scripts
        .OrderBy(obj => obj.Scriptname)
        .GroupBy(obj => obj.Scriptname)
        .ToList();

      // Parameter
      string parameter = type;
      string parameterUpper = parameter.ToUpper();
      switch (parameterUpper)
      {
        case "ONETOOLBAR": break;

        case "SEPARATE":
          _createSeparateToolBar = true;
          break;

        default:
          _createSeparateToolBar = true;

          _groups = _groups
            .Where(obj => obj.Key.ToUpper().Equals(parameterUpper))
            .ToList();

          if (_groups.Count == 0)
          {
            MessageBox.Show("Keine Scripte mit angegebenen Namen gefunden: " + parameter);
            return;
          }
          break;
      }

      // Remove
      if (!_createSeparateToolBar)
      {
        _toolbarName = "ShopForProcess_Generated-Toolbar";
        RemoveToolbar();
        CreateToolbar();
      }

      // Create      
      AddButtonScriptSettings();
      AddButtonsScripts();

      return;
    }



    private void AddButtonsScripts()
    {
      int scriptCount = 0;


      Progress progress = new Progress("SimpleProgress");
      progress.SetAllowCancel(true);
      progress.SetAskOnCancel(false);
      progress.SetNeededSteps(_groups.SelectMany(obj => obj).Count() + 1);
      progress.SetTitle("Meine Progressbar");
      progress.ShowImmediately();


      try
      {
        foreach (var group in _groups)
        {
          scriptCount++;
          string scriptName = group.Key;

          if (progress.Canceled())
          {
            progress.EndPart(true);
            return;
          }

          if (_createSeparateToolBar)
          {
            _toolbarName = "ShopForProcess_" + scriptName;
            RemoveToolbar();
            CreateToolbar();
          }

          // Create buttons
          var scriptVariants = group.OrderBy(obj => obj.Name);
          foreach (var scriptObject in scriptVariants)
          {
            progress.SetActionText(scriptName + ": " + scriptObject.Name);
            progress.Step(1);

            // Standard
            if (!scriptObject.CustomActions.Any())
            {
              string action = GetCommandline(scriptName, scriptObject.Name);
              string imagePath = Path.Combine("$(MD_SCRIPTS)", scriptName, "Images", scriptObject.Scriptname + ".bmp");
              string toolTip = scriptName + " " + scriptObject.Name;
              _toolbar.AddButton(_toolbarName, 0, action, imagePath, toolTip);
            }
            // Custom action
            else
            {
              foreach (var customAction in scriptObject.CustomActions)
              {
                string action = customAction.Split(' ')[0];
                string imagePath =
                  Path.Combine("$(MD_SCRIPTS)", scriptName, "Images", action + ".bmp");
                string toolTip = scriptName + " " + scriptObject.Name + " " + action;
                _toolbar.AddButton(_toolbarName, 0, customAction, imagePath, toolTip);
              }
            }
          }

          // Separator
          if (!_createSeparateToolBar)
          {
            if (!_groups.Last().Equals(group))
            {
              AddSeparator();
            }
          }

          //#if DEBUG
          //          if (scriptCount == 3)
          //          {
          //            return;
          //          }
          //#endif
        }
      }
      finally
      {
        progress.EndPart(true);
      }
      return;
    }

    private void AddButtonScriptSettings()
    {
      if (_createSeparateToolBar)
      {
        _toolbarName = "ShopForProcess_ScriptSettings";
        RemoveToolbar();
        CreateToolbar();
      }

      string action = "ScriptSettings /name:NULL /parameter:NULL";
      string imagePath = Path.Combine("$(MD_SCRIPTS)", "ScriptSettings", "Images", "ScriptSettings" + ".bmp");
      string toolTip = "ScriptSettings";
      _toolbar.AddButton(_toolbarName, 0, action, imagePath, toolTip);

      if (!_createSeparateToolBar)
      {
        if (_groups.Count != 0)
        {
          AddSeparator();
        }
      }
    }

    private void CreateToolbar()
    {
      _toolbar.CreateCustomToolbar(_toolbarName, Eplan.EplApi.Gui.Toolbar.ToolBarDockPos.eToolbarFloat, 0, 0, true);
    }

    private void RemoveToolbar()
    {
      bool isUserToolbar = false;
      if (_toolbar.ExistsToolbar(_toolbarName, ref isUserToolbar))
      {
        _toolbar.RemoveCustomToolbar(_toolbarName);
      }
    }

    private void AddSeparator()
    {
      _toolbar.AddButton(_toolbarName, 1, 0); // Separator
    }

    private string GetCommandline(string scriptname, string name)
    {
      const string QUOTE = "\"";
      string commandline = string.Format(
        "{0} /Name:" + QUOTE + "{1}" + QUOTE, scriptname, name);
      return commandline;

    }

    [Serializable]
    public class ScriptSettingsXmlObject
    {
      public ScriptSettingsXmlObject()
      {
      }

      public ScriptSettingsXmlObject(List<ScriptObject> scripts, string companyName, List<Menu> menus)
      {
        Scripts = scripts;
        CompanyName = companyName;
        Menus = menus;
      }

      public string CompanyName { get; set; }
      public string Language { get; set; }
      public string Version { get; set; }

      public List<ScriptObject> Scripts { get; set; }

      public List<Menu> Menus { get; set; }


      public struct Menu
      {
        [XmlAttribute(DataType = "string", AttributeName = "Name")]
        public string Name { get; set; }
      }
    }

    public class ScriptObject
    {
      public ScriptObject()
      {
        IsAvailableInMenu = true;
        Documentations = new List<Documentation>();
      }

      public ScriptObject(string scriptname, string name, string serial, bool isAvailableInMenu = true)
      {
        Scriptname = scriptname;
        Name = name;
        Serial = serial;
        Settings = new ObservableCollection<Setting>();
        IsAvailableInMenu = isAvailableInMenu;
      }

      [XmlIgnore]
      public List<Documentation> Documentations { get; set; }

      [XmlAttribute(DataType = "string", AttributeName = "Scriptname")]
      public string Scriptname { get; set; }

      [XmlAttribute(DataType = "string", AttributeName = "Name")]
      public string Name { get; set; }

      [XmlIgnore]
      public bool IsAvailableInMenu { get; set; }

      [XmlAttribute(DataType = "string", AttributeName = "Serial")]
      public string Serial { get; set; }

      [XmlArray("Settings")]
      [XmlArrayItem("Setting")]
      public ObservableCollection<Setting> Settings { get; set; }

      [XmlArray("CustomActions")]
      [XmlArrayItem("CustomActions")]
      public List<string> CustomActions { get; set; }

    }

    [XmlRoot("Setting")]
    [XmlInclude(typeof(string))]
    public class Setting
    {
      public Setting()
      {

      }

      public string Value { get; set; }

      [XmlIgnore]
      public string Default { get; set; }

      [XmlAttribute(DataType = "string", AttributeName = "Parameter")]
      public string Parameter { get; set; }

      [XmlIgnore]
      public string Description { get; set; }

      [XmlIgnore]
      public string Tooltip { get; set; }

      [XmlAttribute("Type")]
      public SettingType Type { get; set; }

      public object Clone()
      {
        return MemberwiseClone();
      }
    }

    public enum SettingType
    {
      Bool,
      Directory,
      List,
      Path,
      Text,
      PathFile
    }

    public class Documentation
    {
      public string Language { get; set; }
      public string Url { get; set; }
    }
  }


  class GetProjectProperty
  {
    /* Usage
	private static string GetProjectProperty(string id, string index)
	{
		string value = null;
		ActionCallingContext actionCallingContext = new ActionCallingContext();
		actionCallingContext.AddParameter("id", id);
		actionCallingContext.AddParameter("index", index);
		new CommandLineInterpreter().Execute("GetProjectProperty", actionCallingContext);
		actionCallingContext.GetParameter("value", ref value);
		return value;
	}
    */
    [DeclareAction("GetProjectProperty")]
    public void Action(string id, string index, out string value)
    {
      string pathTemplate = Path.Combine(PathMap.SubstitutePath("$(MD_SCRIPTS)"),
          "ScriptSettings", "Library", "GetProjectProperty", "GetProjectProperty_Template.xml");
      string pathScheme = Path.Combine(PathMap.SubstitutePath("$(MD_SCRIPTS)"),
          "ScriptSettings", "Library", "GetProjectProperty", "GetProjectProperty_Scheme.xml");
      string pathOutput = Path.Combine(PathMap.SubstitutePath("$(MD_SCRIPTS)"),
          "ScriptSettings", "Library", "GetProjectProperty", "GetProjectProperty_Output.txt");

      try
      {
        // Set scheme
        string content = File.ReadAllText(pathTemplate);
        content = content.Replace("GetProjectProperty_ID", id);
        content = content.Replace("GetProjectProperty_INDEX", index);
        File.WriteAllText(pathScheme, content);
        new Settings().ReadSettings(pathScheme);

        // Export
        ActionCallingContext actionCallingContext = new ActionCallingContext();
        actionCallingContext.AddParameter("configscheme", "GetProjectProperty");
        actionCallingContext.AddParameter("destinationfile", pathOutput);
        actionCallingContext.AddParameter("language", "de_DE");
        new CommandLineInterpreter().Execute("label", actionCallingContext);

        // Read
        value = File.ReadAllLines(pathOutput)[0];
      }
      catch (Exception exception)
      {
        MessageBox.Show(exception.Message, "GetProjectProperty", MessageBoxButtons.OK,
            MessageBoxIcon.Error);
        value = "[Error]";
      }

    }
  }

  class GetCurrentScriptPath
  {
    /* Usage
    private static string GetCurrentScriptPath(string scriptName)
    {
        string value = null;
        ActionCallingContext actionCallingContext = new ActionCallingContext();
        actionCallingContext.AddParameter("scriptName", scriptName);
        new CommandLineInterpreter().Execute("GetCurrentScriptPath", actionCallingContext);
        actionCallingContext.GetParameter("value", ref value);
        return value;
    }
    */
    [DeclareAction("GetCurrentScriptPath")]
    public void Action(string scriptName, out string value)
    {
      Settings settings = new Settings();
      var settingsUrlScripts = "STATION.EplanEplApiScriptGui.Scripts";
      int countOfScripts = settings.GetCountOfValues(settingsUrlScripts);
      for (int i = 0; i < countOfScripts; i++)
      {
        string scriptPath = settings.GetStringSetting(settingsUrlScripts, i);
        if (scriptPath.EndsWith(@"\" + scriptName))
        {
          // found
          value = scriptPath;
          return;
        }
      }

      // not found
      value = null;
      return;
    }
  }

  class GetCurrentLoadedScripts
  {
    /* Usage
    private static string GetCurrentLoadedScripts()
    {
        string value = null;
        ActionCallingContext actionCallingContext = new ActionCallingContext();
        new CommandLineInterpreter().Execute("GetCurrentLoadedScripts", actionCallingContext);
        actionCallingContext.GetParameter("value", ref value);
        return value;
    }
    */
    [DeclareAction("GetCurrentLoadedScripts")]
    public void Action(out string value)
    {
      Settings settings = new Settings();
      var settingsUrlScripts = "STATION.EplanEplApiScriptGui.Scripts";
      int countOfScripts = settings.GetCountOfValues(settingsUrlScripts);
      StringBuilder stringBuilder = new StringBuilder();
      for (int i = 0; i < countOfScripts; i++)
      {
        string scriptPath = settings.GetStringSetting(settingsUrlScripts, i);
        stringBuilder.Append(scriptPath);

        // not last one
        if (i != countOfScripts - 1)
        {
          stringBuilder.Append("|");
        }
      }

      // returns list: "\\path\script1.cs|\\path\script2.vb"
      value = stringBuilder.ToString();
    }
  }

  class GetProjectLanguages
  {
    /* Usage
    private static string GetProjectLanguages()
    {
        string value = null;
        ActionCallingContext actionCallingContext = new ActionCallingContext();
        new CommandLineInterpreter().Execute("GetProjectLanguages", actionCallingContext);
        actionCallingContext.GetParameter("value", ref value);
        return value;
    }
    */
    private readonly string TempPath = Path.Combine(
        PathMap.SubstitutePath("$(TMP)"), "GetProjectLanguages.xml");

    [DeclareAction("GetProjectLanguages")]
    public void Action(out string value)
    {
      ActionCallingContext actionCallingContext = new ActionCallingContext();
      actionCallingContext.AddParameter("prj", FullProjectPath());
      actionCallingContext.AddParameter("node", "TRANSLATEGUI");
      actionCallingContext.AddParameter("XMLFile", TempPath);
      new CommandLineInterpreter().Execute("XSettingsExport", actionCallingContext);

      if (File.Exists(TempPath))
      {
        string languagesString = GetValueSettingsXml(TempPath,
            "/Settings/CAT/MOD/Setting[@name='TRANSLATE_LANGUAGES']/Val");

        if (languagesString != null)
        {
          List<string> languages = languagesString.Split(';').ToList();
          languages = languages.Where(obj => !obj.Equals("")).ToList();

          StringBuilder stringBuilder = new StringBuilder();
          for (int i = 0; i < languages.Count; i++)
          {
            var language = languages[i];
            stringBuilder.Append(language);

            // not last one
            if (i != languages.Count - 1)
            {
              stringBuilder.Append("|");
            }
          }

          // returns list: "de_DE|en_EN"
          value = stringBuilder.ToString();
          return;
        }
      }

      value = null;
      return;
    }

    // Returns the EPLAN Project Path
    private static string FullProjectPath()
    {
      ActionCallingContext acc = new ActionCallingContext();
      acc.AddParameter("TYPE", "PROJECT");

      string projectPath = string.Empty;
      new CommandLineInterpreter().Execute("selectionset", acc);
      acc.GetParameter("PROJECT", ref projectPath);

      return projectPath;
    }

    // Read EPLAN XML-ProjectInfo and returns the value
    private static string GetValueSettingsXml(string filename, string url)
    {
      XmlDocument xmlDocument = new XmlDocument();
      xmlDocument.Load(filename);

      XmlNodeList rankListSchemaName = xmlDocument.SelectNodes(url);
      if (rankListSchemaName != null && rankListSchemaName.Count > 0)
      {
        // Get Text from MultiLanguage or not :)
        string value = rankListSchemaName[0].InnerText;
        return value;
      }
      else
      {
        return null;
      }
    }
  }

  class GetProjectVariableLanguage
  {
    /* Usage
		private static string GetProjectVariableLanguage()
		{
			string value = null;
			ActionCallingContext actionCallingContext = new ActionCallingContext();
			new CommandLineInterpreter().Execute("GetProjectVariableLanguage", actionCallingContext);
			actionCallingContext.GetParameter("value", ref value);
			return value;
		}
		*/

    private readonly string TempPath = Path.Combine(
      PathMap.SubstitutePath("$(TMP)"), "GetProjectLanguages.xml");

    [DeclareAction("GetProjectVariableLanguage")]
    public void Action(out string value)
    {
      ActionCallingContext actionCallingContext = new ActionCallingContext();
      actionCallingContext.AddParameter("prj", FullProjectPath());
      actionCallingContext.AddParameter("node", "TRANSLATEGUI");
      actionCallingContext.AddParameter("XMLFile", TempPath);
      new CommandLineInterpreter().Execute("XSettingsExport", actionCallingContext);

      if (File.Exists(TempPath))
      {
        string languagesString = GetValueSettingsXml(TempPath,
          "/Settings/CAT/MOD/Setting[@name='VAR_LANGUAGE']/Val");

        if (languagesString != null)
        {
          value = languagesString;
          return;
        }
      }

      value = null;
      return;
    }

    // Returns the EPLAN Project Path
    private static string FullProjectPath()
    {
      ActionCallingContext acc = new ActionCallingContext();
      acc.AddParameter("TYPE", "PROJECT");

      string projectPath = string.Empty;
      new CommandLineInterpreter().Execute("selectionset", acc);
      acc.GetParameter("PROJECT", ref projectPath);

      return projectPath;
    }

    // Read EPLAN XML-ProjectInfo and returns the value
    private static string GetValueSettingsXml(string filename, string url)
    {
      XmlDocument xmlDocument = new XmlDocument();
      xmlDocument.Load(filename);

      XmlNodeList rankListSchemaName = xmlDocument.SelectNodes(url);
      if (rankListSchemaName != null && rankListSchemaName.Count > 0)
      {
        // Get Text from MultiLanguage or not :)
        string value = rankListSchemaName[0].InnerText;
        return value;
      }
      else
      {
        return null;
      }
    }
  }
}
