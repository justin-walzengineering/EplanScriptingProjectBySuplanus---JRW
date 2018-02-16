using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Scripting;

// Goal:
// This will translate language of project to ES (Spanish, US, Chinese)
// You can see/confirm this worked correctly by going to [Options] > [Settings] > {"Project name"} > [Translation] > [General], then look in Translation column. 
// This does not affect the display language of the project

// Run script in Eplan using [Utilities]>[Scripts]>[Run]
// Then choose the file from the file location. 
// The file will be a .cs extension. 

internal class AddProjectLanguageClass
{
    [Start]
    public void XAfActionSettingProject_Start()
    {
        CommandLineInterpreter oCLI = new CommandLineInterpreter();
        ActionCallingContext oACC = new ActionCallingContext();
        oACC.AddParameter("set", "TRANSLATEGUI.TRANSLATE_LANGUAGES");
        oACC.AddParameter("value", "es_ES;en_US;zh_CN;");

        oCLI.Execute("XAfActionSettingProject", oACC);
        return;
    }
}