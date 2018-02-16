using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Scripting;

public class PageChangeZoomAll
// *********************************************************************************
// Funktion: Ansicht "Ausschnitt ganze Seite" bei Seitenwechsel 
// *********************************************************************************
// �nderungshistorie:
// 2010-08-04   nairolf Script erstellt
// 2011-05-07   nairolf Erweiterung: Zoom bei Seitenwechsel bei Sprung zu Gegenst�ck
// *********************************************************************************

// ********************************************************************************************
// Function: view "full page view" when paging
// ********************************************************************************************
// change history:
// 2010-08-04 nairolf script created
// 2011-05-07 nairolf extension: Zoom at page change at jump to counterpart
// ********************************************************************************************

{
    [DeclareEventHandler("onActionStart.String.XGedNextPageAction")]
    public bool MyEventHandlerFunction1(IEventParameter iEventParameter)
    {
        PageZoomAll();        
        return true;            
    }
    [DeclareEventHandler("onActionStart.String.XGedPrevPageAction")]
    public bool MyEventHandlerFunction2(IEventParameter iEventParameter)
    {
        PageZoomAll();       
        return true;            
    }
    [DeclareEventHandler("onActionStart.String.XPmPageOpenOnePage")]
    public bool MyEventHandlerFunction3(IEventParameter iEventParameter)
    {
        PageZoomAll();
        return true;            
    }
    [DeclareEventHandler("NotifyPageOpened")]
    public bool MyEventHandlerFunction4(IEventParameter iEventParameter)
    {
        PageZoomAll();
        return true;            
    }    
    public void PageZoomAll()
    {
        CommandLineInterpreter oCLI = new CommandLineInterpreter();
        oCLI.Execute("XGedZoomAllAction");                
    }
}