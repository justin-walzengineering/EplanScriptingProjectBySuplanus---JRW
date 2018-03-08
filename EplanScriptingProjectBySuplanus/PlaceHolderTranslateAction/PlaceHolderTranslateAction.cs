// PlaceHolderTranslateAction.cs
//
// extends the context menu of the placeholder object (Tab Values) with the menu item "Translate"
// and to the menu item "Translations remove"
//
// Copyright by Frank Schöneck, 2015
// last change:
// V1.0.0, 23.10.2015, Frank Schöneck, start of project
//
// for Eplan Electric P8, from V2.5
//

using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Scripting;

namespace EplanScriptingProjectBySuplanus.PlaceHolderTranslateAction
{
    public class PlaceHolderTranslate
    {
        [DeclareMenu]
        public void PlaceHolderTranslateContextMenu()
        {
            //Context-Menüeintrag (hier im Platzhalterobjekt)
            Eplan.EplApi.Gui.ContextMenu oContextMenu = new Eplan.EplApi.Gui.ContextMenu();
            Eplan.EplApi.Gui.ContextMenuLocation oContextMenuLocation = new Eplan.EplApi.Gui.ContextMenuLocation("PlaceHolder", "1004");
            oContextMenu.AddMenuItem(oContextMenuLocation, "Übersetzen", "PlaceHolderTranslateAction", false, false);
            oContextMenu.AddMenuItem(oContextMenuLocation, "Übersetzungen entfernen", "PlaceHolderTranslateDeleteAction", false, false);
        }

        [DeclareAction("PlaceHolderTranslateAction")]
        public void PlaceHolderTranslate_Action()
        {
            //Übersetzen
            new CommandLineInterpreter().Execute("EnfTranslateEditAction");
        }

        [DeclareAction("PlaceHolderTranslateDeleteAction")]
        public void PlaceHolderTranslateDelete_Action()
        {
            //Übersetzungen entfernen
            new CommandLineInterpreter().Execute("EnfDeleteEditTranslationsAction");
        }

    }
}
