Public Class CorrectPlug

    ' Goal
    ' This will allow user ability to choose Correct Plugs function inside Plugs navigator menu
    ' Right click on Plug in navigator and see "Correct Plugs..." at bottom of pop up. 

    ' Load script in Eplan using [Utilities]>[Scripts]>[Load]
    ' Then choose the file from the file location. 
    ' The file will be a .vb extension. 

    <DeclareMenu()>
    Public Sub MenuFunction()

        Dim oConMenuLoc As New Eplan.EplApi.Gui.ContextMenuLocation("XpluGVTree", "1004")
        Dim oConMenu As New Eplan.EplApi.Gui.ContextMenu()
        oConMenu.Addmenuitem(oConMenuLoc, "Correct Plug...", "XplugCallAutoCorrectionDlgAction", False, False)

    End Sub
End Class