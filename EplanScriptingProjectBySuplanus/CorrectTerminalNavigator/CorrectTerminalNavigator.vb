Public Class CorrectTerminals

    ' Goal
    ' This will allow user ability to choose Correct Terminals function inside Terminals navigator menu
    ' Right click on Terminal in navigator and see "Correct Terminals..." at bottom of pop up. 

    ' Load script in Eplan using [Utilities]>[Scripts]>[Load]
    ' Then choose the file from the file location. 
    ' The file will be a .vb extension. 

    <DeclareMenu()>
    Public Sub MenuFunction()

        Dim oConMenuLoc As New Eplan.EplApi.Gui.ContextMenuLocation("TSGViewTree", "1004")
        Dim oConMenu As New Eplan.EplApi.Gui.ContextMenu()
        oConMenu.Addmenuitem(oConMenuLoc, "Correct Terminals...", "XTPCallAutoCorrectionDlgAction", False, False)

    End Sub
End Class