Public Class CorrectCable

    ' Goal
    ' This will allow user ability to choose Correct Cable function inside Cable navigator menu
    ' Right click on cable in navigator and see "Correct cable..." at bottom of pop up. 

    ' Load script in Eplan using [Utilities]>[Scripts]>[Load]
    ' Then choose the file from the file location. 
    ' The file will be a .vb extension. 

    <DeclareMenu()>
    Public Sub MenuFunction()

        Dim oConMenuLoc As New Eplan.EplApi.Gui.ContextMenuLocation("XCCablePrjDataTabTree", "1126")
        Dim oConMenu As New Eplan.EplApi.Gui.ContextMenu()
        oConMenu.Addmenuitem(oConMenuLoc, "Correct cable...", "XCActionCorrectCable", False, False)

    End Sub
End Class