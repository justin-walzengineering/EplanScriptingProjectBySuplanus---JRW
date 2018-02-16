Public Class CorrectConnections

    ' Goal
    ' This will allow user ability to choose Correct Connections function inside Connections navigator menu
    ' Right click on Connection in navigator and see "Correct Connections..." at bottom of pop up. 

    ' Load script in Eplan using [Utilities]>[Scripts]>[Load]
    ' Then choose the file from the file location. 
    ' The file will be a .vb extension. 

    <DeclareMenu()>
    Public Sub MenuFunction()

        Dim oConMenuLoc As New Eplan.EplApi.Gui.ContextMenuLocation("XCmPrjDataTreeDialog", "1017")
        Dim oConMenu As New Eplan.EplApi.Gui.ContextMenu()
        oConMenu.Addmenuitem(oConMenuLoc, "Correct Connections...", "XCMCorrectionChoiceAction", False, False)

    End Sub
End Class