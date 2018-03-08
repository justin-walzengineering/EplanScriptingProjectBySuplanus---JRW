Public Class ApriFiles

  <Start()> _
    Public Sub ApriFiles() ' aprifiles = openfiles

        Dim openFileDialog1 As New OpenFileDialog()
        Dim File As String

        openFileDialog1.InitialDirectory = "C:\Program\EPLAN\Electric P8\Projects" ' your path
        openFileDialog1.Filter = "eplan files (*.elk)|*.elk"
        openFileDialog1.FilterIndex = 2
        openFileDialog1.RestoreDirectory = True
        openFileDialog1.Multiselect = True
        openFileDialog1.Title = "Select ELK files to export to DWG - utility by ProjectCenter" 'title of dialog box
        openFileDialog1.ShowDialog()
        MessageBox.Show("Press OK to start") 'start when ready message

        For Each File In OpenFileDialog1.FileNames
            Dim exportContext As New ActionCallingContext()
            exportContext.AddParameter("TYPE", "DWGPROJECT")
            exportContext.AddParameter("PROJECTNAME", File)
            exportContext.AddParameter("DESTINATIONPATH", "C:\DWG")
            Dim commandLineInterpreter As New CommandLineInterpreter()
            CommandLineInterpreter.Execute("export", exportContext)
        Next File

        MessageBox.Show("Export Terminated") 'end of operation messagge
        Return
    End Sub
End Class