Imports System.Windows.Forms

Public Class mainForm

    Private Sub ShowNewForm(ByVal sender As Object, ByVal e As EventArgs) Handles NewToolStripButton.Click

    End Sub


    Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub DataAnalyzerStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles DataAnalyzerStripMenuItem.Click
        Dim openFileForm As New frmOpenFile
        openFileForm.MdiParent = Me
        openFileForm.Show()
    End Sub


End Class
