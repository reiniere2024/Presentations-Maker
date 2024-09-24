Public Class MainMenu


    Private Sub MainMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


    End Sub


    Private Sub ViewBooksToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewBooksToolStripMenuItem.Click

        Dim myform As New RunPresentations

        myform.MdiParent = Me
        myform.Show()


    End Sub


    Private Sub CardsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CardsToolStripMenuItem.Click

        Dim myform As New MakePresentations

        myform.MdiParent = Me
        myform.Show()


    End Sub

    Private Sub MenuStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub
End Class