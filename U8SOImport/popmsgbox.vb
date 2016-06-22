Public Class popmsgbox

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Me.Hide()

    End Sub

    Private Sub popmsgbox_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        msgtext.Text = msg
    End Sub
End Class