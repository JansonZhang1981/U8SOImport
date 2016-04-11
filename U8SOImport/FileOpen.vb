Imports System.Data
Imports System.Data.OleDb
Public Class FileOpen

    Private Sub FileOpen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'conn.ConnectionString = connstr
        conn.Open()
        Dim cmd As New OleDbCommand
        cmd.CommandText = "select ccuscode,ccusname,ccusabbname from Customer"
        cmd.Connection = conn
        Dim myread As OleDbDataReader = cmd.ExecuteReader

        Do While myread.Read
            Dim it As item = New item(myread("ccusname").ToString, myread("ccusabbname").ToString, myread("ccuscode").ToString)

            ComboBox1.Items.Add(it)
        Loop
        conn.Close()
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            TextBox1.Text = OpenFileDialog1.FileName
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MsgBox("导入文件必须选择！", , "警告")
            Return
        ElseIf ComboBox1.SelectedItem Is Nothing Then
            MsgBox("客户必须选择！", , "警告")
            Return
        Else
            filename = TextBox1.Text
            cus = ComboBox1.SelectedItem
            MsgBox(cus.Value)
            Me.Hide()
            '   Dim elForm As New ExcelLoad

            '   elForm.Show()
        End If

    End Sub

End Class