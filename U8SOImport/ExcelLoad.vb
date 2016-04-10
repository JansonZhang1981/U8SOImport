Public Class ExcelLoad

    Private Sub ExcelLoad_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
     
        Dim x As String
        If Not Is64bit() Then
            x = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\12.0\Access Connectivity Engine\Engines\Excel", "TypeGuessRows", Nothing)
            If x <> "0" Then
                My.Computer.Registry.SetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\12.0\Access Connectivity Engine\Engines\Excel", "TypeGuessRows", 0, Microsoft.Win32.RegistryValueKind.DWord)
            End If
        ElseIf Is64bit() Then
            x = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\12.0\Access Connectivity Engine\Engines\Excel", "TypeGuessRows", Nothing)
            If x <> "0" Then
                My.Computer.Registry.SetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\12.0\Access Connectivity Engine\Engines\Excel", "TypeGuessRows", 0, Microsoft.Win32.RegistryValueKind.DWord)
            End If
        End If

        Dim _Connectstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=<FilePath>;Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"""
        Dim excConn As New OleDb.OleDbConnection(_Connectstring.Replace("<FilePath>", FileName))
   
        'Dim dCmd As New OleDb.OleDbCommand
        'dCmd.CommandText = "SELECT 零件号,零件名称,订货数量 FROM [Sheet1$]"
        'dCmd.Connection = excConn
        Dim dt As DataTable

        '上两行打开一个读取excel的链接
        '   MsgBox(_Connectstring)
        Dim mydataset As DataSet = New DataSet
        Using da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter("SELECT 零件号,零件名称,订货数量 FROM [Sheet1$]", excConn)

            Try
                dt = New DataTable
                da.Fill(dt)
            Catch ex As Exception
                '   Console.WriteLine(ex)
                MsgBox("请注意Sheet名是否为Sheet1！")
            End Try

            '   dt = mydataset.Tables("Sheet1")


            DataGridView1.AutoGenerateColumns = True '自动创建列  
            DataGridView1.DataSource = dt
        End Using

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Hide()
        FileOpen.Show()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim th As String = DataGridView1.Item("零件号", DataGridView1.CurrentCell.RowIndex).Value
        Dim inv As New Inventory(th)
        MsgBox(inv.cInvCode)
    End Sub
End Class