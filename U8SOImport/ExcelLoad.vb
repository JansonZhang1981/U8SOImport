Imports System.Data
Imports System.Data.OleDb
Imports UFIDA.U8.MomServiceCommon
Imports UFIDA.U8.U8MOMAPIFramework
Imports UFIDA.U8.U8APIFramework
Imports UFIDA.U8.U8APIFramework.Meta
Imports UFIDA.U8.U8APIFramework.Parameter
Imports MSXML2
Public Class ExcelLoad
    Public SOMains As SOMain()
    Public j As Integer
    Public info As String
    Public excConn As OleDbConnection
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
        excConn = New OleDb.OleDbConnection(_Connectstring.Replace("<FilePath>", filename))
        excConn.Open()
        Dim dCmd As New OleDb.OleDbCommand
        dCmd.CommandText = "select distinct 订单号 from [Sheet1$]  where 订单号 is not null"
        dCmd.Connection = excConn
        Try
            Dim dr As OleDbDataReader
            dr = dCmd.ExecuteReader
            Dim i As Integer = 0
            Do While dr.Read
                Dim d2cmd As New OleDbCommand
                d2cmd.CommandText = "select top 1 *  from [Sheet1$]  where 订单号 ='" + dr("订单号").ToString + "'"
                d2cmd.Connection = excConn
                Dim d2r As OleDbDataReader
                d2r = d2cmd.ExecuteReader
                Do While d2r.Read
                    ReDim Preserve SOMains(i)
                    Dim sm As New SOMain(d2r("订单号").ToString, d2r("到货方").ToString, d2r("要货方").ToString, d2r("到货仓库").ToString, Format(CDate(d2r("订单接收时间").ToString), "yyyy-MM-dd"), Format(CDate(d2r("要求到货时间").ToString), "yyyy-MM-dd"))
                    SOMains(i) = sm
                    i += 1
                Loop
            Loop

        Catch ex As Exception
            MsgBox("请注意Sheet名是否为Sheet1！")
        End Try

        j = 0

        If SOMains.Length > 0 Then
            Button3.Enabled = True
            Button4.Enabled = True
            Button5.Enabled = True
            Button6.Enabled = True
            showSO(j)
        Else
            Button3.Enabled = False
            Button4.Enabled = False
            Button5.Enabled = False
            Button6.Enabled = False
        End If

        '  excConn.Close()


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        excConn.Close()
        Me.Hide()
        FileOpen.Show()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        SOImport()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        showSO(0)
        j = 0
    End Sub
    Public Sub showSO(ByVal i As Integer)
        info = "第" + (i + 1).ToString + "页；共" + SOMains.Length.ToString + "页"
        Label7.Text = info
        TextBox1.Text = SOMains(i).cusSONo
        TextBox2.Text = SOMains(i).yhf
        TextBox3.Text = SOMains(i).dhf
        TextBox4.Text = SOMains(i).dhck
        TextBox5.Text = SOMains(i).sodate
        TextBox6.Text = SOMains(i).dhrq

        Dim dt As DataTable

        '上两行打开一个读取excel的链接
        '   MsgBox(_Connectstring)
        Dim mydataset As DataSet = New DataSet
        Using da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter("select distinct 零件号,零件名称,订货数量 from [Sheet1$] where 订单号='" + SOMains(i).cusSONo + "'", excConn)

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


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If j > 0 Then
            j = j - 1
            showSO(j)
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If j < SOMains.Length - 1 Then
            j = j + 1
            showSO(j)
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        j = SOMains.Length - 1
        showSO(j)
    End Sub
    Private Sub SOImport()
        On Error GoTo ErrHandler
        Dim v As Integer

        For i = 0 To SOMains.Length - 1
            Dim u8EnvCtx As New U8EnvContext
            u8EnvCtx.U8Login = u8login
            '设置上下文参数
            u8EnvCtx.SetApiContext("VoucherType", 12)  '上下文数据类型：int，含义：单据类型：12

            '第三步：构造ApiBroker对象,调用Connect,传入Api的地址标识(Url)，传入上下文
            Dim u8apiBroker As New U8ApiComBroker

            u8apiBroker.Connect("U8API/SaleOrder/Save", u8EnvCtx)
            '方法二是构造BusinessObject对象，具体方法如下：
            Dim domHead As BusinessObject
            domHead = u8apiBroker.GetBoParam("domHead")

            domHead.RowCount = 1 '设置BO对象(表头)行数，只能为一行
            '给BO对象(表头)的字段赋值，值可以是真实类型，也可以是无类型字符串
            '****************************** 以下是必输字段 *****************************
            domHead(0).SetValue("id", "100000002")   '主关键字段，Integer类型
            domHead(0).SetValue("csocode", "200000002")   '订 单 号，String类型
            domHead(0).SetValue("ddate", Format(Now(), "yyyy-MM-dd"))   '订单日期，Date类型
            domHead(0).SetValue("cbustype", "普通销售")   '业务类型，Integer类型
            domHead(0).SetValue("cstname", "普通销售")   '销售类型，String类型
            domHead(0).SetValue("ccusabbname", cus.Value)   '客户简称，String类型
            domHead(0).SetValue("cdepname", "市场部")   '销售部门，String类型
            domHead(0).SetValue("itaxrate", "17")   '税率，Double类型
            domHead(0).SetValue("cexch_name", "人民币")   '币种，String类型
            domHead(0).SetValue("cmaker", u8login.cUserName)   '制单人，String类型
            domHead(0).SetValue("cstcode", "01")   '销售类型编号，String类型
            domHead(0).SetValue("cdepcode", "07")   '部门编码，String类型
            domHead(0).SetValue("ccuscode", "01")   '客户编码，String类型
            domHead(0).SetValue("iexchrate", "1")   '汇率，Double类型
            domHead(0).SetValue("cdefine10", SOMains(i).dhf)   '到货方，String类型
            domHead(0).SetValue("cdefine11", SOMains(i).cusSONo)   '客户订单号，String类型
            domHead(0).SetValue("cdefine12", SOMains(i).yhf)   '要货方，String类型
            domHead(0).SetValue("cdefine13", SOMains(i).dhck)   '到货仓库，String类型

            Dim domBody As BusinessObject
            domBody = u8apiBroker.GetBoParam("domBody")
            ' domBody.RowCount = 10 '设置BO对象(表体)行数为多行
            Dim d3Cmd As New OleDb.OleDbCommand
            d3Cmd.CommandText = "select * from [Sheet1$]  where 订单号 ='" + SOMains(i).cusSONo + "'"
            d3Cmd.Connection = excConn
            Dim d3r As OleDbDataReader
            d3r = d3Cmd.ExecuteReader
            Dim y As Integer = 0
            Do While d3r.Read
                Dim inv As New Inventory(d3r("零件号").ToString)
                Dim quantity As String = d3r("订货数量").ToString
                '****************************** 以下是必输字段 *****************************
                '’  domBody(y).SetValue("isosid", "字段值")   '主关键字段，Integer类型
                domBody(y).SetValue("cinvname", inv.cInvName)   '存货名称，String类型
                domBody(y).SetValue("cinvcode", inv.cInvCode)   '存货编码，String类型
                '  domBody(y).SetValue("autoid", "字段值")   '销售订单 2，Integer类型
                domBody(y).SetValue("iquantity", quantity)   '数量，Double类型
                domBody(y).SetValue("dpredate", "2016-4-20")   '预发货日期，Date类型
                domBody(y).SetValue("dpremodate", "2016-4-19")   '预完工日期，Date类型
                domBody(y).SetValue("id", "100000002")   '主表id，Integer类型
                'domBody(y).SetValue("iinvexchrate", "字段值")   '换算率，Double类型
                'domBody(y).SetValue("cunitid", "字段值")   '销售单位编码，String类型
                'domBody(y).SetValue("cinva_unit", "字段值")   '销售单位，String类型
                domBody(y).SetValue("cinvm_unit", inv.cComUnitCode)   '主计量单位，String类型
                'domBody(y).SetValue("igrouptype", "字段值")   '单位类型，Integer类型
                domBody(y).SetValue("cgroupcode", inv.cGroupCode)   '计量单位组，String类型
                'domBody(y).SetValue("dreleasedate", "字段值")   '预留失效日期，Date类型
                domBody(y).SetValue("editprop", "A")   '编辑属性：A表新增，M表修改，D表删除，String类型
                y += 1
                v = y
            Loop


            '给普通参数VoucherState赋值。此参数的数据类型为Integer，此参数按值传递，表示状态:0增加;1修改
            u8apiBroker.AssignNormalValue("VoucherState", 0)  '参数类型：Integer

            '该参数vNewID为INOUT型普通参数。此参数的数据类型为String，此参数按值传递。在API调用返回时，可以通过GetResult("vNewID")获取其值
            u8apiBroker.AssignNormalValue("vNewID", "000000002")  '参数类型：String

            '给普通参数DomConfig赋值。此参数的数据类型为MSXML2.IXMLDOMDocument2，此参数按引用传递，表示ATO,PTO选配
            Dim CurDom As New DOMDocument
            u8apiBroker.AssignNormalValue("DomConfig", CurDom)  '参数类型：MSXML2.IXMLDOMDocument2

            '第五步：调用API
            If u8apiBroker.InvokeApi() = False Then

                '第六步：错误处理
                If u8apiBroker.ErrorType = ExceptionType.Business Then

                    '处理API业务错误
                ElseIf u8apiBroker.ErrorType = ExceptionType.System Then

                    '处理系统错误
                End If
            Else
                '第七步：获取返回结果

                '获取返回值
                '获取普通返回值。此返回值数据类型为String，此参数按值传递，表示成功为空串
                Dim result As String
                result = CStr(u8apiBroker.GetReturnValue())
                '获取out/inout参数值

                '获取普通INOUT参数vNewID。此返回值数据类型为String，在使用该参数之前，请判断是否为空
                Dim vNewIDRet As String
                vNewIDRet = CStr(u8apiBroker.GetResult("vNewID"))

            End If
            '结束本次调用，释放API资源
            u8apiBroker.Disconnect()

            u8apiBroker = Nothing
        Next
      
        MsgBox("OK")
        Exit Sub
ErrHandler:
        MsgBox(v)
        MsgBox(Err.Description)


    End Sub

   

    Private Sub DataGridView1_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView1.RowPostPaint
        Try

            Dim rectangle As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, Me.DataGridView1.RowHeadersWidth - 4, e.RowBounds.Height)

            TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), Me.DataGridView1.RowHeadersDefaultCellStyle.Font, rectangle, DataGridView1.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.Right)
        Catch ex As Exception

            MsgBox(ex.ToString, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly)

        End Try
        
    End Sub
End Class