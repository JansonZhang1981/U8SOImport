﻿Imports System.Data
Imports System.Data.OleDb
Imports UFIDA.U8.MomServiceCommon
Imports UFIDA.U8.U8MOMAPIFramework
Imports UFIDA.U8.U8APIFramework
Imports UFIDA.U8.U8APIFramework.Meta
Imports UFIDA.U8.U8APIFramework.Parameter
Imports MSXML2
Imports System.Threading
Public Class ExcelLoad2
    Public SMains As SOMain()
    Public j As Integer
    Public info As String
    Public excConn As OleDbConnection
    Public dt As DataTable
    Public newdt As DataTable
    Public fd(5) As String
    Private Sub ExcelLoad_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'GroupBox1.Text = cus.ccusabbname

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
        filename = "C:/123.xls"
        Dim _Connectstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=<FilePath>;Extended Properties=""Excel 8.0;HDR=NO;IMEX=1"""
        excConn = New OleDb.OleDbConnection(_Connectstring.Replace("<FilePath>", filename))
        excConn.Open()
        dt = New DataTable()
        Dim Sql As String = "select * from [Sheet1$] where F1 is not null"
        Dim mycommand As OleDbDataAdapter = New OleDbDataAdapter(Sql, excConn)
        mycommand.Fill(dt)
        fd(0) = dt.Rows(0)("F3").ToString
        fd(1) = dt.Rows(0)("F4").ToString
        fd(2) = dt.Rows(0)("F5").ToString
        fd(3) = dt.Rows(0)("F6").ToString
        fd(4) = dt.Rows(0)("F7").ToString
        fd(5) = dt.Rows(0)("F8").ToString

        Dim dv As DataView = New DataView(dt)
        Dim dt2 As DataTable = dv.ToTable(True, "F1")
        Dim n As Integer = 0
        For i = 0 To dt2.Rows.Count - 1
            If dt2.Rows(i)("F1").ToString <> "到货工厂代码" Then
                ReDim Preserve SMains(n)
                Dim sm As New SOMain("", dt2.Rows(i)("F1").ToString, "", "", "", "")
                SMains(n) = sm
                n += 1
            End If

        Next
        '  newdt = tempdt()

        j = 0

        If SMains.Length > 0 Then
            Button3.Enabled = True
            Button4.Enabled = True
            Button5.Enabled = True
            Button6.Enabled = True
            showSO1(j)
        Else
            Button3.Enabled = False
            Button4.Enabled = False
            Button5.Enabled = False
            Button6.Enabled = False
        End If

       
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        excConn.Close()
        Me.Hide()
        FileOpen.Show()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim t As New Thread(AddressOf showprogressbar)

        t.Start()
        Thread.Sleep(0)
        SOImport()
        t.Abort()
        MsgBox("导入成功", MsgBoxStyle.OkOnly, "提示")


    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        showSO1(0)
        j = 0
    End Sub
    Public Sub showSO1(ByVal i As Integer)
        DataGridView1.Rows.Clear()

        info = "第" + (i + 1).ToString + "页；共" + SMains.Length.ToString + "页"

        Label7.Text = info
        '    TextBox1.Text = SOMains(i).cusSONo

        TextBox3.Text = SMains(i).dhf

        Dim dv As DataView = New DataView(newdt)
        dv.RowFilter = "dhf = '" + SMains(i).dhf + "'"
        Dim dt2 As DataTable = dv.ToTable()
        Dim j As Integer
        For j = 0 To dt2.Rows.Count - 1
            Dim index As Integer = DataGridView1.Rows.Add()
            DataGridView1.Rows(index).Cells(0).Value = dt2.Rows(j)("partno").ToString

            DataGridView1.Rows(index).Cells(1).Value = dt2.Rows(j)("sl").ToString
            DataGridView1.Rows(index).Cells(2).Value = dt2.Rows(j)("sdate").ToString

        Next

    End Sub


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If j > 0 Then
            j = j - 1
            showSO1(j)
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If j < SMains.Length - 1 Then
            j = j + 1
            showSO1(j)
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        j = SMains.Length - 1
        showSO1(j)
    End Sub
    Private Sub SOImport()
        On Error GoTo ErrHandler
        Dim v As Integer

        For i = 0 To SMains.Length - 1
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
            domHead(0).SetValue("ccusabbname", cus.ccusabbname)   '客户简称，String类型
            domHead(0).SetValue("ccuscode", cus.ccuscode)   '客户编码，String类型
            domHead(0).SetValue("ccusname", cus.ccusname)   '客户名称，String类型
            domHead(0).SetValue("cdepname", "市场部")   '销售部门，String类型
            domHead(0).SetValue("itaxrate", "17")   '税率，Double类型
            domHead(0).SetValue("cexch_name", "人民币")   '币种，String类型
            domHead(0).SetValue("cmaker", u8login.cUserName)   '制单人，String类型
            domHead(0).SetValue("cstcode", "01")   '销售类型编号，String类型
            domHead(0).SetValue("cdepcode", "07")   '部门编码，String类型
            domHead(0).SetValue("iexchrate", "1")   '汇率，Double类型
            domHead(0).SetValue("cdefine10", SMains(i).dhf)   '到货方，String类型
         

            Dim domBody As BusinessObject
            domBody = u8apiBroker.GetBoParam("domBody")
            ' domBody.RowCount = 10 '设置BO对象(表体)行数为多行


            Dim dv As DataView = New DataView(newdt)
            dv.RowFilter = "dhf = '" + SMains(i).dhf + "'"
            Dim dt2 As DataTable = dv.ToTable()
            Dim j As Integer
            Dim y As Integer = 0
            For j = 0 To dt2.Rows.Count - 1

                Dim inv As New Inventory(dt2.Rows(j)("partno").ToString)
                Dim quantity As String = dt2.Rows(j)("sl").ToString
                Dim yfhrq As String = dt2.Rows(j)("sdate").ToString


                '****************************** 以下是必输字段 *****************************
                '’  domBody(y).SetValue("isosid", "字段值")   '主关键字段，Integer类型
                domBody(y).SetValue("cinvname", inv.cInvName)   '存货名称，String类型
                domBody(y).SetValue("cinvcode", inv.cInvCode)   '存货编码，String类型
                '  domBody(y).SetValue("autoid", "字段值")   '销售订单 2，Integer类型
                domBody(y).SetValue("iquantity", quantity)   '数量，Double类型
                domBody(y).SetValue("dpredate", yfhrq)   '预发货日期，Date类型
                domBody(y).SetValue("dpremodate", yfhrq)   '预完工日期，Date类型
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

            Next


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

        'MsgBox("导入成功", MsgBoxStyle.OkOnly, "提示")
        Button1.Enabled = False
        excConn.Close()
        Exit Sub
ErrHandler:
        '  MsgBox(v)
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

    Public Sub showprogressbar()

        Dim pr As New waitForm
        If pr.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub

    End Sub

    Public Function tempdt() As DataTable

        Dim tdt As DataTable = New DataTable("temp_table")
        tdt.Columns.Add("dhf", System.Type.GetType("System.String"))
        tdt.Columns.Add("partno", System.Type.GetType("System.String"))
        tdt.Columns.Add("sl", System.Type.GetType("System.String"))

        tdt.Columns.Add("sdate", System.Type.GetType("System.String"))

        Dim dv As DataView = New DataView(dt)
        For i = 0 To SMains.Length - 1
            dv.RowFilter = "F1 = '" + SMains(i).dhf + "'"
            Dim dt2 As DataTable = dv.ToTable()

            Dim j, k As Integer

            For j = 0 To dt2.Rows.Count - 1
                Dim pn As String = dt2.Rows(j)("F2").ToString

                For k = 0 To 5
                    Dim x As Integer = k + 3
                    Dim col As String = "F" + CStr(x)
                    If dt2.Rows(j)(col).ToString <> "" Then

                        Dim newrow As DataRow = tdt.NewRow
                        newrow("dhf") = SMains(i).dhf
                        newrow("partno") = pn
                        newrow("sl") = dt2.Rows(j)(col).ToString
                        newrow("sdate") = fd(k)
                        tdt.Rows.Add(newrow)

                    End If

                Next

            Next

        Next

        Return tdt

    End Function
End Class