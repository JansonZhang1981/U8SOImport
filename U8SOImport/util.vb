Imports UFIDA.U8.MomServiceCommon
Imports UFIDA.U8.U8MOMAPIFramework
Imports UFIDA.U8.U8APIFramework
Imports UFIDA.U8.U8APIFramework.Meta
Imports UFIDA.U8.U8APIFramework.Parameter
Imports MSXML2
Imports System.Data
Imports System.Data.OleDb
Module util
    Public u8login As U8Login.clsLogin
    Public connstr As String
    Public conn As New ADODB.Connection
    Public filename As String
    Public cus As item
    Public msg As String

    Public Function Is64bit() As Boolean
        If Environment.GetEnvironmentVariable("Program Files(x86)") = "" Then

            Return True
        Else

            Return False

        End If

    End Function

    Public Function setAttribute(ByVal nd As IXMLDOMElement, ByVal name As String, ByVal value As Object) As Boolean
        nd.setAttribute(name, CStr(value))
    End Function
    Public Function GetTablename(ByVal i As Integer, ByVal cn As OleDbConnection) As String  'i表示第几个sheet，大于0
        'Dim sConnectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:/123.xls;Extended Properties=Excel 8.0;"
        'Dim cn As OleDbConnection = New OleDbConnection(sConnectionString)

        'cn.Open()
        Dim tb As DataTable = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)

        Return tb.Rows(i - 1)("TABLE_NAME") '第一个


    End Function
    Public Function GetTablenames(ByVal conn As OleDbConnection) As Array
        Dim vList As New List(Of String)
        Try
            'Dim strConn As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=<FilePath>;Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"""
            'Dim conn As OleDbConnection
            'conn = New OleDb.OleDbConnection(strConn.Replace("<FilePath>", "c:/1.xlsx"))
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If


            Dim sheetNames As DataTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})

            conn.Close()
            Dim vName As String = String.Empty
            Dim pOUTPres As New List(Of String)

            For i = 0 To sheetNames.Rows.Count - 1
                If sheetNames.Rows(i)(2).ToString().Trim().Contains("OUTPres") And i > 0 Then
                    If sheetNames.Rows(i)(2).ToString().Trim().Contains(sheetNames.Rows(i - 1)(2).ToString().Trim() + "OUTPres") Then
                        Continue For
                    End If
                End If
                pOUTPres.Add(sheetNames.Rows(i)(2).ToString().Trim())
            Next

            Dim vSheets As String() = pOUTPres.ToArray()
            Dim pSheetName As String = String.Empty



            For i = 0 To vSheets.Length - 1
                Dim pStart As String = vSheets(i).Substring(0, 1)
                Dim pEnd As String = vSheets(i).Substring(vSheets(i).Length - 1, 1)
                If pStart = "'" And pEnd = "'" Then
                    vSheets(i) = vSheets(i).Substring(1, vSheets(i).Length - 2)
                End If

                Dim pChar As Char() = vSheets(i).ToCharArray
                pSheetName = String.Empty
                For j = 0 To pChar.Length - 1
                    If j < pChar.Length - 1 Then
                        If pChar(j).ToString = "'" And pChar(j + 1).ToString = "'" Then
                            pSheetName += pChar(j).ToString
                            j = j + 1
                        Else
                            pSheetName += pChar(j).ToString
                        End If
                    Else
                        pSheetName += pChar(j).ToString
                    End If

                Next
                vSheets(i) = pSheetName
            Next



            For i = 0 To vSheets.Length - 1
                If vList.IndexOf(vSheets(i).ToLower) = -1 Then
                    vList.Add(vSheets(i))
                End If
            Next

            Dim ptList As New List(Of String)
            For j = 0 To vList.Count - 1
                ptList.Add(vList(j))
            Next


            For i = 0 To ptList.Count - 1
                If ptList(i).ToString().Contains("FilterDatabase") Or ptList(i).ToString().Contains("Print_Titles") _
                     Or ptList(i).ToString().Contains("_xlnm#Database") Or ptList(i).ToString().Contains("Print_Area") _
                     Or ptList(i).ToString().Contains("_xlnm.Database") Or ptList(i).ToString().Contains("ExternalData") _
                     Or ptList(i).ToString().Contains("DRUG_IMP_STOCK") Or ptList(i).ToString().Contains("Sheet1$zy") _
                     Or ptList(i).ToString().Contains("Sheet1$xy") Or ptList(i).ToString().Contains("data_xy_zcy") _
                     Or ptList(i).ToString().Contains("Results") Then

                    vList.Remove(ptList(i).ToString)

                End If

            Next

            If vList.Count > 1 Then
                Dim pCheckList As New List(Of String)
                For j = 0 To vList.Count - 1
                    pCheckList.Add(vList(j))
                Next
                conn.Open()
                Dim pComm As New OleDbCommand
                pComm.Connection = conn

                For i = 0 To pCheckList.Count - 1
                    Try
                        pComm.CommandText = String.Format("select count(*) from [{0}] where 1=0", pCheckList(i))
                        pComm.ExecuteNonQuery()
                    Catch ex As Exception
                        If ex.Message.Contains("Microsoft Access 数据库引擎找不到对象") Then
                            vList.Remove(pCheckList(i).ToString)
                        End If

                    End Try
                Next
                conn.Close()
            End If

        Catch ex As Exception

        End Try

        Return vList.ToArray

    End Function
    Function GetFirstSheetNameFromExcelFileName(ByVal numberSheetID As Integer) As String
        If Not System.IO.File.Exists(filename) Then
            Return "文件不存在!"
        End If
        If numberSheetID < 1 Then
            numberSheetID = 1
        End If

        Try
            Dim strFirstSheetName = ""
            Dim obj As New Microsoft.Office.Interop.Excel.Application
            Dim WB As Microsoft.Office.Interop.Excel.Workbook = obj.Workbooks.Open(filename)
    
            Dim xlSheet As Microsoft.Office.Interop.Excel.Worksheet
            xlSheet = WB.Sheets(numberSheetID)
      
            strFirstSheetName = xlSheet.Name
            'xlSheet = Nothing
            'WB.Close()
            'WB = Nothing
            'obj.Quit()
            'obj = Nothing


            obj.Workbooks.Close()

            obj.Quit()
            xlSheet = Nothing
            WB = Nothing
            obj = Nothing

            System.GC.Collect()
           
            Return strFirstSheetName

        Catch ex As Exception
            Return ex.Message

        End Try

    End Function
    
End Module
