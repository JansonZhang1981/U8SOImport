Imports System.Data
Imports System.Data.OleDb
Public Class Inventory
    Public cInvCode As String
    Public cInvAddCode As String
    Public cInvName As String
    Public cGroupCode As String
    Public cComUnitCode As String

    Public Sub New(ByVal th As String)
        Me.cInvAddCode = th
        Dim excCon As New OleDbConnection
        excCon.ConnectionString =U8Login.UfDbName 
        excCon.Open()
        Dim cmd As New OleDbCommand
        cmd.CommandText = "select * from Inventory where cInvAddCode='" + cInvAddCode + "' order by cInvCCode"
        cmd.Connection = excCon
        Dim myread As OleDbDataReader = cmd.ExecuteReader

        If myread.Read Then
            Me.cInvCode = myread("cInvCode").ToString
            Me.cInvName = myread("cInvName").ToString
            Me.cGroupCode = myread("cGroupCode").ToString
            Me.cComUnitCode = myread("cComUnitCode").ToString
        End If

        excCon.Close()
    End Sub

End Class
