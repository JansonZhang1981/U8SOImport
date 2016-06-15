﻿Imports UFIDA.U8.MomServiceCommon
Imports UFIDA.U8.U8MOMAPIFramework
Imports UFIDA.U8.U8APIFramework
Imports UFIDA.U8.U8APIFramework.Meta
Imports UFIDA.U8.U8APIFramework.Parameter
Imports MSXML2
Imports System.Environment
Imports System.Configuration

Public Class LoginForm
    Public sYear, sDate As String

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        u8login = New U8Login.clsLogin
        Dim sSubId As String = "AS"
        Dim sAccID As String = ConfigurationManager.AppSettings("sAccID").ToString
        '   Dim sAccID As String = "(default)@001"
        Dim sUserID As String = UsernameTextBox.Text.Trim
        Dim sPassword As String = PasswordTextBox.Text.Trim
        Dim sServer As String = ConfigurationManager.AppSettings("sServer").ToString
        '  Dim sServer As String = "192.168.5.200"

        Dim sSerial As String = ""
        If Not u8login.Login(sSubId, sAccID, sYear, sUserID, sPassword, sDate, sServer, sSerial) Then
            MsgBox("登陆失败，原因：" + u8login.ShareString)
        Else
            connstr = u8login.UfDbName
            conn = New ADODB.Connection
            'conn.ConnectionString = connstr
            conn.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            conn.Open(connstr)


            Me.Hide()
            FileOpen.Show()
            '   Form1.Show()
        End If

        '  Me.Close()
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        '   MsgBox(u8login.cUserName)
        Me.Close()

    End Sub

    Private Sub LoginForm1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        sYear = Format(Now(), "yyyy")
        sDate = Format(Now(), "yyyy-MM-dd")

    End Sub



End Class
