﻿Imports UFIDA.U8.MomServiceCommon
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

End Module
