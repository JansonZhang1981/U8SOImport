Public Class SOMain
    Public cusSONo As String '客户订单号
    Public dhf As String '到货方
    Public yhf As String '要货方
    Public dhck As String '到货仓库
    Public sodate As String
    Public dhrq As String

    Public Sub New(ByVal cusSONo As String, ByVal dhf As String, ByVal yhf As String, ByVal dhck As String, ByVal sodate As String, ByVal dhrq As String)
        Me.cusSONo = cusSONo
        Me.dhf = dhf
        Me.yhf = yhf
        Me.dhck = dhck
        Me.sodate = sodate
        Me.dhrq = dhrq

    End Sub

End Class
