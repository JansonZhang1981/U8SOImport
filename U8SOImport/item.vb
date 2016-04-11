Public Class item
    Public ccusname As String
    Public ccusabbname As String
    Public ccuscode As String
    Public Sub New(ByVal ccusname As String, ByVal ccusabbname As String, ByVal ccuscode As String)
        Me.ccusname = ccusname
        Me.ccusabbname = ccusname
        Me.ccuscode = ccuscode
    End Sub
    Public Overrides Function ToString() As String
        Return Me.ccusname
    End Function
End Class
