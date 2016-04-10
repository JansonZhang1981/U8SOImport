Public Class item
    Public Text As String
    Public Vaule As String
    Public id As String
    Public Sub New(ByVal iText As String, ByVal iValue As String, ByVal iId As String)
        Text = iText
        Vaule = iValue
        id = iId
    End Sub
    Public Overrides Function ToString() As String
        Return Me.Text
    End Function
End Class
