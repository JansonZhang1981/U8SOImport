Public Class item
    Public Text As String
    Public Value As String
    Public id As String
    Public Sub New(ByVal iText As String, ByVal iValue As String, ByVal iId As String)
        Me.Text = iText
        Me.Value = iValue
        Me.id = iId
    End Sub
    Public Overrides Function ToString() As String
        Return Me.Text
    End Function
End Class
