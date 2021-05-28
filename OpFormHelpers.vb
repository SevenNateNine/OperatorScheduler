Imports Microsoft.VisualBasic

Partial Public Class OperatorMainForm
    ''' <summary>
    ''' adds text to rich textbox
    ''' </summary>
    ''' <param name="Message"></param>
    ''' <param name="timestamp"></param>
    ''' Defaults to current time if none specificed
    Private Sub ConsoleAdd(ByVal Message As String, Optional ByVal timestamp As String = "")
        If timestamp.Length = 0 Then
            timestamp = DateTime.Now.ToString()
        End If
        ConsoleRichTextBox.Text += String.Format("{0}    {1}{2}", timestamp, Message, vbLf)
    End Sub
End Class
