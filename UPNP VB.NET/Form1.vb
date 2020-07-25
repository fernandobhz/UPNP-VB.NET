Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim UPnP As New UPnP
        MsgBox(String.Join(vbCrLf, UPnP.Print.ToArray))
        End

    End Sub
End Class
