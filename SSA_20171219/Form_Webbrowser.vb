Public Class Form_Webbrowser

    Private Sub Form_Webbrowser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.WebBrowser1.Navigate("www.schallschutzausweis.tu-bs.de")

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim str As String = Me.WebBrowser1.Document.ActiveElement.Name
        Me.WebBrowser1.Document.GetElementById("thefile").OuterText = "c:\thefile.xml"

    End Sub
End Class