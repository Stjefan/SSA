Public Class Form_Aussteller

    Private Sub Form_Aussteller_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            Me.TB_Name.Text = Trim(My.Settings.Aussteller_Name)
            Me.TB_Zusatz.Text = Trim(My.Settings.Aussteller_Zusatz)
            Me.TB_Strasse.Text = Trim(My.Settings.Aussteller_Strasse)
            Me.TB_Nr.Text = Trim(My.Settings.Aussteller_Nr)
            Me.TB_PLZ.Text = Trim(My.Settings.Aussteller_PLZ)
            Me.TB_Ort.Text = Trim(My.Settings.Aussteller_Ort)

            Me.PB_Logo.BackgroundImage = GetImageFromString(My.Settings.Logo)
            Me.PB_Signatur.BackgroundImage = GetImageFromString(My.Settings.Signatur)
        Catch ex As Exception
            Form_Main.Cursor = Cursors.Default
            MsgBox("Fehler beim Einlesen der Projektdaten. Das Projekt wurde vermutlich von einer älteren Version erzeugt!", MsgBoxStyle.OkOnly, "Fehlermeldung")
        End Try
    End Sub


    Private Sub Button_OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_OK.Click

        My.Settings.Aussteller_Name = Me.TB_Name.Text
        My.Settings.Aussteller_Zusatz = Me.TB_Zusatz.Text
        My.Settings.Aussteller_Strasse = Me.TB_Strasse.Text
        My.Settings.Aussteller_Nr = Me.TB_Nr.Text
        My.Settings.Aussteller_PLZ = Me.TB_PLZ.Text
        My.Settings.Aussteller_Ort = Me.TB_Ort.Text

        If Not IsNothing(Me.PB_Logo.BackgroundImage) Then
            '    btmLogo = Nothing
            'btmLogo = New Bitmap(Me.PB_Logo.BackgroundImage)
            My.Settings.Logo = GetStringFromImage(Me.PB_Logo.BackgroundImage)
        Else
            'btmLogo = Nothing
            My.Settings.Logo = GetStringFromImage(Nothing)
        End If



        If Not IsNothing(Me.PB_Signatur.BackgroundImage) Then
            'btmSignatur = Nothing
            'btmSignatur = New Bitmap(Me.PB_Signatur.BackgroundImage)
            My.Settings.Signatur = GetStringFromImage(Me.PB_Signatur.BackgroundImage)
        Else
            'btmSignatur = Nothing
            My.Settings.Signatur = GetStringFromImage(Nothing)
        End If
        My.Settings.Save()

        'Konfig_speichern()
        '' ''Logo_speichern(Me.PB_Logo.BackgroundImage)
        '' ''Signatur_speichern(Me.PB_Signatur.BackgroundImage)

        ''''  Konfig_einlesen()

        Form_Main.PB_Logo.BackgroundImage = GetImageFromString(My.Settings.Logo) 'btmLogo

        Form_Main.Cursor = Cursors.Default

        Me.Close()
        Me.Dispose()

        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub Button_Logo_hinzufuegen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Logo_hinzufuegen.Click
        If OFD_Logo.ShowDialog = System.Windows.Forms.DialogResult.OK Then


            Dim sPfadLogo As String = Me.OFD_Logo.FileName

            If Not IsNothing(Me.PB_Logo.Image) Then
                Me.PB_Logo.BackgroundImage.Dispose()
                Me.PB_Logo.BackgroundImage = Nothing
            End If

            'Logo einlesen ohne Skalierung
            Me.PB_Logo.BackgroundImage = New Bitmap(sPfadLogo)    ' System.Drawing.Image.FromFile(sPfadLogo)


        End If
    End Sub

    Private Sub Button_Logo_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Logo_mit.Click
        Me.PB_Logo.BackgroundImage.RotateFlip(RotateFlipType.Rotate90FlipNone)
        Me.PB_Logo.Refresh()
    End Sub
    Private Sub Button_Logo_gegen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Logo_gegen.Click
        Me.PB_Logo.BackgroundImage.RotateFlip(RotateFlipType.Rotate270FlipNone)
        Me.PB_Logo.Refresh()
    End Sub
    Private Sub Button_Logo_entfernen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Logo_entfernen.Click
        If MsgBox("Soll das Logo wirklich entfernt werden?", MsgBoxStyle.YesNo, "Logo entfernen") = MsgBoxResult.Yes Then
            Me.PB_Logo.BackgroundImage.Dispose()
            Me.PB_Logo.BackgroundImage = Nothing
        End If
    End Sub

    Private Sub Button_Signatur_hinzufuegen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Signatur_hinzufuegen.Click
        If Me.OFD_Signatur.ShowDialog = System.Windows.Forms.DialogResult.OK Then

            Dim sPfadSignatur As String = Me.OFD_Signatur.FileName

            'Form_Haupt auf Bildschirmgröße maximieren
            If Not IsNothing(Me.PB_Signatur.Image) Then
                Me.PB_Signatur.BackgroundImage.Dispose()
                Me.PB_Signatur.BackgroundImage = Nothing
            End If

            'Lageplan einlesen ohne Skalierung
            Me.PB_Signatur.BackgroundImage = System.Drawing.Image.FromFile(sPfadSignatur)

        End If
    End Sub

    Private Sub Button_Signatur_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Signatur_mit.Click
        Me.PB_Signatur.BackgroundImage.RotateFlip(RotateFlipType.Rotate90FlipNone)
        Me.PB_Signatur.Refresh()
    End Sub

    Private Sub Button_Signatur_gegen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Signatur_gegen.Click
        Me.PB_Signatur.BackgroundImage.RotateFlip(RotateFlipType.Rotate270FlipNone)
        Me.PB_Signatur.Refresh()
    End Sub

    Private Sub Button_Signatur_entfernen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Signatur_entfernen.Click
        If MsgBox("Soll das Signatur wirklich entfernt werden?", MsgBoxStyle.YesNo, "Logo entfernen") = MsgBoxResult.Yes Then
            Me.PB_Signatur.BackgroundImage.Dispose()
            Me.PB_Signatur.BackgroundImage = Nothing
        End If
    End Sub

End Class