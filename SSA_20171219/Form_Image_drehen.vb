Public Class Form_Image_drehen
    '   Public PB As PictureBox

    Private Sub BT_im_Uhrzeigersinn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_im_Uhrzeigersinn.Click

        'Lageplan im Uhrzeigersinn drehen

        Form_Main.PB_Grundriss.BackgroundImage.RotateFlip(RotateFlipType.Rotate90FlipNone)
        Form_Main.PB_Grundriss.Refresh()
        Form_Main.PB_DB_Grundriss.Refresh()
        '   btmGrundriss.RotateFlip(RotateFlipType.Rotate90FlipNone)

    End Sub

    Private Sub BT_gegen_Uhrzeigersinn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_gegen_Uhrzeigersinn.Click

        'Lageplan im Uhrzeigersinn drehen
        Form_Main.PB_Grundriss.BackgroundImage.RotateFlip(RotateFlipType.Rotate270FlipNone)
        Form_Main.PB_Grundriss.Refresh()
        Form_Main.PB_DB_Grundriss.Refresh()
        'btmGrundriss.RotateFlip(RotateFlipType.Rotate270FlipNone)

    End Sub

    Private Sub BT_Abbruch_Uhrzeigersinn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Abbruch_Uhrzeigersinn.Click

        'Abbruch-Button bei Abfrage Lageplan drehen
        btmGrundriss = New Bitmap(Form_Main.PB_Grundriss.BackgroundImage)

        Me.Close()

    End Sub

    
    'Private Sub Form_Image_drehen_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
    '    Form_Main.Update_SSA_Deckblatt()
    'End Sub
End Class