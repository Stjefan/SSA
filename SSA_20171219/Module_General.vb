Module Module_General
    'Implements System.Text

    Public Sub Info_Anzeigen()
        'Dim WebBrowser As WebBrowser = Form_Webbrowser.WebBrowser1 'CType(WebBro, WebBrowser)
        Process.Start("iexplore.exe", "www.schallschutzausweis.tu-bs.de")

        'Form_Webbrowser.Show()
        'WebBrowser.Navigate("www.schallschutzausweis.tu-bs.de")
    End Sub

    Public Sub Senden_Daten()

        Dim str As String
        str = Built_xml()

        Clipboard.SetText(str)

        If Left(str, 7) = "Fehler:" Then
            MsgBox(str, vbOKOnly, "Fehler")
        Else
            UploadFile("www.schallschutzausweis.tu-bs.de/upload.php", str) '"c:\thefile.xml")
        End If
    End Sub

    Public Sub Projekt_schliessen()
        'Es ist gerade ein Projekt geöffnet. Dieses muss geschlossen werden (und evtl. gespeichert), bevor ein neues erzeugt werden kann.
        Dim resultMsg As MsgBoxResult = MsgBox("Soll das Projekt " + stProjekt_Name + " vor dem Schließen gespeichert werden?", _
                                                    MsgBoxStyle.YesNoCancel, "Projekt")
        If resultMsg = MsgBoxResult.Yes Then
            Projektdaten_speichern()
        End If
        If resultMsg <> MsgBoxResult.Cancel Then
            'Das Projekt soll beendet/geschlossen -> die aktuellen Arrays und Variablen müssen zurückgesetzt werden.
            Loesche_Projektdaten()
            Form_Main.Update_Anzeige()
        End If
    End Sub
    Public Sub Drucken()
        With Form_Main.PrintDialog1
            .Document = Form_Main.PrintDocument1
            .AllowSomePages = True
            If .ShowDialog = System.Windows.Forms.DialogResult.OK Then
                Form_Main.PrintDocument1.Print()
            End If
        End With
    End Sub
    Public Sub Projekt_neu()
        Dim bCancel As Boolean = False     'Falls der Anwender ein Projekt geöffnet hat und über das Menü den Punkt "Projekt neu" gewählt hat, aber
        'bei der Rückfrage, ob das aktuelle Projekt gespeichert werden soll den Button "Abbruch" wählt, wird kein neues Projekt eingerichtet.
        If Projekt.bStatus_Projekt Then
            'Es ist gerade ein Projekt geöffnet. Dieses muss geschlossen werden (und evtl. gespeichert), bevor ein neues erzeugt werden kann.
            Dim resultMsg As MsgBoxResult = MsgBox("Das Projekt " + stProjekt_Name + " ist noch geöffnet. Soll das Projekt gespeichert werden?", _
                                                MsgBoxStyle.YesNoCancel, "Projekt")
            If resultMsg = MsgBoxResult.Yes Then
                Projektdaten_speichern()
            End If
            If resultMsg = MsgBoxResult.Cancel Then
                bCancel = True
            Else    'Das Projekt soll beendet/geschlossen -> die aktuellen Arrays und Variablen müssen zurückgesetzt werden.
                Loesche_Projektdaten()
                Form_Main.Update_Anzeige()
            End If
        End If

        If bCancel = False Then
            'Abfrage zu neuem Projekt: Projektname und Dateiverzeichnis in neuem Fenster
            Dim WForm As New Form_Projekt_einrichten
            'WForm.ShowDialog()
            Dim resultDlg As DialogResult = WForm.ShowDialog
            If resultDlg = System.Windows.Forms.DialogResult.OK And stProjekt_Name <> "" And stProjekt_Pfad <> "" Then
                'Es ist ein neues Projekt erzeugt
                Projekt.bStatus_Projekt = True
                Form_Main.Text = "Schallschutzausweis 7Label 2018 - " + stProjekt_Name '+ " - PMKG"    'Projekt.Name + " - PMKG"

                Projektdaten_speichern()

            End If
        End If
    End Sub

    Public Sub Projekt_laden()
        Dim bCancel As Boolean = False     'Falls der Anwender ein Projekt geöffnet hat und über das Menü den Punkt "Projekt neu" gewählt hat, aber
        'bei der Rückfrage, ob das aktuelle Projekt gespeichert werden soll den Button "Abbruch" wählt, wird kein neues Projekt eingerichtet.
        If Projekt.bStatus_Projekt Then
            'Es ist gerade ein Projekt geöffnet. Dieses muss geschlossen werden (und evtl. gespeichert), bevor ein neues erzeugt werden kann.
            Dim resultMsg As MsgBoxResult = MsgBox("Das Projekt " + stProjekt_Name + " ist noch geöffnet. Soll das Projekt gespeichert werden?", _
                                                MsgBoxStyle.YesNoCancel, "Projekt")
            If resultMsg = MsgBoxResult.Yes Then
                Projektdaten_speichern()
            End If
            If resultMsg = MsgBoxResult.Cancel Then
                bCancel = True
            Else    'Das Projekt soll beendet/geschlossen -> die aktuellen Arrays und Variablen müssen zurückgesetzt werden.
                'Loesche_Projektdaten()
            End If
        End If

        If bCancel = False Then
            'Abfrage zu neuem Projekt: Projektname und Dateiverzeichnis in neuem Fenster
            'vorhandenes Projekt laden
            'If stProjekt_Pfad <> "" Then Form_Haupt.OFD_Projekt.InitialDirectory = stProjekt_Pfad
            Dim result As DialogResult = Form_Main.OFD_Projekt.ShowDialog()
            If result = DialogResult.OK Then
                Dim tmpName As String = stProjekt_Name
                Dim tmpPfad As String = stProjekt_Pfad
                If Projekt.bStatus_Projekt Then Loesche_Projektdaten()
                stProjekt_Name = tmpName
                stProjekt_Pfad = tmpPfad
                Form_Main.Text = "Schallschutzausweis 7Label 2018 - " + stProjekt_Name
                Projektdaten_einlesen()
                Form_Main.Update_Anzeige()
            End If
        End If
        ' Form_Main.Update_Anzeige()
    End Sub

    Public Sub Projekt_Excel_laden()
        Dim result As DialogResult = Form_Main.OFD_XLS.ShowDialog()
        If result = DialogResult.OK Then
            Dim tmpName As String = stProjekt_Name
            Dim tmpPfad As String = stProjekt_Pfad
            If Projekt.bStatus_Projekt Then Loesche_Projektdaten()
            stProjekt_Name = tmpName
            stProjekt_Pfad = tmpPfad
            Form_Main.Text = "Schallschutzausweis 7Label 2018 - " + stProjekt_Name
            Projektdaten_XLS_einlesen(Form_Main.OFD_XLS.FileName)
            Form_Main.Update_Anzeige()
        End If
    End Sub

    Public Sub Grundriss_hinzufuegen()
        With Form_Main
            If .OFD_Grundriss.ShowDialog = System.Windows.Forms.DialogResult.OK Then

                Dim sPfadGrundriss As String = .OFD_Grundriss.FileName

                'Form_Haupt auf Bildschirmgröße maximieren
                If Not IsNothing(.PB_Grundriss.BackgroundImage) Then
                    .PB_Grundriss.BackgroundImage.Dispose()
                    .PB_Grundriss.BackgroundImage = Nothing
                    .PB_DB_Grundriss.BackgroundImage.Dispose()
                    .PB_DB_Grundriss.BackgroundImage = Nothing
                    btmGrundriss = Nothing
                End If

                'Lageplan einlesen ohne Skalierung
                btmGrundriss = New Bitmap(System.Drawing.Image.FromFile(sPfadGrundriss))
                .PB_Grundriss.BackgroundImage = btmGrundriss 'System.Drawing.Image.FromFile(sPfadGrundriss)
                .PB_DB_Grundriss.BackgroundImage = btmGrundriss

                'Bild/Lageplan drehen? b_IM_gedreht könnte auch direkt in Form_Image_drehen initialisiert werden.
                Dim frm_Image_drehen As New Form_Image_drehen
                frm_Image_drehen.ShowDialog()
                ' Form_Main.Update_SSA_Deckblatt()
            End If
        End With

    End Sub

End Module
