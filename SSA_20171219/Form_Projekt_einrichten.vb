Option Explicit On
Option Strict On

'Imports System.IO
Imports Microsoft.VisualBasic

Public Class Form_Projekt_einrichten

    Dim tmpProjekt_Pfad As String
    Dim tmpProjekt_Name As String

    'Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
    '    MyBase.OnLoad(e)

    '    'Form2 zu Datenabfrage für neues Projekt öffnen


    'End Sub

    Private Sub TB_Projektverzeichnis_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TB_Projektverzeichnis.Click
        ' Beim Klick in die Textbox "Dateiverzeichnis für Projekt" wird der FolderBrowserDialog "FBD_Projekt_einrichten" erzeugt und angezeigt.
        ' Der Anwender wählt aus dem Verzeichnisbaum den Ordner, in den das neue Projekt gespeichert werden soll. In die Textbox wird
        ' der absolute Pfad angezeigt; - der Projektname wird als Dateiname übernommen.

        'Verzeichnis auswählen

        FBD_Projekt_einrichten.ShowDialog()

        FBD_Projekt_einrichten.SelectedPath = Trim(FBD_Projekt_einrichten.SelectedPath)

        tmpProjekt_Pfad = FBD_Projekt_einrichten.SelectedPath

        ' tmpProjekt_Pfad = tmpProjekt_Pfad & "\" & tmpProjekt_Name

        TB_Projektverzeichnis.Text = Trim(tmpProjekt_Pfad) '& "\" & tmpProjekt_Name)  'tmpProjekt_Pfad)

    End Sub

    Private Sub BT_Verzeichnis_durchsuchen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BT_Verzeichnis_durchsuchen.Click
        ' Beim Klick auf den Button "BT_Verzeichnis_durchsuchen" wird der FolderBrowserDialog "FBD_Projekt_einrichten" erzeugt und angezeigt.
        ' Der Code der Sub entspricht der Sub "TB_Projektverzeichnis_Click" -> Beschreibung siehe oben.

        'Verzeichnis auswählen

        FBD_Projekt_einrichten.ShowDialog()

        FBD_Projekt_einrichten.SelectedPath = Trim(FBD_Projekt_einrichten.SelectedPath)

        tmpProjekt_Pfad = FBD_Projekt_einrichten.SelectedPath

        'tmpProjekt_Pfad = tmpProjekt_Pfad & "\" & tmpProjekt_Name

        TB_Projektverzeichnis.Text = Trim(tmpProjekt_Pfad) ' & "\" & tmpProjekt_Name)  'tmpProjekt_Pfad)

    End Sub

    Private Sub TB_Projektname_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TB_Projektname.Leave
        ' Der vom Anwender angegebene Name wird beim verlassen der Textbox in der globalen Variablen 
        ' tmpProjekt_Name abgelegt.

        'Name neues Projekt übernehmen

        tmpProjekt_Name = Trim(TB_Projektname.Text)

    End Sub


    Private Sub BT_Ok_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BT_Ok.Click
        'Überprüfung, ob das angegebene Projekt schon existiert. Falls ja, nachfagen, ob es überschrieben werden soll, oder nicht.
        If tmpProjekt_Pfad <> "" Then
            tmpProjekt_Pfad = TB_Projektverzeichnis.Text
            Dim fil As New IO.FileInfo(tmpProjekt_Pfad & "\" & tmpProjekt_Name & "\" & tmpProjekt_Name & ".PRJ")
            Dim bSpeichern As Boolean = False

            If fil.Exists Then
                Dim resultMsg As MsgBoxResult = MsgBox(tmpProjekt_Pfad & "\" & tmpProjekt_Name & "\" & tmpProjekt_Name & _
                                                ".PRJ besteht bereits. Soll dieses Projekt ersetzt werden?", MsgBoxStyle.YesNo, "Projekt")
                If resultMsg = MsgBoxResult.Yes Then
                    bSpeichern = True
                End If
            Else

                My.Computer.FileSystem.CreateDirectory(tmpProjekt_Pfad & "\" & tmpProjekt_Name)
                'DATEN-Verzeichnis und *.DAT-Dateien erstellen
                My.Computer.FileSystem.CreateDirectory(tmpProjekt_Pfad & "\" & tmpProjekt_Name & "\Daten")

                bSpeichern = True
            End If

            If bSpeichern Then

                stProjekt_Pfad = tmpProjekt_Pfad & "\" & tmpProjekt_Name
                stProjekt_Name = tmpProjekt_Name

                Projekt.bStatus_Projekt = True
                ' Beim Klick auf den Button "Ok" wird diese Form ("Form_Projekt_einrichten") geschlossen.
                Me.Close()
            End If
        Else
            MsgBox("Das Projekt kann nicht angelegt werden. Bitte geben Sie einen gültigen Projektpfad ein!", MsgBoxStyle.OkOnly, "Eingabefehler")
        End If
    End Sub


    Private Sub BT_Abbruch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BT_Abbruch.Click
        ' Beim Klick auf den Button "Abbruch" werden den globalen Variablen stProjekt_Name und stProjektPfad
        ' ein leerer String übergeben und diese Form ("Form_Projekt_einrichten") geschlossen.
        'stProjekt_Name = ""
        'stProjekt_Pfad = ""

        Me.Close()


    End Sub


    Private Sub Form_Projekt_einrichten_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        tmpProjekt_Pfad = TB_Projektverzeichnis.Text
    End Sub
End Class

