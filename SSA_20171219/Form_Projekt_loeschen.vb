Option Explicit On
Option Strict On

'Imports System.IO

Public Class Form_Projekt_loeschen

    Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
        MyBase.OnLoad(e)

        'Form zu Projekt löschen öffnen
    End Sub

    Private Sub TB_Projektverzeichnis_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TB_Projektverzeichnis.Click

        'Verzeichnis auswählen

        FBD_Projekt_loeschen.ShowDialog()

        TB_Projektverzeichnis.Text = Trim(FBD_Projekt_loeschen.SelectedPath)

    End Sub

    Private Sub BT_Verzeichnis_durchsuchen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BT_Verzeichnis_durchsuchen.Click

        'Verzeichnis auswählen

        FBD_Projekt_loeschen.ShowDialog()

        TB_Projektverzeichnis.Text = Trim(FBD_Projekt_loeschen.SelectedPath)
    End Sub

    Private Sub BT_Ok_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BT_Delete.Click
        Dim stProjekt_Pfad_Delete As String
        stProjekt_Pfad_Delete = TB_Projektverzeichnis.Text  'Trim(FBD_Projekt_loeschen.SelectedPath)
        Dim dir As New IO.DirectoryInfo(stProjekt_Pfad_Delete)

        If dir.Exists Then  'Len(stProjekt_Pfad_Delete) > 0 Then

            My.Computer.FileSystem.DeleteDirectory(stProjekt_Pfad_Delete, FileIO.UIOption.AllDialogs, FileIO.RecycleOption.SendToRecycleBin)
        Else
            MsgBox("Der angegebene Pfad existiert nicht.", MsgBoxStyle.OkOnly, "Pfadangabe")
        End If
        Me.Close()

    End Sub


    Private Sub BT_Abbruch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BT_Abbruch.Click

        'stProjekt_Name = ""
        'stProjekt_Pfad = ""

        Me.Close()
    End Sub



End Class

