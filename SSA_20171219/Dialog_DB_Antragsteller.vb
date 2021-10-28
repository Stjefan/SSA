Imports System.Windows.Forms

Public Class Dialog_DB_Antragsteller

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Me.ListView1.SelectedItems.Count > 0 Then
            Dim lvi As ListViewItem = Me.ListView1.SelectedItems(0)

            Form_Main.TB_Antragsteller_Name.Text = lvi.SubItems(0).Text
            Form_Main.TB_Antragsteller_Zusatz.Text = lvi.SubItems(1).Text
            Form_Main.TB_Antragsteller_Strasse.Text = lvi.SubItems(2).Text
            Form_Main.TB_Antragsteller_Nr.Text = lvi.SubItems(3).Text
            Form_Main.TB_Antragsteller_PLZ.Text = lvi.SubItems(4).Text
            Form_Main.TB_Antragsteller_Ort.Text = lvi.SubItems(5).Text
        End If

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub Dialog_DB_Antragsteller_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Load_DB_Antragsteller()
        
    End Sub
    Private Sub Load_DB_Antragsteller()
        Me.ListView1.Items.Clear()
        If Not IsNothing(DB_Antragsteller) Then
            For i As Integer = 0 To DB_Antragsteller.Count - 1 '30
                Dim lvi As ListViewItem = Me.ListView1.Items.Add(CType(DB_Antragsteller(i), AdresseData).Name)
                lvi.SubItems.Add(CType(DB_Antragsteller(i), AdresseData).Zusatz)
                lvi.SubItems.Add(CType(DB_Antragsteller(i), AdresseData).Strasse)
                lvi.SubItems.Add(CType(DB_Antragsteller(i), AdresseData).Nr)
                If CType(DB_Antragsteller(i), AdresseData).PLZ = "" Then
                    lvi.SubItems.Add("")
                Else
                    lvi.SubItems.Add(CType(DB_Antragsteller(i), AdresseData).PLZ)
                End If
                lvi.SubItems.Add(CType(DB_Antragsteller(i), AdresseData).Ort)
                Dim iIt As Integer = Me.ListView1.Items.Count
            Next
        End If
    End Sub
    Private Sub ListView1_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.DoubleClick
        If Me.ListView1.SelectedItems.Count > 0 Then
            Dim lvi As ListViewItem = Me.ListView1.SelectedItems(0)

            Form_Main.TB_Antragsteller_Name.Text = lvi.SubItems(0).Text
            Form_Main.TB_Antragsteller_Zusatz.Text = lvi.SubItems(1).Text
            Form_Main.TB_Antragsteller_Strasse.Text = lvi.SubItems(2).Text
            Form_Main.TB_Antragsteller_Nr.Text = lvi.SubItems(3).Text
            Form_Main.TB_Antragsteller_PLZ.Text = lvi.SubItems(4).Text
            Form_Main.TB_Antragsteller_Ort.Text = lvi.SubItems(5).Text
        End If
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Button_Entfernen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Entfernen.Click
        If Me.ListView1.SelectedItems.Count > 0 Then
            Dim lvi As ListViewItem = Me.ListView1.SelectedItems(0)

            Dim iEl As Integer = -1
            For i As Integer = 0 To 30
                With CType(DB_Antragsteller(i), AdresseData)
                    If lvi.SubItems(0).Text = .Name And lvi.SubItems(1).Text = .Zusatz And lvi.SubItems(2).Text = .Strasse And _
                                        lvi.SubItems(3).Text = .Nr And (lvi.SubItems(4).Text = CStr(.PLZ) Or lvi.SubItems(4).Text = "" And .PLZ = "") And lvi.SubItems(5).Text = .Ort Then
                        iEl = i
                        Exit For
                    End If
                End With
            Next


            If iEl > -1 Then
                DB_Antragsteller.RemoveAt(iEl)
                'For j As Integer = iEl To 29
                '    Dim tAdresse As AdresseData
                '    With tAdresse 'CType(My.Settings.DB_Antragsteller(j), AdresseData)
                '        .Name = CType(My.Settings.DB_Antragsteller_SET(j + 1), AdresseData).Name
                '        .Zusatz = CType(My.Settings.DB_Antragsteller_SET(j + 1), AdresseData).Zusatz
                '        .Strasse = CType(My.Settings.DB_Antragsteller_SET(j + 1), AdresseData).Strasse
                '        .Nr = CType(My.Settings.DB_Antragsteller_SET(j + 1), AdresseData).Nr
                '        .PLZ = CType(My.Settings.DB_Antragsteller_SET(j + 1), AdresseData).PLZ
                '        .Ort = CType(My.Settings.DB_Antragsteller_SET(j + 1), AdresseData).Ort
                '    End With
                '    My.Settings.DB_Antragsteller_SET(j) = tAdresse
                'Next

                'Dim newAdresse As AdresseData
                'My.Settings.DB_Antragsteller_SET(30) = newAdresse
                'With DB_Antragsteller(30)
                '    .Name = ""
                '    .Zusatz = ""
                '    .Strasse = ""
                '    .Nr = ""
                '    .PLZ = ""
                '    .Ort = ""
                'End With

                Load_DB_Antragsteller()
            End If
        End If
    End Sub
End Class
