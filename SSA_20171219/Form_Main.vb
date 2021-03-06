Public Class Form_Main
    <VBFixedString(100)> Private tStr As String = ""
    Private iPage As Integer
    Private wPage As Integer
    Private hPage As Integer
    Private xPage As Integer
    Private yPage As Integer
    Private bUpdate_Anzeige As Boolean
    'Private bMove As Boolean
    'Private ptMouse As System.Drawing.Point
    <VBFixedString(55)> Dim Leer555 As String = ""

    Private Sub Form_Main_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'DB_Antragsteller_speichern()
        My.Settings.DB_Antragsteller_SET = GetStringFromArraylist(DB_Antragsteller)
        My.Settings.Save()

        ' Dim iAnt As Integer = My.Settings.DB_Antragsteller_SET.Count

        If Projekt.bStatus_Projekt Then
            If MsgBox("Soll das Projekt gespeichert werden?", MsgBoxStyle.YesNo, "Speichern") = MsgBoxResult.Yes Then
                If stProjekt_Pfad <> "" And stProjekt_Name <> "" Then
                    Projektdaten_speichern()
                Else
                    Projektdaten_speichernUnter()
                End If
            End If
        End If
    End Sub

    Private Sub ToolStripMenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem7.Click
        'My.Settings.ShowSplash = Not My.Settings.ShowSplash
    End Sub
    Private Sub TB_Datum_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TB_Datum.KeyPress
        Dim i As Integer = AscW(e.KeyChar)
        If i = 11 Then
            My.Settings.ShowSplash = Not My.Settings.ShowSplash
            My.Settings.Save()
        End If
    End Sub
    Private Sub Form_Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Dim tImg As Image = Form_Begruessung.MainLayoutPanel.BackgroundImage 'Me.Panel_eSSA.BackgroundImage
        'tImg.Save("\\Server-kf\kuf\Auftrag\10\105\1050\10500\zo Bearbeitung\Begruessung.png")

        LizenznehmerSW = LizenznehmerPR.Replace("&", "&&")
        ' Dim iAnt As Integer = My.Settings.DB_Antragsteller.Count

        '   Panel_SSA.BackgroundImage = My.Resources.Vorlage_Eingabeblatt_FullVers
        '  AddHandler Form_Main.Button_GC_W.click, AddressOf Button_GC_Click
        If btDemo > versNormal Then Form_Begruessung.Version.Text = "DEMO-Version 20140915"

        Dim res As System.Windows.Forms.DialogResult = System.Windows.Forms.DialogResult.OK
        If My.Settings.ShowSplash Then res = Form_Begruessung.ShowDialog


        If Not IsNothing(My.Settings.DB_Antragsteller_SET) Then
            DB_Antragsteller = GetArraylistFromString(My.Settings.DB_Antragsteller_SET) '.AddRange(GetArraylistFromString(My.Settings.DB_Antragsteller_SET))
        End If

        If res = System.Windows.Forms.DialogResult.OK Then 'Form_Begruessung.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Me.TB_Datum.Text = Now.Day & "." & Now.Month & "." & Now.Year
            Me.Label_gueltigBis.Text = IIf(Now.Day < 9, "0" & Now.Day, Now.Day).ToString & "." & IIf(Now.Month < 9, "0" & Now.Month, Now.Month).ToString & "." & Now.AddYears(10).Year
            'Konfig_einlesen()

            Projekt_Default()
            Update_Anzeige()

            Me.Update_SSA()

            Dim ctrl As Control = Me.TabPage1
            Add_NUD_Events(ctrl)

            Me.TB_Datum.Select()

            If btDemo = versDEMO Then
                Me.TSMI_Druckvorschau.Enabled = False
                Me.TSMI_SeiteEinrichten.Enabled = False
                Me.TSMI_SchallschutzausweisDrucken.Enabled = False
                Me.TSMI_DEGA.Enabled = False
                Me.TSMI_DirekteHilfe.Enabled = False
                Me.TSMI_Handbuch.Enabled = False
                'Me.TSMI_Datenuebertragung.Enabled = False
                Me.TSMI_ExcelProjektLaden.Enabled = False
                Me.Untermenue_Projekt_Speichern.Enabled = False
                Me.Untermenue_Projekt_SpeichernUnter.Enabled = False
                Me.Untermenue_Projekt_Neu.Enabled = False
                Me.Untermenue_Projekt_Schliessen.Enabled = False
                Me.Untermenue_Projekt_Laden.Enabled = False
                Me.Untermenue_Projekt_Entfernen.Enabled = False

                Me.TSB_Drucken.Enabled = False
                Me.TSB_Speichern.Enabled = False
                Me.TSB_Neu.Enabled = False
                Me.TSB_Oeffnen.Enabled = False
                Me.TSB_DEGA.Enabled = False
                Me.TSB_Hilfe.Enabled = False
            End If
        Else
            Me.Close()
        End If

        'LEER100 = ""
        'LEER512 = ""
        'LEER2 = ""
        'LEER5 = ""
        'LEER10 = ""

    End Sub
    Private Sub Add_NUD_Events(ByVal ctrl As Control)
        For Each cCtrl As Control In ctrl.Controls
            If TypeOf (cCtrl) Is NumericUpDown Then
                AddHandler cCtrl.MouseUp, AddressOf NUD_MouseUp
                AddHandler cCtrl.MouseWheel, AddressOf NUD_MouseWheel
                AddHandler cCtrl.KeyUp, AddressOf nud_KeyUp
                'AddHandler CType(cCtrl, NumericUpDown).valuechangd, AddressOf NUD_TextChangd
            End If

            If Not IsNothing(cCtrl.Controls) Then Add_NUD_Events(cCtrl)
        Next
    End Sub
    Private Sub nud_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim tmpNUD As NumericUpDown = CType(sender, NumericUpDown)
        If e.KeyCode = Keys.Tab Then tmpNUD.Select(0, tmpNUD.Text.Length)
    End Sub
    Private Sub NUD_MouseWheel(ByVal Sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Me.TabPage1.Focus()
    End Sub
    Private Sub NUD_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Dim tmpNUD As NumericUpDown = CType(sender, NumericUpDown)
        If e.Button = System.Windows.Forms.MouseButtons.Left Then tmpNUD.Select(0, tmpNUD.Text.Length)
    End Sub

    Private Sub OFD_Projekt_FileOk(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OFD_Projekt.FileOk

        OFD_Projekt.FileName = Trim(OFD_Projekt.FileName)

        Dim shA As Short
        Dim shB As Short

        shA = CShort(Len(OFD_Projekt.FileName))

        For shB = shA To 1 Step -1

            If Mid(Trim(OFD_Projekt.FileName), shB, 1) = "\" Then
                stProjekt_Pfad = Microsoft.VisualBasic.Left(OFD_Projekt.FileName, shB - 1)
                stProjekt_Name = Microsoft.VisualBasic.Mid(OFD_Projekt.FileName, shB + 1, shA - shB - 4)
                Exit For
            End If

        Next shB

    End Sub

    Public Sub Projekt_Default()
        'With Projekt
        btmGrundriss = Nothing

        '.Grundriss = ""

        ''Status
        '.bStatus_Grundriss = False

        ''Rubrik I
        '.Antragsteller.Name = ""
        '.Antragsteller.Zusatz = ""
        '.Antragsteller.Strasse = ""
        '.Antragsteller.Nr = ""
        '.Antragsteller.PLZ = ""
        '.Antragsteller.Ort = ""
        'End With
        'With Projekt.Gebaeude.Adresse
        '    .Name = ""
        '    .Zusatz = ""
        '    .Strasse = ""
        '    .Nr = ""
        '    .PLZ = ""
        '    .Ort = ""
        'End With
        'With Projekt.Gebaeude
        '    .Baujahr = 0
        '    .Gebaeudetyp = ""
        '    .Kosten = 0
        '    .Wohneinheiten = ""
        'End With
        'With Projekt.Wohnung
        '    .Geschoss.Bezeichnung = ""
        '    .Geschoss.Nr = 0
        '    .Geschoss.Typ = 0
        'End With
        'Projekt.Datum = ""

        ''Rubrik II
        'With Projekt.Standort
        '    .abgewandFreibereich = 0
        '    .Aussenlaermpegel = 0
        '    .Gebietscharakter = 0
        'End With
        ''Rubrik III
        'Projekt.LS_Wand.Untersuchung = 0
        'With Projekt.LS_Wand.Messung
        '    .Anzahl = 0
        '    '.Messanteil = 0
        '    '.Messung_1.Messverfahren = 0
        '    .Messung_1.Pegel = 0
        '    .Messung_1.C = ""
        '    .Messung_2.Pegel = 0
        '    .Messung_2.C = ""
        '    .Messung_3.Pegel = 0
        '    .Messung_3.C = ""
        '    .Messung_4.Pegel = 0
        '    .Messung_4.C = ""
        '    .Messung_5.Pegel = 0
        '    .Messung_5.C = ""
        '    .Messung_6.Pegel = 0
        '    .Messung_6.C = ""
        'End With
        'With Projekt.LS_Wand.Prognose
        '    .Pegel = 0
        '    .C = ""
        'End With

        'Projekt.LS_Decke.Untersuchung = 0
        'With Projekt.LS_Decke.Messung
        '    .Anzahl = 0
        '    .Messung_1.Pegel = 0
        '    .Messung_1.C = ""
        '    .Messung_2.Pegel = 0
        '    .Messung_2.C = ""
        '    .Messung_3.Pegel = 0
        '    .Messung_3.C = ""
        '    .Messung_4.Pegel = 0
        '    .Messung_4.C = ""
        '    .Messung_5.Pegel = 0
        '    .Messung_5.C = ""
        '    .Messung_6.Pegel = 0
        '    .Messung_6.C = ""
        'End With
        'With Projekt.LS_Decke.Prognose
        '    .Pegel = 0
        '    .C = ""
        'End With

        'Projekt.TS_Decke.Untersuchung = 0
        'With Projekt.TS_Decke.Messung
        '    .Anzahl = 0
        '    .Messung_1.Pegel.Pegel = 0
        '    .Messung_1.Pegel.C = ""
        '    .Messung_2.Pegel.Pegel = 0
        '    .Messung_2.Pegel.C = ""
        '    .Messung_3.Pegel.Pegel = 0
        '    .Messung_3.Pegel.C = ""
        '    .Messung_4.Pegel.Pegel = 0
        '    .Messung_4.Pegel.C = ""
        '    .Messung_5.Pegel.Pegel = 0
        '    .Messung_5.Pegel.C = ""
        '    .Messung_6.Pegel.Pegel = 0
        '    .Messung_6.Pegel.C = ""
        'End With
        'With Projekt.TS_Decke.Prognose.Pegel
        '    .Pegel = 0
        '    .C = ""
        'End With

        'Projekt.TS_TPLH.Untersuchung = 0
        'With Projekt.TS_TPLH.Messung
        '    .Anzahl = 0
        '    .Messung_1.Pegel.Pegel = 0
        '    .Messung_1.Pegel.C = ""
        '    .Messung_2.Pegel.Pegel = 0
        '    .Messung_2.Pegel.C = ""
        '    .Messung_3.Pegel.Pegel = 0
        '    .Messung_3.Pegel.C = ""
        '    .Messung_4.Pegel.Pegel = 0
        '    .Messung_4.Pegel.C = ""
        '    .Messung_5.Pegel.Pegel = 0
        '    .Messung_5.Pegel.C = ""
        '    .Messung_6.Pegel.Pegel = 0
        '    .Messung_6.Pegel.C = ""
        'End With
        'With Projekt.TS_TPlH.Prognose
        '    .Treppe.Pegel.Pegel = 0
        '    .Treppe.Pegel.C = ""
        '    .Podest.Pegel.Pegel = 0
        '    .Podest.Pegel.C = ""
        '    .Laube.Pegel.Pegel = 0
        '    .Laube.Pegel.C = ""
        '    .Hausflur.Pegel.Pegel = 0
        '    .Hausflur.Pegel.C = ""
        'End With

        'Projekt.TS_BLT.Untersuchung = 0
        'With Projekt.TS_BLT.Messung
        '    .Anzahl = 0
        '    .Messung_1.Pegel.Pegel = 0
        '    .Messung_1.Pegel.C = ""
        '    .Messung_2.Pegel.Pegel = 0
        '    .Messung_2.Pegel.C = ""
        '    .Messung_3.Pegel.Pegel = 0
        '    .Messung_3.Pegel.C = ""
        '    .Messung_4.Pegel.Pegel = 0
        '    .Messung_4.Pegel.C = ""
        '    .Messung_5.Pegel.Pegel = 0
        '    .Messung_5.Pegel.C = ""
        '    .Messung_6.Pegel.Pegel = 0
        '    .Messung_6.Pegel.C = ""
        'End With
        'With Projekt.TS_BLT.Prognose
        '    .Balkon.Pegel.Pegel = 0
        '    .Balkon.Pegel.C = ""
        '    .Loggia.Pegel.Pegel = 0
        '    .Loggia.Pegel.C = ""
        '    .Terrasse.Pegel.Pegel = 0
        '    .Terrasse.Pegel.C = ""
        'End With

        'With Projekt.Tueren
        '    .L = 0
        '    .Ort = 0
        '    .Untersuchung = 0
        'End With

        'Projekt.Aussenbauteile = 0

        'With Projekt.Wasser
        '    .L_Intervall = 0
        '    .LcLa_erfuellt = 0
        '    .Untersuchung = 0
        'End With

        ''With Projekt.Nutzer
        ''    .L_Intervall = 0
        ''    .Untersuchung = 0
        ''End With
        ''With Projekt.Koerper
        ''    .Untersuchung = 0
        ''    .L_Intervall = 0
        ''End With
        'With Projekt.NutzerKoerper
        '    .BemerkungKoerper = ""
        '    .BemerkungNutzer = ""

        '    .L_Intervall = 0
        '    .MessungPrognose = 0
        '    .Untersuchung = 0
        'End With
        'Projekt.Nachbarn = 0
        'Projekt.anordnungRaeume = 0
        'Projekt.lauteRaeume = 0
        Projekt.Name = ""
        Projekt.Pfad = ""
        Projekt.Grundriss = ""

        Projekt.bStatus_Grundriss = False
        Projekt.bStatus_Projekt = False


        With Projekt
            .Antragsteller.Name = ""
            .Antragsteller.Zusatz = ""
            .Antragsteller.Strasse = ""
            .Antragsteller.Nr = ""
            .Antragsteller.PLZ = ""
            .Antragsteller.Ort = ""

            .Gebaeude.Adresse.Name = ""
            .Gebaeude.Adresse.Zusatz = ""
            .Gebaeude.Adresse.Strasse = ""
            .Gebaeude.Adresse.Nr = ""
            .Gebaeude.Adresse.PLZ = ""
            .Gebaeude.Adresse.Ort = ""
            .Gebaeude.Gebaeudetyp = ""
            .Gebaeude.Baujahr = 0
            .Gebaeude.Wohneinheiten = ""
            .Gebaeude.Kosten = 0

            .Wohnung.Wohnungsbezeichnung = ""
            .Wohnung.Geschoss.Nr = 0
            .Wohnung.Geschoss.Typ = 0
            .Wohnung.Geschoss.Bezeichnung = ""
            .Wohnung.Raeume = ""
            .Wohnung.Wohnflaeche = 1
        End With
        With Projekt.Standort
            'Standort
            .Gebietscharakter = 0
            .Bem_Gebietscharakter = ""
            'Außenlärm
            .Aussenlaermpegel = 0
            .Bem_Aussenlaermpegel = ""
            'Freibereich
            .abgewandFreibereich = 0
        End With
        'LS Wände
        With Projekt.LS_Wand
            .Untersuchung = 0
            .Prognose.Pegel = 0
            .Prognose.C = ""
            '.Bemerkung_LS = ""
        End With
        With Projekt.LS_Wand.Messung

            .Messung_1.C = ""
            .Messung_1.Pegel = 0

            .Messung_2.C = ""
            .Messung_2.Pegel = 0

            .Messung_3.C = ""
            .Messung_3.Pegel = 0

            .Messung_4.C = ""
            .Messung_4.Pegel = 0

            .Messung_5.C = ""
            .Messung_5.Pegel = 0

            .Messung_6.C = ""
            .Messung_6.Pegel = 0
        End With

        'LS Decken
        With Projekt.LS_Decke
            .Untersuchung = 0
            .Prognose.Pegel = 0
            .Prognose.C = ""
            '.Bemerkung_LS = ""
        End With
        With Projekt.LS_Decke.Messung

            .Messung_1.C = ""
            .Messung_1.Pegel = 0

            .Messung_2.C = ""
            .Messung_2.Pegel = 0

            .Messung_3.C = ""
            .Messung_3.Pegel = 0

            .Messung_4.C = ""
            .Messung_4.Pegel = 0

            .Messung_5.C = ""
            .Messung_5.Pegel = 0

            .Messung_6.C = ""
            .Messung_6.Pegel = 0
        End With

        'TS Decken
        With Projekt.TS_Decke
            .Untersuchung = 0
            .fEstrich = 0

            .Prognose.Bodenbelag = 0
            .Prognose.Pegel.Pegel = 0
            .Prognose.Pegel.C = ""
            '.Bemerkung_TS_Decke = ""
        End With
        With Projekt.TS_Decke.Messung
            .Messung_1.Bodenbelag = 0
            .Messung_1.Pegel.C = ""
            .Messung_1.Pegel.Pegel = 0

            .Messung_2.Bodenbelag = 0
            .Messung_2.Pegel.C = ""
            .Messung_2.Pegel.Pegel = 0

            .Messung_3.Bodenbelag = 0
            .Messung_3.Pegel.C = ""
            .Messung_3.Pegel.Pegel = 0

            .Messung_4.Bodenbelag = 0
            .Messung_4.Pegel.C = ""
            .Messung_4.Pegel.Pegel = 0

            .Messung_5.Bodenbelag = 0
            .Messung_5.Pegel.C = ""
            .Messung_5.Pegel.Pegel = 0

            .Messung_6.Bodenbelag = 0
            .Messung_6.Pegel.C = ""
            .Messung_6.Pegel.Pegel = 0
        End With

        'TS TPH
        With Projekt.TS_TPLH
            .Untersuchung = 0

            .Prognose.Treppe.Bodenbelag = 0
            .Prognose.Treppe.Pegel.Pegel = 0
            .Prognose.Treppe.Pegel.C = ""
            .Prognose.Podest.Bodenbelag = 0
            .Prognose.Podest.Pegel.Pegel = 0
            .Prognose.Podest.Pegel.C = ""
            .Prognose.Laube.Bodenbelag = 0
            .Prognose.Laube.Pegel.Pegel = 0
            .Prognose.Laube.Pegel.C = ""
            .Prognose.Hausflur.Bodenbelag = 0
            .Prognose.Hausflur.Pegel.Pegel = 0
            .Prognose.Hausflur.Pegel.C = ""
            '.Bemerkung_TS_TPH = ""
        End With
        With Projekt.TS_TPLH.Messung
            .Anzahl = 0

            .Messung_1.Bodenbelag = 0
            .Messung_1.Pegel.C = ""
            .Messung_1.Pegel.Pegel = 0

            .Messung_2.Bodenbelag = 0
            .Messung_2.Pegel.C = ""
            .Messung_2.Pegel.Pegel = 0

            .Messung_3.Bodenbelag = 0
            .Messung_3.Pegel.C = ""
            .Messung_3.Pegel.Pegel = 0

            .Messung_4.Bodenbelag = 0
            .Messung_4.Pegel.C = ""
            .Messung_4.Pegel.Pegel = 0

            .Messung_5.Bodenbelag = 0
            .Messung_5.Pegel.C = ""
            .Messung_5.Pegel.Pegel = 0

            .Messung_6.Bodenbelag = 0
            .Messung_6.Pegel.C = ""
            .Messung_6.Pegel.Pegel = 0
        End With
        'TS BLLT
        With Projekt.TS_BLT
            .Untersuchung = 0
            '.Bemerkung_TS_BLLT=""
            .Prognose.Balkon.Bodenbelag = 0
            .Prognose.Balkon.Pegel.Pegel = 0
            .Prognose.Balkon.Pegel.C = ""
            .Prognose.Loggia.Bodenbelag = 0
            .Prognose.Loggia.Pegel.Pegel = 0
            .Prognose.Loggia.Pegel.C = ""
            .Prognose.Terrasse.Bodenbelag = 0
            .Prognose.Terrasse.Pegel.Pegel = 0
            .Prognose.Terrasse.Pegel.C = ""
            '.Bemerkung_TS_BLLT = ""
        End With
        With Projekt.TS_BLT.Messung
            .Anzahl = 0

            .Messung_1.Bodenbelag = 0
            .Messung_1.Pegel.C = ""
            .Messung_1.Pegel.Pegel = 0

            .Messung_2.Bodenbelag = 0
            .Messung_2.Pegel.C = ""
            .Messung_2.Pegel.Pegel = 0

            .Messung_3.Bodenbelag = 0
            .Messung_3.Pegel.C = ""
            .Messung_3.Pegel.Pegel = 0

            .Messung_4.Bodenbelag = 0
            .Messung_4.Pegel.C = ""
            .Messung_4.Pegel.Pegel = 0

            .Messung_5.Bodenbelag = 0
            .Messung_5.Pegel.C = ""
            .Messung_5.Pegel.Pegel = 0

            .Messung_6.Bodenbelag = 0
            .Messung_6.Pegel.C = ""
            .Messung_6.Pegel.Pegel = 0
        End With

        'Tueren
        With Projekt.Tueren
            .Ort = 0
            .Untersuchung = 0
            .L = 0
            '.Bemerkung_Tueren = ""
        End With

        'LS Außenbauteile
        Projekt.Aussenbauteile = 0
        Projekt.Bem_Aussenbauteile = ""
        'Wasser + haustechn. Anlagen
        'With Projekt.Wasser
        '    .Untersuchung = 0
        '    .L_Intervall = 0
        '    .LcLa_erfuellt = 0
        '    .Bem_Wasser = ""
        'End With
        With Projekt.Wasser
            .L_Intervall = 0
            .LcLa_erfuellt = 0
            .Untersuchung = 0
        End With

        With Projekt.NutzerKoerper
            .BemerkungKoerper = ""
            .BemerkungNutzer = ""

            .L_Intervall = 0
            .MessungPrognose = 0
            .Untersuchung = 0
        End With
        '' Nutzer, Körper
        'With Projekt.NutzerKoerper
        '    .Untersuchung = 0
        '    .MessungPrognose = 0
        '    .L_Intervall = 0
        '    .BemerkungKoerper = ""
        '    .BemerkungNutzer = ""
        '    '.Bemerkung = ""
        'End With
        
        'Nachbarn
        Projekt.Nachbarn = 0
        Projekt.Bemerkung_Nachbarn = ""
        'Anordnung lauter Räume
        Projekt.anordnungRaeume = 0
        Projekt.Bem_anordnungRaeume = ""
        'laute Räume angrenzend
        Projekt.lauteRaeume = 0
        Projekt.Bem_lauteRaeume = ""
        'eigener Wohnbereich
        Projekt.eigenerWohnbereich = 0
        'Projekt.Bemerkung_eigenerWohnbereich = ""
        Projekt.NHZ = 0

        If Not IsNothing(Me.PB_Grundriss.BackgroundImage) Then
            ' Form_Main.PB_Grundriss.BackgroundImage.Dispose()
            Me.PB_Grundriss.BackgroundImage = Nothing
        End If

    End Sub

    Public Sub Update_Anzeige()
        bUpdate_Anzeige = True
        With Projekt
            Me.Text = "Schallschutzausweis 7Label - " + stProjekt_Name

            If Not IsNothing(PB_Grundriss.BackgroundImage) Then
                PB_Grundriss.BackgroundImage.Dispose()
                PB_Grundriss.BackgroundImage = Nothing
            End If
            If Not IsNothing(Me.PB_DB_Grundriss.BackgroundImage) Then
                Me.PB_DB_Grundriss.BackgroundImage.Dispose()
                Me.PB_DB_Grundriss.BackgroundImage = Nothing
            End If

            'If Not IsNothing(btmGrundriss) Then
            Me.PB_Grundriss.BackgroundImage = btmGrundriss
            Me.PB_DB_Grundriss.BackgroundImage = btmGrundriss

            ' End If

            'Lageplan einlesen ohne Skalierung
            'Dim fil As New IO.FileInfo(stProjekt_Pfad & "\Daten\Grundriss.jpg")
            ''Ermitteln, ob die Datei "Grundriss" im Projektordner schon existiert
            'If fil.Exists Then PB_Grundriss.backgroundImage = System.Drawing.Image.FromFile(stProjekt_Pfad & "\Daten\Grundriss.jpg") 'Projekt.Grundriss)

            Me.TB_Datum.Text = .Datum

            Me.TB_Antragsteller_Name.Text = .Antragsteller.Name
            Me.TB_Antragsteller_Zusatz.Text = .Antragsteller.Zusatz
            Me.TB_Antragsteller_Strasse.Text = .Antragsteller.Strasse
            Me.TB_Antragsteller_Nr.Text = .Antragsteller.Nr
            Me.TB_Antragsteller_PLZ.Text = .Antragsteller.PLZ
            Me.TB_Antragsteller_Ort.Text = .Antragsteller.Ort

            Me.TB_Gebaeude_Name.Text = .Gebaeude.Adresse.Name
            Me.TB_Gebaeude_Zusatz.Text = .Gebaeude.Adresse.Zusatz
            Me.TB_Gebaeude_Strasse.Text = .Gebaeude.Adresse.Strasse
            Me.TB_Gebaeude_Nr.Text = .Gebaeude.Adresse.Nr
            Me.TB_Gebaeude_PLZ.Text = .Gebaeude.Adresse.PLZ
            Me.TB_Gebaeude_Ort.Text = .Gebaeude.Adresse.Ort
            Me.CB_Gebaeude_Gebaeudetyp.Text = .Gebaeude.Gebaeudetyp
            Me.TB_Gebaeude_Baujahr.Text = .Gebaeude.Baujahr.ToString
            If .Gebaeude.Baujahr = 0 Then Me.TB_Gebaeude_Baujahr.Text = ""
            Me.NUD_Gebaeude_Wohneinheiten.Text = .Gebaeude.Wohneinheiten
            Me.NUD_Gebaeude_Kosten.Value = CDec(.Gebaeude.Kosten)

            Me.TB_Wohnung_Wohnungsbezeichnung.Text = .Wohnung.Wohnungsbezeichnung
            Me.NUD_Wohnung_Geschoss.Value = .Wohnung.Geschoss.Nr
            Me.CB_Wohnung_Geschoss.SelectedIndex = .Wohnung.Geschoss.Typ - 1
            Me.TB_Wohnung_Geschoss.Text = .Wohnung.Geschoss.Bezeichnung
            Me.NUD_Wohnung_Raeume.Text = .Wohnung.Raeume
            If CDec(.Wohnung.Wohnflaeche) < Me.NUD_Wohnung_Wohnflaeche.Minimum Then
                Me.NUD_Wohnung_Wohnflaeche.Value = Me.NUD_Wohnung_Wohnflaeche.Minimum
                MsgBox("Die Wohnfläche wurde auf das Minimum von " & CStr(Me.NUD_Wohnung_Wohnflaeche.Minimum) & " geändert.", MsgBoxStyle.OkOnly, "Info")
            ElseIf CDec(.Wohnung.Wohnflaeche) > Me.NUD_Wohnung_Wohnflaeche.Maximum Then
                Me.NUD_Wohnung_Wohnflaeche.Value = Me.NUD_Wohnung_Wohnflaeche.Maximum
                MsgBox("Die Wohnfläche wurde auf das Maximum von " & Me.NUD_Wohnung_Wohnflaeche.Maximum & " geändert.", MsgBoxStyle.OkOnly, "Info")
            Else
                Me.NUD_Wohnung_Wohnflaeche.Value = CDec(.Wohnung.Wohnflaeche)
            End If

            If btDemo > versNormal Then
                Me.Label_Geb_Lizenznehmer.Text = "Demoversion"
                Me.Label_d_Lizenznehmer.Text = "Demoversion"
                Me.Label_e_Lizenznehmer.Text = "Demoversion"
            Else
                Me.Label_Geb_Lizenznehmer.Text = LizenznehmerSW
                Me.Label_d_Lizenznehmer.Text = LizenznehmerSW
                Me.Label_e_Lizenznehmer.Text = LizenznehmerSW
            End If
        End With
        With Projekt.Standort
            'Standort
            Me.TB_Bem_Gebietscharakter.Text = .Bem_Gebietscharakter
            Select Case .Gebietscharakter
                Case GC_WR
                    ' Button_GC_WR_Click(Nothing, Nothing)
                    Button_GC_WR_Click(Me.Button_GC_WR, Nothing)
                Case GC_WA
                    Button_GC_WA_Click(Me.Button_GC_WA, Nothing)
                Case GC_MIWB
                    Button_GC_MIWB_Click(Me.Button_GC_MIWB, Nothing)
                Case GC_GE
                    Button_GC_GE_Click(Me.Button_GC_GE, Nothing)
                Case GC_GI
                    Button_GC_GI_Click(Me.Button_GC_GI, Nothing)
                Case Else
                    Me.Button_GC_WR.BackColor = colBlue
                    Me.Button_GC_WA.BackColor = colBlue
                    Me.Button_GC_MIWB.BackColor = colBlue
                    Me.Button_GC_GE.BackColor = colBlue
                    Me.Button_GC_GI.BackColor = colBlue
            End Select
            Me.Update_Klasse_GC()
            'Außenlärm
            Me.TB_Bem_Aussenlaermsituation.Text = .Bem_Aussenlaermpegel
            Select Case .Aussenlaermpegel
                Case AP_bis55
                    Button_AP_bis55_Click(Me.Button_AP_bis55, Nothing)
                Case AP_56_60
                    Button_AP_56_60_Click(Me.Button_AP_56_60, Nothing)
                Case AP_61_65
                    Button_AP_61_65_Click(Me.Button_AP_61_65, Nothing)
                Case AP_66_70
                    Button_AP_66_70_Click(Me.Button_AP_66_70, Nothing)
                Case AP_71_75
                    Button_AP_71_75_Click(Me.Button_AP_71_75, Nothing)
                Case AP_76
                    Button_AP_76_Click(Me.Button_AP_76, Nothing)
                Case Else
                    Me.Button_AP_bis55.BackColor = colBlue
                    Me.Button_AP_56_60.BackColor = colBlue
                    Me.Button_AP_61_65.BackColor = colBlue
                    Me.Button_AP_66_70.BackColor = colBlue
                    Me.Button_AP_71_75.BackColor = colBlue
                    Me.Button_AP_76.BackColor = colBlue
            End Select
            Me.Update_Klasse_AP()
            'Freibereich
            Select Case .abgewandFreibereich
                Case aF_ja
                    Button_aF_ja_Click(Me.Button_aF_ja, Nothing)
                Case aF_nein
                    Button_aF_nein_Click(Me.Button_aF_nein, Nothing)
                Case Else
                    Me.Button_aF_ja.BackColor = colBlue
                    Me.Button_aF_nein.BackColor = colBlue
            End Select
        End With
        'LS Wände
        With Projekt.LS_Wand
            If .Untersuchung = Prognose Then
                'Button_Change_Color(Me.Button_LS_W_Prognose)
                Button_LS_W_Prognose_Click(Me.Button_LS_W_Prognose, Nothing)
            ElseIf .Untersuchung = Messung Then
                'Button_Change_Color(Me.Button_LS_W_Messung)
                Button_LS_W_Messung_Click(Me.Button_LS_W_Messung, Nothing)
            ElseIf .Untersuchung = nv Then
                'Button_Change_Color(Me.Button_LS_W_nv)
                Button_LS_W_nv_Click(Nothing, Nothing)
            Else
                Me.Button_LS_W_Prognose.BackColor = colBlue
                Me.Button_LS_W_Messung.BackColor = colBlue
                Me.Button_LS_W_nv.BackColor = colBlue
                Me.Panel_SSA_LS_W_Messung_b.Height = 0
                Me.Panel_SSA_LS_W_Prognose_b.Height = 0
            End If
            Me.NUD_LS_W_Prog_R.Value = CDec(.Prognose.Pegel)
            Me.NUD_LS_W_Prog_C.Text = .Prognose.C
            
            Update_Anzeige_Messung_LS_W()
        End With
        Me.Update_Klasse_LS_W()
        'LS Decken
        With Projekt.LS_Decke
            If .Untersuchung = Prognose Then
                'Button_Change_Color(Me.Button_LS_D_Prognose)
                Button_LS_D_Prognose_Click(Nothing, Nothing)
            ElseIf .Untersuchung = Messung Then
                'Button_Change_Color(Me.Button_LS_D_Messung)
                Button_LS_D_Messung_Click(Nothing, Nothing)
            ElseIf .Untersuchung = nv Then
                'Button_Change_Color(Me.Button_LS_D_nv)
                Button_LS_D_nv_Click(Nothing, Nothing)
            Else
                Me.Button_LS_D_Prognose.BackColor = colBlue
                Me.Button_LS_D_Messung.BackColor = colBlue
                Me.Button_LS_D_nv.BackColor = colBlue
                Me.Panel_SSA_LS_D_Messung_b.Height = 0
                Me.Panel_SSA_LS_D_Prognose_b.Height = 0
            End If
            Me.NUD_LS_D_Prog_R.Value = CDec(.Prognose.Pegel)
            Me.NUD_LS_D_Prog_C.Text = .Prognose.C
           
            Update_Anzeige_Messung_LS_D()
        End With
        Me.Update_Klasse_LS_D()
        'TS Decken
        With Projekt.TS_Decke
            If .Untersuchung = Prognose Then
                'Button_Change_Color(Me.Button_TS_D_Prognose)
                Button_TS_D_Prognose_Click(Nothing, Nothing)
            ElseIf .Untersuchung = Messung Then
                'Button_Change_Color(Me.Button_TS_D_Messung)
                Button_TS_D_Messung_Click(Nothing, Nothing)
            ElseIf .Untersuchung = nv Then
                'Button_Change_Color(Me.Button_TS_D_nv)
                Button_TS_D_nv_Click(Nothing, Nothing)
            Else
                Me.Button_TS_D_Prognose.BackColor = colBlue
                Me.Button_TS_D_Messung.BackColor = colBlue
                Me.Button_TS_D_nv.BackColor = colBlue
                Me.Panel_SSA_TS_D_Messung_b.Height = 0
                Me.Panel_SSA_TS_D_Prognose_b.Height = 0
            End If
            '# fEstrich
            If Projekt.TS_Decke.fEstrich = fE_kleiner50 Then
                Button_Change_Color(Me.Button_TS_D_P_kleiner)
            ElseIf Projekt.TS_Decke.fEstrich = fE_groesser50 Then
                Button_Change_Color(Me.Button_TS_D_P_groesser)
            Else
                Me.Button_TS_D_MV_groesser.BackColor = colBlue
                Me.Button_TS_D_MV_kleiner.BackColor = colBlue
            End If
            '#Bodenbelag
            If Projekt.TS_Decke.Prognose.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_D_Prog_mit)
            ElseIf Projekt.TS_Decke.Prognose.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_D_Prog_ohne)
            Else
                Me.Button_TS_D_Prog_mit.BackColor = colBlue
                Me.Button_TS_D_Prog_ohne.BackColor = colBlue
            End If
            Me.NUD_TS_D_Prog_L.Value = CDec(.Prognose.Pegel.Pegel)
            Me.NUD_TS_D_Prog_C.Text = .Prognose.Pegel.C

            Update_Anzeige_Messung_TS_D()
        End With
        Me.Update_Klasse_TS_D()
        'TS TPH
        With Projekt.TS_TPLH
            If .Untersuchung = Prognose Then
                'Button_Change_Color(Me.Button_TS_TPH_Prognose)
                Button_TS_TPH_Prognose_Click(Nothing, Nothing)
            ElseIf .Untersuchung = Messung Then
                'Button_Change_Color(Me.Button_TS_TPH_Messung)
                Button_TS_TPH_Messung_Click(Nothing, Nothing)
            ElseIf .Untersuchung = nv Then
                'Button_Change_Color(Me.Button_TS_TPH_nv)
                Button_TS_TPH_nv_Click(Nothing, Nothing)
            Else
                Me.Button_TS_TPH_Prognose.BackColor = colBlue
                Me.Button_TS_TPH_Messung.BackColor = colBlue
                Me.Button_TS_TPH_nv.BackColor = colBlue
                Me.Panel_SSA_TS_TPHf_Messung_b.Height = 0
                Me.Panel_SSA_TS_TPHf_Prognose_b.Height = 0
            End If
            '#Bodenbelag
            If Projekt.TS_TPLH.Prognose.Treppe.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_Treppe_Prog_mit)
            ElseIf Projekt.TS_TPLH.Prognose.Treppe.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_Treppe_Prog_ohne)
            Else
                Me.Button_TS_Treppe_Prog_mit.BackColor = colBlue
                Me.Button_TS_Treppe_Prog_ohne.BackColor = colBlue
            End If
            Me.NUD_TS_TPLH_Prog_T_L.Text = .Prognose.Treppe.Pegel.Pegel.ToString
            Me.NUD_TS_TPLH_Prog_T_C.Text = .Prognose.Treppe.Pegel.C
            '#Bodenbelag
            If Projekt.TS_TPLH.Prognose.Podest.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_Podest_Prog_mit)
            ElseIf Projekt.TS_TPLH.Prognose.Podest.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_Podest_Prog_ohne)
            Else
                Me.Button_TS_Podest_Prog_mit.BackColor = colBlue
                Me.Button_TS_Podest_Prog_ohne.BackColor = colBlue
            End If
            Me.NUD_TS_TPLH_Prog_P_L.Text = .Prognose.Podest.Pegel.Pegel.ToString
            Me.NUD_TS_TPLH_Prog_P_C.Text = .Prognose.Podest.Pegel.C
            '#Bodenbelag
            If Projekt.TS_TPLH.Prognose.Laube.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_Laubengang_Prog_mit)
            ElseIf Projekt.TS_TPLH.Prognose.Laube.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_Laubengang_Prog_ohne)
            Else
                Me.Button_TS_Laubengang_Prog_mit.BackColor = colBlue
                Me.Button_TS_Laubengang_Prog_ohne.BackColor = colBlue
            End If
            Me.NUD_TS_TPLH_Prog_L_L.Text = (.Prognose.Laube.Pegel.Pegel).ToString
            Me.NUD_TS_TPLH_Prog_L_C.Text = .Prognose.Laube.Pegel.C
            '#Bodenbelag
            If Projekt.TS_TPLH.Prognose.Hausflur.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_Hausflur_Prog_mit)
            ElseIf Projekt.TS_TPLH.Prognose.Hausflur.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_Hausflur_Prog_ohne)
            Else
                Me.Button_TS_Hausflur_Prog_mit.BackColor = colBlue
                Me.Button_TS_Hausflur_Prog_ohne.BackColor = colBlue
            End If
            Me.NUD_TS_TPLH_Prog_H_L.Text = (.Prognose.Hausflur.Pegel.Pegel).ToString
            Me.NUD_TS_TPLH_Prog_H_C.Text = .Prognose.Hausflur.Pegel.C

            Update_Anzeige_Messung_TS_TPH()
        End With
        Me.Update_Klasse_TS_TPLH()
        'TS BLLT
        With Projekt.TS_BLT
            If .Untersuchung = Prognose Then
                'Button_Change_Color(Me.Button_TS_BLLT_Prognose)
                Button_TS_BLLT_Prognose_Click(Nothing, Nothing)
            ElseIf .Untersuchung = Messung Then
                'Button_Change_Color(Me.Button_TS_BLLT_Messung)
                Button_TS_BLLT_Messung_Click(Nothing, Nothing)
            ElseIf .Untersuchung = nv Then
                'Button_Change_Color(Me.Button_TS_BLLT_nv)
                Button_TS_BLLT_nv_Click(Nothing, Nothing)
            Else
                Me.Button_TS_BLLT_Prognose.BackColor = colBlue
                Me.Button_TS_BLLT_Messung.BackColor = colBlue
                Me.Button_TS_BLLT_nv.BackColor = colBlue
                Me.Panel_SSA_TS_BLLT_Messung_b.Height = 0
                Me.Panel_SSA_TS_BLLT_Prognose_b.Height = 0
            End If
            '#Bodenbelag Balkon
            If Projekt.TS_BLT.Prognose.Balkon.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_Balkon_Prog_mit)
            ElseIf Projekt.TS_BLT.Prognose.Balkon.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_Balkon_Prog_ohne)
            Else
                Me.Button_TS_Balkon_Prog_mit.BackColor = colBlue
                Me.Button_TS_Balkon_Prog_ohne.BackColor = colBlue
            End If
            Me.NUD_TS_BLLT_Prog_Ba_L.Text = .Prognose.Balkon.Pegel.Pegel.ToString
            Me.NUD_TS_BLLT_Prog_Ba_C.Text = .Prognose.Balkon.Pegel.C
            '#Bodenbelag Loggia
            If Projekt.TS_BLT.Prognose.Loggia.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_Loggia_Prog_mit)
            ElseIf Projekt.TS_BLT.Prognose.Loggia.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_Loggia_Prog_ohne)
            Else
                Me.Button_TS_Loggia_Prog_mit.BackColor = colBlue
                Me.Button_TS_Loggia_Prog_ohne.BackColor = colBlue
            End If
            Me.NUD_TS_BLLT_Prog_Lo_L.Text = (.Prognose.Loggia.Pegel.Pegel).ToString
            Me.NUD_TS_BLLT_Prog_Lo_C.Text = .Prognose.Loggia.Pegel.C
            '#Bodenbelag Terrasse
            If Projekt.TS_BLT.Prognose.Terrasse.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_Terrasse_Prog_mit)
            ElseIf Projekt.TS_BLT.Prognose.Terrasse.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_Terrasse_Prog_ohne)
            Else
                Me.Button_TS_Terrasse_Prog_mit.BackColor = colBlue
                Me.Button_TS_Terrasse_Prog_ohne.BackColor = colBlue
            End If
            Me.NUD_TS_BLLT_Prog_Te_L.Text = (.Prognose.Terrasse.Pegel.Pegel).ToString
            Me.NUD_TS_BLLT_Prog_Te_C.Text = .Prognose.Terrasse.Pegel.C

            Update_Anzeige_Messung_TS_BLT()
        End With
        Me.Update_Klasse_TS_BLT()
        'Tueren
        With Projekt.Tueren
            If .Ort = Diele Then
                '                Button_Change_Color(Me.Button_Tueren_Diele)
                Button_Tueren_Diele_Click(Nothing, Nothing)
            ElseIf .Ort = Aufenthalt Then
                '               Button_Change_Color(Me.Button_Tuere_Aufenthaltsraum)
                Button_Tuere_Aufenthaltsraum_Click(Nothing, Nothing)
            ElseIf .Ort = nv Then
                'Button_Change_Color(Me.Button_Tuere_nv)
                Button_Tuere_nv_Click(Nothing, Nothing)
            Else
                Me.Button_Tueren_Diele.BackColor = colBlue
                Me.Button_Tuere_Aufenthaltsraum.BackColor = colBlue
                Me.Button_Tuere_nv.BackColor = colBlue
            End If
            If .Untersuchung = Messung Then
                Button_Tuere_Messung_Click(Nothing, Nothing)
            ElseIf .Untersuchung = Pruefzeugnis Then
                Button_Tuere_Pruefzeugnis_Click(Nothing, Nothing)
            Else
                Me.Button_Tuere_Messung.BackColor = colBlue
                Me.Button_Tuere_Pruefzeugnis.BackColor = colBlue
            End If
            Me.NUD_Tuer_R.Value = CDec(.L)
        End With
        Me.Update_Klasse_Tueren()
        'LS Außenbauteile
        Select Case Projekt.Aussenbauteile
            Case ohneNachweis
                Button_Aussenbauteile_oN_Click(Nothing, Nothing)
                'Case FensterMitDichtung
                '    Button_Aussenbauteile_FmD_Click(Nothing, Nothing)
            Case DINerfuellt
                Button_Aussenbauteile_DIN_Click(Nothing, Nothing)
            Case DINPlusErfuellt
                Button_Aussenbauteile_DINPlus_Click(Nothing, Nothing)
            Case Else
                Me.Button_Aussenbauteile_DINPlus.BackColor = colBlue
                Me.Button_Aussenbauteile_DIN.BackColor = colBlue
                'Me.Button_Aussenbauteile_FmD.BackColor = colBlue
                Me.Button_Aussenbauteile_oN.BackColor = colBlue
        End Select
        Me.Update_Klasse_Aussenbauteile()
        'Wasser + haustechn. Anlagen
        With Projekt.Wasser
            If .Untersuchung = Messung Then
                Button_Wasser_Messung_Click(Nothing, Nothing)
            ElseIf .Untersuchung = Prognose Then
                Button_Wasser_Prognose_Click(Nothing, Nothing)
            Else
                Me.Button_Wasser_Prognose.BackColor = colBlue
                Me.Button_Wasser_Messung.BackColor = colBlue
            End If
            Me.CB_Wasser_L.Text = Get_Str_Wasser_L() 'CDec(.Laf)
            If .LcLa_erfuellt = erfuellt Then
                Button_Wasser_erfuellt_Click(Nothing, Nothing)
            ElseIf .LcLa_erfuellt = keineAngabe Then
                Button_Wasser_kA_Click(Nothing, Nothing)
            Else
                Me.Button_Wasser_kA.BackColor = colBlue
                Me.Button_Wasser_erfuellt.BackColor = colBlue
            End If
        End With
        Me.Update_Klasse_Wasser()
        'Nutzer
        With Projekt.NutzerKoerper
            If .Untersuchung = nv Then
                'Button_Change_Color(Me.Button_Nutzer_on)
                Button_NutzerKoerper_nv_Click(Nothing, Nothing)
            ElseIf .Untersuchung = Koerper Then
                'Button_Change_Color(Me.Button_Nutzer_Prognose)
                'Button_NutzerKoerper_Prognose_Click(Nothing, Nothing)
                Button_Koerper_Click(Nothing, Nothing)
                Me.CB_Koerper_L.Text = Get_Str_Koerper_L()
                If .MessungPrognose = Messung Then
                    Button_NutzerKoerper_Messung_Click(Nothing, Nothing)
                ElseIf .MessungPrognose = Prognose Then
                    Button_NutzerKoerper_Prognose_Click(Nothing, Nothing)
                Else
                    Me.Button_NutzerKoerper_Messung.BackColor = colBlue
                    Me.Button_NutzerKoerper_Prognose.BackColor = colBlue
                End If
            ElseIf .Untersuchung = Nutzer Then
                '                Button_Change_Color(Me.Button_Nutzer_Messung)
                'Button_Nutzer_Messung_Click(Nothing, Nothing)
                Button_Nutzer_Click(Nothing, Nothing)
                Me.CB_Nutzer_L.Text = Get_Str_Nutzer_L()
                If .MessungPrognose = Messung Then
                    Button_NutzerKoerper_Messung_Click(Nothing, Nothing)
                ElseIf .MessungPrognose = Prognose Then
                    Button_NutzerKoerper_Prognose_Click(Nothing, Nothing)
                Else
                    Me.Button_NutzerKoerper_Messung.BackColor = colBlue
                    Me.Button_NutzerKoerper_Prognose.BackColor = colBlue
                End If
            Else
                Me.Button_NutzerKoerper_nv.BackColor = colBlue
                Me.Button_Nutzer.BackColor = colBlue
                Me.Button_Koerper.BackColor = colBlue
                Me.CB_Koerper_L.Text = ""
                Me.CB_Nutzer_L.Text = ""
                Me.Panel_SSA_NutzerKoerper_b.Height = 0
                Me.Panel_SSA_Nutzer_Pegel_b.Height = 0
                Me.Panel_SSA_Koerper_Pegel_b.Height = 0
            End If
        End With
        Me.Update_Klasse_NutzerKoerper()
        'Körper
        'With Projekt.Koerper
        '    If .Untersuchung = nv Then
        '        'Button_Change_Color(Me.Button_Koerper_on)
        '        Button_Koerper_on_Click(Nothing, Nothing)
        '    ElseIf .Untersuchung = Prognose Then
        '        'Button_Change_Color(Me.Button_Koerper_Prognose)
        '        Button_Koerper_Prognose_Click(Nothing, Nothing)
        '    ElseIf .Untersuchung = Messung Then
        '        'Button_Change_Color(Me.Button_Koerper_Messung)
        '        Button_Koerper_Messung_Click(Nothing, Nothing)
        '    Else
        '        Me.Button_Koerper_on.BackColor = colBlue
        '        Me.Button_NutzerKoerper_Prognose.BackColor = colBlue
        '        Me.Button_NutzerKoerper_Messung.BackColor = colBlue
        '    End If
        '    Me.CB_Koerper_L.Text = Get_Str_Koerper_L()
        'End With

        'Nachbarn
        Select Case Projekt.Nachbarn
            Case 1
                Button_Nachbar_1_Click(Nothing, Nothing)
            Case 2
                Button_Nachbar_2_Click(Nothing, Nothing)
            Case 3
                Button_Nachbar_3_Click(Nothing, Nothing)
            Case 4
                Button_Nachbar_4_Click(Nothing, Nothing)
            Case 5
                Button_Nachbar_5_Click(Nothing, Nothing)
            Case Else
                Me.Button_Nachbar_1.BackColor = colBlue
                Me.Button_Nachbar_2.BackColor = colBlue
                Me.Button_Nachbar_3.BackColor = colBlue
                Me.Button_Nachbar_4.BackColor = colBlue
                Me.Button_Nachbar_5.BackColor = colBlue
        End Select
        Me.Update_Klasse_Nachbarn()
        'Anordnung lauter Räume
        Select Case Projekt.anordnungRaeume
            Case guenstig
                Button_anordRaeume_g_Click(Nothing, Nothing)
            Case unguenstig
                Button_anordRaeume_ug_Click(Nothing, Nothing)
            Case Else
                Me.Button_anordRaeume_ug.BackColor = colBlue
                Me.Button_anordRaeume_g.BackColor = colBlue
        End Select

        'laute Räume angrenzend
        Select Case Projekt.lauteRaeume
            Case keineAngrenzend
                Button_lauteRaeume_keine_Click(Nothing, Nothing)
            Case L_35_45
                Button_lauteRaeume_35_Click(Nothing, Nothing)
            Case L_30_40
                Button_lauteRaeume_30_Click(Nothing, Nothing)
            Case L_25_35
                Button_lauteRaeume_25_Click(Nothing, Nothing)
            Case Else
                Me.Button_lauteRaeume_keine.BackColor = colBlue
                Me.Button_lauteRaeume_25.BackColor = colBlue
                Me.Button_lauteRaeume_30.BackColor = colBlue
                Me.Button_lauteRaeume_35.BackColor = colBlue
        End Select
        Me.Update_Klasse_lauteRaeume()
        'NHZ
        If Projekt.NHZ = NHZ_010 Then
            Button_NHZ_010_Click(Nothing, Nothing)
        ElseIf Projekt.NHZ = NHZ_020 Then
            Button_NHZ_020_Click(Nothing, Nothing)
        ElseIf Projekt.NHZ = NHZ_keine Then
            Button_NHZ_Keine_Click(Nothing, Nothing)
        Else
            Me.Button_NHZ_010.BackColor = colBlue
            Me.Button_NHZ_020.BackColor = colBlue
            Me.Button_NHZ_Keine.BackColor = colBlue
        End If
        Me.Update_Klasse_NHZ()
        'eigener Wohnbereich
        If Projekt.eigenerWohnbereich = keineEmpfehlung Then
            Button_EW_kE_Click(Nothing, Nothing)
        ElseIf Projekt.eigenerWohnbereich = EW1 Then
            Button_EW_1_Click(Nothing, Nothing)
        ElseIf Projekt.eigenerWohnbereich = EW2 Then
            Button_EW_2_Click(Nothing, Nothing)
        Else
            Me.Button_EW_kE.BackColor = colBlue
            Me.Button_EW_1.BackColor = colBlue
            Me.Button_EW_2.BackColor = colBlue
        End If
        bUpdate_Anzeige = False
        Update_SSA_Bewertung()
    End Sub
    Public Sub Update_Anzeige_Messung_TS_BLT()
        With Projekt.TS_BLT.Messung
            'Messung 1
            If .Messung_1.Bauteil = Balkon Then
                Button_Change_Color(Me.Button_TS_BLLT_BT_Ba_1)
            ElseIf .Messung_1.Bauteil = Laube Then
                Button_Change_Color(Me.Button_TS_TPLH_BT_La_1)
            ElseIf .Messung_1.Bauteil = Loggie Then
                Button_Change_Color(Me.Button_TS_BLLT_BT_Lo_1)
            ElseIf .Messung_1.Bauteil = Terrasse Then
                Button_Change_Color(Me.Button_TS_BLLT_BT_Te_1)
            Else
                Me.Button_TS_BLLT_BT_Ba_1.BackColor = colBlue
                Me.Button_TS_TPLH_BT_La_1.BackColor = colBlue
                Me.Button_TS_BLLT_BT_Lo_1.BackColor = colBlue
            End If
            If .Messung_1.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_BLLT_MV_1_mit)
            ElseIf .Messung_1.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_BLLT_MV_1_ohne)
            Else
                Me.Button_TS_BLLT_MV_1_mit.BackColor = colBlue
                Me.Button_TS_BLLT_MV_1_ohne.BackColor = colBlue
            End If

            Me.NUD_TS_BLLT_Mes_L_1.Value = CDec(.Messung_1.Pegel.Pegel)
            Me.NUD_TS_BLLT_Mes_C_1.Text = .Messung_1.Pegel.C

            'Messung 2
            If .Messung_2.Bauteil = Balkon Then
                Button_Change_Color(Me.Button_TS_BLLT_BT_Ba_2)
            ElseIf .Messung_2.Bauteil = Laube Then
                Button_Change_Color(Me.Button_TS_TPLH_BT_La_2)
            ElseIf .Messung_2.Bauteil = Loggie Then
                Button_Change_Color(Me.Button_TS_BLLT_BT_Lo_2)
            ElseIf .Messung_2.Bauteil = Terrasse Then
                Button_Change_Color(Me.Button_TS_BLLT_BT_Te_2)
            Else
                Me.Button_TS_BLLT_BT_Ba_2.BackColor = colBlue
                Me.Button_TS_TPLH_BT_La_2.BackColor = colBlue
                Me.Button_TS_BLLT_BT_Lo_2.BackColor = colBlue
            End If
            If .Messung_2.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_BLLT_MV_2_mit)
            ElseIf .Messung_2.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_BLLT_MV_2_ohne)
            Else
                Me.Button_TS_BLLT_MV_2_mit.BackColor = colBlue
                Me.Button_TS_BLLT_MV_2_ohne.BackColor = colBlue
            End If

            Me.NUD_TS_BLLT_Mes_L_2.Value = CDec(.Messung_2.Pegel.Pegel)
            Me.NUD_TS_BLLT_Mes_C_2.Text = .Messung_2.Pegel.C

            'Messung 3
            If .Messung_3.Bauteil = Balkon Then
                Button_Change_Color(Me.Button_TS_BLLT_BT_Ba_3)
            ElseIf .Messung_3.Bauteil = Laube Then
                Button_Change_Color(Me.Button_TS_TPLH_BT_La_3)
            ElseIf .Messung_3.Bauteil = Loggie Then
                Button_Change_Color(Me.Button_TS_BLLT_BT_Lo_3)
            ElseIf .Messung_3.Bauteil = Terrasse Then
                Button_Change_Color(Me.Button_TS_BLLT_BT_Te_3)
            Else
                Me.Button_TS_BLLT_BT_Ba_3.BackColor = colBlue
                Me.Button_TS_TPLH_BT_La_3.BackColor = colBlue
                Me.Button_TS_BLLT_BT_Lo_3.BackColor = colBlue
            End If
            If .Messung_3.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_BLLT_MV_3_mit)
            ElseIf .Messung_3.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_BLLT_MV_3_ohne)
            Else
                Me.Button_TS_BLLT_MV_3_mit.BackColor = colBlue
                Me.Button_TS_BLLT_MV_3_ohne.BackColor = colBlue
            End If
            Me.NUD_TS_BLLT_Mes_L_3.Value = CDec(.Messung_3.Pegel.Pegel)
            Me.NUD_TS_BLLT_Mes_C_3.Text = .Messung_3.Pegel.C

            'Messung 4
            If .Messung_4.Bauteil = Balkon Then
                Button_Change_Color(Me.Button_TS_BLLT_BT_Ba_4)
            ElseIf .Messung_4.Bauteil = Laube Then
                Button_Change_Color(Me.Button_TS_TPLH_BT_La_4)
            ElseIf .Messung_4.Bauteil = Loggie Then
                Button_Change_Color(Me.Button_TS_BLLT_BT_Lo_4)
            ElseIf .Messung_4.Bauteil = Terrasse Then
                Button_Change_Color(Me.Button_TS_BLLT_BT_Te_4)
            Else
                Me.Button_TS_BLLT_BT_Ba_4.BackColor = colBlue
                Me.Button_TS_TPLH_BT_La_4.BackColor = colBlue
                Me.Button_TS_BLLT_BT_Lo_4.BackColor = colBlue
            End If
            If .Messung_4.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_BLLT_MV_4_mit)
            ElseIf .Messung_4.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_BLLT_MV_4_ohne)
            Else
                Me.Button_TS_BLLT_MV_4_mit.BackColor = colBlue
                Me.Button_TS_BLLT_MV_4_ohne.BackColor = colBlue
            End If
            Me.NUD_TS_BLLT_Mes_L_4.Value = CDec(.Messung_4.Pegel.Pegel)
            Me.NUD_TS_BLLT_Mes_C_4.Text = .Messung_4.Pegel.C

            'Messung 5
            If .Messung_5.Bauteil = Balkon Then
                Button_Change_Color(Me.Button_TS_BLLT_BT_Ba_5)
            ElseIf .Messung_5.Bauteil = Laube Then
                Button_Change_Color(Me.Button_TS_TPLH_BT_La_5)
            ElseIf .Messung_5.Bauteil = Loggie Then
                Button_Change_Color(Me.Button_TS_BLLT_BT_Lo_5)
            ElseIf .Messung_5.Bauteil = Terrasse Then
                Button_Change_Color(Me.Button_TS_BLLT_BT_Te_5)
            Else
                Me.Button_TS_BLLT_BT_Ba_5.BackColor = colBlue
                Me.Button_TS_TPLH_BT_La_5.BackColor = colBlue
                Me.Button_TS_BLLT_BT_Lo_5.BackColor = colBlue
            End If
            If .Messung_5.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_BLLT_MV_5_mit)
            ElseIf .Messung_5.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_BLLT_MV_5_ohne)
            Else
                Me.Button_TS_BLLT_MV_5_mit.BackColor = colBlue
                Me.Button_TS_BLLT_MV_5_ohne.BackColor = colBlue
            End If
            Me.NUD_TS_BLLT_Mes_L_5.Value = CDec(.Messung_5.Pegel.Pegel)
            Me.NUD_TS_BLLT_Mes_C_5.Text = .Messung_5.Pegel.C

            'Messung 6
            If .Messung_6.Bauteil = Balkon Then
                Button_Change_Color(Me.Button_TS_BLLT_BT_Ba_6)
            ElseIf .Messung_6.Bauteil = Laube Then
                Button_Change_Color(Me.Button_TS_TPLH_BT_La_6)
            ElseIf .Messung_6.Bauteil = Loggie Then
                Button_Change_Color(Me.Button_TS_BLLT_BT_Lo_6)
            ElseIf .Messung_6.Bauteil = Terrasse Then
                Button_Change_Color(Me.Button_TS_BLLT_BT_Te_6)
            Else
                Me.Button_TS_BLLT_BT_Ba_6.BackColor = colBlue
                Me.Button_TS_TPLH_BT_La_6.BackColor = colBlue
                Me.Button_TS_BLLT_BT_Lo_6.BackColor = colBlue
            End If
            If .Messung_6.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_BLLT_MV_6_mit)
            ElseIf .Messung_6.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_BLLT_MV_6_ohne)
            Else
                Me.Button_TS_BLLT_MV_6_mit.BackColor = colBlue
                Me.Button_TS_BLLT_MV_6_ohne.BackColor = colBlue
            End If
            Me.NUD_TS_BLLT_Mes_L_6.Value = CDec(.Messung_6.Pegel.Pegel)
            Me.NUD_TS_BLLT_Mes_C_6.Text = .Messung_6.Pegel.C

            Messungen_anzeigen(CType(Me.Panel_SSA_TS_BLLT_Messung.Controls.Find("Button_TS_BLLT_M_" & (.Anzahl + 1).ToString, True)(0), System.Windows.Forms.Button))
            'Select Case .Anzahl
            '    Case 1, 0
            '        Me.Button_TS_BLLT_M_1_Click(Nothing, Nothing)
            '    Case 2
            '        Me.Button_TS_BLLT_M_2_Click(Nothing, Nothing)
            '    Case 3
            '        Me.Button_TS_BLLT_M_3_Click(Nothing, Nothing)
            '    Case 4
            '        Me.Button_TS_BLLT_M_4_Click(Nothing, Nothing)
            '    Case 5
            '        Me.Button_TS_BLLT_M_5_Click(Nothing, Nothing)
            '    Case 6
            '        Me.Button_TS_BLLT_M_6_Click(Nothing, Nothing)
            'End Select
        End With
    End Sub
    Public Sub Update_Anzeige_Messung_TS_TPH()
        With Projekt.TS_TPLH.Messung
            'Messung 1
            If .Messung_1.Bauteil = Treppe Then
                Button_Change_Color(Me.Button_TS_TPH_BT_T_1)
            ElseIf .Messung_1.Bauteil = Podest Then
                Button_Change_Color(Me.Button_TS_TPH_BT_P_1)
            ElseIf .Messung_1.Bauteil = Hausflur Then
                Button_Change_Color(Me.Button_TS_TPH_BT_H_1)
            Else
                Me.Button_TS_TPH_BT_T_1.BackColor = colBlue
                Me.Button_TS_TPH_BT_P_1.BackColor = colBlue
                Me.Button_TS_TPH_BT_H_1.BackColor = colBlue
            End If
            If .Messung_1.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_TPH_MV_1_mit)
            ElseIf .Messung_1.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_TPH_MV_1_ohne)
            Else
                Me.Button_TS_TPH_MV_1_mit.BackColor = colBlue
                Me.Button_TS_TPH_MV_1_ohne.BackColor = colBlue
            End If

            Me.NUD_TS_TPLH_Mes_L_1.Value = CDec(.Messung_1.Pegel.Pegel)
            Me.NUD_TS_TPLH_Mes_C_1.Text = .Messung_1.Pegel.C

            'Messung 2
            If .Messung_2.Bauteil = Treppe Then
                Button_Change_Color(Me.Button_TS_TPH_BT_T_2)
            ElseIf .Messung_2.Bauteil = Podest Then
                Button_Change_Color(Me.Button_TS_TPH_BT_P_2)
            ElseIf .Messung_2.Bauteil = Hausflur Then
                Button_Change_Color(Me.Button_TS_TPH_BT_H_2)
            Else
                Me.Button_TS_TPH_BT_T_2.BackColor = colBlue
                Me.Button_TS_TPH_BT_P_2.BackColor = colBlue
                Me.Button_TS_TPH_BT_H_2.BackColor = colBlue
            End If
            If .Messung_2.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_TPH_MV_2_mit)
            ElseIf .Messung_2.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_TPH_MV_2_ohne)
            Else
                Me.Button_TS_TPH_MV_2_mit.BackColor = colBlue
                Me.Button_TS_TPH_MV_2_ohne.BackColor = colBlue
            End If

            Me.NUD_TS_TPLH_Mes_L_2.Value = CDec(.Messung_2.Pegel.Pegel)
            Me.NUD_TS_TPLH_Mes_C_2.Text = .Messung_2.Pegel.C

            'Messung 3
            If .Messung_3.Bauteil = Treppe Then
                Button_Change_Color(Me.Button_TS_TPH_BT_T_3)
            ElseIf .Messung_3.Bauteil = Podest Then
                Button_Change_Color(Me.Button_TS_TPH_BT_P_3)
            ElseIf .Messung_3.Bauteil = Hausflur Then
                Button_Change_Color(Me.Button_TS_TPH_BT_H_3)
            Else
                Me.Button_TS_TPH_BT_T_3.BackColor = colBlue
                Me.Button_TS_TPH_BT_P_3.BackColor = colBlue
                Me.Button_TS_TPH_BT_H_3.BackColor = colBlue
            End If
            If .Messung_3.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_TPH_MV_3_mit)
            ElseIf .Messung_3.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_TPH_MV_3_ohne)
            Else
                Me.Button_TS_TPH_MV_3_mit.BackColor = colBlue
                Me.Button_TS_TPH_MV_3_ohne.BackColor = colBlue
            End If

            Me.NUD_TS_TPLH_Mes_L_3.Value = CDec(.Messung_3.Pegel.Pegel)
            Me.NUD_TS_TPLH_Mes_C_3.Text = .Messung_3.Pegel.C

            'Messung 4
            If .Messung_4.Bauteil = Treppe Then
                Button_Change_Color(Me.Button_TS_TPH_BT_T_4)
            ElseIf .Messung_4.Bauteil = Podest Then
                Button_Change_Color(Me.Button_TS_TPH_BT_P_4)
            ElseIf .Messung_4.Bauteil = Hausflur Then
                Button_Change_Color(Me.Button_TS_TPH_BT_H_4)
            Else
                Me.Button_TS_TPH_BT_T_4.BackColor = colBlue
                Me.Button_TS_TPH_BT_P_4.BackColor = colBlue
                Me.Button_TS_TPH_BT_H_4.BackColor = colBlue
            End If
            If .Messung_4.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_TPH_MV_4_mit)
            ElseIf .Messung_4.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_TPH_MV_4_ohne)
            Else
                Me.Button_TS_TPH_MV_4_mit.BackColor = colBlue
                Me.Button_TS_TPH_MV_4_ohne.BackColor = colBlue
            End If

            Me.NUD_TS_TPLH_Mes_L_4.Value = CDec(.Messung_4.Pegel.Pegel)
            Me.NUD_TS_TPLH_Mes_C_4.Text = .Messung_4.Pegel.C

            'Messung 5
            If .Messung_5.Bauteil = Treppe Then
                Button_Change_Color(Me.Button_TS_TPH_BT_T_5)
            ElseIf .Messung_5.Bauteil = Podest Then
                Button_Change_Color(Me.Button_TS_TPH_BT_P_5)
            ElseIf .Messung_5.Bauteil = Hausflur Then
                Button_Change_Color(Me.Button_TS_TPH_BT_H_5)
            Else
                Me.Button_TS_TPH_BT_T_5.BackColor = colBlue
                Me.Button_TS_TPH_BT_P_5.BackColor = colBlue
                Me.Button_TS_TPH_BT_H_5.BackColor = colBlue
            End If
            If .Messung_5.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_TPH_MV_5_mit)
            ElseIf .Messung_5.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_TPH_MV_5_ohne)
            Else
                Me.Button_TS_TPH_MV_5_mit.BackColor = colBlue
                Me.Button_TS_TPH_MV_5_ohne.BackColor = colBlue
            End If

            Me.NUD_TS_TPLH_Mes_L_5.Value = CDec(.Messung_5.Pegel.Pegel)
            Me.NUD_TS_TPLH_Mes_C_5.Text = .Messung_5.Pegel.C

            'Messung 6
            If .Messung_6.Bauteil = Treppe Then
                Button_Change_Color(Me.Button_TS_TPH_BT_T_6)
            ElseIf .Messung_6.Bauteil = Podest Then
                Button_Change_Color(Me.Button_TS_TPH_BT_P_6)
            ElseIf .Messung_6.Bauteil = Hausflur Then
                Button_Change_Color(Me.Button_TS_TPH_BT_H_6)
            Else
                Me.Button_TS_TPH_BT_T_6.BackColor = colBlue
                Me.Button_TS_TPH_BT_P_6.BackColor = colBlue
                Me.Button_TS_TPH_BT_H_6.BackColor = colBlue
            End If
            If .Messung_6.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_TPH_MV_6_mit)
            ElseIf .Messung_6.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_TPH_MV_6_ohne)
            Else
                Me.Button_TS_TPH_MV_6_mit.BackColor = colBlue
                Me.Button_TS_TPH_MV_6_ohne.BackColor = colBlue
            End If

            Me.NUD_TS_TPLH_Mes_L_6.Value = CDec(.Messung_6.Pegel.Pegel)
            Me.NUD_TS_TPLH_Mes_C_6.Text = .Messung_6.Pegel.C

            Messungen_anzeigen(CType(Me.Panel_SSA_TS_TPHf_Messung.Controls.Find("Button_TS_TPH_M_" & (.Anzahl + 1).ToString, True)(0), System.Windows.Forms.Button))
            'Select Case .Anzahl
            '    Case 1, 0
            '        Me.Button_TS_TPH_M_1_Click(Nothing, Nothing)
            '    Case 2
            '        Me.Button_TS_TPH_M_2_Click(Nothing, Nothing)
            '    Case 3
            '        Me.Button_TS_TPH_M_3_Click(Nothing, Nothing)
            '    Case 4
            '        Me.Button_TS_TPH_M_4_Click(Nothing, Nothing)
            '    Case 5
            '        Me.Button_TS_TPH_M_5_Click(Nothing, Nothing)
            '    Case 6
            '        Me.Button_TS_TPH_M_6_Click(Nothing, Nothing)
            'End Select
        End With
    End Sub
    Public Sub Update_Anzeige_Messung_TS_D()
        '# fEstrich
        If Projekt.TS_Decke.fEstrich = fE_kleiner50 Then
            Button_Change_Color(Me.Button_TS_D_MV_kleiner)
            Button_Change_Color(Me.Button_TS_D_P_kleiner)
        ElseIf Projekt.TS_Decke.fEstrich = fE_groesser50 Then
            Button_Change_Color(Me.Button_TS_D_MV_groesser)
            Button_Change_Color(Me.Button_TS_D_P_groesser)
        Else
            Me.Button_TS_D_MV_groesser.BackColor = colBlue
            Me.Button_TS_D_MV_kleiner.BackColor = colBlue
            Me.Button_TS_D_P_groesser.BackColor = colBlue
            Me.Button_TS_D_P_kleiner.BackColor = colBlue
        End If
        'Button_TS_D_MV_kleiner
        'Button_TS_D_P_kleiner

        With Projekt.TS_Decke.Messung
            'Messung 1
            If .Messung_1.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_D_MV_1_mit)
            ElseIf .Messung_1.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_D_MV_1_ohne)
            Else
                Me.Button_TS_D_MV_1_mit.BackColor = colBlue
                Me.Button_TS_D_MV_1_ohne.BackColor = colBlue
            End If
            Me.NUD_TS_D_Mes_L_1.Value = CDec(.Messung_1.Pegel.Pegel)
            Me.NUD_TS_D_Mes_C_1.Text = .Messung_1.Pegel.C

            'Messung 2
            If .Messung_2.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_D_MV_2_mit)
            ElseIf .Messung_2.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_D_MV_2_ohne)
            Else
                Me.Button_TS_D_MV_2_mit.BackColor = colBlue
                Me.Button_TS_D_MV_2_ohne.BackColor = colBlue
            End If
            Me.NUD_TS_D_Mes_L_2.Value = CDec(.Messung_2.Pegel.Pegel)
            Me.NUD_TS_D_Mes_C_2.Text = .Messung_2.Pegel.C

            'Messung 3
            If .Messung_3.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_D_MV_3_mit)
            ElseIf .Messung_3.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_D_MV_3_ohne)
            Else
                Me.Button_TS_D_MV_3_mit.BackColor = colBlue
                Me.Button_TS_D_MV_3_ohne.BackColor = colBlue
            End If
            Me.NUD_TS_D_Mes_L_3.Value = CDec(.Messung_3.Pegel.Pegel)
            Me.NUD_TS_D_Mes_C_3.Text = .Messung_3.Pegel.C

            'Messung 4
            If .Messung_4.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_D_MV_4_mit)
            ElseIf .Messung_4.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_D_MV_4_ohne)
            Else
                Me.Button_TS_D_MV_4_mit.BackColor = colBlue
                Me.Button_TS_D_MV_4_ohne.BackColor = colBlue
            End If
            Me.NUD_TS_D_Mes_L_4.Value = CDec(.Messung_4.Pegel.Pegel)
            Me.NUD_TS_D_Mes_C_4.Text = .Messung_4.Pegel.C

            'Messung 5
            If .Messung_5.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_D_MV_5_mit)
            ElseIf .Messung_5.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_D_MV_5_ohne)
            Else
                Me.Button_TS_D_MV_5_mit.BackColor = colBlue
                Me.Button_TS_D_MV_5_ohne.BackColor = colBlue
            End If
            Me.NUD_TS_D_Mes_L_5.Value = CDec(.Messung_5.Pegel.Pegel)
            Me.NUD_TS_D_Mes_C_5.Text = .Messung_5.Pegel.C

            'Messung 6
            If .Messung_6.Bodenbelag = Be_mit Then
                Button_Change_Color(Me.Button_TS_D_MV_6_mit)
            ElseIf .Messung_6.Bodenbelag = Be_ohne Then
                Button_Change_Color(Me.Button_TS_D_MV_6_ohne)
            Else
                Me.Button_TS_D_MV_6_mit.BackColor = colBlue
                Me.Button_TS_D_MV_6_ohne.BackColor = colBlue
            End If
            Me.NUD_TS_D_Mes_L_6.Value = CDec(.Messung_6.Pegel.Pegel)
            Me.NUD_TS_D_Mes_C_6.Text = .Messung_6.Pegel.C

            Messungen_anzeigen(CType(Me.Panel_SSA_TS_D_Messung.Controls.Find("Button_TS_D_M_" & (.Anzahl + 1).ToString, True)(0), System.Windows.Forms.Button))
            'Select Case .Anzahl
            '    Case 1, 0
            '        Me.Button_TS_D_M_1_Click(Nothing, Nothing)
            '    Case 2
            '        Me.Button_TS_D_M_2_Click(Nothing, Nothing)
            '    Case 3
            '        Me.Button_TS_D_M_3_Click(Nothing, Nothing)
            '    Case 4
            '        Me.Button_TS_D_M_4_Click(Nothing, Nothing)
            '    Case 5
            '        Me.Button_TS_D_M_5_Click(Nothing, Nothing)
            '    Case 6
            '        Me.Button_TS_D_M_6_Click(Nothing, Nothing)
            'End Select
        End With
    End Sub
    Public Sub Update_Anzeige_Messung_LS_D()
        With Projekt.LS_Decke.Messung
            'Messung 1
            Me.NUD_LS_D_Mes_R_1.Value = CDec(.Messung_1.Pegel)
            Me.NUD_LS_D_Mes_C_1.Text = .Messung_1.C

            'Messung 2
            Me.NUD_LS_D_Mes_R_2.Value = CDec(.Messung_2.Pegel)
            Me.NUD_LS_D_Mes_C_2.Text = .Messung_2.C

            'Messung 3
            Me.NUD_LS_D_Mes_R_3.Value = CDec(.Messung_3.Pegel)
            Me.NUD_LS_D_Mes_C_3.Text = .Messung_3.C

            'Messung 4
            Me.NUD_LS_D_Mes_R_4.Value = CDec(.Messung_4.Pegel)
            Me.NUD_LS_D_Mes_C_4.Text = .Messung_4.C

            'Messung 5
            Me.NUD_LS_D_Mes_R_5.Value = CDec(.Messung_5.Pegel)
            Me.NUD_LS_D_Mes_C_5.Text = .Messung_5.C

            'Messung 6
            Me.NUD_LS_D_Mes_R_6.Value = CDec(.Messung_6.Pegel)
            Me.NUD_LS_D_Mes_C_6.Text = .Messung_6.C

            'Select Case .Anzahl
            '    Case 1, 0
            '        Me.Button_LS_D_M_1_Click(Nothing, Nothing)
            '    Case 2
            '        Me.Button_LS_D_M_2_Click(Nothing, Nothing)
            '    Case 3
            '        Me.Button_LS_D_M_3_Click(Nothing, Nothing)
            '    Case 4
            '        Me.Button_LS_D_M_4_Click(Nothing, Nothing)
            '    Case 5
            '        Me.Button_LS_D_M_5_Click(Nothing, Nothing)
            '    Case 6
            '        Me.Button_LS_D_M_6_Click(Nothing, Nothing)
            'End Select
            Messungen_anzeigen(CType(Me.Panel_SSA_LS_D_Messung.Controls.Find("Button_LS_D_M_" & (.Anzahl + 1).ToString, True)(0), System.Windows.Forms.Button))
        End With
    End Sub
    Public Sub Update_Anzeige_Messung_LS_W()
        With Projekt.LS_Wand.Messung

            'Messung 1
            Me.NUD_LS_W_Mes_R_1.Value = CDec(.Messung_1.Pegel)
            Me.NUD_LS_W_Mes_C_1.Text = .Messung_1.C

            'Messung 2
            Me.NUD_LS_W_Mes_R_2.Value = CDec(.Messung_2.Pegel)
            Me.NUD_LS_W_Mes_C_2.Text = .Messung_2.C

            'Messung 3
            Me.NUD_LS_W_Mes_R_3.Value = CDec(.Messung_3.Pegel)
            Me.NUD_LS_W_Mes_C_3.Text = .Messung_3.C

            'Messung 4
            Me.NUD_LS_W_Mes_R_4.Value = CDec(.Messung_4.Pegel)
            Me.NUD_LS_W_Mes_C_4.Text = .Messung_4.C

            'Messung 5
            Me.NUD_LS_W_Mes_R_5.Value = CDec(.Messung_5.Pegel)
            Me.NUD_LS_W_Mes_C_5.Text = .Messung_5.C

            'Messung 6
            Me.NUD_LS_W_Mes_R_6.Value = CDec(.Messung_6.Pegel)
            Me.NUD_LS_W_Mes_C_6.Text = .Messung_6.C

            'Select Case .Anzahl
            '    Case 1, 0
            '        Me.Button_LS_W_M_1_Click(Nothing, Nothing)
            '    Case 2
            '        Me.Button_LS_W_M_2_Click(Nothing, Nothing)
            '    Case 3
            '        Me.Button_LS_W_M_3_Click(Nothing, Nothing)
            '    Case 4
            '        Me.Button_LS_W_M_4_Click(Nothing, Nothing)
            '    Case 5
            '        Me.Button_LS_W_M_5_Click(Nothing, Nothing)
            '    Case 6
            '        Me.Button_LS_W_M_6_Click(Nothing, Nothing)
            'End Select
            Messungen_anzeigen(CType(Me.Panel_SSA_LS_W_Messung.Controls.Find("Button_LS_W_M_" & (.Anzahl + 1).ToString, True)(0), System.Windows.Forms.Button))
        End With

    End Sub

#Region "Untermenue_Projekt"
    Private Sub Untermenue_Projekt_Neu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Untermenue_Projekt_Neu.Click
        Projekt_neu()
    End Sub

    Private Sub Untermenue_Projekt_Laden_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Untermenue_Projekt_Laden.Click
        Projekt_laden()
    End Sub

    Private Sub TSMI_ExcelProjektLaden_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSMI_ExcelProjektLaden.Click
        Projekt_Excel_laden()
    End Sub

    Private Sub Untermenue_Projekt_Entfernen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Untermenue_Projekt_Entfernen.Click
        Dim WForm As New Form_Projekt_loeschen
        WForm.ShowDialog()
    End Sub

    Private Sub Untermenue_Projekt_Schliessen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Untermenue_Projekt_Schliessen.Click
        Projekt_schliessen()
    End Sub

    Private Sub Untermenue_Projekt_Speichern_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Untermenue_Projekt_Speichern.Click
        If stProjekt_Pfad <> "" And stProjekt_Name <> "" Then
            Projektdaten_speichern()
        Else
            Projektdaten_speichernUnter()
        End If

    End Sub

    Private Sub Untermenue_Projekt_SpeichernUnter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Untermenue_Projekt_SpeichernUnter.Click
        Projektdaten_speichernUnter()
    End Sub

    Private Sub Untermenue_Projekt_Beenden_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Untermenue_Projekt_Beenden.Click
        Me.Close()
    End Sub
#End Region

#Region "Functions"
    Private Sub Button_Change_Color(ByVal clickButton As System.Windows.Forms.Button)
        For Each ctrl As Control In clickButton.Parent.Controls
            If TypeOf (ctrl) Is System.Windows.Forms.Button And InStr(ctrl.Name, "Button_Tab") = 0 Then ctrl.BackColor = colBlue
            ' If TypeOf (ctrl) Is Button Then ctrl.BackColor = colBlue
        Next
        clickButton.BackColor = colRed
    End Sub

    Private Sub Messungen_anzeigen(ByVal mButton As System.Windows.Forms.Button)
        'Name des Buttons f. Anzahl Messungen:      Button_LS_W_M_3
        'Name des Panels f. die Messung:            Panel_LS_W_M_3

        Dim nameButton As String = mButton.Name

        Dim iAnzMes As Integer = CInt(Microsoft.VisualBasic.Right(nameButton, 1))
        Dim namePanel As String = Microsoft.VisualBasic.Replace(nameButton, "Button", "Panel")
        namePanel = Microsoft.VisualBasic.Left(namePanel, namePanel.Length - 1)

        Dim iMes As Integer

        For iMes = 1 To iAnzMes
            For Each ctrl As Control In mButton.Parent.Parent.Controls
                If ctrl.Name = namePanel & iMes Then
                    ctrl.Show()
                    Exit For
                End If
            Next
        Next

        For iMes = iAnzMes + 1 To 6
            For Each ctrl As Control In mButton.Parent.Parent.Controls
                If ctrl.Name = namePanel & iMes Then
                    Messung_Clear(CType(ctrl, Panel))
                    ctrl.Hide()
                    Exit For
                End If
            Next
        Next

    End Sub
    Private Sub Messung_Clear(ByVal pnlMessung As Panel)
        'Steuerelemente zurücksetzen
        For Each ctrl As Control In pnlMessung.Controls
            If TypeOf (ctrl) Is System.Windows.Forms.Button Then
                ctrl.BackColor = colBlue
            ElseIf TypeOf (ctrl) Is NumericUpDown Then
                If ctrl.Name.Contains("_C_") Then
                    ctrl.Text = ""
                Else
                    ctrl.Text = "0"
                End If
            End If
        Next

        Dim katMessung As String = Microsoft.VisualBasic.Left(pnlMessung.Name, pnlMessung.Name.Length - 1)

        'Variablen zurücksetzen
        If katMessung = "Panel_LS_W_M_" Then
            If pnlMessung.Name = "Panel_LS_W_M_1" Then
                Projekt.LS_Wand.Messung.Messung_1.Pegel = 0
                Projekt.LS_Wand.Messung.Messung_1.C = ""
            ElseIf pnlMessung.Name = "Panel_LS_W_M_2" Then
                Projekt.LS_Wand.Messung.Messung_2.Pegel = 0
                Projekt.LS_Wand.Messung.Messung_2.C = ""
            ElseIf pnlMessung.Name = "Panel_LS_W_M_3" Then
                Projekt.LS_Wand.Messung.Messung_3.Pegel = 0
                Projekt.LS_Wand.Messung.Messung_3.C = ""
            ElseIf pnlMessung.Name = "Panel_LS_W_M_4" Then
                Projekt.LS_Wand.Messung.Messung_4.Pegel = 0
                Projekt.LS_Wand.Messung.Messung_4.C = ""
            ElseIf pnlMessung.Name = "Panel_LS_W_M_5" Then
                Projekt.LS_Wand.Messung.Messung_5.Pegel = 0
                Projekt.LS_Wand.Messung.Messung_5.C = ""
            ElseIf pnlMessung.Name = "Panel_LS_W_M_6" Then
                Projekt.LS_Wand.Messung.Messung_6.Pegel = 0
                Projekt.LS_Wand.Messung.Messung_6.C = ""
            End If
        ElseIf katMessung = "Panel_LS_D_M_" Then
            If (pnlMessung.Name = "Panel_LS_D_M_1") Then
                Projekt.LS_Decke.Messung.Messung_1.Pegel = 0
                Projekt.LS_Decke.Messung.Messung_1.C = ""
            ElseIf (pnlMessung.Name = "Panel_LS_D_M_2") Then
                Projekt.LS_Decke.Messung.Messung_2.Pegel = 0
                Projekt.LS_Decke.Messung.Messung_2.C = ""
            ElseIf (pnlMessung.Name = "Panel_LS_D_M_3") Then
                Projekt.LS_Decke.Messung.Messung_3.Pegel = 0
                Projekt.LS_Decke.Messung.Messung_3.C = ""
            ElseIf (pnlMessung.Name = "Panel_LS_D_M_4") Then
                Projekt.LS_Decke.Messung.Messung_4.Pegel = 0
                Projekt.LS_Decke.Messung.Messung_4.C = ""
            ElseIf (pnlMessung.Name = "Panel_LS_D_M_5") Then
                Projekt.LS_Decke.Messung.Messung_5.Pegel = 0
                Projekt.LS_Decke.Messung.Messung_5.C = ""
            ElseIf (pnlMessung.Name = "Panel_LS_D_M_6") Then
                Projekt.LS_Decke.Messung.Messung_6.Pegel = 0
                Projekt.LS_Decke.Messung.Messung_6.C = ""
            End If
        ElseIf katMessung = "Panel_TS_D_M_" Then
            If (pnlMessung.Name = "Panel_TS_D_M_1") Then
                Projekt.TS_Decke.Messung.Messung_1.Pegel.Pegel = 0
                Projekt.TS_Decke.Messung.Messung_1.Pegel.C = ""
            ElseIf (pnlMessung.Name = "Panel_TS_D_M_2") Then
                Projekt.TS_Decke.Messung.Messung_2.Pegel.Pegel = 0
                Projekt.TS_Decke.Messung.Messung_2.Pegel.C = ""
            ElseIf (pnlMessung.Name = "Panel_TS_D_M_3") Then
                Projekt.TS_Decke.Messung.Messung_3.Pegel.Pegel = 0
                Projekt.TS_Decke.Messung.Messung_3.Pegel.C = ""
            ElseIf (pnlMessung.Name = "Panel_TS_D_M_4") Then
                Projekt.TS_Decke.Messung.Messung_4.Pegel.Pegel = 0
                Projekt.TS_Decke.Messung.Messung_4.Pegel.C = ""
            ElseIf (pnlMessung.Name = "Panel_TS_D_M_5") Then
                Projekt.TS_Decke.Messung.Messung_5.Pegel.Pegel = 0
                Projekt.TS_Decke.Messung.Messung_5.Pegel.C = ""
            ElseIf (pnlMessung.Name = "Panel_TS_D_M_6") Then
                Projekt.TS_Decke.Messung.Messung_6.Pegel.Pegel = 0
                Projekt.TS_Decke.Messung.Messung_6.Pegel.C = ""
            End If
        ElseIf katMessung = "Panel_TS_TPH_M_" Then
            If (pnlMessung.Name = "Panel_TS_TPH_M_1") Then
                Projekt.TS_TPLH.Messung.Messung_1.Pegel.Pegel = 0
                Projekt.TS_TPLH.Messung.Messung_1.Pegel.C = ""
            ElseIf (pnlMessung.Name = "Panel_TS_TPH_M_2") Then
                Projekt.TS_TPLH.Messung.Messung_2.Pegel.Pegel = 0
                Projekt.TS_TPLH.Messung.Messung_2.Pegel.C = ""
            ElseIf (pnlMessung.Name = "Panel_TS_TPH_M_3") Then
                Projekt.TS_TPLH.Messung.Messung_3.Pegel.Pegel = 0
                Projekt.TS_TPLH.Messung.Messung_3.Pegel.C = ""
            ElseIf (pnlMessung.Name = "Panel_TS_TPH_M_4") Then
                Projekt.TS_TPLH.Messung.Messung_4.Pegel.Pegel = 0
                Projekt.TS_TPLH.Messung.Messung_4.Pegel.C = ""
            ElseIf (pnlMessung.Name = "Panel_TS_TPH_M_5") Then
                Projekt.TS_TPLH.Messung.Messung_5.Pegel.Pegel = 0
                Projekt.TS_TPLH.Messung.Messung_5.Pegel.C = ""
            ElseIf (pnlMessung.Name = "Panel_TS_TPH_M_6") Then
                Projekt.TS_TPLH.Messung.Messung_6.Pegel.Pegel = 0
                Projekt.TS_TPLH.Messung.Messung_6.Pegel.C = ""
            End If
        ElseIf katMessung = "Panel_TS_BLLT_M_" Then
            If (pnlMessung.Name = "Panel_TS_BLLT_M_1") Then
                Projekt.TS_BLT.Messung.Messung_1.Pegel.Pegel = 0
                Projekt.TS_BLT.Messung.Messung_1.Pegel.C = ""
            ElseIf (pnlMessung.Name = "Panel_TS_BLLT_M_2") Then
                Projekt.TS_BLT.Messung.Messung_2.Pegel.Pegel = 0
                Projekt.TS_BLT.Messung.Messung_2.Pegel.C = ""
            ElseIf (pnlMessung.Name = "Panel_TS_BLLT_M_3") Then
                Projekt.TS_BLT.Messung.Messung_3.Pegel.Pegel = 0
                Projekt.TS_BLT.Messung.Messung_3.Pegel.C = ""
            ElseIf (pnlMessung.Name = "Panel_TS_BLLT_M_4") Then
                Projekt.TS_BLT.Messung.Messung_4.Pegel.Pegel = 0
                Projekt.TS_BLT.Messung.Messung_4.Pegel.C = ""
            ElseIf (pnlMessung.Name = "Panel_TS_BLLT_M_5") Then
                Projekt.TS_BLT.Messung.Messung_5.Pegel.Pegel = 0
                Projekt.TS_BLT.Messung.Messung_5.Pegel.C = ""
            ElseIf (pnlMessung.Name = "Panel_TS_BLLT_M_6") Then
                Projekt.TS_BLT.Messung.Messung_6.Pegel.Pegel = 0
                Projekt.TS_BLT.Messung.Messung_6.Pegel.C = ""
            End If
        End If
    End Sub

    Private Sub NUD_Validating(ByVal meNUD As NumericUpDown)
        With meNUD
            If .Value > .Maximum Then .Value = .Maximum
            If .Value < .Minimum Then .Value = .Minimum
        End With
    End Sub
#End Region

#Region "Datum"


    Private Sub TB_Datum_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Datum.TextChanged
        If IsDate(Me.TB_Datum.Text) Then
            Dim tmpdate As Date = CDate(Me.TB_Datum.Text)
            Projekt.Datum = Me.TB_Datum.Text
            Me.Label_gueltigBis.Text = IIf(tmpdate.Day < 10, "0" & tmpdate.Day, tmpdate.Day).ToString & "." & IIf(tmpdate.Month < 10, "0" & tmpdate.Month, tmpdate.Month).ToString & "." & tmpdate.AddYears(10).Year.ToString
        Else
            Projekt.Datum = ""
            Me.Label_gueltigBis.Text = ""
        End If
    End Sub
    Private Sub TB_Datum_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Datum.Leave
        If Me.TB_Datum.Text <> "" And IsDate(Me.TB_Datum.Text) = False Then
            MsgBox("Ungültige Datumsangabe!", MsgBoxStyle.OkOnly, "Eingabefehler")
            Me.TB_Datum.Focus()
        Else
            Projekt.Datum = Me.TB_Datum.Text
        End If

        Me.Update_SSA_Datum()
    End Sub
#End Region
#Region "Antragsteller"
    Private Sub TB_Antragsteller_Name_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Antragsteller_Name.TextChanged
        Projekt.Antragsteller.Name = Me.TB_Antragsteller_Name.Text

        Me.Update_SSA_Antragsteller()
    End Sub

    Private Sub TB_Antragsteller_Zusatz_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Antragsteller_Zusatz.TextChanged
        Projekt.Antragsteller.Zusatz = Me.TB_Antragsteller_Zusatz.Text

        Me.Update_SSA_Antragsteller()
    End Sub

    Private Sub TB_Antragsteller_Strasse_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Antragsteller_Strasse.TextChanged
        Projekt.Antragsteller.Strasse = Me.TB_Antragsteller_Strasse.Text

        Me.Update_SSA_Antragsteller()
    End Sub

    Private Sub TB_Antragsteller_Nr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Antragsteller_Nr.TextChanged
        Projekt.Antragsteller.Nr = Me.TB_Antragsteller_Nr.Text

        Me.Update_SSA_Antragsteller()
    End Sub

    Private Sub TB_Antragsteller_PLZ_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Antragsteller_PLZ.TextChanged
        If IsNumeric(Me.TB_Antragsteller_PLZ.Text) Then
            Projekt.Antragsteller.PLZ = Me.TB_Antragsteller_PLZ.Text
        Else
            Projekt.Antragsteller.PLZ = ""
        End If

        Me.Update_SSA_Antragsteller()
    End Sub
    Private Sub TB_Antragsteller_PLZ_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Antragsteller_PLZ.Leave
        Dim bPlz As Boolean = False
        If IsNumeric(Me.TB_Antragsteller_PLZ.Text) Then
            If Me.TB_Antragsteller_PLZ.Text.Length <> 5 Or CInt(Me.TB_Antragsteller_PLZ.Text) < 1000 Or CInt(Me.TB_Antragsteller_PLZ.Text) > 99999 Then
                bPlz = True
            End If
        ElseIf Me.TB_Antragsteller_PLZ.Text <> "" Then
            bPlz = True
        End If

        If bPlz Then
            MsgBox("Die eingegebene Postleitzahl ist ungültig!", MsgBoxStyle.OkOnly, "Eingabefehler")
            Me.TB_Antragsteller_PLZ.Text = ""
            Me.TB_Antragsteller_PLZ.Focus()
        End If

        Me.Update_SSA_Antragsteller()
    End Sub

    Private Sub TB_Antragsteller_Ort_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Antragsteller_Ort.TextChanged
        Projekt.Antragsteller.Ort = Me.TB_Antragsteller_Ort.Text

        Me.Update_SSA_Antragsteller()
    End Sub
#End Region
#Region "Gebäude"
    Private Sub TB_Gebaeude_Name_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Gebaeude_Name.TextChanged
        Projekt.Gebaeude.Adresse.Name = Me.TB_Gebaeude_Name.Text

        Me.Update_SSA_Gebaeude()
    End Sub

    Private Sub TB_Gebaeude_Zusatz_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Gebaeude_Zusatz.TextChanged
        Projekt.Gebaeude.Adresse.Zusatz = Me.TB_Gebaeude_Zusatz.Text

        Me.Update_SSA_Gebaeude()
    End Sub

    Private Sub TB_Gebaeude_Strasse_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Gebaeude_Strasse.TextChanged
        Projekt.Gebaeude.Adresse.Strasse = Me.TB_Gebaeude_Strasse.Text

        Me.Update_SSA_Gebaeude()
    End Sub

    Private Sub TB_Gebaeude_Nr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Gebaeude_Nr.TextChanged
        Projekt.Gebaeude.Adresse.Nr = Me.TB_Gebaeude_Nr.Text

        Me.Update_SSA_Gebaeude()
    End Sub

    Private Sub TB_Gebaeude_PLZ_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Gebaeude_PLZ.TextChanged
        If IsNumeric(Me.TB_Gebaeude_PLZ.Text) Then
            Projekt.Gebaeude.Adresse.PLZ = Me.TB_Gebaeude_PLZ.Text
        Else
            Projekt.Gebaeude.Adresse.PLZ = ""
        End If

        Me.Update_SSA_Gebaeude()
    End Sub
    Private Sub TB_Gebaeude_PLZ_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Gebaeude_PLZ.Leave
        Dim bPlz As Boolean = False
        If IsNumeric(Me.TB_Gebaeude_PLZ.Text) Then
            If Me.TB_Gebaeude_PLZ.Text.Length <> 5 Or CInt(Me.TB_Gebaeude_PLZ.Text) < 1000 Or CInt(Me.TB_Gebaeude_PLZ.Text) > 99999 Then
                bPlz = True
            End If
        ElseIf Me.TB_Gebaeude_PLZ.Text <> "" Then
            bPlz = True
        End If

        If bPlz Then
            MsgBox("Die eingegebene Postleitzahl ist ungültig!", MsgBoxStyle.OkOnly, "Eingabefehler")
            Me.TB_Gebaeude_PLZ.Text = ""
            Me.TB_Gebaeude_PLZ.Focus()
        End If

        Me.Update_SSA_Gebaeude()
    End Sub

    Private Sub TB_Gebaeude_Ort_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Gebaeude_Ort.TextChanged
        Projekt.Gebaeude.Adresse.Ort = Me.TB_Gebaeude_Ort.Text

        Me.Update_SSA_Gebaeude()
    End Sub

    Private Sub CB_Gebaeude_Gebaeudetyp_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Gebaeude_Gebaeudetyp.SelectedIndexChanged
        Projekt.Gebaeude.Gebaeudetyp = Me.CB_Gebaeude_Gebaeudetyp.Text

        Me.Update_SSA_Gebaeude()
    End Sub
    Private Sub TB_Gebaeude_Baujahr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Gebaeude_Baujahr.TextChanged
        If IsNumeric(Me.TB_Gebaeude_Baujahr.Text) Then
            Projekt.Gebaeude.Baujahr = CInt(Me.TB_Gebaeude_Baujahr.Text)
        Else
            Projekt.Gebaeude.Baujahr = 1800
        End If

        Me.Update_SSA_Gebaeude()
    End Sub
    Private Sub TB_Gebaeude_Baujahr_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Gebaeude_Baujahr.Leave
        Dim bJahr As Boolean = False
        If IsNumeric(Me.TB_Gebaeude_Baujahr.Text) Then
            If Me.TB_Gebaeude_Baujahr.Text.Length <> 4 Or CInt(Me.TB_Gebaeude_Baujahr.Text) <= 2200 Or CInt(Me.TB_Gebaeude_Baujahr.Text) >= 1800 Then
                bJahr = True
            End If
        ElseIf Me.TB_Gebaeude_Baujahr.Text <> "" Then
            bJahr = True
        End If

        If bJahr = False Then
            MsgBox("Das eingegebene Baujahr ist ungültig!", MsgBoxStyle.OkOnly, "Eingabefehler")
            Me.TB_Gebaeude_Baujahr.Text = ""
            Me.TB_Gebaeude_Baujahr.Focus()
        End If

        Me.Update_SSA_Gebaeude()
    End Sub

    Private Sub nud_Gebaeude_Wohneinheiten_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_Gebaeude_Wohneinheiten.TextChanged
        If Me.TabPage1.Focused = False Then
            Projekt.Gebaeude.Wohneinheiten = Me.NUD_Gebaeude_Wohneinheiten.Text
        Else
            Me.NUD_Gebaeude_Wohneinheiten.Text = Projekt.Gebaeude.Wohneinheiten
        End If

        Me.Update_SSA_Gebaeude()
    End Sub
#End Region
#Region "Wohnung"
    Private Sub TB_Wohnung_Wohnungsbezeichnung_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Wohnung_Wohnungsbezeichnung.TextChanged
        Projekt.Wohnung.Wohnungsbezeichnung = Me.TB_Wohnung_Wohnungsbezeichnung.Text

        Me.Update_SSA_Gebaeude()
    End Sub

    Private Sub CB_Wohnung_Geschoss_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Wohnung_Geschoss.SelectedIndexChanged
        Projekt.Wohnung.Geschoss.Typ = CByte(Me.CB_Wohnung_Geschoss.SelectedIndex + 1)
        If Me.CB_Wohnung_Geschoss.Text = "EG" Then NUD_Wohnung_Geschoss.Value = 0

        Me.Update_SSA_Gebaeude()
    End Sub

    Private Sub NUD_Wohnung_Geschoss_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_Wohnung_Geschoss.TextChanged
        If Me.CB_Wohnung_Geschoss.Text = "EG" And Me.NUD_Wohnung_Geschoss.Value <> 0 Then
            Me.NUD_Wohnung_Geschoss.Value = 0
        Else
            If Me.TabPage1.Focused = False Then
                If Me.NUD_Wohnung_Geschoss.Text = "" Then Me.NUD_Wohnung_Geschoss.Text = "0"
                Projekt.Wohnung.Geschoss.Nr = CInt(Me.NUD_Wohnung_Geschoss.Text)
            Else
                Me.NUD_Wohnung_Geschoss.Value = Projekt.Wohnung.Geschoss.Nr
            End If
        End If

        Me.Update_SSA_Gebaeude()
    End Sub
    Private Sub NUD_Wohnung_Geschoss_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_Wohnung_Geschoss.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub TB_Wohnung_Geschoss_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Wohnung_Geschoss.TextChanged
        Projekt.Wohnung.Geschoss.Bezeichnung = Me.TB_Wohnung_Geschoss.Text

        Me.Update_SSA_Gebaeude()
    End Sub

    Private Sub NUD_Wohnung_Raeume_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_Wohnung_Raeume.TextChanged

        If Me.TabPage1.Focused = False Then
            If Me.NUD_Wohnung_Raeume.Text = "" Then Me.NUD_Wohnung_Raeume.Text = "0"
            Projekt.Wohnung.Raeume = Me.NUD_Wohnung_Raeume.Text
        Else
            Me.NUD_Wohnung_Raeume.Text = Projekt.Wohnung.Raeume
        End If

        Me.Update_SSA_Gebaeude()
    End Sub
    Private Sub NUD_Wohnung_Raeume_validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_Wohnung_Raeume.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_Wohnung_Wohnflaeche_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_Wohnung_Wohnflaeche.TextChanged
        If Me.TabPage1.Focused = False Then 'CType(sender, NumericUpDown).Focused = False And 
            If Me.NUD_Wohnung_Wohnflaeche.Text = "" Then Me.NUD_Wohnung_Wohnflaeche.Text = "1"
            Projekt.Wohnung.Wohnflaeche = CSng(Me.NUD_Wohnung_Wohnflaeche.Text)
        Else
            Me.NUD_Wohnung_Wohnflaeche.Value = CDec(Projekt.Wohnung.Wohnflaeche)
        End If
    End Sub
    Private Sub NUD_Wohnung_Wohnflaeche_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_Wohnung_Wohnflaeche.Validating
        NUD_Validating(CType(sender, NumericUpDown))

        Me.Update_SSA_Gebaeude()
    End Sub

    Private Sub NUD_Gebaeude_Kosten_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_Gebaeude_Kosten.TextChanged
        If Me.TabPage1.Focused = False Then
            If Me.NUD_Gebaeude_Kosten.Text = "" Then Me.NUD_Gebaeude_Kosten.Text = "0"
            Projekt.Gebaeude.Kosten = CInt(Me.NUD_Gebaeude_Kosten.Text)
        Else
            Me.NUD_Gebaeude_Kosten.Value = CDec(Projekt.Gebaeude.Kosten)
        End If
    End Sub
    Private Sub NUD_Gebaeude_Kosten_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_Gebaeude_Kosten.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub
#End Region

#Region "GC - Gebietscharakter"
    Private Sub Button_GC_WR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_GC_WR.Click
        If CType(sender, System.Windows.Forms.Button).BackColor = colRed Then
            CType(sender, System.Windows.Forms.Button).BackColor = colBlue

            Projekt.Standort.Gebietscharakter = 0
        Else
            Button_Change_Color(Button_GC_WR)

            Projekt.Standort.Gebietscharakter = GC_WR
        End If
    
        Update_Klasse_GC()
    End Sub

    Private Sub Button_GC_WA_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_GC_WA.Click
        If CType(sender, System.Windows.Forms.Button).BackColor = colRed Then
            CType(sender, System.Windows.Forms.Button).BackColor = colBlue

            Projekt.Standort.Gebietscharakter = 0
        Else
            Button_Change_Color(Me.Button_GC_WA)

            Projekt.Standort.Gebietscharakter = GC_WA
        End If

        Update_Klasse_GC()
    End Sub

    Private Sub Button_GC_MIWB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_GC_MIWB.Click
        If CType(sender, System.Windows.Forms.Button).BackColor = colRed Then
            CType(sender, System.Windows.Forms.Button).BackColor = colBlue

            Projekt.Standort.Gebietscharakter = 0
        Else
            Button_Change_Color(Me.Button_GC_MIWB)

            Projekt.Standort.Gebietscharakter = GC_MIWB
        End If

        Update_Klasse_GC()
    End Sub

    Private Sub Button_GC_GE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_GC_GE.Click
        If CType(sender, System.Windows.Forms.Button).BackColor = colRed Then
            CType(sender, System.Windows.Forms.Button).BackColor = colBlue

            Projekt.Standort.Gebietscharakter = 0
        Else
            Button_Change_Color(Me.Button_GC_GE)

            Projekt.Standort.Gebietscharakter = GC_GE
        End If

        Update_Klasse_GC()
    End Sub

    Private Sub Button_GC_GI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_GC_GI.Click
        If CType(sender, System.Windows.Forms.Button).BackColor = colRed Then
            CType(sender, System.Windows.Forms.Button).BackColor = colBlue

            Projekt.Standort.Gebietscharakter = 0
        Else
            Button_Change_Color(Me.Button_GC_GI)

            Projekt.Standort.Gebietscharakter = GC_GI
        End If
        Update_Klasse_GC()
    End Sub

    Private Sub Update_Klasse_GC()
        Me.Button_Klasse_GC.Text = Get_Klasse_GC()
        Me.Button_Klasse_GC.BackColor = Get_Klasse_Color(Me.Button_Klasse_GC.Text)

        Me.Update_SSA_GC()
    End Sub
#End Region
#Region "AP - Außenlärmpegel"
    Private Sub Button_AP_bis55_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_AP_bis55.Click
        If CType(sender, System.Windows.Forms.Button).BackColor = colRed Then
            CType(sender, System.Windows.Forms.Button).BackColor = colBlue

            Projekt.Standort.Aussenlaermpegel = 0
        Else
            Button_Change_Color(Button_AP_bis55)

            Projekt.Standort.Aussenlaermpegel = AP_bis55
        End If

        Update_Klasse_AP()
    End Sub

    Private Sub Button_AP_56_60_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_AP_56_60.Click
        If CType(sender, System.Windows.Forms.Button).BackColor = colRed Then
            CType(sender, System.Windows.Forms.Button).BackColor = colBlue

            Projekt.Standort.Aussenlaermpegel = 0
        Else
            Button_Change_Color(Button_AP_56_60)

            Projekt.Standort.Aussenlaermpegel = AP_56_60
        End If

        Update_Klasse_AP()
    End Sub

    Private Sub Button_AP_61_65_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_AP_61_65.Click
        If CType(sender, System.Windows.Forms.Button).BackColor = colRed Then
            CType(sender, System.Windows.Forms.Button).BackColor = colBlue

            Projekt.Standort.Aussenlaermpegel = 0
        Else
            Button_Change_Color(Button_AP_61_65)

            Projekt.Standort.Aussenlaermpegel = AP_61_65
        End If
        Update_Klasse_AP()
    End Sub

    Private Sub Button_AP_66_70_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_AP_66_70.Click
        If CType(sender, System.Windows.Forms.Button).BackColor = colRed Then
            CType(sender, System.Windows.Forms.Button).BackColor = colBlue

            Projekt.Standort.Aussenlaermpegel = 0
        Else
            Button_Change_Color(Button_AP_66_70)

            Projekt.Standort.Aussenlaermpegel = AP_66_70
        End If
        Update_Klasse_AP()
    End Sub

    Private Sub Button_AP_71_75_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_AP_71_75.Click
        If CType(sender, System.Windows.Forms.Button).BackColor = colRed Then
            CType(sender, System.Windows.Forms.Button).BackColor = colBlue

            Projekt.Standort.Aussenlaermpegel = 0
        Else
            Button_Change_Color(Button_AP_71_75)

            Projekt.Standort.Aussenlaermpegel = AP_71_75
        End If
        Update_Klasse_AP()
    End Sub

    Private Sub Button_AP_76_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_AP_76.Click
        If CType(sender, System.Windows.Forms.Button).BackColor = colRed Then
            CType(sender, System.Windows.Forms.Button).BackColor = colBlue

            Projekt.Standort.Aussenlaermpegel = 0
        Else
            Button_Change_Color(Button_AP_76)

            Projekt.Standort.Aussenlaermpegel = AP_76
        End If
        Update_Klasse_AP()
    End Sub

    Private Sub Update_Klasse_AP()
        Me.Button_Klasse_AP.Text = Get_Klasse_AP()
        Me.Button_Klasse_AP.BackColor = Get_Klasse_Color(Me.Button_Klasse_AP.Text)

        Me.Update_SSA_AL()
    End Sub
#End Region
#Region "aF - abgewandt Freibereich"

    Private Sub Button_aF_ja_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_aF_ja.Click
        If CType(sender, System.Windows.Forms.Button).BackColor = colRed Then
            CType(sender, System.Windows.Forms.Button).BackColor = colBlue

            Projekt.Standort.abgewandFreibereich = 0
        Else
            Button_Change_Color(Button_aF_ja)

            Projekt.Standort.abgewandFreibereich = aF_ja
        End If
        Me.Update_SSA_AL()
    End Sub

    Private Sub Button_aF_nein_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_aF_nein.Click
        If CType(sender, System.Windows.Forms.Button).BackColor = colRed Then
            CType(sender, System.Windows.Forms.Button).BackColor = colBlue

            Projekt.Standort.abgewandFreibereich = 0
        Else
            Button_Change_Color(Button_aF_nein)

            Projekt.Standort.abgewandFreibereich = aF_nein
        End If
        Me.Update_SSA_AL()
    End Sub
#End Region

#Region "LS Wände - Luftschall Wände"
    Private Sub Update_Klasse_LS_W()
        Me.Button_Klasse_LS_W.Text = Get_Klasse_LS_Waende()
        Me.Button_Klasse_LS_W.BackColor = Get_Klasse_Color(Me.Button_Klasse_LS_W.Text)

        Me.Update_SSA_LS_W()
    End Sub
#Region "Untersuchung"
    Private Sub Button_LS_W_Prognose_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_W_Prognose.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_37", CType(sender, Control))
    End Sub
    Private Sub Button_LS_W_Prognose_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_W_Prognose.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide() 'Show_Help("ssa_help_37", (CType(sender, Button).Location))
    End Sub

    Private Sub Button_LS_W_Prognose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_W_Prognose.Click

        Button_Change_Color(Button_LS_W_Prognose)

        Panel_SSA_LS_W_Messung_b.Height = 0   '90
        Panel_SSA_LS_W_Prognose_b.Height = 37    '0

        Projekt.LS_Wand.Untersuchung = Prognose

        Update_Klasse_LS_W()

    End Sub

    Private Sub Button_LS_W_Messung_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_W_Messung.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_37", CType(sender, Control))
    End Sub
    Private Sub Button_LS_W_Messung_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_W_Messung.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub
    Private Sub Button_LS_W_Messung_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_W_Messung.Click
        Button_Change_Color(Button_LS_W_Messung)

        Panel_SSA_LS_W_Messung_b.Height = 55  '0
        Panel_SSA_LS_W_Prognose_b.Height = 0  '36

        Projekt.LS_Wand.Untersuchung = Messung

        Update_Klasse_LS_W()
    End Sub

    Private Sub Button_LS_W_nv_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_W_nv.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_36", CType(sender, Control))
    End Sub
    Private Sub Button_ls_w_nv_mouseleave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_W_nv.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub
    Private Sub Button_LS_W_nv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_W_nv.Click
        Button_Change_Color(Button_LS_W_nv)

        Panel_SSA_LS_W_Messung_b.Height = 0  '90
        Panel_SSA_LS_W_Prognose_b.Height = 0  '36

        Projekt.LS_Wand.Untersuchung = nv

        Update_Klasse_LS_W()
    End Sub
#End Region
#Region "Anzahl Messungen"
    Private Sub Button_LS_W_M_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_W_M_1.Click
        Button_Change_Color(Button_LS_W_M_1)

        Projekt.LS_Wand.Messung.Anzahl = 0
        Messungen_anzeigen(Button_LS_W_M_1)

        Update_Klasse_LS_W()
    End Sub

    Private Sub Button_LS_W_M_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_W_M_2.Click
        Button_Change_Color(Button_LS_W_M_2)

        Projekt.LS_Wand.Messung.Anzahl = 1
        Messungen_anzeigen(Button_LS_W_M_2)

        Update_Klasse_LS_W()
    End Sub

    Private Sub Button_LS_W_M_3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_W_M_3.Click
        Button_Change_Color(Button_LS_W_M_3)

        Projekt.LS_Wand.Messung.Anzahl = 2
        Messungen_anzeigen(Button_LS_W_M_3)

        Update_Klasse_LS_W()
    End Sub

    Private Sub Button_LS_W_M_4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_W_M_4.Click
        Button_Change_Color(Button_LS_W_M_4)

        Projekt.LS_Wand.Messung.Anzahl = 3
        Messungen_anzeigen(Button_LS_W_M_4)

        Update_Klasse_LS_W()
    End Sub

    Private Sub Button_LS_W_M_5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_W_M_5.Click
        Button_Change_Color(Button_LS_W_M_5)

        Projekt.LS_Wand.Messung.Anzahl = 4
        Messungen_anzeigen(Button_LS_W_M_5)

        Update_Klasse_LS_W()
    End Sub

    Private Sub Button_LS_W_M_6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_W_M_6.Click
        Button_Change_Color(Button_LS_W_M_6)

        Projekt.LS_Wand.Messung.Anzahl = 5
        Messungen_anzeigen(Button_LS_W_M_6)

        Update_Klasse_LS_W()
    End Sub
#End Region
#Region "NUD - Messungen"
#Region "Validating"
    Private Sub NUD_LS_W_Mes_R_1_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_W_Mes_R_1.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_LS_W_Mes_C_1_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_W_Mes_C_1.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.LS_Wand.Messung.Messung_1.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_LS_W_Mes_R_2_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_W_Mes_R_2.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_LS_W_Mes_C_2_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_W_Mes_C_2.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.LS_Wand.Messung.Messung_2.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_LS_W_Mes_R_3_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_W_Mes_R_3.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_LS_W_Mes_C_3_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_W_Mes_C_3.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.LS_Wand.Messung.Messung_3.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_LS_W_Mes_R_4_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_W_Mes_R_4.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_LS_W_Mes_C_4_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_W_Mes_C_4.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.LS_Wand.Messung.Messung_4.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_LS_W_Mes_R_5_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_W_Mes_R_5.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_LS_W_Mes_C_5_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_W_Mes_C_5.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.LS_Wand.Messung.Messung_5.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_LS_W_Mes_R_6_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_W_Mes_R_6.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_LS_W_Mes_C_6_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_W_Mes_C_6.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.LS_Wand.Messung.Messung_6.C = CType(sender, NumericUpDown).Text
    End Sub
#End Region
#Region "Value Changed"
    Private Sub NUD_LS_W_Mes_R_1_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_W_Mes_R_1.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.LS_Wand.Messung.Messung_1.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.LS_Wand.Messung.Messung_1.Pegel)
        End If

        Update_Klasse_LS_W()
    End Sub

    Private Sub NUD_LS_W_Mes_C_1_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_W_Mes_C_1.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.LS_Wand.Messung.Messung_1.C = menud.Text
        Else
            menud.Text = Projekt.LS_Wand.Messung.Messung_1.C
        End If

        Update_Klasse_LS_W()
    End Sub

    Private Sub NUD_LS_W_Mes_R_2_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_W_Mes_R_2.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.LS_Wand.Messung.Messung_2.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.LS_Wand.Messung.Messung_2.Pegel)
        End If

        Update_Klasse_LS_W()
    End Sub

    Private Sub NUD_LS_W_Mes_C_2_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_W_Mes_C_2.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.LS_Wand.Messung.Messung_2.C = menud.Text
        Else
            menud.Text = Projekt.LS_Wand.Messung.Messung_2.C
        End If

        Update_Klasse_LS_W()
    End Sub

    Private Sub NUD_LS_W_Mes_R_3_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_W_Mes_R_3.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.LS_Wand.Messung.Messung_3.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.LS_Wand.Messung.Messung_3.Pegel)
        End If

        Update_Klasse_LS_W()
    End Sub

    Private Sub NUD_LS_W_Mes_C_3_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_W_Mes_C_3.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.LS_Wand.Messung.Messung_3.C = menud.Text
        Else
            menud.Text = Projekt.LS_Wand.Messung.Messung_3.C
        End If

        Update_Klasse_LS_W()
    End Sub

    Private Sub NUD_LS_W_Mes_R_4_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_W_Mes_R_4.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.LS_Wand.Messung.Messung_4.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.LS_Wand.Messung.Messung_4.Pegel)
        End If

        Update_Klasse_LS_W()
    End Sub

    Private Sub NUD_LS_W_Mes_C_4_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_W_Mes_C_4.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.LS_Wand.Messung.Messung_4.C = menud.Text
        Else
            menud.Text = Projekt.LS_Wand.Messung.Messung_4.C
        End If

        Update_Klasse_LS_W()
    End Sub

    Private Sub NUD_LS_W_Mes_R_5_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_W_Mes_R_5.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.LS_Wand.Messung.Messung_5.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.LS_Wand.Messung.Messung_5.Pegel)
        End If

        Update_Klasse_LS_W()
    End Sub

    Private Sub NUD_LS_W_Mes_C_5_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_W_Mes_C_5.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.LS_Wand.Messung.Messung_5.C = menud.Text
        Else
            menud.Text = Projekt.LS_Wand.Messung.Messung_5.C
        End If

        Update_Klasse_LS_W()
    End Sub

    Private Sub NUD_LS_W_Mes_R_6_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_W_Mes_R_6.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.LS_Wand.Messung.Messung_6.Pegel = CSng(menud.Value)
        Else
            menud.Value = CDec(Projekt.LS_Wand.Messung.Messung_6.Pegel)
        End If

        Update_Klasse_LS_W()
    End Sub

    Private Sub NUD_LS_W_Mes_C_6_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_W_Mes_C_6.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.LS_Wand.Messung.Messung_6.C = menud.Text
        Else
            menud.Text = Projekt.LS_Wand.Messung.Messung_6.C
        End If

        Update_Klasse_LS_W()
    End Sub
#End Region
#End Region
#Region "NUD - Prognose"
#Region "Validating"
    Private Sub NUD_LS_W_Prog_R_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_W_Prog_R.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_LS_W_Prog_C_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_W_Prog_C.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.LS_Wand.Prognose.C = CType(sender, NumericUpDown).Text
    End Sub
#End Region
#Region "TextChanged"
    Private Sub NUD_LS_W_Prog_R_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_W_Prog_R.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.LS_Wand.Prognose.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.LS_Wand.Prognose.Pegel)
        End If
        Update_Klasse_LS_W()
    End Sub

    Private Sub NUD_LS_W_Prog_C_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_W_Prog_C.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.LS_Wand.Prognose.C = menud.Text
        Else
            menud.Text = Projekt.LS_Wand.Prognose.C
        End If

        Update_Klasse_LS_W()
    End Sub
#End Region
#End Region
#End Region

#Region "LS Decken - Luftschall Decken"
    Private Sub Update_Klasse_LS_D()
        Me.Button_Klasse_LS_D.Text = Get_Klasse_LS_Decken()
        Me.Button_Klasse_LS_D.BackColor = Get_Klasse_Color(Me.Button_Klasse_LS_D.Text)

        Me.Update_SSA_LS_D()
    End Sub
#Region "Untersuchung"
    Private Sub Button_LS_D_Prognose_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_D_Prognose.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_37", CType(sender, Control))
    End Sub
    Private Sub Button_LS_D_Prognose_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_D_Prognose.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub
    Private Sub Button_LS_D_Prognose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_D_Prognose.Click
        Button_Change_Color(Button_LS_D_Prognose)

        Panel_SSA_LS_D_Messung_b.Height = 0   '90
        Panel_SSA_LS_D_Prognose_b.Height = 37

        Projekt.LS_Decke.Untersuchung = Prognose

        Update_Klasse_LS_D()
    End Sub

    Private Sub Button_LS_D_Messung_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_D_Messung.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_37", CType(sender, Control))
    End Sub
    Private Sub Button_LS_D_Messung_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_D_Messung.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub
    Private Sub Button_LS_D_Messung_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_D_Messung.Click
        Button_Change_Color(Button_LS_D_Messung)

        Panel_SSA_LS_D_Messung_b.Height = 55
        Panel_SSA_LS_D_Prognose_b.Height = 0   '36

        Projekt.LS_Decke.Untersuchung = Messung

        Update_Klasse_LS_D()
    End Sub

    Private Sub Button_LS_D_nv_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_D_nv.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_36", CType(sender, Control))
    End Sub
    Private Sub Button_LS_D_nv_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_D_nv.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub
    Private Sub Button_LS_D_nv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_D_nv.Click
        Button_Change_Color(Button_LS_D_nv)

        Panel_SSA_LS_D_Messung_b.Height = 0   '90
        Panel_SSA_LS_D_Prognose_b.Height = 0   '36

        Projekt.LS_Decke.Untersuchung = nv

        Update_Klasse_LS_D()
    End Sub
#End Region

#Region "Anzahl Messungen"
    Private Sub Button_LS_D_M_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_D_M_1.Click
        Button_Change_Color(Button_LS_D_M_1)

        Projekt.LS_Decke.Messung.Anzahl = 0
        Messungen_anzeigen(Button_LS_D_M_1)

        Update_Klasse_LS_D()
    End Sub

    Private Sub Button_LS_D_M_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_D_M_2.Click
        Button_Change_Color(Button_LS_D_M_2)

        Projekt.LS_Decke.Messung.Anzahl = 1
        Messungen_anzeigen(Button_LS_D_M_2)

        Update_Klasse_LS_D()
    End Sub

    Private Sub Button_LS_D_M_3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_D_M_3.Click
        Button_Change_Color(Button_LS_D_M_3)

        Projekt.LS_Decke.Messung.Anzahl = 2
        Messungen_anzeigen(Button_LS_D_M_3)

        Update_Klasse_LS_D()
    End Sub

    Private Sub Button_LS_D_M_4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_D_M_4.Click
        Button_Change_Color(Button_LS_D_M_4)

        Projekt.LS_Decke.Messung.Anzahl = 3
        Messungen_anzeigen(Button_LS_D_M_4)

        Update_Klasse_LS_D()
    End Sub

    Private Sub Button_LS_D_M_5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_D_M_5.Click
        Button_Change_Color(Button_LS_D_M_5)

        Projekt.LS_Decke.Messung.Anzahl = 4
        Messungen_anzeigen(Button_LS_D_M_5)

        Update_Klasse_LS_D()
    End Sub

    Private Sub Button_LS_D_M_6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_LS_D_M_6.Click
        Button_Change_Color(Button_LS_D_M_6)

        Projekt.LS_Decke.Messung.Anzahl = 5
        Messungen_anzeigen(Button_LS_D_M_6)

        Update_Klasse_LS_D()
    End Sub
#End Region
#Region "NUD - Messung"
#Region "Validating"
    Private Sub NUD_LS_D_Mes_R_1_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_D_Mes_R_1.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_LS_D_Mes_C_1_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_D_Mes_C_1.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.LS_Decke.Messung.Messung_1.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_LS_D_Mes_R_2_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_D_Mes_R_2.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_LS_D_Mes_C_2_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_D_Mes_C_2.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.LS_Decke.Messung.Messung_2.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_LS_D_Mes_R_3_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_D_Mes_R_3.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_LS_D_Mes_C_3_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_D_Mes_C_3.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.LS_Decke.Messung.Messung_3.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_LS_D_Mes_R_4_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_D_Mes_R_4.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_LS_D_Mes_C_4_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_D_Mes_C_4.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.LS_Decke.Messung.Messung_4.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_LS_D_Mes_R_5_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_D_Mes_R_5.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_LS_D_Mes_C_5_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_D_Mes_C_5.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.LS_Decke.Messung.Messung_5.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_LS_D_Mes_R_6_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_D_Mes_R_6.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_LS_D_Mes_C_6_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_D_Mes_C_6.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.LS_Decke.Messung.Messung_6.C = CType(sender, NumericUpDown).Text
    End Sub
#End Region
#Region "TextChanged"
    Private Sub NUD_LS_D_Mes_R_1_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_D_Mes_R_1.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.LS_Decke.Messung.Messung_1.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.LS_Decke.Messung.Messung_1.Pegel)
        End If

        Update_Klasse_LS_D()
    End Sub

    Private Sub NUD_LS_D_Mes_C_1_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_D_Mes_C_1.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.LS_Decke.Messung.Messung_1.C = menud.Text
        Else
            menud.Text = Projekt.LS_Decke.Messung.Messung_1.C
        End If

        Update_Klasse_LS_D()
    End Sub

    Private Sub NUD_LS_D_Mes_R_2_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_D_Mes_R_2.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.LS_Decke.Messung.Messung_2.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.LS_Decke.Messung.Messung_2.Pegel)
        End If

        Update_Klasse_LS_D()
    End Sub

    Private Sub NUD_LS_D_Mes_C_2_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_D_Mes_C_2.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.LS_Decke.Messung.Messung_2.C = menud.Text
        Else
            menud.Text = Projekt.LS_Decke.Messung.Messung_2.C
        End If

        Update_Klasse_LS_D()
    End Sub

    Private Sub NUD_LS_D_Mes_R_3_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_D_Mes_R_3.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.LS_Decke.Messung.Messung_3.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.LS_Decke.Messung.Messung_3.Pegel)
        End If

        Update_Klasse_LS_D()
    End Sub

    Private Sub NUD_LS_D_Mes_C_3_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_D_Mes_C_3.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.LS_Decke.Messung.Messung_3.C = menud.Text
        Else
            menud.Text = Projekt.LS_Decke.Messung.Messung_3.C
        End If

        Update_Klasse_LS_D()
    End Sub

    Private Sub NUD_LS_D_Mes_R_4_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_D_Mes_R_4.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.LS_Decke.Messung.Messung_4.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.LS_Decke.Messung.Messung_4.Pegel)
        End If

        Update_Klasse_LS_D()
    End Sub

    Private Sub NUD_LS_D_Mes_C_4_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_D_Mes_C_4.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.LS_Decke.Messung.Messung_4.C = menud.Text
        Else
            menud.Text = Projekt.LS_Decke.Messung.Messung_4.C
        End If

        Update_Klasse_LS_D()
    End Sub

    Private Sub NUD_LS_D_Mes_R_5_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_D_Mes_R_5.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.LS_Decke.Messung.Messung_5.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.LS_Decke.Messung.Messung_5.Pegel)
        End If

        Update_Klasse_LS_D()
    End Sub

    Private Sub NUD_LS_D_Mes_C_5_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_D_Mes_C_5.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.LS_Decke.Messung.Messung_5.C = menud.Text
        Else
            menud.Text = Projekt.LS_Decke.Messung.Messung_5.C
        End If

        Update_Klasse_LS_D()
    End Sub

    Private Sub NUD_LS_D_Mes_R_6_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_D_Mes_R_6.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.LS_Decke.Messung.Messung_6.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.LS_Decke.Messung.Messung_6.Pegel)
        End If

        Update_Klasse_LS_D()
    End Sub

    Private Sub NUD_LS_D_Mes_C_6_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_D_Mes_C_6.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.LS_Decke.Messung.Messung_6.C = menud.Text
        Else
            menud.Text = Projekt.LS_Decke.Messung.Messung_6.C
        End If

        Update_Klasse_LS_D()
    End Sub
#End Region
#End Region
#Region "NUD - Prognose"
#Region "Validating"
    Private Sub NUD_LS_D_Prog_R_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_D_Prog_R.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_LS_D_Prog_C_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_LS_D_Prog_C.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.LS_Decke.Prognose.C = CType(sender, NumericUpDown).Text
    End Sub
#End Region
#Region "TextChanged"
    Private Sub NUD_LS_D_Prog_R_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_D_Prog_R.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.LS_Decke.Prognose.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.LS_Decke.Prognose.Pegel)
        End If

        Update_Klasse_LS_D()
    End Sub

    Private Sub NUD_LS_D_Prog_C_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_LS_D_Prog_C.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.LS_Decke.Prognose.C = menud.Text
        Else
            menud.Text = Projekt.LS_Decke.Prognose.C
        End If

        Update_Klasse_LS_D()
    End Sub
#End Region
#End Region
#End Region

#Region "TS Decken - Trittschall Decken"
    Private Sub Update_Klasse_TS_D()
        Me.Button_Klasse_TS_D.Text = Get_Klasse_TS_Decken()
        Me.Button_Klasse_TS_D.BackColor = Get_Klasse_Color(Me.Button_Klasse_TS_D.Text)

        Me.Update_SSA_TS_D()
    End Sub
#Region "Untersuchung"
    Private Sub Button_TS_D_Prognose_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_Prognose.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_37", CType(sender, Control))
    End Sub
    Private Sub Button_TS_D_Prognose_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_Prognose.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub
    Private Sub Button_TS_D_Prognose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_Prognose.Click
        Button_Change_Color(Button_TS_D_Prognose)

        Me.Panel_SSA_TS_D_Messung_b.Height = 0  '90
        Me.Panel_SSA_TS_D_Prognose_b.Height = 73

        Projekt.TS_Decke.Untersuchung = Prognose

        Update_Klasse_TS_D()
    End Sub

    Private Sub Button_TS_D_Messung_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_Messung.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_37", CType(sender, Control))
    End Sub
    Private Sub Button_TS_D_Messung_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_Messung.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub
    Private Sub Button_TS_D_Messung_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_Messung.Click
        Button_Change_Color(Button_TS_D_Messung)

        Me.Panel_SSA_TS_D_Messung_b.Height = 91
        Me.Panel_SSA_TS_D_Prognose_b.Height = 0   '36

        Projekt.TS_Decke.Untersuchung = Messung


        Update_Klasse_TS_D()
    End Sub

    Private Sub Button_TS_D_nv_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_nv.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_36", CType(sender, Control))
    End Sub
    Private Sub Button_TS_D_nv_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_nv.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub
    Private Sub Button_TS_D_nv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_nv.Click
        Button_Change_Color(Button_TS_D_nv)

        Me.Panel_SSA_TS_D_Messung_b.Height = 0   '90
        Me.Panel_SSA_TS_D_Prognose_b.Height = 0   '36

        Projekt.TS_Decke.Untersuchung = nv

        Update_Klasse_TS_D()
    End Sub

    Private Sub Button_TS_D_MV_kleiner_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_MV_kleiner.Click
        Button_Change_Color(Button_TS_D_MV_kleiner)

        Projekt.TS_Decke.fEstrich = fE_kleiner50

        'Me.NUD_TS_D_Mes_C_1.Text = ""
        'Me.NUD_TS_D_Mes_C_1.Visible = False

        Update_Klasse_TS_D()
    End Sub

    Private Sub Button_TS_D_MV_groesser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_MV_groesser.Click
        Button_Change_Color(Button_TS_D_MV_groesser)

        Projekt.TS_Decke.fEstrich = fE_groesser50

        'Me.NUD_TS_D_Mes_C_1.Text = ""
        'Me.NUD_TS_D_Mes_C_1.Visible = False

        Update_Klasse_TS_D()
    End Sub
#End Region
#Region "Messung Estrich"
    Private Sub Button_TS_D_P_kleiner_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_P_kleiner.Click
        Button_Change_Color(Button_TS_D_P_kleiner)

        Projekt.TS_Decke.fEstrich = fE_kleiner50
        Update_Klasse_TS_D()
    End Sub

    Private Sub Button_TS_D_P_groesser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_P_groesser.Click
        Button_Change_Color(Button_TS_D_P_groesser)

        Projekt.TS_Decke.fEstrich = fE_groesser50

        Update_Klasse_TS_D()
    End Sub
#End Region
#Region "Prognose mit/ohone"
    Private Sub Button_TS_D_Prog_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_Prog_mit.Click
        Button_Change_Color(Button_TS_D_Prog_mit)

        Projekt.TS_Decke.Prognose.Bodenbelag = Be_mit

        Update_Klasse_TS_D()
    End Sub

    Private Sub Button_TS_D_Prog_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_Prog_ohne.Click
        Button_Change_Color(Button_TS_D_Prog_ohne)

        Projekt.TS_Decke.Prognose.Bodenbelag = Be_ohne

        Update_Klasse_TS_D()

    End Sub

#End Region
#Region "Messung mit/ohne"
    Private Sub Button_TS_D_MV_1_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_MV_1_mit.Click
        Button_Change_Color(Button_TS_D_MV_1_mit)

        Projekt.TS_Decke.Messung.Messung_1.Bodenbelag = Be_mit

        Update_Klasse_TS_D()
    End Sub

    Private Sub Button_TS_D_MV_1_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_MV_1_ohne.Click
        Button_Change_Color(Button_TS_D_MV_1_ohne)

        Projekt.TS_Decke.Messung.Messung_1.Bodenbelag = Be_ohne

        Update_Klasse_TS_D()
    End Sub

    Private Sub Button_TS_D_MV_2_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_MV_2_mit.Click
        Button_Change_Color(Button_TS_D_MV_2_mit)

        Projekt.TS_Decke.Messung.Messung_2.Bodenbelag = Be_mit

        Update_Klasse_TS_D()
    End Sub

    Private Sub Button_TS_D_MV_2_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_MV_2_ohne.Click
        Button_Change_Color(Button_TS_D_MV_2_ohne)

        Projekt.TS_Decke.Messung.Messung_2.Bodenbelag = Be_ohne

        Update_Klasse_TS_D()
    End Sub

    Private Sub Button_TS_D_MV_3_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_MV_3_mit.Click
        Button_Change_Color(Button_TS_D_MV_3_mit)

        Projekt.TS_Decke.Messung.Messung_3.Bodenbelag = Be_mit

        Update_Klasse_TS_D()
    End Sub

    Private Sub Button_TS_D_MV_3_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_MV_3_ohne.Click
        Button_Change_Color(Button_TS_D_MV_3_ohne)

        Projekt.TS_Decke.Messung.Messung_3.Bodenbelag = Be_ohne

        Update_Klasse_TS_D()
    End Sub

    Private Sub Button_TS_D_MV_4_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_MV_4_mit.Click
        Button_Change_Color(Button_TS_D_MV_4_mit)

        Projekt.TS_Decke.Messung.Messung_4.Bodenbelag = Be_mit

        Update_Klasse_TS_D()
    End Sub

    Private Sub Button_TS_D_MV_4_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_MV_4_ohne.Click
        Button_Change_Color(Button_TS_D_MV_4_ohne)

        Projekt.TS_Decke.Messung.Messung_4.Bodenbelag = Be_ohne

        Update_Klasse_TS_D()
    End Sub

    Private Sub Button_TS_D_MV_5_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_MV_5_mit.Click
        Button_Change_Color(Button_TS_D_MV_5_mit)

        Projekt.TS_Decke.Messung.Messung_5.Bodenbelag = Be_mit

        Update_Klasse_TS_D()
    End Sub

    Private Sub Button_TS_D_MV_5_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_MV_5_ohne.Click
        Button_Change_Color(Button_TS_D_MV_5_ohne)

        Projekt.TS_Decke.Messung.Messung_5.Bodenbelag = Be_ohne

        Update_Klasse_TS_D()
    End Sub

    Private Sub Button_TS_D_MV_6_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_MV_6_mit.Click
        Button_Change_Color(Button_TS_D_MV_6_mit)

        Projekt.TS_Decke.Messung.Messung_6.Bodenbelag = Be_mit

        Update_Klasse_TS_D()
    End Sub

    Private Sub Button_TS_D_MV_6_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_MV_6_ohne.Click
        Button_Change_Color(Button_TS_D_MV_6_ohne)

        Projekt.TS_Decke.Messung.Messung_6.Bodenbelag = Be_ohne

        Update_Klasse_TS_D()
    End Sub
#End Region
#Region "Anzahl Messungen"
    Private Sub Button_TS_D_M_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_M_1.Click
        Button_Change_Color(Button_TS_D_M_1)

        Projekt.TS_Decke.Messung.Anzahl = 0
        Messungen_anzeigen(Button_TS_D_M_1)

        Update_Klasse_TS_D()
    End Sub

    Private Sub Button_TS_D_M_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_M_2.Click
        Button_Change_Color(Button_TS_D_M_2)

        Projekt.TS_Decke.Messung.Anzahl = 1
        Messungen_anzeigen(Button_TS_D_M_2)

        Update_Klasse_TS_D()
    End Sub

    Private Sub Button_TS_D_M_3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_M_3.Click
        Button_Change_Color(Button_TS_D_M_3)

        Projekt.TS_Decke.Messung.Anzahl = 2
        Messungen_anzeigen(Button_TS_D_M_3)

        Update_Klasse_TS_D()
    End Sub

    Private Sub Button_TS_D_M_4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_M_4.Click
        Button_Change_Color(Button_TS_D_M_4)

        Projekt.TS_Decke.Messung.Anzahl = 3
        Messungen_anzeigen(Button_TS_D_M_4)

        Update_Klasse_TS_D()
    End Sub

    Private Sub Button_TS_D_M_5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_M_5.Click
        Button_Change_Color(Button_TS_D_M_5)

        Projekt.TS_Decke.Messung.Anzahl = 4
        Messungen_anzeigen(Button_TS_D_M_5)

        Update_Klasse_TS_D()
    End Sub

    Private Sub Button_TS_D_M_6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_D_M_6.Click
        Button_Change_Color(Button_TS_D_M_6)

        Projekt.TS_Decke.Messung.Anzahl = 5
        Messungen_anzeigen(Button_TS_D_M_6)

        Update_Klasse_TS_D()
    End Sub
#End Region
#Region "NUD - Messungen"
#Region "Validating"
    Private Sub NUD_TS_D_Mes_L_1_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_D_Mes_L_1.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_D_Mes_C_1_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_D_Mes_C_1.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_Decke.Messung.Messung_1.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_TS_D_Mes_L_2_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_D_Mes_L_2.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_D_Mes_C_2_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_D_Mes_C_2.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_Decke.Messung.Messung_2.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_TS_D_Mes_L_3_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_D_Mes_L_3.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_D_Mes_C_3_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_D_Mes_C_3.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_Decke.Messung.Messung_3.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_TS_D_Mes_L_4_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_D_Mes_L_4.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_D_Mes_C_4_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_D_Mes_C_4.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_Decke.Messung.Messung_4.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_TS_D_Mes_L_5_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_D_Mes_L_5.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_D_Mes_C_5_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_D_Mes_C_5.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_Decke.Messung.Messung_5.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_TS_D_Mes_L_6_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_D_Mes_L_6.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_D_Mes_C_6_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_D_Mes_C_6.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_Decke.Messung.Messung_6.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub
#End Region
    Public Sub Check_Bodenbelag(ByRef but_mit As System.Windows.Forms.Button, ByRef but_ohne As System.Windows.Forms.Button, ByVal Untersuchung As String, ByVal Pegel As Single, ByVal Bauteil As Byte)
        Select Case Untersuchung
            Case "TS_D"
                If (Pegel <= 50 And Pegel > 45) Or Pegel <= 35 Then
                    'Button_Change_Color(but_ohne)
                    but_ohne.PerformClick()
                    but_mit.Enabled = False
                Else
                    but_mit.Enabled = True
                End If
            Case "TS_TPLH"
                If (Pegel <= 53 And Pegel > 48 And Bauteil <> Hausflur) Or (Pegel <= 50 And Pegel > 48 And Bauteil = Hausflur) Or Pegel <= 38 Then
                    but_ohne.PerformClick()
                    but_mit.Enabled = False
                Else
                    but_mit.Enabled = True
                End If
            Case "TS_BLT"
                If (Pegel <= 50 And Pegel > 48 And Bauteil <> Balkon) Or (Pegel <= 58 And Pegel > 48 And Bauteil = Balkon) Or Pegel <= 38 Then
                    but_ohne.PerformClick()
                    but_mit.Enabled = False
                Else
                    but_mit.Enabled = True
                End If
        End Select
    End Sub
#Region "TextChanged"
    Private Sub NUD_TS_D_Mes_L_1_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_D_Mes_L_1.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_Decke.Messung.Messung_1.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_Decke.Messung.Messung_1.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_D_MV_1_mit, Me.Button_TS_D_MV_1_ohne, "TS_D", _
                Projekt.TS_Decke.Messung.Messung_1.Pegel.Pegel, 0)

        Update_Klasse_TS_D()
    End Sub

    Private Sub NUD_TS_D_Mes_C_1_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_D_Mes_C_1.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_Decke.Messung.Messung_1.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_Decke.Messung.Messung_1.Pegel.C
        End If

        Update_Klasse_TS_D()
    End Sub

    Private Sub NUD_TS_D_Mes_L_2_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_D_Mes_L_2.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_Decke.Messung.Messung_2.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_Decke.Messung.Messung_2.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_D_MV_2_mit, Me.Button_TS_D_MV_2_ohne, "TS_D", _
                Projekt.TS_Decke.Messung.Messung_2.Pegel.Pegel, 0)

        Update_Klasse_TS_D()
    End Sub

    Private Sub NUD_TS_D_Mes_C_2_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_D_Mes_C_2.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_Decke.Messung.Messung_2.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_Decke.Messung.Messung_2.Pegel.C
        End If

        Update_Klasse_TS_D()
    End Sub

    Private Sub NUD_TS_D_Mes_L_3_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_D_Mes_L_3.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_Decke.Messung.Messung_3.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_Decke.Messung.Messung_3.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_D_MV_3_mit, Me.Button_TS_D_MV_3_ohne, "TS_D", _
                Projekt.TS_Decke.Messung.Messung_3.Pegel.Pegel, 0)

        Update_Klasse_TS_D()
    End Sub

    Private Sub NUD_TS_D_Mes_C_3_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_D_Mes_C_3.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_Decke.Messung.Messung_3.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_Decke.Messung.Messung_3.Pegel.C
        End If

        Update_Klasse_TS_D()
    End Sub

    Private Sub NUD_TS_D_Mes_L_4_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_D_Mes_L_4.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_Decke.Messung.Messung_4.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_Decke.Messung.Messung_4.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_D_MV_4_mit, Me.Button_TS_D_MV_4_ohne, "TS_D", _
                Projekt.TS_Decke.Messung.Messung_4.Pegel.Pegel, 0)

        Update_Klasse_TS_D()
    End Sub

    Private Sub NUD_TS_D_Mes_C_4_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_D_Mes_C_4.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_Decke.Messung.Messung_4.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_Decke.Messung.Messung_4.Pegel.C
        End If

        Update_Klasse_TS_D()
    End Sub

    Private Sub NUD_TS_D_Mes_L_5_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_D_Mes_L_5.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_Decke.Messung.Messung_5.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_Decke.Messung.Messung_5.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_D_MV_5_mit, Me.Button_TS_D_MV_5_ohne, "TS_D", _
                Projekt.TS_Decke.Messung.Messung_5.Pegel.Pegel, 0)

        Update_Klasse_TS_D()
    End Sub

    Private Sub NUD_TS_D_Mes_C_5_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_D_Mes_C_5.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_Decke.Messung.Messung_5.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_Decke.Messung.Messung_5.Pegel.C
        End If

        Update_Klasse_TS_D()
    End Sub

    Private Sub NUD_TS_D_Mes_L_6_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_D_Mes_L_6.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_Decke.Messung.Messung_6.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_Decke.Messung.Messung_6.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_D_MV_6_mit, Me.Button_TS_D_MV_6_ohne, "TS_D", _
               Projekt.TS_Decke.Messung.Messung_6.Pegel.Pegel, 0)

        Update_Klasse_TS_D()
    End Sub

    Private Sub NUD_TS_D_Mes_C_6_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_D_Mes_C_6.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_Decke.Messung.Messung_6.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_Decke.Messung.Messung_6.Pegel.C
        End If

        Update_Klasse_TS_D()
    End Sub
#End Region
#End Region
#Region "NUD - Prognose"
#Region "Validating"
    Private Sub NUD_TS_D_Prog_L_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_D_Prog_L.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_D_Prog_C_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_D_Prog_C.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_Decke.Prognose.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub
#End Region
#Region "TextChanged"
    Private Sub NUD_TS_D_Prog_L_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_D_Prog_L.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_Decke.Prognose.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_Decke.Prognose.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_D_Prog_mit, Me.Button_TS_D_Prog_ohne, "TS_D", _
                Projekt.TS_Decke.Prognose.Pegel.Pegel, 0)

        Update_Klasse_TS_D()
    End Sub

    Private Sub NUD_TS_D_Prog_C_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_D_Prog_C.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_Decke.Prognose.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_Decke.Prognose.Pegel.C
        End If

        Update_Klasse_TS_D()
    End Sub
#End Region
#End Region
#End Region

#Region "TS TPH - Trittschall Treppen - Podeste - Hausflure"
    Private Sub Update_Klasse_TS_TPLH()
        Me.Button_Klasse_TS_TPH.Text = Get_Klasse_TS_TPLH()
        Me.Button_Klasse_TS_TPH.BackColor = Get_Klasse_Color(Me.Button_Klasse_TS_TPH.Text)

        Me.Update_SSA_TS_TPLH()
    End Sub
#Region "Untersuchung"
    Private Sub Button_TS_TPH_Prognose_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_Prognose.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help_2("ssa_help_37", "ssa_help_59", CType(sender, Control))
    End Sub
    Private Sub Button_TS_TPH_Prognose_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_Prognose.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub
    Private Sub Button_TS_TPH_Prognose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_Prognose.Click
        Button_Change_Color(Button_TS_TPH_Prognose)

        Me.Panel_SSA_TS_TPHf_Messung_b.Height = 0    '124
        Me.Panel_SSA_TS_TPHf_Prognose_b.Height = 80

        Projekt.TS_TPLH.Untersuchung = Prognose

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_TPH_Messung_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_Messung.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help_2("ssa_help_37", "ssa_help_59", CType(sender, Control))
    End Sub
    Private Sub Button_TS_TPH_Messung_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_Messung.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub
    Private Sub Button_TS_TPH_Messung_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_Messung.Click
        Button_Change_Color(Button_TS_TPH_Messung)

        Me.Panel_SSA_TS_TPHf_Messung_b.Height = 124
        Me.Panel_SSA_TS_TPHf_Prognose_b.Height = 0    '62

        Projekt.TS_TPLH.Untersuchung = Messung

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_TPH_nv_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_nv.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_36", CType(sender, Control))
    End Sub
    Private Sub Button_TS_TPH_nv_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_nv.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub
    Private Sub Button_TS_TPH_nv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_nv.Click
        Button_Change_Color(Button_TS_TPH_nv)

        Me.Panel_SSA_TS_TPHf_Messung_b.Height = 0 '124
        Me.Panel_SSA_TS_TPHf_Prognose_b.Height = 0    '62

        Projekt.TS_TPLH.Untersuchung = nv

        Update_Klasse_TS_TPLH()
    End Sub
#End Region

#Region "TPH Prognose mit/ohne"

    Private Sub Button_TS_Treppe_Prog_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_Treppe_Prog_mit.Click
        Button_Change_Color(Button_TS_Treppe_Prog_mit)

        Projekt.TS_TPLH.Prognose.Treppe.Bodenbelag = Be_mit

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_Treppe_Prog_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_Treppe_Prog_ohne.Click
        Button_Change_Color(Button_TS_Treppe_Prog_ohne)

        Projekt.TS_TPLH.Prognose.Treppe.Bodenbelag = Be_ohne

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_Podest_Prog_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_Podest_Prog_mit.Click
        Button_Change_Color(Button_TS_Podest_Prog_mit)

        Projekt.TS_TPLH.Prognose.Podest.Bodenbelag = Be_mit

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_Podest_Prog_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_Podest_Prog_ohne.Click
        Button_Change_Color(Button_TS_Podest_Prog_ohne)

        Projekt.TS_TPLH.Prognose.Podest.Bodenbelag = Be_ohne

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_Laubengang_Prog_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_Laubengang_Prog_mit.Click
        Button_Change_Color(Button_TS_Laubengang_Prog_mit)

        Projekt.TS_TPLH.Prognose.Laube.Bodenbelag = Be_mit

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_Laubengang_Prog_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_Laubengang_Prog_ohne.Click
        Button_Change_Color(Button_TS_Laubengang_Prog_ohne)

        Projekt.TS_TPLH.Prognose.Laube.Bodenbelag = Be_ohne

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_Hausflur_Prog_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_Hausflur_Prog_mit.Click
        Button_Change_Color(Button_TS_Hausflur_Prog_mit)

        Projekt.TS_TPLH.Prognose.Hausflur.Bodenbelag = Be_mit

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_Button_TS_Hausflur_Prog_ohne_Prog_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_Hausflur_Prog_ohne.Click
        Button_Change_Color(Button_TS_Hausflur_Prog_ohne)

        Projekt.TS_TPLH.Prognose.Hausflur.Bodenbelag = Be_ohne

        Update_Klasse_TS_TPLH()
    End Sub
#End Region
#Region "TPH Messung mit/ohne"
    Private Sub Button_TS_THP_MV_1_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_MV_1_mit.Click
        Button_Change_Color(Button_TS_TPH_MV_1_mit)

        Projekt.TS_TPlH.Messung.Messung_1.Bodenbelag = Be_mit

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_THP_MV_1_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_MV_1_ohne.Click
        Button_Change_Color(Button_TS_TPH_MV_1_ohne)

        Projekt.TS_TPlH.Messung.Messung_1.Bodenbelag = Be_ohne

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_THP_MV_2_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_MV_2_mit.Click
        Button_Change_Color(Button_TS_TPH_MV_2_mit)

        Projekt.TS_TPlH.Messung.Messung_2.Bodenbelag = Be_mit

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_THP_MV_2_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_MV_2_ohne.Click
        Button_Change_Color(Button_TS_TPH_MV_2_ohne)

        Projekt.TS_TPlH.Messung.Messung_2.Bodenbelag = Be_ohne

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_THP_MV_3_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_MV_3_mit.Click
        Button_Change_Color(Button_TS_TPH_MV_3_mit)

        Projekt.TS_TPlH.Messung.Messung_3.Bodenbelag = Be_mit

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_THP_MV_3_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_MV_3_ohne.Click
        Button_Change_Color(Button_TS_TPH_MV_3_ohne)

        Projekt.TS_TPlH.Messung.Messung_3.Bodenbelag = Be_ohne

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_THP_MV_4_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_MV_4_mit.Click
        Button_Change_Color(Button_TS_TPH_MV_4_mit)

        Projekt.TS_TPlH.Messung.Messung_4.Bodenbelag = Be_mit

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_THP_MV_4_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_MV_4_ohne.Click
        Button_Change_Color(Button_TS_TPH_MV_4_ohne)

        Projekt.TS_TPlH.Messung.Messung_4.Bodenbelag = Be_ohne

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_THP_MV_5_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_MV_5_mit.Click
        Button_Change_Color(Button_TS_TPH_MV_5_mit)

        Projekt.TS_TPlH.Messung.Messung_5.Bodenbelag = Be_mit

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_THP_MV_5_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_MV_5_ohne.Click
        Button_Change_Color(Button_TS_TPH_MV_5_ohne)

        Projekt.TS_TPlH.Messung.Messung_5.Bodenbelag = Be_ohne

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_THP_MV_6_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_MV_6_mit.Click
        Button_Change_Color(Button_TS_TPH_MV_6_mit)

        Projekt.TS_TPlH.Messung.Messung_6.Bodenbelag = Be_mit

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_THP_MV_6_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_MV_6_ohne.Click
        Button_Change_Color(Button_TS_TPH_MV_6_ohne)

        Projekt.TS_TPlH.Messung.Messung_6.Bodenbelag = Be_ohne

        Update_Klasse_TS_TPLH()
    End Sub
#End Region
#Region "Messung"
#Region "Anzahl Messungen"
    Private Sub Button_TS_TPH_M_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_M_1.Click
        Button_Change_Color(Button_TS_TPH_M_1)

        Projekt.TS_TPlH.Messung.Anzahl = 0
        Messungen_anzeigen(Button_TS_TPH_M_1)

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_TPH_M_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_M_2.Click
        Button_Change_Color(Button_TS_TPH_M_2)

        Projekt.TS_TPlH.Messung.Anzahl = 1
        Messungen_anzeigen(Button_TS_TPH_M_2)

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_TPH_M_3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_M_3.Click
        Button_Change_Color(Button_TS_TPH_M_3)

        Projekt.TS_TPlH.Messung.Anzahl = 2
        Messungen_anzeigen(Button_TS_TPH_M_3)

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_TPH_M_4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_M_4.Click
        Button_Change_Color(Button_TS_TPH_M_4)

        Projekt.TS_TPlH.Messung.Anzahl = 3
        Messungen_anzeigen(Button_TS_TPH_M_4)

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_TPH_M_5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_M_5.Click
        Button_Change_Color(Button_TS_TPH_M_5)

        Projekt.TS_TPlH.Messung.Anzahl = 4
        Messungen_anzeigen(Button_TS_TPH_M_5)

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub Button_TS_TPH_M_6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_M_6.Click
        Button_Change_Color(Button_TS_TPH_M_6)

        Projekt.TS_TPlH.Messung.Anzahl = 5
        Messungen_anzeigen(Button_TS_TPH_M_6)

        Update_Klasse_TS_TPLH()
    End Sub
#End Region
#Region "Bauteil - TPLH"
    Private Sub Button_TS_TPH_BT_T_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_BT_T_1.Click
        Button_Change_Color(Button_TS_TPH_BT_T_1)

        Projekt.TS_TPlH.Messung.Messung_1.Bauteil = Treppe
    End Sub

    Private Sub Button_TS_TPH_BT_P_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_BT_P_1.Click
        Button_Change_Color(Button_TS_TPH_BT_P_1)

        Projekt.TS_TPlH.Messung.Messung_1.Bauteil = Podest
    End Sub

    Private Sub Button_TS_TPH_BT_H_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_BT_H_1.Click
        Button_Change_Color(Button_TS_TPH_BT_H_1)

        Projekt.TS_TPlH.Messung.Messung_1.Bauteil = Hausflur
    End Sub

    Private Sub Button_TS_TPH_BT_T_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_BT_T_2.Click
        Button_Change_Color(Button_TS_TPH_BT_T_2)

        Projekt.TS_TPlH.Messung.Messung_2.Bauteil = Treppe
    End Sub

    Private Sub Button_TS_TPH_BT_P_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_BT_P_2.Click
        Button_Change_Color(Button_TS_TPH_BT_P_2)

        Projekt.TS_TPlH.Messung.Messung_2.Bauteil = Podest
    End Sub

    Private Sub Button_TS_TPH_BT_H_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_BT_H_2.Click
        Button_Change_Color(Button_TS_TPH_BT_H_2)

        Projekt.TS_TPlH.Messung.Messung_2.Bauteil = Hausflur
    End Sub

    Private Sub Button_TS_TPH_BT_T_3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_BT_T_3.Click
        Button_Change_Color(Button_TS_TPH_BT_T_3)

        Projekt.TS_TPlH.Messung.Messung_3.Bauteil = Treppe
    End Sub

    Private Sub Button_TS_TPH_BT_P_3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_BT_P_3.Click
        Button_Change_Color(Button_TS_TPH_BT_P_3)

        Projekt.TS_TPlH.Messung.Messung_3.Bauteil = Podest
    End Sub

    Private Sub Button_TS_TPH_BT_H_3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_BT_H_3.Click
        Button_Change_Color(Button_TS_TPH_BT_H_3)

        Projekt.TS_TPlH.Messung.Messung_3.Bauteil = Hausflur
    End Sub

    Private Sub Button_TS_TPH_BT_T_4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_BT_T_4.Click
        Button_Change_Color(Button_TS_TPH_BT_T_4)

        Projekt.TS_TPlH.Messung.Messung_4.Bauteil = Treppe
    End Sub

    Private Sub Button_TS_TPH_BT_P_4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_BT_P_4.Click
        Button_Change_Color(Button_TS_TPH_BT_P_4)

        Projekt.TS_TPlH.Messung.Messung_4.Bauteil = Podest
    End Sub

    Private Sub Button_TS_TPH_BT_H_4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_BT_H_4.Click
        Button_Change_Color(Button_TS_TPH_BT_H_4)

        Projekt.TS_TPlH.Messung.Messung_4.Bauteil = Hausflur
    End Sub

    Private Sub Button_TS_TPH_BT_T_5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_BT_T_5.Click
        Button_Change_Color(Button_TS_TPH_BT_T_5)

        Projekt.TS_TPlH.Messung.Messung_5.Bauteil = Treppe
    End Sub

    Private Sub Button_TS_TPH_BT_P_5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_BT_P_5.Click
        Button_Change_Color(Button_TS_TPH_BT_P_5)

        Projekt.TS_TPlH.Messung.Messung_5.Bauteil = Podest
    End Sub

    Private Sub Button_TS_TPH_BT_H_5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_BT_H_5.Click
        Button_Change_Color(Button_TS_TPH_BT_H_5)

        Projekt.TS_TPlH.Messung.Messung_5.Bauteil = Hausflur
    End Sub

    Private Sub Button_TS_TPH_BT_T_6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_BT_T_6.Click
        Button_Change_Color(Button_TS_TPH_BT_T_6)

        Projekt.TS_TPlH.Messung.Messung_6.Bauteil = Treppe
    End Sub

    Private Sub Button_TS_TPH_BT_P_6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_BT_P_6.Click
        Button_Change_Color(Button_TS_TPH_BT_P_6)

        Projekt.TS_TPlH.Messung.Messung_6.Bauteil = Podest
    End Sub

    Private Sub Button_TS_TPH_BT_H_6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPH_BT_H_6.Click
        Button_Change_Color(Button_TS_TPH_BT_H_6)

        Projekt.TS_TPlH.Messung.Messung_6.Bauteil = Hausflur
    End Sub

    Private Sub Button_TS_TPLH_BT_La_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPLH_BT_La_1.Click
        Button_Change_Color(Button_TS_TPLH_BT_La_1)

        Projekt.TS_TPLH.Messung.Messung_1.Bauteil = Laube
    End Sub

    Private Sub Button_TS_TPLH_BT_La_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPLH_BT_La_2.Click
        Button_Change_Color(Button_TS_TPLH_BT_La_2)

        Projekt.TS_TPLH.Messung.Messung_2.Bauteil = Laube
    End Sub

    Private Sub Button_TS_TPLH_BT_La_3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPLH_BT_La_3.Click
        Button_Change_Color(Button_TS_TPLH_BT_La_3)

        Projekt.TS_TPLH.Messung.Messung_3.Bauteil = Laube
    End Sub

    Private Sub Button_TS_TPLH_BT_La_4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPLH_BT_La_4.Click
        Button_Change_Color(Button_TS_TPLH_BT_La_4)

        Projekt.TS_TPLH.Messung.Messung_4.Bauteil = Laube
    End Sub

    Private Sub Button_TS_TPLH_BT_La_5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPLH_BT_La_5.Click
        Button_Change_Color(Button_TS_TPLH_BT_La_5)

        Projekt.TS_TPLH.Messung.Messung_5.Bauteil = Laube
    End Sub
    Private Sub Button_TS_TPLH_BT_La_6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_TPLH_BT_La_6.Click
        Button_Change_Color(Button_TS_TPLH_BT_La_6)

        Projekt.TS_TPLH.Messung.Messung_6.Bauteil = Laube
    End Sub
#End Region
#Region "NUD - Messungen"
#Region "Validating"
    Private Sub NUD_TS_TPH_Mes_L_1_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_TPLH_Mes_L_1.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_TPH_Mes_C_1_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_TPLH_Mes_C_1.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_TPlH.Messung.Messung_1.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_TS_TPH_Mes_L_2_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_TPLH_Mes_L_2.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_TPH_Mes_C_2_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_TPLH_Mes_C_2.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_TPlH.Messung.Messung_2.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_TS_TPH_Mes_L_3_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_TPLH_Mes_L_3.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_TPH_Mes_C_3_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_TPLH_Mes_C_3.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_TPlH.Messung.Messung_3.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_TS_TPH_Mes_L_4_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_TPLH_Mes_L_4.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_TPH_Mes_C_4_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_TPLH_Mes_C_4.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_TPlH.Messung.Messung_4.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_TS_TPH_Mes_L_5_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_TPLH_Mes_L_5.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_TPH_Mes_C_5_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_TPLH_Mes_C_5.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_TPlH.Messung.Messung_5.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_TS_TPH_Mes_L_6_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_TPLH_Mes_L_6.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_TPH_Mes_C_6_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_TPLH_Mes_C_6.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_TPlH.Messung.Messung_6.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub
#End Region
#Region "TextChanged"
    Private Sub NUD_TS_TPH_Mes_L_1_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_TPLH_Mes_L_1.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_TPlH.Messung.Messung_1.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_TPlH.Messung.Messung_1.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_TPH_MV_1_mit, Me.Button_TS_TPH_MV_1_ohne, "TS_TPLH", _
                Projekt.TS_TPLH.Messung.Messung_1.Pegel.Pegel, Projekt.TS_TPLH.Messung.Messung_1.Bauteil)

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub NUD_TS_TPH_Mes_C_1_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_TPLH_Mes_C_1.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_TPlH.Messung.Messung_1.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_TPlH.Messung.Messung_1.Pegel.C
        End If

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub NUD_TS_TPH_Mes_L_2_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_TPLH_Mes_L_2.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_TPlH.Messung.Messung_2.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_TPlH.Messung.Messung_2.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_TPH_MV_2_mit, Me.Button_TS_TPH_MV_2_ohne, "TS_TPLH", _
                 Projekt.TS_TPLH.Messung.Messung_2.Pegel.Pegel, Projekt.TS_TPLH.Messung.Messung_2.Bauteil)

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub NUD_TS_TPH_Mes_C_2_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_TPLH_Mes_C_2.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_TPlH.Messung.Messung_2.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_TPlH.Messung.Messung_2.Pegel.C
        End If

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub NUD_TS_TPH_Mes_L_3_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_TPLH_Mes_L_3.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_TPlH.Messung.Messung_3.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_TPlH.Messung.Messung_3.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_TPH_MV_3_mit, Me.Button_TS_TPH_MV_3_ohne, "TS_TPLH", _
                Projekt.TS_TPLH.Messung.Messung_3.Pegel.Pegel, Projekt.TS_TPLH.Messung.Messung_3.Bauteil)

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub NUD_TS_TPH_Mes_C_3_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_TPLH_Mes_C_3.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_TPlH.Messung.Messung_3.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_TPlH.Messung.Messung_3.Pegel.C
        End If

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub NUD_TS_TPH_Mes_L_4_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_TPLH_Mes_L_4.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_TPlH.Messung.Messung_4.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_TPlH.Messung.Messung_4.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_TPH_MV_4_mit, Me.Button_TS_TPH_MV_4_ohne, "TS_TPLH", _
                Projekt.TS_TPLH.Messung.Messung_4.Pegel.Pegel, Projekt.TS_TPLH.Messung.Messung_4.Bauteil)

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub NUD_TS_TPH_Mes_C_4_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_TPLH_Mes_C_4.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_TPlH.Messung.Messung_4.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_TPlH.Messung.Messung_4.Pegel.C
        End If

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub NUD_TS_TPH_Mes_L_5_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_TPLH_Mes_L_5.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_TPlH.Messung.Messung_5.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_TPlH.Messung.Messung_5.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_TPH_MV_5_mit, Me.Button_TS_TPH_MV_5_ohne, "TS_TPLH", _
                Projekt.TS_TPLH.Messung.Messung_5.Pegel.Pegel, Projekt.TS_TPLH.Messung.Messung_5.Bauteil)

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub NUD_TS_TPH_Mes_C_5_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_TPLH_Mes_C_5.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_TPlH.Messung.Messung_5.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_TPlH.Messung.Messung_5.Pegel.C
        End If

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub NUD_TS_TPH_Mes_L_6_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_TPLH_Mes_L_6.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_TPlH.Messung.Messung_6.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_TPlH.Messung.Messung_6.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_TPH_MV_6_mit, Me.Button_TS_TPH_MV_6_ohne, "TS_TPLH", _
                Projekt.TS_TPLH.Messung.Messung_6.Pegel.Pegel, Projekt.TS_TPLH.Messung.Messung_6.Bauteil)

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub NUD_TS_TPH_Mes_C_6_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_TPLH_Mes_C_6.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_TPlH.Messung.Messung_6.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_TPlH.Messung.Messung_6.Pegel.C
        End If

        Update_Klasse_TS_TPLH()
    End Sub
#End Region
#End Region
#End Region
#Region "Prognose"
    '#Region "Bauteil - TPH"
    '    Private Sub Button_TS_TPH_BT_T_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '        Button_Change_Color(Button_TS_TPH_BT_T)

    '        Projekt.TS_TPH.Prognose.Bauteil = Treppe
    '    End Sub

    '    Private Sub Button_TS_TPH_BT_P_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '        Button_Change_Color(Button_TS_TPH_BT_P)

    '        Projekt.TS_TPH.Prognose.Bauteil = Podest
    '    End Sub

    '    Private Sub Button_TS_TPH_BT_H_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '        Button_Change_Color(Button_TS_TPH_BT_H)

    '        Projekt.TS_TPH.Prognose.Bauteil = Hausflur
    '    End Sub
    '#End Region
#Region "NUD - Prognose"
#Region "Validating"
    Private Sub NUD_TS_TPH_Prog_T_L_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_TPLH_Prog_T_L.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_TPH_Prog_T_C_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_TPLH_Prog_T_C.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_TPlH.Prognose.Treppe.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_TS_TPH_Prog_P_L_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_TPLH_Prog_P_L.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_TPH_Prog_P_C_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_TPLH_Prog_P_C.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_TPlH.Prognose.Podest.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_TS_TPLH_Prog_La_L_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_TPLH_Prog_L_L.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_TPLH_Prog_La_C_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_TPLH_Prog_L_C.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_TPLH.Prognose.Laube.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub
    Private Sub NUD_TS_TPH_Prog_H_L_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_TPLH_Prog_H_L.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_TPH_Prog_H_C_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_TPLH_Prog_H_C.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_TPlH.Prognose.Hausflur.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub
#End Region
#Region "TextChanged"
    Private Sub NUD_TS_TPH_Prog_T_L_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_TPLH_Prog_T_L.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_TPlH.Prognose.Treppe.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_TPlH.Prognose.Treppe.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_Treppe_Prog_mit, Me.Button_TS_Treppe_Prog_ohne, "TS_TPLH", _
                Projekt.TS_TPLH.Prognose.Treppe.Pegel.Pegel, Treppe)

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub NUD_TS_TPH_Prog_T_C_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_TPLH_Prog_T_C.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_TPlH.Prognose.Treppe.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_TPlH.Prognose.Treppe.Pegel.C
        End If

        Update_Klasse_TS_TPLH()
    End Sub
    Private Sub NUD_TS_TPH_Prog_P_L_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_TPLH_Prog_P_L.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_TPLH.Prognose.Podest.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_TPLH.Prognose.Podest.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_Podest_Prog_mit, Me.Button_TS_Podest_Prog_ohne, "TS_TPLH", _
                Projekt.TS_TPLH.Prognose.Podest.Pegel.Pegel, Podest)

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub NUD_TS_TPH_Prog_P_C_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_TPLH_Prog_P_C.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_TPLH.Prognose.Podest.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_TPLH.Prognose.Podest.Pegel.C
        End If

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub NUD_TS_BLLT_Prog_La_L_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_TPLH_Prog_L_L.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_TPLH.Prognose.Laube.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_TPLH.Prognose.Laube.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_Laubengang_Prog_mit, Me.Button_TS_Laubengang_Prog_ohne, "TS_TPLH", _
                Projekt.TS_TPLH.Prognose.Laube.Pegel.Pegel, Laube)

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub NUD_TS_BLLT_Prog_La_C_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_TPLH_Prog_L_C.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_TPLH.Prognose.Laube.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_TPLH.Prognose.Laube.Pegel.C
        End If

        Update_Klasse_TS_TPLH()
    End Sub
    Private Sub NUD_TS_TPH_Prog_H_L_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_TPLH_Prog_H_L.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_TPLH.Prognose.Hausflur.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_TPLH.Prognose.Hausflur.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_Hausflur_Prog_mit, Me.Button_TS_Hausflur_Prog_ohne, "TS_TPLH", _
                Projekt.TS_TPLH.Prognose.Hausflur.Pegel.Pegel, Hausflur)

        Update_Klasse_TS_TPLH()
    End Sub

    Private Sub NUD_TS_TPH_Prog_H_C_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_TPLH_Prog_H_C.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_TPLH.Prognose.Hausflur.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_TPLH.Prognose.Hausflur.Pegel.C
        End If

        Update_Klasse_TS_TPLH()
    End Sub
#End Region
#End Region
#End Region
#End Region

#Region "TS BLLT - Trittschall Balkon - Laubengang - Loggie - Terrasse"
    Private Sub Update_Klasse_TS_BLT()
        Me.Button_Klasse_TS_BLLT.Text = Get_Klasse_TS_BLT()
        Me.Button_Klasse_TS_BLLT.BackColor = Get_Klasse_Color(Me.Button_Klasse_TS_BLLT.Text)

        Me.Update_SSA_TS_BLT()
    End Sub
#Region "Untersuchung"
    Private Sub Button_TS_BLLT_Prognose_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_Prognose.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help_2("ssa_help_37", "ssa_help_59", CType(sender, Control))
    End Sub
    Private Sub Button_TS_BLLT_Prognose_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_Prognose.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub
    Private Sub Button_TS_BLLT_Prognose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_Prognose.Click
        Button_Change_Color(Button_TS_BLLT_Prognose)

        Me.Panel_SSA_TS_BLLT_Messung_b.Height = 0     '124
        Me.Panel_SSA_TS_BLLT_Prognose_b.Height = 80

        Projekt.TS_BLT.Untersuchung = Prognose

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub Button_TS_BLLT_Messung_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_Messung.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help_2("ssa_help_37", "ssa_help_59", CType(sender, Control))
    End Sub
    Private Sub Button_TS_BLLT_Messung_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_Messung.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub
    Private Sub Button_TS_BLLT_Messung_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_Messung.Click
        Button_Change_Color(Button_TS_BLLT_Messung)

        Me.Panel_SSA_TS_BLLT_Messung_b.Height = 124
        Me.Panel_SSA_TS_BLLT_Prognose_b.Height = 0    '62

        Projekt.TS_BLT.Untersuchung = Messung

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub Button_TS_BLLT_nv_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_nv.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_36", CType(sender, Control))
    End Sub
    Private Sub Button_TS_BLLT_nv_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_nv.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub
    Private Sub Button_TS_BLLT_nv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_nv.Click
        Button_Change_Color(Button_TS_BLLT_nv)

        Me.Panel_SSA_TS_BLLT_Messung_b.Height = 0 '124
        Me.Panel_SSA_TS_BLLT_Prognose_b.Height = 0    '62

        Projekt.TS_BLT.Untersuchung = nv

        Update_Klasse_TS_BLT()
    End Sub
#End Region


#Region "BLT Prognose mit/ohne"

    Private Sub Button_TS_Balkon_Prog_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_Balkon_Prog_mit.Click
        Button_Change_Color(Button_TS_Balkon_Prog_mit)

        Projekt.TS_BLT.Prognose.Balkon.Bodenbelag = Be_mit

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub Button_TS_Balkon_Prog_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_Balkon_Prog_ohne.Click
        Button_Change_Color(Button_TS_Balkon_Prog_ohne)

        Projekt.TS_BLT.Prognose.Balkon.Bodenbelag = Be_ohne

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub Button_TS_Loggia_Prog_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_Loggia_Prog_mit.Click
        Button_Change_Color(Button_TS_Loggia_Prog_mit)

        Projekt.TS_BLT.Prognose.Loggia.Bodenbelag = Be_mit

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub Button_TS_Loggia_Prog_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_Loggia_Prog_ohne.Click
        Button_Change_Color(Button_TS_Loggia_Prog_ohne)

        Projekt.TS_BLT.Prognose.Loggia.Bodenbelag = Be_ohne

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub Button_TS_Terrasse_Prog_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_Terrasse_Prog_mit.Click
        Button_Change_Color(Button_TS_Terrasse_Prog_mit)

        Projekt.TS_BLT.Prognose.Terrasse.Bodenbelag = Be_mit

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub Button_TS_Terrasse_Prog_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_Terrasse_Prog_ohne.Click
        Button_Change_Color(Button_TS_Terrasse_Prog_ohne)

        Projekt.TS_BLT.Prognose.Terrasse.Bodenbelag = Be_ohne

        Update_Klasse_TS_BLT()
    End Sub
#End Region
#Region "TPH Messung mit/ohne"
    Private Sub Button_TS_BLLT_MV_1_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_MV_1_mit.Click
        Button_Change_Color(Button_TS_BLLT_MV_1_mit)

        Projekt.TS_BLT.Messung.Messung_1.Bodenbelag = Be_mit

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub Button_TS_BLLT_MV_1_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_MV_1_ohne.Click
        Button_Change_Color(Button_TS_BLLT_MV_1_ohne)

        Projekt.TS_BLT.Messung.Messung_1.Bodenbelag = Be_ohne

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub Button_TS_BLLT_MV_2_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_MV_2_mit.Click
        Button_Change_Color(Button_TS_BLLT_MV_2_mit)

        Projekt.TS_BLT.Messung.Messung_2.Bodenbelag = Be_mit

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub Button_TS_BLLT_MV_2_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_MV_2_ohne.Click
        Button_Change_Color(Button_TS_BLLT_MV_2_ohne)

        Projekt.TS_BLT.Messung.Messung_2.Bodenbelag = Be_ohne

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub Button_TS_BLLT_MV_3_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_MV_3_mit.Click
        Button_Change_Color(Button_TS_BLLT_MV_3_mit)

        Projekt.TS_BLT.Messung.Messung_3.Bodenbelag = Be_mit

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub Button_TS_BLLT_MV_3_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_MV_3_ohne.Click
        Button_Change_Color(Button_TS_BLLT_MV_3_ohne)

        Projekt.TS_BLT.Messung.Messung_3.Bodenbelag = Be_ohne

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub Button_TS_BLLT_MV_4_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_MV_4_mit.Click
        Button_Change_Color(Button_TS_BLLT_MV_4_mit)

        Projekt.TS_BLT.Messung.Messung_4.Bodenbelag = Be_mit

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub Button_TS_BLLT_MV_4_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_MV_4_ohne.Click
        Button_Change_Color(Button_TS_BLLT_MV_4_ohne)

        Projekt.TS_BLT.Messung.Messung_4.Bodenbelag = Be_ohne

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub Button_TS_BLLT_MV_5_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_MV_5_mit.Click
        Button_Change_Color(Button_TS_BLLT_MV_5_mit)

        Projekt.TS_BLT.Messung.Messung_5.Bodenbelag = Be_mit

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub Button_TS_BLLT_MV_5_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_MV_5_ohne.Click
        Button_Change_Color(Button_TS_BLLT_MV_5_ohne)

        Projekt.TS_BLT.Messung.Messung_5.Bodenbelag = Be_ohne

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub Button_TS_BLLT_MV_6_mit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_MV_6_mit.Click
        Button_Change_Color(Button_TS_BLLT_MV_6_mit)

        Projekt.TS_BLT.Messung.Messung_6.Bodenbelag = Be_mit

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub Button_TS_BLLT_MV_6_ohne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_MV_6_ohne.Click
        Button_Change_Color(Button_TS_BLLT_MV_6_ohne)

        Projekt.TS_BLT.Messung.Messung_6.Bodenbelag = Be_ohne

        Update_Klasse_TS_BLT()
    End Sub
#End Region

#Region "Messung"
#Region "Anzahl Messungen"
    Private Sub Button_TS_BLLT_M_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_M_1.Click
        Button_Change_Color(Button_TS_BLLT_M_1)

        Projekt.TS_BLT.Messung.Anzahl = 0
        Messungen_anzeigen(Button_TS_BLLT_M_1)

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub Button_TS_BLLT_M_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_M_2.Click
        Button_Change_Color(Button_TS_BLLT_M_2)

        Projekt.TS_BLT.Messung.Anzahl = 1
        Messungen_anzeigen(Button_TS_BLLT_M_2)

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub Button_TS_BLLT_M_3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_M_3.Click
        Button_Change_Color(Button_TS_BLLT_M_3)

        Projekt.TS_BLT.Messung.Anzahl = 2
        Messungen_anzeigen(Button_TS_BLLT_M_3)

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub Button_TS_BLLT_M_4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_M_4.Click
        Button_Change_Color(Button_TS_BLLT_M_4)

        Projekt.TS_BLT.Messung.Anzahl = 3
        Messungen_anzeigen(Button_TS_BLLT_M_4)

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub Button_TS_BLLT_M_5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_M_5.Click
        Button_Change_Color(Button_TS_BLLT_M_5)

        Projekt.TS_BLT.Messung.Anzahl = 4
        Messungen_anzeigen(Button_TS_BLLT_M_5)

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub Button_TS_BLLT_M_6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_M_6.Click
        Button_Change_Color(Button_TS_BLLT_M_6)

        Projekt.TS_BLT.Messung.Anzahl = 5
        Messungen_anzeigen(Button_TS_BLLT_M_6)

        Update_Klasse_TS_BLT()
    End Sub
#End Region
#Region "Bauteil - BLT"
    Private Sub Button_TS_BLLT_BT_Ba_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_BT_Ba_1.Click
        Button_Change_Color(Button_TS_BLLT_BT_Ba_1)

        Projekt.TS_BLT.Messung.Messung_1.Bauteil = Balkon
    End Sub

    Private Sub Button_TS_BLLT_BT_Lo_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_BT_Lo_1.Click
        Button_Change_Color(Button_TS_BLLT_BT_Lo_1)

        Projekt.TS_BLT.Messung.Messung_1.Bauteil = Loggie
    End Sub

    Private Sub Button_TS_BLLT_BT_Te_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_BT_Te_1.Click
        Button_Change_Color(Button_TS_BLLT_BT_Te_1)

        Projekt.TS_BLT.Messung.Messung_1.Bauteil = Terrasse
    End Sub

    Private Sub Button_TS_BLLT_BT_Ba_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_BT_Ba_2.Click
        Button_Change_Color(Button_TS_BLLT_BT_Ba_2)

        Projekt.TS_BLT.Messung.Messung_2.Bauteil = Balkon
    End Sub

    Private Sub Button_TS_BLLT_BT_Lo_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_BT_Lo_2.Click
        Button_Change_Color(Button_TS_BLLT_BT_Lo_2)

        Projekt.TS_BLT.Messung.Messung_2.Bauteil = Loggie
    End Sub

    Private Sub Button_TS_BLLT_BT_Te_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_BT_Te_2.Click
        Button_Change_Color(Button_TS_BLLT_BT_Te_2)

        Projekt.TS_BLT.Messung.Messung_2.Bauteil = Terrasse
    End Sub

    Private Sub Button_TS_BLLT_BT_Ba_3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_BT_Ba_3.Click
        Button_Change_Color(Button_TS_BLLT_BT_Ba_3)

        Projekt.TS_BLT.Messung.Messung_3.Bauteil = Balkon
    End Sub

    Private Sub Button_TS_BLLT_BT_Lo_3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_BT_Lo_3.Click
        Button_Change_Color(Button_TS_BLLT_BT_Lo_3)

        Projekt.TS_BLT.Messung.Messung_3.Bauteil = Loggie
    End Sub

    Private Sub Button_TS_BLLT_BT_Te_3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_BT_Te_3.Click
        Button_Change_Color(Button_TS_BLLT_BT_Te_3)

        Projekt.TS_BLT.Messung.Messung_3.Bauteil = Terrasse
    End Sub


    Private Sub Button_TS_BLLT_BT_Ba_4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_BT_Ba_4.Click
        Button_Change_Color(Button_TS_BLLT_BT_Ba_4)

        Projekt.TS_BLT.Messung.Messung_4.Bauteil = Balkon
    End Sub

    Private Sub Button_TS_BLLT_BT_Lo_4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_BT_Lo_4.Click
        Button_Change_Color(Button_TS_BLLT_BT_Lo_4)

        Projekt.TS_BLT.Messung.Messung_4.Bauteil = Loggie
    End Sub

    Private Sub Button_TS_BLLT_BT_Te_4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_BT_Te_4.Click
        Button_Change_Color(Button_TS_BLLT_BT_Te_4)

        Projekt.TS_BLT.Messung.Messung_4.Bauteil = Terrasse
    End Sub


    Private Sub Button_TS_BLLT_BT_Ba_5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_BT_Ba_5.Click
        Button_Change_Color(Button_TS_BLLT_BT_Ba_5)

        Projekt.TS_BLT.Messung.Messung_5.Bauteil = Balkon
    End Sub

    Private Sub Button_TS_BLLT_BT_Lo_5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_BT_Lo_5.Click
        Button_Change_Color(Button_TS_BLLT_BT_Lo_5)

        Projekt.TS_BLT.Messung.Messung_5.Bauteil = Loggie
    End Sub

    Private Sub Button_TS_BLLT_BT_Te_5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_BT_Te_5.Click
        Button_Change_Color(Button_TS_BLLT_BT_Te_5)

        Projekt.TS_BLT.Messung.Messung_5.Bauteil = Terrasse
    End Sub


    Private Sub Button_TS_BLLT_BT_Ba_6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_BT_Ba_6.Click
        Button_Change_Color(Button_TS_BLLT_BT_Ba_6)

        Projekt.TS_BLT.Messung.Messung_6.Bauteil = Balkon
    End Sub

    Private Sub Button_TS_BLLT_BT_Lo_6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_BT_Lo_6.Click
        Button_Change_Color(Button_TS_BLLT_BT_Lo_6)

        Projekt.TS_BLT.Messung.Messung_6.Bauteil = Loggie
    End Sub

    Private Sub Button_TS_BLLT_BT_Te_6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_TS_BLLT_BT_Te_6.Click
        Button_Change_Color(Button_TS_BLLT_BT_Te_6)

        Projekt.TS_BLT.Messung.Messung_6.Bauteil = Terrasse
    End Sub
#End Region
#Region "NUD - Messungen"
#Region "Validating"
    Private Sub NUD_TS_BLLT_Mes_L_1_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_BLLT_Mes_L_1.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_BLLT_Mes_C_1_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_BLLT_Mes_C_1.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_BLT.Messung.Messung_1.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_TS_BLLT_Mes_L_2_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_BLLT_Mes_L_2.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_BLLT_Mes_C_2_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_BLLT_Mes_C_2.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_BLT.Messung.Messung_2.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_TS_BLLT_Mes_L_3_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_BLLT_Mes_L_3.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_BLLT_Mes_C_3_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_BLLT_Mes_C_3.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_BLT.Messung.Messung_3.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_TS_BLLT_Mes_L_4_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_BLLT_Mes_L_4.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_BLLT_Mes_C_4_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_BLLT_Mes_C_4.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_BLT.Messung.Messung_4.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_TS_BLLT_Mes_L_5_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_BLLT_Mes_L_5.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_BLLT_Mes_C_5_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_BLLT_Mes_C_5.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_BLT.Messung.Messung_5.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_TS_BLLT_Mes_L_6_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_BLLT_Mes_L_6.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_BLLT_Mes_C_6_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_BLLT_Mes_C_6.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_BLT.Messung.Messung_6.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub
#End Region
#Region "TextChanged"
    Private Sub NUD_TS_BLLT_Mes_L_1_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_BLLT_Mes_L_1.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_BLT.Messung.Messung_1.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_BLT.Messung.Messung_1.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_BLLT_MV_1_mit, Me.Button_TS_BLLT_MV_1_ohne, "TS_BLT", _
                Projekt.TS_BLT.Messung.Messung_1.Pegel.Pegel, Projekt.TS_BLT.Messung.Messung_1.Bauteil)

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub NUD_TS_BLLT_Mes_C_1_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_BLLT_Mes_C_1.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_BLT.Messung.Messung_1.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_BLT.Messung.Messung_1.Pegel.C
        End If

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub NUD_TS_BLLT_Mes_L_2_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_BLLT_Mes_L_2.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_BLT.Messung.Messung_2.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_BLT.Messung.Messung_2.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_BLLT_MV_2_mit, Me.Button_TS_BLLT_MV_2_ohne, "TS_BLT", _
                Projekt.TS_BLT.Messung.Messung_2.Pegel.Pegel, Projekt.TS_BLT.Messung.Messung_2.Bauteil)

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub NUD_TS_BLLT_Mes_C_2_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_BLLT_Mes_C_2.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_BLT.Messung.Messung_2.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_BLT.Messung.Messung_2.Pegel.C
        End If

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub NUD_TS_BLLT_Mes_L_3_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_BLLT_Mes_L_3.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_BLT.Messung.Messung_3.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_BLT.Messung.Messung_3.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_BLLT_MV_3_mit, Me.Button_TS_BLLT_MV_3_ohne, "TS_BLT", _
                Projekt.TS_BLT.Messung.Messung_3.Pegel.Pegel, Projekt.TS_BLT.Messung.Messung_3.Bauteil)

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub NUD_TS_BLLT_Mes_C_3_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_BLLT_Mes_C_3.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_BLT.Messung.Messung_3.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_BLT.Messung.Messung_3.Pegel.C
        End If

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub NUD_TS_BLLT_Mes_L_4_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_BLLT_Mes_L_4.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_BLT.Messung.Messung_4.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_BLT.Messung.Messung_4.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_BLLT_MV_4_mit, Me.Button_TS_BLLT_MV_4_ohne, "TS_BLT", _
                Projekt.TS_BLT.Messung.Messung_4.Pegel.Pegel, Projekt.TS_BLT.Messung.Messung_4.Bauteil)

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub NUD_TS_BLLT_Mes_C_4_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_BLLT_Mes_C_4.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_BLT.Messung.Messung_4.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_BLT.Messung.Messung_4.Pegel.C
        End If

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub NUD_TS_BLLT_Mes_L_5_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_BLLT_Mes_L_5.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_BLT.Messung.Messung_5.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_BLT.Messung.Messung_5.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_BLLT_MV_5_mit, Me.Button_TS_BLLT_MV_5_ohne, "TS_BLT", _
                Projekt.TS_BLT.Messung.Messung_5.Pegel.Pegel, Projekt.TS_BLT.Messung.Messung_5.Bauteil)

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub NUD_TS_BLLT_Mes_C_5_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_BLLT_Mes_C_5.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_BLT.Messung.Messung_5.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_BLT.Messung.Messung_5.Pegel.C
        End If

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub NUD_TS_BLLT_Mes_L_6_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_BLLT_Mes_L_6.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_BLT.Messung.Messung_6.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_BLT.Messung.Messung_6.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_BLLT_MV_6_mit, Me.Button_TS_BLLT_MV_6_ohne, "TS_BLT", _
                Projekt.TS_BLT.Messung.Messung_6.Pegel.Pegel, Projekt.TS_BLT.Messung.Messung_6.Bauteil)

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub NUD_TS_BLLT_Mes_C_6_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_BLLT_Mes_C_6.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_BLT.Messung.Messung_6.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_BLT.Messung.Messung_6.Pegel.C
        End If

        Update_Klasse_TS_BLT()
    End Sub
#End Region
#End Region
#End Region
#Region "Prognose"
    '#Region "Bauteil - BLLT"
    '    Private Sub Button_TS_BLLT_BT_La_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '        Button_Change_Color(Button_TS_BLLT_BT_La)

    '        Projekt.TS_BLLT.Prognose.Bauteil = Laube
    '    End Sub

    '    Private Sub Button_TS_BLLT_BT_Ba_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '        Button_Change_Color(Button_TS_BLLT_BT_Ba)

    '        Projekt.TS_BLLT.Prognose.Bauteil = Balkon
    '    End Sub

    '    Private Sub Button_TS_BLLT_BT_Lo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '        Button_Change_Color(Button_TS_BLLT_BT_Lo)

    '        Projekt.TS_BLLT.Prognose.Bauteil = Loggie
    '    End Sub

    '    Private Sub Button_TS_BLLT_BT_Te_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '        Button_Change_Color(Button_TS_BLLT_BT_Te)

    '        Projekt.TS_BLLT.Prognose.Bauteil = Terrasse
    '    End Sub
    '#End Region
#Region "NUD - Prognosse"
#Region "Validating"
    Private Sub NUD_TS_BLLT_Prog_Ba_L_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_BLLT_Prog_Ba_L.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_BLLT_Prog_Ba_C_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_BLLT_Prog_Ba_C.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_BLT.Prognose.Balkon.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub


    Private Sub NUD_TS_BLLT_Prog_Lo_L_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_BLLT_Prog_Lo_L.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_BLLT_Prog_Lo_C_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_BLLT_Prog_Lo_C.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_BLT.Prognose.Loggia.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub

    Private Sub NUD_TS_BLLT_Prog_Te_L_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_BLLT_Prog_Te_L.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub

    Private Sub NUD_TS_BLLT_Prog_Te_C_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_TS_BLLT_Prog_Te_C.Validating
        NUD_Validating(CType(sender, NumericUpDown))
        Projekt.TS_BLT.Prognose.Terrasse.Pegel.C = CType(sender, NumericUpDown).Text
    End Sub
#End Region
#Region "TextChanged"
    Private Sub NUD_TS_BLLT_Prog_Ba_L_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_BLLT_Prog_Ba_L.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_BLT.Prognose.Balkon.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_BLT.Prognose.Balkon.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_Balkon_Prog_mit, Me.Button_TS_Balkon_Prog_ohne, "TS_BLT", _
                Projekt.TS_BLT.Prognose.Balkon.Pegel.Pegel, Balkon)

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub NUD_TS_BLLT_Prog_Ba_C_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_BLLT_Prog_Ba_C.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_BLT.Prognose.Balkon.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_BLT.Prognose.Balkon.Pegel.C
        End If

        Update_Klasse_TS_BLT()
    End Sub


    Private Sub NUD_TS_BLLT_Prog_Lo_L_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_BLLT_Prog_Lo_L.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_BLT.Prognose.Loggia.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_BLT.Prognose.Loggia.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_Loggia_Prog_mit, Me.Button_TS_Loggia_Prog_ohne, "TS_BLT", _
                Projekt.TS_BLT.Prognose.Loggia.Pegel.Pegel, Loggie)

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub NUD_TS_BLLT_Prog_Lo_C_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_BLLT_Prog_Lo_C.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_BLT.Prognose.Loggia.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_BLT.Prognose.Loggia.Pegel.C
        End If

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub NUD_TS_BLLT_Prog_Te_L_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_BLLT_Prog_Te_L.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.TS_BLT.Prognose.Terrasse.Pegel.Pegel = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.TS_BLT.Prognose.Terrasse.Pegel.Pegel)
        End If

        Check_Bodenbelag(Me.Button_TS_Terrasse_Prog_mit, Me.Button_TS_Terrasse_Prog_ohne, "TS_BLT", _
                Projekt.TS_BLT.Prognose.Terrasse.Pegel.Pegel, Terrasse)

        Update_Klasse_TS_BLT()
    End Sub

    Private Sub NUD_TS_BLLT_Prog_Te_C_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_BLLT_Prog_Te_C.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            Projekt.TS_BLT.Prognose.Terrasse.Pegel.C = menud.Text
        Else
            menud.Text = Projekt.TS_BLT.Prognose.Terrasse.Pegel.C
        End If

        Update_Klasse_TS_BLT()
    End Sub
#End Region
#End Region
#End Region
#End Region

#Region "Türen"
    Private Sub Update_Klasse_Tueren()
        Me.Button_Klasse_Tueren.Text = Get_Klasse_Tueren()
        Me.Button_Klasse_Tueren.BackColor = Get_Klasse_Color(Me.Button_Klasse_Tueren.Text)

        Me.Update_SSA_Tueren()
    End Sub
#Region "Diele"
    Private Sub Button_Tueren_Diele_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tueren_Diele.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_21", CType(sender, Control))
    End Sub
    Private Sub Button_Tueren_Diele_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tueren_Diele.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub
    Private Sub Button_Tueren_Diele_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tueren_Diele.Click
        Button_Change_Color(Button_Tueren_Diele)

        Me.Panel_SSA_Tuere_Untersuchung_b.Height = 54

        Projekt.Tueren.Ort = Diele

        Update_Klasse_Tueren()
    End Sub
#End Region
#Region "Aufenthaltsraum"
    Private Sub Button_Tuere_Aufenthaltsraum_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tuere_Aufenthaltsraum.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_21", CType(sender, Control))
    End Sub
    Private Sub Button_Tuere_Aufenthaltsraum_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tuere_Aufenthaltsraum.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub
    Private Sub Button_Tuere_Aufenthaltsraum_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tuere_Aufenthaltsraum.Click
        Button_Change_Color(Button_Tuere_Aufenthaltsraum)

        Me.Panel_SSA_Tuere_Untersuchung_b.Height = 54

        Projekt.Tueren.Ort = Aufenthalt

        Update_Klasse_Tueren()
    End Sub
#End Region
#Region "nv"
    Private Sub Button_Tuere_nv_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tuere_nv.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help_2("ssa_help_36", "ssa_help_21", CType(sender, Control))
    End Sub
    Private Sub Button_Tuere_nv_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tuere_nv.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub
    Private Sub Button_Tuere_nv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tuere_nv.Click
        Button_Change_Color(Button_Tuere_nv)

        Me.Panel_SSA_Tuere_Untersuchung_b.Height = 0    '54

        Projekt.Tueren.Ort = nv

        Update_Klasse_Tueren()
    End Sub
#End Region
#Region "Untersuchung"
    Private Sub Button_Tuere_Pruefzeugnis_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tuere_Pruefzeugnis.Click
        Button_Change_Color(Button_Tuere_Pruefzeugnis)

        Projekt.Tueren.Untersuchung = Pruefzeugnis

        Update_Klasse_Tueren()
    End Sub
    Private Sub Button_Tuere_Messung_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tuere_Messung.Click
        Button_Change_Color(Button_Tuere_Messung)

        Projekt.Tueren.Untersuchung = Messung

        Update_Klasse_Tueren()
    End Sub
#End Region
#Region "Pegel"
    Private Sub NUD_Tuer_R_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles NUD_Tuer_R.Validating
        NUD_Validating(CType(sender, NumericUpDown))
    End Sub
    Private Sub NUD_Tuer_R_TextChangd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_Tuer_R.TextChanged
        Dim menud As NumericUpDown = CType(sender, NumericUpDown)
        If Me.TabPage1.Focused = False Then
            If menud.Text = "" Then menud.Text = "0"
            Projekt.Tueren.L = CSng(menud.Text)
        Else
            menud.Value = CDec(Projekt.Tueren.L)
        End If

        Update_Klasse_Tueren()
    End Sub
#End Region
#End Region
#Region "Aussenbauteile"
    Private Sub Update_Klasse_Aussenbauteile()
        Me.Button_Klasse_Aussenbauteile.Text = Get_Klasse_Aussenbauteile()
        Me.Button_Klasse_Aussenbauteile.BackColor = Get_Klasse_Color(Me.Button_Klasse_Aussenbauteile.Text)

        Me.Update_SSA_Aussenbauteile()
    End Sub

    Private Sub Button_Aussenbauteile_oN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Aussenbauteile_oN.Click
        Button_Change_Color(Button_Aussenbauteile_oN)

        Projekt.Aussenbauteile = ohneNachweis

        Update_Klasse_Aussenbauteile()
    End Sub

    'Private Sub Button_Aussenbauteile_FmD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Button_Change_Color(Button_Aussenbauteile_FmD)

    '    Projekt.Aussenbauteile = FensterMitDichtung

    '    Update_Klasse_Aussenbauteile()
    'End Sub

    Private Sub Button_Aussenbauteile_DIN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Aussenbauteile_DIN.Click
        Button_Change_Color(Button_Aussenbauteile_DIN)

        Projekt.Aussenbauteile = DINerfuellt

        Update_Klasse_Aussenbauteile()
    End Sub

    Private Sub Button_Aussenbauteile_DINPlus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Aussenbauteile_DINPlus.Click
        Button_Change_Color(Button_Aussenbauteile_DINPlus)

        Projekt.Aussenbauteile = DINPlusErfuellt

        Update_Klasse_Aussenbauteile()
    End Sub
#End Region
#Region "Wasserinstallationen und haustechnische Anlagen"
    Private Sub Update_Klasse_Wasser()
        Me.Button_Klasse_Wasser.Text = Get_Klasse_Wasser()
        Me.Button_Klasse_Wasser.BackColor = Get_Klasse_Color(Me.Button_Klasse_Wasser.Text)

        Me.Update_SSA_Wasser()
    End Sub
    Private Sub Button_Wasser_Prognose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Wasser_Prognose.Click
        Button_Change_Color(Button_Wasser_Prognose)

        Projekt.Wasser.Untersuchung = Prognose

        Update_Klasse_Wasser()
    End Sub
    Private Sub Button_Wasser_Messung_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Wasser_Messung.Click
        Button_Change_Color(Button_Wasser_Messung)

        Projekt.Wasser.Untersuchung = Messung

        Update_Klasse_Wasser()
    End Sub
    Private Sub Button_Wasser_erfuellt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Wasser_erfuellt.Click
        Button_Change_Color(Button_Wasser_erfuellt)

        Projekt.Wasser.LcLa_erfuellt = erfuellt

        Update_Klasse_Wasser()
    End Sub
    Private Sub Button_Wasser_kA_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Wasser_kA.Click
        Button_Change_Color(Button_Wasser_kA)

        Projekt.Wasser.LcLa_erfuellt = keineAngabe

        Update_Klasse_Wasser()
    End Sub

    Private Sub CB_Wasser_L_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Wasser_L.SelectedIndexChanged
        With Projekt.Wasser
            Select Case Me.CB_Wasser_L.Text
                Case "≤ 20"
                    .L_Intervall = HT_kleiner20
                Case "20 < L ≤ 24"
                    .L_Intervall = HT_20_24
                Case "24 < L ≤ 27"
                    .L_Intervall = HT_24_27
                Case "27 < L ≤ 30"
                    .L_Intervall = HT_27_30
                Case "30 < L ≤ 35"
                    .L_Intervall = HT_30_35
                Case "> 35"
                    .L_Intervall = HT_groesser35
                Case Else
                    .L_Intervall = 0
            End Select
        End With
        Update_Klasse_Wasser()
    End Sub
#End Region
#Region "Nutzergeräusche Körperschallentkopplung"
    Private Sub Update_Klasse_NutzerKoerper()
        Me.Button_Klasse_NutzerKoerper.Text = Get_Klasse_NutzerKoerper()
        Me.Button_Klasse_NutzerKoerper.BackColor = Get_Klasse_Color(Me.Button_Klasse_NutzerKoerper.Text)

        Me.Update_SSA_Nutzer_Koerper()
    End Sub
#Region "Buttons nv, Nutzer, Körper"
    Private Sub Button_NutzerKoerper_nv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_NutzerKoerper_nv.Click
        Button_Change_Color(Button_NutzerKoerper_nv)

        ' Me.CB_Nutzer_L.Text = ""
        Me.CB_Nutzer_L.SelectedIndex = -1
        Me.Panel_SSA_Nutzer_Pegel_b.Height = 0    '18

        Me.CB_Koerper_L.SelectedIndex = -1
        Me.Panel_SSA_Koerper_Pegel_b.Height = 0    '18

        Me.Panel_SSA_NutzerKoerper_b.Height = 0

        Projekt.NutzerKoerper.Untersuchung = nv

        Update_Klasse_NutzerKoerper()
    End Sub
    Private Sub Button_Nutzer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Nutzer.Click
        Button_Change_Color(Button_Nutzer)

        Me.Panel_SSA_Nutzer_Pegel_b.Height = 18

        Me.CB_Koerper_L.SelectedIndex = -1
        Me.Panel_SSA_Koerper_Pegel_b.Height = 0    '18


        Me.Panel_SSA_NutzerKoerper_b.Height = 39

        Projekt.NutzerKoerper.Untersuchung = nutzer

        Update_Klasse_NutzerKoerper()
    End Sub
    Private Sub Button_Koerper_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Koerper.Click
        Button_Change_Color(Button_Koerper)

        Me.CB_Nutzer_L.SelectedIndex = -1
        Me.Panel_SSA_Nutzer_Pegel_b.Height = 0    '18

        Me.Panel_SSA_Koerper_Pegel_b.Height = 18

        Me.Panel_SSA_NutzerKoerper_b.Height = 39

        Projekt.NutzerKoerper.Untersuchung = koerper

        Update_Klasse_NutzerKoerper()
    End Sub
#End Region
#Region "Buttons Prognose/Messung"
    Private Sub Button_NutzerKoerper_Prognose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_NutzerKoerper_Prognose.Click
        Button_Change_Color(Button_NutzerKoerper_Prognose)

        'Me.Panel_SSA_Koerper_Pegel_b.Height = 19

        Projekt.NutzerKoerper.MessungPrognose = Prognose

        Update_Klasse_NutzerKoerper()
    End Sub
    Private Sub Button_NutzerKoerper_Messung_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_NutzerKoerper_Messung.Click
        Button_Change_Color(Button_NutzerKoerper_Messung)

        'Me.Panel_SSA_Koerper_Pegel_b.Height = 19

        Projekt.NutzerKoerper.MessungPrognose = Messung

        Update_Klasse_NutzerKoerper()
    End Sub
#End Region
    Private Sub CB_Nutzer_L_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Nutzer_L.SelectedIndexChanged
        With Projekt.NutzerKoerper
            Select Case Me.CB_Nutzer_L.Text
                Case "≤ 20"
                    .L_Intervall = kleiner20
                Case "20 < L ≤ 25"
                    .L_Intervall = L_20_25
                Case "25 < L ≤ 30"
                    .L_Intervall = L_25_30
                Case "30 < L ≤ 35"
                    .L_Intervall = L_30_35
                Case "35 < L ≤ 40"
                    .L_Intervall = L_35_40
                Case "40 < L ≤ 45"
                    .L_Intervall = L_40_45
                Case "> 45"
                    .L_Intervall = groesser45
                Case Else
                    .L_Intervall = 0
            End Select
        End With

        Update_Klasse_NutzerKoerper()
    End Sub

    Private Sub CB_Koerper_L_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Koerper_L.SelectedIndexChanged
        With Projekt.NutzerKoerper
            Select Case Me.CB_Koerper_L.Text
                Case "≤ 38"
                    .L_Intervall = kleiner20
                Case "38 < L ≤ 43"
                    .L_Intervall = L_20_25
                Case "43 < L ≤ 48"
                    .L_Intervall = L_25_30
                Case "48 < L ≤ 53"
                    .L_Intervall = L_30_35
                Case "53 < L ≤ 58"
                    .L_Intervall = L_35_40
                Case "58 < L ≤ 63"
                    .L_Intervall = L_40_45
                Case "> 63"
                    .L_Intervall = groesser45
                Case Else
                    .L_Intervall = 0
            End Select
        End With

        Update_Klasse_NutzerKoerper()
    End Sub
#End Region
#Region "Körperschallentkopplung"
    Private Sub Button_Koerper_Prognose_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_NutzerKoerper_Prognose.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_NutzerKoerper_Prognose_19", CType(sender, Control))
    End Sub

    Private Sub Button_Koerper_Messung_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_NutzerKoerper_Messung.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_20", CType(sender, Control))
    End Sub

    'Private Sub Update_Klasse_Koerper()
    '    Me.Button_Klasse_Koerper.Text = Get_Klasse_Koerper()

    '    Me.Update_SSA_Nutzer_Koerper()
    'End Sub
    'Private Sub Button_Koerper_on_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Button_Change_Color(Button_Koerper_on)

    '    ' Me.CB_Koerper_L.Text = ""
    '    Me.CB_Koerper_L.SelectedIndex = -1
    '    Me.Panel_SSA_Koerper_Pegel_b.Height = 0   '18

    '    Projekt.Koerper.Untersuchung = nv

    '    Update_Klasse_Koerper()
    'End Sub
#End Region
#Region "Nachbarwohneinheiten"
    Private Sub Update_Klasse_Nachbarn()
        Me.Button_Klasse_Nachbarn.Text = Get_Klasse_Nachbarn()
        Me.Button_Klasse_Nachbarn.BackColor = Get_Klasse_Color(Me.Button_Klasse_Nachbarn.Text)

        Me.Update_SSA_Nachbarn()
    End Sub
    Private Sub Button_Nachbar_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Nachbar_1.Click
        Button_Change_Color(Button_Nachbar_1)

        Projekt.Nachbarn = 1

        Update_Klasse_Nachbarn()
    End Sub
    Private Sub Button_Nachbar_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Nachbar_2.Click
        Button_Change_Color(Button_Nachbar_2)

        Projekt.Nachbarn = 2

        Update_Klasse_Nachbarn()
    End Sub
    Private Sub Button_Nachbar_3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Nachbar_3.Click
        Button_Change_Color(Button_Nachbar_3)

        Projekt.Nachbarn = 3

        Update_Klasse_Nachbarn()
    End Sub
    Private Sub Button_Nachbar_4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Nachbar_4.Click
        Button_Change_Color(Button_Nachbar_4)

        Projekt.Nachbarn = 4

        Update_Klasse_Nachbarn()
    End Sub
    Private Sub Button_Nachbar_5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Nachbar_5.Click
        Button_Change_Color(Button_Nachbar_5)

        Projekt.Nachbarn = 5

        Update_Klasse_Nachbarn()
    End Sub
#End Region
#Region "Anordnung lauter Räume schalltechnisch"
    Private Sub Button_anordRaeume_g_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_anordRaeume_g.Click
        Button_Change_Color(Button_anordRaeume_g)

        Projekt.anordnungRaeume = guenstig

        Me.Update_SSA_anordRaeume()
    End Sub

    Private Sub Button_anordRaeume_ug_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_anordRaeume_ug.Click
        Button_Change_Color(Button_anordRaeume_ug)

        Projekt.anordnungRaeume = unguenstig

        Me.Update_SSA_anordRaeume()
    End Sub
#End Region
#Region "laute Räume angrenzend"
    Private Sub Update_Klasse_lauteRaeume()
        Me.Button_Klasse_lauteRaeume.Text = Get_Klasse_lauteRaeume()
        Me.Button_Klasse_lauteRaeume.BackColor = Get_Klasse_Color(Me.Button_Klasse_lauteRaeume.Text)

        Me.Update_SSA_lauteRaeume()
    End Sub

    Private Sub Button_lauteRaeume_keine_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_lauteRaeume_keine.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_39", CType(sender, Control))
    End Sub
    Private Sub Button_lauteRaeume_keine_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_lauteRaeume_keine.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub
    Private Sub Button_lauteRaeume_keine_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_lauteRaeume_keine.Click
        Button_Change_Color(Button_lauteRaeume_keine)

        Projekt.lauteRaeume = keineAngrenzend

        Update_Klasse_lauteRaeume()
    End Sub

    Private Sub Button_lauteRaeume_25_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_lauteRaeume_25.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_17", CType(sender, Control))
    End Sub
    Private Sub Button_lauteRaeume_25_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_lauteRaeume_25.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub
    Private Sub Button_lauteRaeume_25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_lauteRaeume_25.Click
        Button_Change_Color(Button_lauteRaeume_25)

        Projekt.lauteRaeume = L_25_35

        Update_Klasse_lauteRaeume()
    End Sub

    Private Sub Button_lauteRaeume_30_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_lauteRaeume_30.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_17", CType(sender, Control))
    End Sub
    Private Sub Button_lauteRaeume_30_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_lauteRaeume_30.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub
    Private Sub Button_lauteRaeume_30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_lauteRaeume_30.Click
        Button_Change_Color(Button_lauteRaeume_30)

        Projekt.lauteRaeume = L_30_40

        Update_Klasse_lauteRaeume()
    End Sub

    Private Sub Button_lauteRaeume_35_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_lauteRaeume_35.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_17", CType(sender, Control))
    End Sub
    Private Sub Button_lauteRaeume_35_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_lauteRaeume_35.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub
    Private Sub Button_lauteRaeume_35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_lauteRaeume_35.Click
        Button_Change_Color(Button_lauteRaeume_35)

        Projekt.lauteRaeume = L_35_45

        Update_Klasse_lauteRaeume()
    End Sub
#End Region
#Region "Treppenhaus Nachhallzeit"
    Private Sub Update_Klasse_NHZ()
        Me.Button_Klasse_NHZ.Text = Get_Klasse_NHZ()

        Me.Update_SSA_NHZ()
    End Sub

    Private Sub Button_NHZ_020_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_NHZ_020.Click
        Button_Change_Color(Button_NHZ_020)

        Projekt.NHZ = NHZ_020

        Update_Klasse_NHZ()
    End Sub
    Private Sub Button_NHZ_010_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_NHZ_010.Click
        Button_Change_Color(Button_NHZ_010)

        Projekt.NHZ = NHZ_010

        Update_Klasse_NHZ()
    End Sub
    Private Sub Button_NHZ_Keine_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_NHZ_Keine.Click
        Button_Change_Color(Button_NHZ_Keine)

        Projekt.NHZ = NHZ_keine

        Update_Klasse_NHZ()
    End Sub
#End Region
#Region "eigener Wohnbereich"
    Private Sub Update_Klasse_EW()
        Me.Button_Klasse_EW.Text = Get_Klasse_EW()

        Me.Update_SSA_EW()
    End Sub
    Private Sub Button_EW_kE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_EW_kE.Click
        Button_Change_Color(Button_EW_kE)

        Projekt.eigenerWohnbereich = keineEmpfehlung

        Update_Klasse_EW()
    End Sub

    Private Sub Button_EW_1_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_EW_1.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_EW1_57", CType(sender, Control))
    End Sub
     Private Sub Button_EW_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_EW_1.Click
        Button_Change_Color(Button_EW_1)

        Projekt.eigenerWohnbereich = EW1

        Update_Klasse_EW()
    End Sub

    Private Sub Button_EW_2_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_EW_2.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_EW2_56", CType(sender, Control))
    End Sub
     Private Sub Button_EW_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_EW_2.Click
        Button_Change_Color(Button_EW_2)

        Projekt.eigenerWohnbereich = EW2

        Update_Klasse_EW()
    End Sub
    Private Sub Button_EW_3_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_EW_3.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_EW3_56_1", CType(sender, Control))
    End Sub
    Private Sub Button_EW_3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_EW_3.Click
        Button_Change_Color(Button_EW_3)

        Projekt.eigenerWohnbereich = EW3

        Update_Klasse_EW()
    End Sub
#End Region

#Region "Bemerkungen"

    Private Sub TB_Bemerkung_LS_W_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Bemerkung_LS_W.TextChanged
        Projekt.LS_Wand.Bemerkung_LS = CType(sender, System.Windows.Forms.TextBox).Text
    End Sub

    Private Sub TB_Bemerkung_LS_D_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Bemerkung_LS_D.TextChanged
        Projekt.LS_Decke.Bemerkung_LS = CType(sender, System.Windows.Forms.TextBox).Text
    End Sub

    Private Sub TB_Bemerkung_TS_D_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Bemerkung_TS_D.TextChanged
        Projekt.TS_Decke.Bemerkung_TS_Decke = CType(sender, System.Windows.Forms.TextBox).Text
    End Sub

    Private Sub TB_Bemerkung_TS_TPH_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Bemerkung_TS_TPH.TextChanged
        Projekt.TS_TPLH.Bemerkung_TS_TPH = CType(sender, System.Windows.Forms.TextBox).Text
    End Sub

    Private Sub TB_Bemerkung_TS_BLLT_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Bemerkung_TS_BLLT.TextChanged
        Projekt.TS_BLT.Bemerkung_TS_BLLT = CType(sender, System.Windows.Forms.TextBox).Text
    End Sub

    Private Sub TB_Bemerkung_Tueren_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Bemerkung_Tueren.TextChanged
        Projekt.Tueren.Bemerkung_Tueren = CType(sender, System.Windows.Forms.TextBox).Text
    End Sub

    Private Sub TB_Bemerkung_Nutzer_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Bemerkung_Nutzer.TextChanged
        Projekt.NutzerKoerper.BemerkungNutzer = CType(sender, System.Windows.Forms.TextBox).Text
    End Sub

    Private Sub TB_Bemerkung_Koerper_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Bemerkung_Koerper.TextChanged
        Projekt.NutzerKoerper.BemerkungKoerper = CType(sender, System.Windows.Forms.TextBox).Text
    End Sub

    Private Sub TB_Bemerkung_Nachbarn_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Bemerkung_Nachbarn.TextChanged
        Projekt.Bemerkung_Nachbarn = CType(sender, System.Windows.Forms.TextBox).Text
    End Sub

    Private Sub TB_Bemerkung_EW_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Bemerkung_EW.TextChanged
        Projekt.Bemerkung_eigenerWohnbereich = CType(sender, System.Windows.Forms.TextBox).Text
    End Sub


    Private Sub TB_Bem_Gebietscharakter_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Bem_Gebietscharakter.TextChanged
        Projekt.Standort.Bem_Gebietscharakter = CType(sender, System.Windows.Forms.TextBox).Text
    End Sub

    Private Sub TB_Bem_Aussenlaermsituation_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Bem_Aussenlaermsituation.TextChanged
        Projekt.Standort.Bem_Aussenlaermpegel = CType(sender, System.Windows.Forms.TextBox).Text
    End Sub


    Private Sub TB_Bem_Aussenbauteile_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Bem_Aussenbauteile.TextChanged
        Projekt.Bem_Aussenbauteile = CType(sender, System.Windows.Forms.TextBox).Text
    End Sub

    Private Sub TB_Bem_WasserHausanlagen_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Bem_WasserHausanlagen.TextChanged
        Projekt.Wasser.Bem_Wasser = CType(sender, System.Windows.Forms.TextBox).Text
    End Sub

    Private Sub TB_Bem_AnordnungRaeume_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Bem_AnordnungRaeume.TextChanged
        Projekt.Bem_anordnungRaeume = CType(sender, System.Windows.Forms.TextBox).Text
    End Sub

    Private Sub TB_Bem_AV_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Bem_AV.TextChanged
        Projekt.Bemerkung_AV = CType(sender, System.Windows.Forms.TextBox).Text
    End Sub
    Private Sub TB_Bem_lauteRaeume_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Bem_lauteRaeume.TextChanged
        Projekt.Bem_lauteRaeume = CType(sender, System.Windows.Forms.TextBox).Text
    End Sub
#End Region
    Private Sub TSMI_DirekteHilfe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSMI_DirekteHilfe.Click
        If Me.Cursor = Cursors.Help Then
            Me.Cursor = Cursors.Default
        Else
            Me.Cursor = Cursors.Help
        End If

    End Sub

    Private Sub PB_Grundriss_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PB_Grundriss.Click
        If Me.Cursor = Cursors.Help Then
            Me.Cursor = Cursors.Default
            Me.TSMI_DirekteHilfe.Checked = False
        Else
            If Not IsNothing(Me.PB_Grundriss.BackgroundImage) Then
                'zoom
                Form_Grundriss.BackgroundImage = Me.PB_Grundriss.BackgroundImage
                Form_Grundriss.ShowDialog()

            Else
                Grundriss_hinzufuegen()
            End If
        End If
    End Sub

    Private Sub TSMI_Hinzufuegen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSMI_Grundriss_Hinzufuegen.Click
        Grundriss_hinzufuegen()
    End Sub
    Private Sub TSMI_Grundriss_Drehen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSMI_Grundriss_Drehen.Click
        Dim frm_Image_drehen As New Form_Image_drehen
        frm_Image_drehen.ShowDialog()
    End Sub

    Private Sub TSMI_Grundriss_Entfernen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSMI_Grundriss_Entfernen.Click
        Me.PB_Grundriss.BackgroundImage.Dispose()
        Me.PB_Grundriss.BackgroundImage = Nothing

    End Sub
    Private Sub TSMI_Bearbeiten_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles TSMI_Bearbeiten.Paint
        If Not IsNothing(Me.PB_Grundriss.BackgroundImage) Then
            Me.TSMI_Grundriss_Drehen.Enabled = True
            Me.TSMI_Grundriss_Entfernen.Enabled = True
        Else
            Me.TSMI_Grundriss_Drehen.Enabled = False
            Me.TSMI_Grundriss_Entfernen.Enabled = False
        End If
    End Sub

    Private Sub TSMI_Aussteller_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSMI_Aussteller.Click
        Dim frm_Konfig As New Form_Aussteller
        frm_Konfig.ShowDialog()

        Me.Update_SSA()
    End Sub

    Private Sub TSMI_EingabeLeeren_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSMI_EingabeLeeren.Click
        Projekt_Default()
        Update_Anzeige()
    End Sub

    Private Sub TSMI_SeiteEinrichten_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSMI_SeiteEinrichten.Click
        PageSetupDialog1.Document = PrintDocument1

        PageSetupDialog1.ShowDialog()
    End Sub

    Private Sub TSMI_Druckvorschau_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSMI_Druckvorschau.Click
        PrintPreviewDialog1.Document = PrintDocument1
        If PrintPreviewDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            PrintDocument1.Print()
        End If
    End Sub

    Private Sub TSMI_SchallschutzausweisDrucken_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSMI_SchallschutzausweisDrucken.Click
        Drucken()
    End Sub

    Private Sub TSMI_Info_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Info_Anzeigen()
    End Sub
    Private Sub TSMI_Datenuebertragung_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Form_Webbrowser.Show()
        Senden_Daten()
    End Sub

    Private Sub PrintDocument1_BeginPrint(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles PrintDocument1.BeginPrint
        '        iPage = 1
        iPage = PrintDocument1.PrinterSettings.FromPage
        If iPage = 0 Then
            iPage = 1
            PrintDocument1.PrinterSettings.ToPage = 4
        End If

    End Sub
    Private Function Get_Zoll_u_width(ByVal Pixel As Integer) As Integer
        Return CInt(Pixel * wPage / My.Resources.Deckblatt.Width) '1145)
    End Function
    Private Function Get_Zoll_u_height(ByVal Pixel As Integer) As Integer
        Return CInt(Pixel * hPage / My.Resources.Deckblatt.Height) '1985)
    End Function
    Private Function Get_Zoll_d_Width(ByVal Pixel As Integer) As Integer
        '        Return CInt(Pixel * 20 / (1163 * 0.0254))
        Return CInt(Pixel * wPage / My.Resources.ssa_detailliert.Width) '1163)
    End Function
    Private Function Get_Zoll_d_Height(ByVal Pixel As Integer) As Integer
        '        Return CInt(Pixel * 28.6 / (1705 * 0.0254))
        Return CInt(Pixel * hPage / My.Resources.ssa_detailliert.Height) '1755)
    End Function
    Private Function Get_Zoll_e_Width(ByVal Pixel As Integer) As Integer
        Return CInt(Pixel * wPage / My.Resources.SSA_einfach.Width) '1116)
    End Function
    Private Function Get_Zoll_e_Height(ByVal Pixel As Integer) As Integer
        Return CInt(Pixel * hPage / My.Resources.SSA_einfach.Height) '1752)
    End Function
    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim gr As Graphics = e.Graphics

        Dim printrect As System.Drawing.Rectangle = e.MarginBounds
        xPage = printrect.X
        yPage = printrect.Y
        wPage = printrect.Width
        hPage = printrect.Height

        Dim Arial_Norm_10 As New System.Drawing.Font("Arial", 10)

        If iPage = 1 And PrintDocument1.PrinterSettings.ToPage >= 1 Then

            gr.DrawImage(My.Resources.Deckblatt, printrect, 0, 0, My.Resources.Deckblatt.Width, My.Resources.Deckblatt.Height, GraphicsUnit.Pixel)

            Print_SSA_Deckblatt(gr)

        ElseIf iPage = 2 And PrintDocument1.PrinterSettings.ToPage >= 2 Then

            gr.DrawImage(My.Resources.ssa_detailliert, printrect, 0, 0, My.Resources.ssa_detailliert.Width, My.Resources.ssa_detailliert.Height, GraphicsUnit.Pixel)

            Print_SSA_detailiert(gr)

        ElseIf iPage = 3 And PrintDocument1.PrinterSettings.ToPage >= 3 Then

            gr.DrawImage(My.Resources.SSA_einfach, printrect, 0, 0, My.Resources.SSA_einfach.Width, My.Resources.SSA_einfach.Height, GraphicsUnit.Pixel)

            Print_SSA_einfach(gr)
        ElseIf iPage = 4 And PrintDocument1.PrinterSettings.ToPage >= 4 Then
            gr.DrawImage(My.Resources.SSA_Hinweise, printrect, 0, 0, My.Resources.SSA_Hinweise.Width, My.Resources.SSA_Hinweise.Height, GraphicsUnit.Pixel)

        End If

        If iPage < 4 And iPage + 1 <= PrintDocument1.PrinterSettings.ToPage Then
            iPage = iPage + 1
            e.HasMorePages = True
        Else
            e.HasMorePages = False
        End If
    End Sub

    Private Sub Print_SSA_Deckblatt(ByVal gr As Graphics)
        '        Dim logoRect As New System.Drawing.Rectangle(xPage + Get_Zoll_u_width(972), yPage + Get_Zoll_u_height(14), _
        '                   Get_Zoll_u_height(87), Get_Zoll_u_height(87))


        If Not IsNothing(PB_Geb_Logo.BackgroundImage) Then
            Dim imScale As Double
            If 140 / PB_Geb_Logo.BackgroundImage.Width < 89 / PB_Geb_Logo.BackgroundImage.Height Then
                imScale = 140 / PB_Geb_Logo.BackgroundImage.Width
            Else
                imScale = 89 / PB_Geb_Logo.BackgroundImage.Height
            End If
            ''972; 13   
            Dim x As Integer = CInt(972 + 140 / 2 - PB_Geb_Logo.BackgroundImage.Width * imScale)
            Dim y As Integer = CInt(13 + 89 / 2 - PB_Geb_Logo.BackgroundImage.Height * imScale / 2)
            Dim w As Integer = CInt(PB_Geb_Logo.BackgroundImage.Width * imScale)
            Dim h As Integer = CInt(PB_Geb_Logo.BackgroundImage.Height * imScale)
            Dim logoRect As New System.Drawing.Rectangle(xPage + Get_Zoll_u_width(x), yPage + Get_Zoll_u_height(y), _
                Get_Zoll_u_height(w), Get_Zoll_u_height(h))
            gr.DrawImage(PB_Geb_Logo.BackgroundImage, logoRect)
        End If



        With Projekt
            If Not IsNothing(.Datum) Then
                '682; 1660  149; 20
                If .Datum <> "" Then
                    Print_Center_Center(.Datum, _
                            Get_Zoll_u_width(725), Get_Zoll_u_height(1908), Get_Zoll_u_width(149), Get_Zoll_u_height(20), gr, False, 6)
                    'Print_Center_Center(CType(.Datum, Date).Day & "." & CType(.Datum, Date).Month & "." & CType(.Datum, Date).Year, _
                    '        Get_Zoll_u_width(706), Get_Zoll_u_height(1860), Get_Zoll_u_width(149), Get_Zoll_u_height(20), gr, False, 6)
                    Dim tDatum As Date = CType(.Datum, Date)
                    Print_Left_Center_Norm_6(IIf(tDatum.Day < 10, "0" & tDatum.Day, tDatum.Day).ToString & "." & IIf(tDatum.Month < 10, "0" & tDatum.Month, tDatum.Month).ToString & "." & tDatum.AddYears(10).Year, _
                            Get_Zoll_u_width(112), Get_Zoll_u_height(148), Get_Zoll_u_height(30), gr)
                End If
            End If
        End With

        'Gebäude
        With Projekt.Gebaeude.Adresse
            Dim tmpPlz As String = Trim(.PLZ) & " "
            'If tmpPlz = "0 " Then tmpPlz = ""
            Print_Left_Center_Norm_6(Trim(.Strasse) & " " & Trim(.Nr) & Chr(13) & Chr(10) & tmpPlz & Trim(.Ort), _
                                Get_Zoll_u_width(282), Get_Zoll_u_height(217), Get_Zoll_u_height(47), gr)
        End With
        'Gebäudeinfos
        With Projekt.Gebaeude
            Print_Left_Center_Norm_6(.Gebaeudetyp, Get_Zoll_u_width(282), Get_Zoll_u_height(269), Get_Zoll_u_height(30), gr)
            Dim tmpBaujahr As String = Trim(.Baujahr.ToString)
            If tmpBaujahr = "1800" Then tmpBaujahr = ""
            Print_Left_Center_Norm_6(tmpBaujahr, Get_Zoll_u_width(282), Get_Zoll_u_height(305), Get_Zoll_u_height(30), gr)
            Print_Left_Center_Norm_6(.Wohneinheiten, Get_Zoll_u_width(282), Get_Zoll_u_height(341), Get_Zoll_u_height(30), gr)
        End With
        With Projekt.Wohnung
            Print_Left_Center_Norm_6(.Wohnungsbezeichnung, Get_Zoll_u_width(282), Get_Zoll_u_height(377), Get_Zoll_u_height(30), gr)
            Print_Left_Center_Norm_6(Get_Str_Geschoss(), Get_Zoll_u_width(282), Get_Zoll_u_height(413), Get_Zoll_u_height(30), gr)
            Print_Left_Center_Norm_6(.Raeume, Get_Zoll_u_width(282), Get_Zoll_u_height(449), Get_Zoll_u_height(30), gr)
            Dim tmpWohnfl As String = Trim(.Wohnflaeche.ToString)
            If tmpWohnfl = "0" Then tmpWohnfl = ""
            Print_Left_Center_Norm_6(tmpWohnfl, Get_Zoll_u_width(282), Get_Zoll_u_height(485), Get_Zoll_u_height(30), gr)
        End With
        With My.Settings 'Aussteller.Adresse
            If .Aussteller_Name <> "" Then Print_Left_Center_Norm_6(Trim(.Aussteller_Name), Get_Zoll_u_width(135), Get_Zoll_u_height(1800), Get_Zoll_u_height(20), gr)
            If .Aussteller_Zusatz <> "" Then Print_Left_Center_Norm_6(Trim(.Aussteller_Zusatz), Get_Zoll_u_width(135), Get_Zoll_u_height(1835), Get_Zoll_u_height(20), gr)
            If .Aussteller_Strasse <> "" Then Print_Left_Center_Norm_6(Trim(.Aussteller_Strasse) & " " & Trim(.Aussteller_Nr), Get_Zoll_u_width(135), Get_Zoll_u_height(1871), Get_Zoll_u_height(20), gr)
            Dim tmpPlz As String = Trim(.Aussteller_PLZ) & " "
            'If tmpPlz = "0 " Then tmpPlz = ""
            If tmpPlz <> "" Or .Aussteller_Ort <> "" Then Print_Left_Center_Norm_6(tmpPlz & Trim(.Aussteller_Ort), Get_Zoll_u_width(135), Get_Zoll_u_height(1907), Get_Zoll_e_Height(20), gr)
        End With

        Dim grundrissRect As System.Drawing.Rectangle '(xPage + Get_Zoll_u_width(768), yPage + Get_Zoll_u_height(220), Get_Zoll_u_width(321), Get_Zoll_u_height(293))
        Dim GSize As Size
        Dim tmpScale As Single

        If Not IsNothing(Me.PB_Grundriss.BackgroundImage) Then
            GSize = New Size(Me.PB_Grundriss.BackgroundImage.Width, Me.PB_Grundriss.BackgroundImage.Height)
            '748; 220   374; 293
            If GSize.Width / 374 > GSize.Height / 293 Then
                'Anpassung an Breite
                tmpScale = CSng(374 / GSize.Width)
                'grundrissRect = New Rectangle(xPage + Get_Zoll_u_width(768), yPage + Get_Zoll_u_height(220) + CInt((Get_Zoll_u_height(293) - CInt(Get_Zoll_u_height(293) * tmpScale)) / 2), Get_Zoll_u_width(321), CInt(Get_Zoll_u_height(293) * tmpScale))
                grundrissRect = New System.Drawing.Rectangle(xPage + Get_Zoll_u_width(748), _
                        yPage + Get_Zoll_u_height(220) + CInt((Get_Zoll_u_height(293) - Get_Zoll_u_width(CInt(GSize.Height * tmpScale))) / 2), _
                        Get_Zoll_u_width(374), Get_Zoll_u_width(CInt(GSize.Height * tmpScale)))
            Else 'Anpassung an Höhe
                tmpScale = CSng(293 / GSize.Height)
                '                grundrissRect = New Rectangle(xPage + Get_Zoll_u_width(768) + CInt((Get_Zoll_u_width(321) - CInt(Get_Zoll_u_width(321) * tmpScale)) / 2), yPage + Get_Zoll_u_height(220), CInt(Get_Zoll_u_width(321) * tmpScale), Get_Zoll_u_height(293))
                grundrissRect = New System.Drawing.Rectangle(xPage + Get_Zoll_u_width(748) + CInt((Get_Zoll_u_width(374) - Get_Zoll_u_height(CInt(GSize.Width * tmpScale))) / 2), _
                        yPage + Get_Zoll_u_height(220), Get_Zoll_u_height(CInt(GSize.Width * tmpScale)), Get_Zoll_u_height(293))
            End If
            gr.DrawImage(Me.PB_Grundriss.BackgroundImage, grundrissRect)
        End If

        'Labelfarben programmieren (Vorlage hat beim Ausdruck einen anderen Farbton
        Print_Klasse("A*", Get_Zoll_u_width(7), Get_Zoll_u_height(713), Get_Zoll_u_width(81), Get_Zoll_u_height(38), gr, True, 6)
        Print_Klasse("A", Get_Zoll_u_width(7), Get_Zoll_u_height(749), Get_Zoll_u_width(81), Get_Zoll_u_height(38), gr, True, 6)
        Print_Klasse("B", Get_Zoll_u_width(7), Get_Zoll_u_height(785), Get_Zoll_u_width(81), Get_Zoll_u_height(38), gr, True, 6)
        Print_Klasse("C", Get_Zoll_u_width(7), Get_Zoll_u_height(821), Get_Zoll_u_width(81), Get_Zoll_u_height(38), gr, True, 6)
        Print_Klasse("D", Get_Zoll_u_width(7), Get_Zoll_u_height(857), Get_Zoll_u_width(81), Get_Zoll_u_height(38), gr, True, 6)
        Print_Klasse("E", Get_Zoll_u_width(7), Get_Zoll_u_height(893), Get_Zoll_u_width(81), Get_Zoll_u_height(38), gr, True, 6)
        Print_Klasse("F", Get_Zoll_u_width(7), Get_Zoll_u_height(929), Get_Zoll_u_width(81), Get_Zoll_u_height(38), gr, True, 6)

        Print_Klasse("A*", Get_Zoll_u_width(7), Get_Zoll_u_height(1029), Get_Zoll_u_width(80), Get_Zoll_u_height(41), gr, True, 6)
        Print_Klasse("A", Get_Zoll_u_width(7), Get_Zoll_u_height(1069), Get_Zoll_u_width(80), Get_Zoll_u_height(41), gr, True, 6)
        Print_Klasse("B", Get_Zoll_u_width(7), Get_Zoll_u_height(1109), Get_Zoll_u_width(80), Get_Zoll_u_height(62), gr, True, 6)
        Print_Klasse("C", Get_Zoll_u_width(7), Get_Zoll_u_height(1170), Get_Zoll_u_width(80), Get_Zoll_u_height(41), gr, True, 6)
        Print_Klasse("D", Get_Zoll_u_width(7), Get_Zoll_u_height(1210), Get_Zoll_u_width(80), Get_Zoll_u_height(160), gr, True, 6)
        Print_Klasse("E", Get_Zoll_u_width(7), Get_Zoll_u_height(1369), Get_Zoll_u_width(80), Get_Zoll_u_height(61), gr, True, 6)
        Print_Klasse("F", Get_Zoll_u_width(7), Get_Zoll_u_height(1429), Get_Zoll_u_width(80), Get_Zoll_u_height(62), gr, True, 6)

        If Not IsNothing(GetImageFromString(My.Settings.Signatur)) Then
            Dim newBtmSig As New Bitmap(GetImageFromString(My.Settings.Signatur))
            GSize = New Size(newBtmSig.Width, newBtmSig.Height)

            '900; 1783  207; 106    961; 1828
            If GSize.Width / 207 > GSize.Height / 106 Then
                'Anpassung an Breite
                tmpScale = CSng(207 / GSize.Width)
                grundrissRect = New System.Drawing.Rectangle(xPage + Get_Zoll_u_width(961), _
                        yPage + Get_Zoll_u_height(1828) + CInt((Get_Zoll_u_height(106) - Get_Zoll_u_width(CInt(GSize.Height * tmpScale))) / 2), _
                        Get_Zoll_u_width(207), Get_Zoll_u_width(CInt(GSize.Height * tmpScale)))
            Else 'Anpassung asn Höhe
                tmpScale = CSng(106 / GSize.Height)
                grundrissRect = New System.Drawing.Rectangle(xPage + Get_Zoll_u_width(961) + CInt((Get_Zoll_u_width(207) - Get_Zoll_u_height(CInt(GSize.Width * tmpScale))) / 2), _
                        yPage + Get_Zoll_u_height(1828), Get_Zoll_u_height(CInt(GSize.Width * tmpScale)), Get_Zoll_u_height(106))
            End If
            gr.DrawImage(newBtmSig, grundrissRect)
        End If
        '5; 1957
        If btDemo > versNormal Then
            Print_Left_Center_Norm_6("Demoversion", Get_Zoll_u_width(5), Get_Zoll_u_height(1999), Get_Zoll_u_height(23), gr)
        Else
            Print_Left_Center_Norm_6(LizenznehmerPR, Get_Zoll_u_width(5), Get_Zoll_u_height(1999), Get_Zoll_u_height(23), gr)
        End If

        If btDemo > versNormal Then Print_Demoversion_quer(gr, 150, -100, 120)

    End Sub

    Private Sub Print_SSA_detailiert(ByVal gr As Graphics)
        '    'Antragsteller
        With Projekt.Antragsteller
            Print_Left_Center_Norm_6(Trim(.Name), Get_Zoll_d_Width(188), Get_Zoll_d_Height(84), Get_Zoll_d_Height(20), gr)
            Print_Left_Center_Norm_6(Trim(.Zusatz), Get_Zoll_d_Width(188), Get_Zoll_d_Height(105), Get_Zoll_d_Height(20), gr)
            Print_Left_Center_Norm_6(Trim(.Strasse) & " " & Trim(.Nr), Get_Zoll_d_Width(188), Get_Zoll_d_Height(126), Get_Zoll_d_Height(20), gr)
            Dim tmpPlz As String = Trim(.PLZ) & " "
            'If tmpPlz = "0 " Then tmpPlz = ""
            Print_Left_Center_Norm_6(tmpPlz & Trim(.Ort), Get_Zoll_d_Width(188), Get_Zoll_d_Height(155), Get_Zoll_d_Height(20), gr)
        End With
        'Gebäude
        With Projekt.Gebaeude.Adresse
            Print_Left_Center_Norm_6(Trim(.Name), Get_Zoll_d_Width(671), Get_Zoll_d_Height(84), Get_Zoll_d_Height(20), gr)
            Print_Left_Center_Norm_6(Trim(.Zusatz), Get_Zoll_d_Width(671), Get_Zoll_d_Height(105), Get_Zoll_d_Height(20), gr)
            Print_Left_Center_Norm_6(Trim(.Strasse) & " " & Trim(.Nr), Get_Zoll_d_Width(671), Get_Zoll_d_Height(126), Get_Zoll_d_Height(20), gr)
            Dim tmpPlz As String = Trim(.PLZ) & " "
            'If tmpPlz = "0 " Then tmpPlz = ""
            Print_Left_Center_Norm_6(tmpPlz & Trim(.Ort), Get_Zoll_d_Width(671), Get_Zoll_d_Height(155), Get_Zoll_d_Height(20), gr)
        End With
        'Wohnungsbezeichnung
        Print_Center_Center(Projekt.Wohnung.Wohnungsbezeichnung, Get_Zoll_d_Width(999), Get_Zoll_d_Height(126), Get_Zoll_d_Width(151), Get_Zoll_d_Height(50), gr, False, 6)
        'Labelfarben programmieren (Vorlage hat beim Ausdruck einen anderen Farbton)
        Print_Klasse("F", Get_Zoll_d_Width(1), Get_Zoll_d_Height(207), Get_Zoll_d_Width(167), Get_Zoll_d_Height(29), gr, True, 6)
        Print_Klasse("E", Get_Zoll_d_Width(167), Get_Zoll_d_Height(207), Get_Zoll_d_Width(167), Get_Zoll_d_Height(29), gr, True, 6)
        Print_Klasse("D", Get_Zoll_d_Width(333), Get_Zoll_d_Height(207), Get_Zoll_d_Width(167), Get_Zoll_d_Height(29), gr, True, 6)
        Print_Klasse("C", Get_Zoll_d_Width(499), Get_Zoll_d_Height(207), Get_Zoll_d_Width(167), Get_Zoll_d_Height(29), gr, True, 6)
        Print_Klasse("B", Get_Zoll_d_Width(665), Get_Zoll_d_Height(207), Get_Zoll_d_Width(166), Get_Zoll_d_Height(29), gr, True, 6)
        Print_Klasse("A", Get_Zoll_d_Width(830), Get_Zoll_d_Height(207), Get_Zoll_d_Width(166), Get_Zoll_d_Height(29), gr, True, 6)
        Print_Klasse("A*", Get_Zoll_d_Width(995), Get_Zoll_d_Height(207), Get_Zoll_d_Width(167), Get_Zoll_d_Height(29), gr, True, 6)

        'Aussteller
        With My.Settings 'Aussteller.Adresse
            Print_Left_Center_Norm_6(Trim(.Aussteller_Name), Get_Zoll_d_Width(88), Get_Zoll_d_Height(1594), Get_Zoll_d_Height(20), gr)
            Print_Left_Center_Norm_6(Trim(.Aussteller_Zusatz), Get_Zoll_d_Width(88), Get_Zoll_d_Height(1615), Get_Zoll_d_Height(20), gr)
            Print_Left_Center_Norm_6(Trim(.Aussteller_Strasse) & " " & Trim(.Aussteller_Nr), Get_Zoll_d_Width(88), Get_Zoll_d_Height(1636), Get_Zoll_d_Height(20), gr)
            Dim tmpPlz As String = Trim(.Aussteller_PLZ) & " "
            'If tmpPlz = "0 " Then tmpPlz = ""
            Print_Left_Center_Norm_6(tmpPlz & Trim(.Aussteller_Ort), Get_Zoll_d_Width(88), Get_Zoll_d_Height(1665), Get_Zoll_d_Height(20), gr)

        End With

        'Rubrik Standort und Außenlärm
        Print_Center_Center(Get_Str_GC(), Get_Zoll_d_Width(253), Get_Zoll_d_Height(332), Get_Zoll_d_Width(439), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_AP(), Get_Zoll_d_Width(253), Get_Zoll_d_Height(406), Get_Zoll_d_Width(291), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_aF(), Get_Zoll_d_Width(546), Get_Zoll_d_Height(406), Get_Zoll_d_Width(146), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Left_Center_Norm_6(Me.TB_Bem_Gebietscharakter.Text, Get_Zoll_d_Width(992), Get_Zoll_d_Height(332), Get_Zoll_d_Height(35), gr) 'Projekt.Standort.Bem_Gebietscharakter, Get_Zoll_d_Width(992), Get_Zoll_d_Height(332), Get_Zoll_d_Height(35), gr)
        Print_Left_Center_Norm_6(Me.TB_Bem_Aussenlaermsituation.Text, Get_Zoll_d_Width(992), Get_Zoll_d_Height(406), Get_Zoll_d_Height(35), gr) 'Trim(Projekt.Standort.Bem_Aussenlaermpegel),

        'Baulicher Schallschutz
        'If Projekt.LS_Wand.Untersuchung = Messung Then    'Me.Label_LS_W_M.Text = "X" Then
        '    gr.FillRectangle(Brushes.Silver, xPage + Get_Zoll_d_Width(399), yPage + Get_Zoll_d_Height(549), Get_Zoll_d_Width(71), Get_Zoll_d_Height(34))
        '    gr.FillRectangle(Brushes.Silver, xPage + Get_Zoll_d_Width(473), yPage + Get_Zoll_d_Height(549), Get_Zoll_d_Width(71), Get_Zoll_d_Height(34))
        'End If
        Print_Center_Center(Get_Str_LS_W_P(), Get_Zoll_d_Width(253), Get_Zoll_d_Height(548), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_LS_W_M(), Get_Zoll_d_Width(326), Get_Zoll_d_Height(548), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        'Print_Center_Center(Get_Str_LS_W_fE(), Get_Zoll_d_Width(399), Get_Zoll_d_Height(548), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        'Print_Center_Center(Get_Str_LS_W_Be(), Get_Zoll_d_Width(472), Get_Zoll_d_Height(548), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_LS_W_C(), Get_Zoll_d_Width(546), Get_Zoll_d_Height(548), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_LS_W_R(), Get_Zoll_d_Width(620), Get_Zoll_d_Height(548), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Left_Center_Norm_6(Projekt.LS_Wand.Bemerkung_LS, Get_Zoll_d_Width(992), Get_Zoll_d_Height(548), Get_Zoll_d_Height(35), gr)
        'Print_Left_Center_Norm_6(Get_Str_LS_W_B(), Get_Zoll_d_Width(992), Get_Zoll_d_Height(548), Get_Zoll_d_Height(35), gr)

        'If Projekt.LS_Decke.Untersuchung = Messung Then    'If Me.Label_LS_D_M.Text = "X" Then
        '    gr.FillRectangle(Brushes.Silver, xPage + Get_Zoll_d_Width(399), yPage + Get_Zoll_d_Height(586), Get_Zoll_d_Width(71), Get_Zoll_d_Height(34))
        '    gr.FillRectangle(Brushes.Silver, xPage + Get_Zoll_d_Width(473), yPage + Get_Zoll_d_Height(586), Get_Zoll_d_Width(71), Get_Zoll_d_Height(34))
        'End If
        Print_Center_Center(Get_Str_LS_D_P(), Get_Zoll_d_Width(253), Get_Zoll_d_Height(585), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_LS_D_M(), Get_Zoll_d_Width(326), Get_Zoll_d_Height(585), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        'Print_Center_Center(Get_Str_LS_D_MA(), Get_Zoll_d_Width(399), Get_Zoll_d_Height(585), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        'Print_Center_Center(Get_Str_LS_D_Be(), Get_Zoll_d_Width(472), Get_Zoll_d_Height(585), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_LS_D_C(), Get_Zoll_d_Width(546), Get_Zoll_d_Height(585), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_LS_D_R(), Get_Zoll_d_Width(620), Get_Zoll_d_Height(585), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Left_Center_Norm_6(Projekt.LS_Decke.Bemerkung_LS, Get_Zoll_d_Width(992), Get_Zoll_d_Height(585), Get_Zoll_d_Height(35), gr)
        'Print_Left_Center_Norm_6(Get_Str_LS_D_B(), Get_Zoll_d_Width(992), Get_Zoll_d_Height(585), Get_Zoll_d_Height(35), gr)

        'If Projekt.TS_Decke.Untersuchung = Messung Then    'If Me.Label_TS_D_M.Text = "X" Then
        gr.FillRectangle(Brushes.Silver, xPage + Get_Zoll_d_Width(399), yPage + Get_Zoll_d_Height(659), Get_Zoll_d_Width(71), Get_Zoll_d_Height(34))
        gr.FillRectangle(Brushes.Silver, xPage + Get_Zoll_d_Width(473), yPage + Get_Zoll_d_Height(659), Get_Zoll_d_Width(71), Get_Zoll_d_Height(34))

        Print_Header_TS_D(Get_Header_TS_D_fE(), Get_Zoll_d_Width(405), Get_Zoll_d_Height(630), Get_Zoll_d_Width(62), Get_Zoll_d_Height(27), gr, False, 5)

        'End If
        Print_Center_Center(Get_Str_TS_D_P(), Get_Zoll_d_Width(253), Get_Zoll_d_Height(659), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_TS_D_M(), Get_Zoll_d_Width(326), Get_Zoll_d_Height(659), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_TS_D_fE(), Get_Zoll_d_Width(399), Get_Zoll_d_Height(659), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_TS_D_Be(), Get_Zoll_d_Width(472), Get_Zoll_d_Height(659), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_TS_D_C(), Get_Zoll_d_Width(546), Get_Zoll_d_Height(659), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_TS_D_L(), Get_Zoll_d_Width(620), Get_Zoll_d_Height(659), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Left_Center_Norm_6(Projekt.TS_Decke.Bemerkung_TS_Decke, Get_Zoll_d_Width(992), Get_Zoll_d_Height(659), Get_Zoll_d_Height(35), gr)
        'Print_Left_Center_Norm_6(Get_Str_TS_D_B(), Get_Zoll_d_Width(992), Get_Zoll_d_Height(659), Get_Zoll_d_Height(35), gr)

        'If Projekt.TS_TPH.Untersuchung = Messung Then    'If Me.Label_TS_TPH_M.Text = "X" Then
        gr.FillRectangle(Brushes.Silver, xPage + Get_Zoll_d_Width(473), yPage + Get_Zoll_d_Height(696), Get_Zoll_d_Width(71), Get_Zoll_d_Height(34))
        'End If
        Print_Center_Center(Get_Str_TS_TPLH_P(), Get_Zoll_d_Width(253), Get_Zoll_d_Height(696), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_TS_TPLH_M(), Get_Zoll_d_Width(326), Get_Zoll_d_Height(696), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_TS_TPLH_Be(), Get_Zoll_d_Width(472), Get_Zoll_d_Height(696), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_TS_TPLH_C(), Get_Zoll_d_Width(546), Get_Zoll_d_Height(696), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_TS_TPLH_L(), Get_Zoll_d_Width(620), Get_Zoll_d_Height(696), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Left_Center_Norm_6(Projekt.TS_TPLH.Bemerkung_TS_TPH, Get_Zoll_d_Width(992), Get_Zoll_d_Height(696), Get_Zoll_d_Height(35), gr)
        'Print_Left_Center_Norm_6(Get_Str_TS_TPH_B(), Get_Zoll_d_Width(992), Get_Zoll_d_Height(696), Get_Zoll_d_Height(35), gr)

        'If Projekt.TS_BLLT.Untersuchung = Messung Then    'If Me.Label_TS_BLLT_M.Text = "X" Then
        gr.FillRectangle(Brushes.Silver, xPage + Get_Zoll_d_Width(473), yPage + Get_Zoll_d_Height(733), Get_Zoll_d_Width(71), Get_Zoll_d_Height(34))
        'End If
        Print_Center_Center(Get_Str_TS_BLT_P(), Get_Zoll_d_Width(253), Get_Zoll_d_Height(733), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_TS_BLT_M(), Get_Zoll_d_Width(326), Get_Zoll_d_Height(733), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_TS_BLT_Be(), Get_Zoll_d_Width(472), Get_Zoll_d_Height(733), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_TS_BLT_C(), Get_Zoll_d_Width(546), Get_Zoll_d_Height(733), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_TS_BLT_L(), Get_Zoll_d_Width(620), Get_Zoll_d_Height(733), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Left_Center_Norm_6(Projekt.TS_BLT.Bemerkung_TS_BLLT, Get_Zoll_d_Width(992), Get_Zoll_d_Height(733), Get_Zoll_d_Height(35), gr)
        'Print_Left_Center_Norm_6(Get_Str_TS_BLLT_B(), Get_Zoll_d_Width(992), Get_Zoll_d_Height(733), Get_Zoll_d_Height(35), gr)

        Print_Center_Center(Get_Str_Tueren_P_D(), Get_Zoll_d_Width(253), Get_Zoll_d_Height(815), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_Tueren_M_D(), Get_Zoll_d_Width(326), Get_Zoll_d_Height(815), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_Tueren_R_D(), Get_Zoll_d_Width(546), Get_Zoll_d_Height(815), Get_Zoll_d_Width(146), Get_Zoll_d_Height(35), gr, False, 6)

        Print_Center_Center(Get_Str_Tueren_P_A(), Get_Zoll_d_Width(253), Get_Zoll_d_Height(852), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_Tueren_M_A(), Get_Zoll_d_Width(326), Get_Zoll_d_Height(852), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_Tueren_R_A(), Get_Zoll_d_Width(546), Get_Zoll_d_Height(852), Get_Zoll_d_Width(146), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Left_Center_Norm_6(Projekt.Tueren.Bemerkung_Tueren, Get_Zoll_d_Width(992), Get_Zoll_d_Height(852), Get_Zoll_d_Height(35), gr)
        'Print_Left_Center_Norm_6(Get_Str_Tueren_B(), Get_Zoll_d_Width(992), Get_Zoll_d_Height(852), Get_Zoll_d_Height(35), gr)

        Print_Center_Center(Get_Str_Aussenbauteile_oN(), Get_Zoll_d_Width(253), Get_Zoll_d_Height(963), Get_Zoll_d_Width(145), Get_Zoll_d_Height(35), gr, False, 6)
        '  Print_Center_Center(Get_Str_Aussenbauteile_FmD(), Get_Zoll_d_Width(326), Get_Zoll_d_Height(963), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_Aussenbauteile_DIN(), Get_Zoll_d_Width(399), Get_Zoll_d_Height(963), Get_Zoll_d_Width(145), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_Aussenbauteile_DINPlus(), Get_Zoll_d_Width(546), Get_Zoll_d_Height(963), Get_Zoll_d_Width(146), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Left_Center_Norm_6(Projekt.Bem_Aussenbauteile, Get_Zoll_d_Width(992), Get_Zoll_d_Height(963), Get_Zoll_d_Height(35), gr)

        Print_Center_Center(Get_Str_Wasser_P(), Get_Zoll_d_Width(253), Get_Zoll_d_Height(1037), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_Wasser_M(), Get_Zoll_d_Width(326), Get_Zoll_d_Height(1037), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_Wasser_LcLa(), Get_Zoll_d_Width(399), Get_Zoll_d_Height(1037), Get_Zoll_d_Width(145), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_Wasser_L(), Get_Zoll_d_Width(546), Get_Zoll_d_Height(1037), Get_Zoll_d_Width(146), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Left_Center_Norm_6(Projekt.Wasser.Bem_Wasser, Get_Zoll_d_Width(992), Get_Zoll_d_Height(1037), Get_Zoll_d_Height(35), gr)

        Print_Center_Center(Get_Str_Nutzer_nv(), Get_Zoll_d_Width(253), Get_Zoll_d_Height(1111), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_Nutzer_P(), Get_Zoll_d_Width(326), Get_Zoll_d_Height(1111), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_Nutzer_M(), Get_Zoll_d_Width(399), Get_Zoll_d_Height(1111), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_Nutzer_L(), Get_Zoll_d_Width(546), Get_Zoll_d_Height(1111), Get_Zoll_d_Width(146), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Left_Center_Norm_6(Projekt.NutzerKoerper.BemerkungNutzer, Get_Zoll_d_Width(992), Get_Zoll_d_Height(1111), Get_Zoll_d_Height(35), gr)
        'Print_Left_Center_Norm_6(Get_Str_Nutzer_B(), Get_Zoll_d_Width(992), Get_Zoll_d_Height(1111), Get_Zoll_d_Height(35), gr)

        Print_Center_Center(Get_Str_Koerper_nv(), Get_Zoll_d_Width(253), Get_Zoll_d_Height(1148), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_Koerper_P(), Get_Zoll_d_Width(326), Get_Zoll_d_Height(1148), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_Koerper_M(), Get_Zoll_d_Width(399), Get_Zoll_d_Height(1148), Get_Zoll_d_Width(72), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_Koerper_L(), Get_Zoll_d_Width(546), Get_Zoll_d_Height(1148), Get_Zoll_d_Width(146), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Left_Center_Norm_6(Projekt.NutzerKoerper.BemerkungKoerper, Get_Zoll_d_Width(992), Get_Zoll_d_Height(1148), Get_Zoll_d_Height(35), gr)
        'Print_Left_Center_Norm_6(Get_Str_Koerper_B(), Get_Zoll_d_Width(992), Get_Zoll_d_Height(1148), Get_Zoll_d_Height(35), gr)

        Print_Center_Center(Get_Str_Nachbar(), Get_Zoll_d_Width(546), Get_Zoll_d_Height(1185), Get_Zoll_d_Width(146), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Left_Center_Norm_6(Projekt.Bemerkung_Nachbarn, Get_Zoll_d_Width(992), Get_Zoll_d_Height(1185), Get_Zoll_d_Height(35), gr)
        'Print_Left_Center_Norm_6(Get_Str_Nachbar_B(), Get_Zoll_d_Width(992), Get_Zoll_d_Height(1185), Get_Zoll_d_Height(35), gr)

        Print_Center_Center(Get_Str_anordRaeume_ug(), Get_Zoll_d_Width(326), Get_Zoll_d_Height(1222), Get_Zoll_d_Width(146), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_anordRaeume_g(), Get_Zoll_d_Width(546), Get_Zoll_d_Height(1222), Get_Zoll_d_Width(146), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Left_Center_Norm_6(Projekt.Bem_anordnungRaeume, Get_Zoll_d_Width(992), Get_Zoll_d_Height(1222), Get_Zoll_d_Height(35), gr)

        Print_Center_Center(Get_Str_lauteRaeume_keine(), Get_Zoll_d_Width(546), Get_Zoll_d_Height(1259), Get_Zoll_d_Width(146), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_lauteRaeume_L_25_35(), Get_Zoll_d_Width(546), Get_Zoll_d_Height(1297), Get_Zoll_d_Width(146), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_lauteRaeume_L_30_40(), Get_Zoll_d_Width(546), Get_Zoll_d_Height(1334), Get_Zoll_d_Width(146), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_lauteRaeume_L_35_45(), Get_Zoll_d_Width(546), Get_Zoll_d_Height(1372), Get_Zoll_d_Width(146), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Left_Center_Norm_6(Projekt.Bem_lauteRaeume, Get_Zoll_d_Width(992), Get_Zoll_d_Height(1372), Get_Zoll_d_Height(35), gr)
        'Print_Left_Center_Norm_6(Get_Str_Nachbar_B(), Get_Zoll_d_Width(992), Get_Zoll_d_Height(1185), Get_Zoll_d_Height(35), gr)

        Print_Center_Center(Get_Str_NHZ(), Get_Zoll_d_Width(253), Get_Zoll_d_Height(1410), Get_Zoll_d_Width(438), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Left_Center_Norm_6(Projekt.Bemerkung_AV, Get_Zoll_d_Width(992), Get_Zoll_d_Height(1410), Get_Zoll_d_Height(35), gr)

        Print_Center_Center(Get_Str_EW_1(), Get_Zoll_d_Width(399), Get_Zoll_d_Height(1485), Get_Zoll_d_Width(73), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_EW_2(), Get_Zoll_d_Width(473), Get_Zoll_d_Height(1485), Get_Zoll_d_Width(73), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_EW_3(), Get_Zoll_d_Width(546), Get_Zoll_d_Height(1485), Get_Zoll_d_Width(73), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Str_EW_kE(), Get_Zoll_d_Width(620), Get_Zoll_d_Height(1485), Get_Zoll_d_Width(73), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Left_Center_Norm_6(Projekt.Bemerkung_eigenerWohnbereich, Get_Zoll_d_Width(992), Get_Zoll_d_Height(1485), Get_Zoll_d_Height(35), gr)
        'Print_Left_Center_Norm_6(Get_Str_EW_B(), Get_Zoll_d_Width(992), Get_Zoll_d_Height(1447), Get_Zoll_d_Height(35), gr)

        '    '########### Punktevergabe  ####################
        Print_Center_Center(Get_Pkte_GC(), Get_Zoll_d_Width(700), Get_Zoll_d_Height(332), Get_Zoll_d_Width(68), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Pkte_Aussenlaerm(), Get_Zoll_d_Width(700), Get_Zoll_d_Height(406), Get_Zoll_d_Width(68), Get_Zoll_d_Height(35), gr, False, 6)
        '698; 548   68; 35  
        Print_Center_Center(Get_Pkte_LS_W(), Get_Zoll_d_Width(700), Get_Zoll_d_Height(548), Get_Zoll_d_Width(68), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Pkte_LS_D(), Get_Zoll_d_Width(700), Get_Zoll_d_Height(548 + 37), Get_Zoll_d_Width(68), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Pkte_TS_D(), Get_Zoll_d_Width(700), Get_Zoll_d_Height(548 + 111), Get_Zoll_d_Width(68), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Pkte_TS_TPLH(), Get_Zoll_d_Width(700), Get_Zoll_d_Height(548 + 148), Get_Zoll_d_Width(68), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Pkte_TS_BLT(), Get_Zoll_d_Width(700), Get_Zoll_d_Height(548 + 185), Get_Zoll_d_Width(68), Get_Zoll_d_Height(35), gr, False, 6)

        Print_Center_Center(Get_Pkte_Tueren_D(), Get_Zoll_d_Width(700), Get_Zoll_d_Height(548 + 267), Get_Zoll_d_Width(68), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Pkte_Tueren_A(), Get_Zoll_d_Width(700), Get_Zoll_d_Height(548 + 304), Get_Zoll_d_Width(68), Get_Zoll_d_Height(35), gr, False, 6)

        Print_Center_Center(Get_Pkte_Aussenbauteile(), Get_Zoll_d_Width(700), Get_Zoll_d_Height(548 + 415), Get_Zoll_d_Width(68), Get_Zoll_d_Height(35), gr, False, 6)

        Print_Center_Center(Get_Pkte_Wasser(), Get_Zoll_d_Width(700), Get_Zoll_d_Height(548 + 489), Get_Zoll_d_Width(68), Get_Zoll_d_Height(35), gr, False, 6)

        'Print_Center_Center(Get_Pkte_Nutzer(), Get_Zoll_d_Width(700), Get_Zoll_d_Height(548 + 563), Get_Zoll_d_Width(68), Get_Zoll_d_Height(35), gr, False, 6)
        'Print_Center_Center(Get_Pkte_Koerper(), Get_Zoll_d_Width(700), Get_Zoll_d_Height(548 + 600), Get_Zoll_d_Width(68), Get_Zoll_d_Height(35), gr, False, 6)

        Print_Center_Center(Get_Pkte_NutzerKoerper(), Get_Zoll_d_Width(700), Get_Zoll_d_Height(548 + 600), Get_Zoll_d_Width(68), Get_Zoll_d_Height(35), gr, False, 6)

        Print_Center_Center(Get_Pkte_Nachbarn(), Get_Zoll_d_Width(700), Get_Zoll_d_Height(548 + 637), Get_Zoll_d_Width(68), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Pkte_anordRaeume(), Get_Zoll_d_Width(700), Get_Zoll_d_Height(548 + 674), Get_Zoll_d_Width(68), Get_Zoll_d_Height(35), gr, False, 6)

        Print_Center_Center(Get_Pkte_lauteRaeume_keine(), Get_Zoll_d_Width(700), Get_Zoll_d_Height(548 + 711), Get_Zoll_d_Width(68), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Pkte_lauteRaeume_25(), Get_Zoll_d_Width(700), Get_Zoll_d_Height(548 + 749), Get_Zoll_d_Width(68), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Pkte_lauteRaeume_30(), Get_Zoll_d_Width(700), Get_Zoll_d_Height(548 + 786), Get_Zoll_d_Width(68), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Pkte_lauteRaeume_35(), Get_Zoll_d_Width(700), Get_Zoll_d_Height(548 + 824), Get_Zoll_d_Width(68), Get_Zoll_d_Height(35), gr, False, 6)

        Print_Center_Center(Get_Pkte_NHZ(), Get_Zoll_d_Width(700), Get_Zoll_d_Height(548 + 862), Get_Zoll_d_Width(68), Get_Zoll_d_Height(35), gr, False, 6)

        Print_Center_Center(Get_Pkte_EW(), Get_Zoll_d_Width(700), Get_Zoll_d_Height(548 + 936), Get_Zoll_d_Width(68), Get_Zoll_d_Height(35), gr, False, 6)

        '    '########### Klasse ####################
        Print_Klasse(Get_Klasse_GC, Get_Zoll_d_Width(776), Get_Zoll_d_Height(332), Get_Zoll_d_Width(78), Get_Zoll_d_Height(34), gr, False, 6)

        Print_Klasse(Get_Klasse_AP(), Get_Zoll_d_Width(776), Get_Zoll_d_Height(406), Get_Zoll_d_Width(78), Get_Zoll_d_Height(34), gr, False, 6)

        Print_Klasse(Get_Klasse_LS_Waende(), Get_Zoll_d_Width(776), Get_Zoll_d_Height(548), Get_Zoll_d_Width(78), Get_Zoll_d_Height(34), gr, False, 6)

        Print_Klasse(Get_Klasse_LS_Decken(), Get_Zoll_d_Width(776), Get_Zoll_d_Height(585), Get_Zoll_d_Width(78), Get_Zoll_d_Height(34), gr, False, 6)

        Print_Klasse(Get_Klasse_TS_Decken(), Get_Zoll_d_Width(776), Get_Zoll_d_Height(659), Get_Zoll_d_Width(78), Get_Zoll_d_Height(34), gr, False, 6)

        Print_Klasse(Get_Klasse_TS_TPLH(), Get_Zoll_d_Width(776), Get_Zoll_d_Height(696), Get_Zoll_d_Width(78), Get_Zoll_d_Height(34), gr, False, 6)

        Print_Klasse(Get_Klasse_TS_BLT(), Get_Zoll_d_Width(776), Get_Zoll_d_Height(733), Get_Zoll_d_Width(78), Get_Zoll_d_Height(34), gr, False, 6)

        Print_Klasse(Get_Klasse_Tueren_A(), Get_Zoll_d_Width(776), Get_Zoll_d_Height(815), Get_Zoll_d_Width(78), Get_Zoll_d_Height(34), gr, False, 6)

        Print_Klasse(Get_Klasse_Tueren_D(), Get_Zoll_d_Width(776), Get_Zoll_d_Height(852), Get_Zoll_d_Width(78), Get_Zoll_d_Height(34), gr, False, 6)

        Print_Klasse(Get_Klasse_Aussenbauteile(), Get_Zoll_d_Width(776), Get_Zoll_d_Height(963), Get_Zoll_d_Width(78), Get_Zoll_d_Height(34), gr, False, 6)

        Print_Klasse(Get_Klasse_Wasser(), Get_Zoll_d_Width(776), Get_Zoll_d_Height(1037), Get_Zoll_d_Width(78), Get_Zoll_d_Height(34), gr, False, 6)

        'Print_Center_Center(Get_Klasse_Nutzer(), Get_Zoll_d_Width(776), Get_Zoll_d_Height(1111), Get_Zoll_d_Width(78), Get_Zoll_d_Height(34), gr, False, 6)
        'Print_Center_Center((Get_Klasse_Koerper()), Get_Zoll_d_Width(776), Get_Zoll_d_Height(1148), Get_Zoll_d_Width(78), Get_Zoll_d_Height(34), gr, False, 6)

        'Print_Center_Center((Get_Klasse_NutzerKoerper()), Get_Zoll_d_Width(776), Get_Zoll_d_Height(1148), Get_Zoll_d_Width(78), Get_Zoll_d_Height(34), gr, False, 6)
        Print_Klasse((Get_Klasse_NutzerKoerper()), Get_Zoll_d_Width(776), Get_Zoll_d_Height(1148), Get_Zoll_d_Width(78), Get_Zoll_d_Height(34), gr, False, 6)

        'Print_Center_Center(Get_Klasse_Nachbarn(), Get_Zoll_d_Width(776), Get_Zoll_d_Height(1185), Get_Zoll_d_Width(78), Get_Zoll_d_Height(34), gr, False, 6)
        Print_Klasse(Get_Klasse_Nachbarn(), Get_Zoll_d_Width(776), Get_Zoll_d_Height(1185), Get_Zoll_d_Width(78), Get_Zoll_d_Height(34), gr, False, 6)

        Print_Klasse(Get_Klasse_lauteRaeume_kE(), Get_Zoll_d_Width(776), Get_Zoll_d_Height(1259), Get_Zoll_d_Width(78), Get_Zoll_d_Height(34), gr, False, 6)

        Print_Klasse(Get_Klasse_lauteRaeume_25(), Get_Zoll_d_Width(776), Get_Zoll_d_Height(1297), Get_Zoll_d_Width(78), Get_Zoll_d_Height(34), gr, False, 6)

        Print_Klasse(Get_Klasse_lauteRaeume_30(), Get_Zoll_d_Width(776), Get_Zoll_d_Height(1334), Get_Zoll_d_Width(78), Get_Zoll_d_Height(34), gr, False, 6)

        Print_Klasse(Get_Klasse_lauteRaeume_35(), Get_Zoll_d_Width(776), Get_Zoll_d_Height(1372), Get_Zoll_d_Width(78), Get_Zoll_d_Height(34), gr, False, 6)

        Print_Center_Center(Get_Klasse_NHZ(), Get_Zoll_d_Width(776), Get_Zoll_d_Height(1410), Get_Zoll_d_Width(78), Get_Zoll_d_Height(34), gr, False, 6)

        Print_Center_Center(Get_Klasse_EW(), Get_Zoll_d_Width(776), Get_Zoll_d_Height(1485), Get_Zoll_d_Width(78), Get_Zoll_d_Height(34), gr, False, 6)

        Dim Get_Calc_Pkte_I As String = Me.Calc_Pkte_I
        Dim Get_Calc_Pkte_II As String = Me.Calc_Pkte_II
        Print_Klasse_Color(Get_Klasse_I, Get_Zoll_d_Width(699), Get_Zoll_d_Height(1618), Get_Zoll_d_Width(156), Get_Zoll_d_Height(54), gr)
        Print_Klasse_Color(Get_Klasse_II, Get_Zoll_d_Width(699), Get_Zoll_d_Height(1679), Get_Zoll_d_Width(156), Get_Zoll_d_Height(55), gr)
        Print_Center_Center(Get_Calc_Pkte_I, Get_Zoll_d_Width(699), Get_Zoll_d_Height(445), Get_Zoll_d_Width(156), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Calc_Pkte_II, Get_Zoll_d_Width(699), Get_Zoll_d_Height(1522), Get_Zoll_d_Width(156), Get_Zoll_d_Height(35), gr, False, 6)
        Print_Center_Center(Get_Calc_Pkte_I, Get_Zoll_d_Width(699), Get_Zoll_d_Height(1618), Get_Zoll_d_Width(156), Get_Zoll_d_Height(48), gr, True, 20)
        Print_Center_Center(Get_Calc_Pkte_II, Get_Zoll_d_Width(699), Get_Zoll_d_Height(1679), Get_Zoll_d_Width(156), Get_Zoll_d_Height(48), gr, True, 20)

        Print_Klasse_Color(Get_Klasse_I, Get_Zoll_d_Width(863), Get_Zoll_d_Height(1618), Get_Zoll_d_Width(120), Get_Zoll_d_Height(54), gr)
        Print_Klasse_Color(Get_Klasse_II, Get_Zoll_d_Width(863), Get_Zoll_d_Height(1679), Get_Zoll_d_Width(120), Get_Zoll_d_Height(55), gr)
        Print_Center_Center(Get_Klasse_I(), Get_Zoll_d_Width(862), Get_Zoll_d_Height(308), Get_Zoll_d_Width(121), Get_Zoll_d_Height(172), gr, True, 14)
        Print_Center_Center(Get_Klasse_II(), Get_Zoll_d_Width(862), Get_Zoll_d_Height(489), Get_Zoll_d_Width(121), Get_Zoll_d_Height(1027), gr, True, 14)
        '863; 1580  120; 54
        Print_Center_Center(Get_Klasse_I(), Get_Zoll_d_Width(863), Get_Zoll_d_Height(1618), Get_Zoll_d_Width(120), Get_Zoll_d_Height(54), gr, True, 20)
        '863; 1641  120; 54
        Print_Center_Center(Get_Klasse_II, Get_Zoll_d_Width(863), Get_Zoll_d_Height(1679), Get_Zoll_d_Width(120), Get_Zoll_d_Height(54), gr, True, 20)
        With Projekt
            If Not IsNothing(.Datum) Then
                '104; 1680  64; 16
                If .Datum <> "" Then
                    'Print_Left_Center_Norm_6(CType(.Datum, Date).Day & "." & CType(.Datum, Date).Month & "." & CType(.Datum, Date).Year, _
                    '        Get_Zoll_d_Width(104), Get_Zoll_d_Height(1680), Get_Zoll_d_Height(16), gr)
                    Print_Left_Center_Norm_6(.Datum,
                            Get_Zoll_d_Width(88), Get_Zoll_d_Height(1698), Get_Zoll_d_Height(16), gr)
                    '994; 1665  159; 30
                    'Print_Center_Center(CType(.Datum, Date).Day & "." & CType(.Datum, Date).Month & "." & CType(.Datum, Date).AddYears(10).Year, _
                    '        Get_Zoll_d_Width(994), Get_Zoll_d_Height(1665), Get_Zoll_d_Width(159), Get_Zoll_d_Height(30), gr, False, 6)
                    Dim tDatum As Date = CType(.Datum, Date)
                    Print_Center_Center(IIf(tDatum.Day < 10, "0" & tDatum.Day, tDatum.Day).ToString & "." & IIf(tDatum.Month < 10, "0" & tDatum.Month, tDatum.Month).ToString & "." & tDatum.AddYears(10).Year, _
                                Get_Zoll_d_Width(994), Get_Zoll_d_Height(1703), Get_Zoll_d_Width(159), Get_Zoll_d_Height(30), gr, False, 6)
                End If
            End If
        End With
        Dim GSize As Size
        Dim tmpscale As Single
        Dim sigRect As System.Drawing.Rectangle
        If Not IsNothing(GetImageFromString(My.Settings.Signatur)) Then

            Dim newBtmSig As New Bitmap(GetImageFromString(My.Settings.Signatur))
            GSize = New Size(newBtmSig.Width, newBtmSig.Height)

            If GSize.Width / 165 > GSize.Height / 55 Then
                'Anpassung an Breite
                tmpscale = CSng(165 / GSize.Width)
                sigRect = New System.Drawing.Rectangle(xPage + Get_Zoll_d_Width(991), _
                        yPage + Get_Zoll_d_Height(1617) + CInt((Get_Zoll_d_Height(55) - Get_Zoll_d_Width(CInt(GSize.Height * tmpscale))) / 2), _
                        Get_Zoll_d_Width(165), Get_Zoll_d_Width(CInt(GSize.Height * tmpscale)))
            Else 'Anpassung asn Höhe
                tmpscale = CSng(55 / GSize.Height)
                sigRect = New System.Drawing.Rectangle(xPage + Get_Zoll_d_Width(991) + CInt((Get_Zoll_d_Width(165) - Get_Zoll_d_Height(CInt(GSize.Width * tmpscale))) / 2), _
                        yPage + Get_Zoll_d_Height(1617), Get_Zoll_d_Height(CInt(GSize.Width * tmpscale)), Get_Zoll_d_Height(55))
            End If
            gr.DrawImage(newBtmSig, sigRect)
        End If
        '5; 1733
        If btDemo > versNormal Then
            Print_Left_Center_Norm_6("Demoversion", Get_Zoll_d_Width(5), Get_Zoll_d_Height(1771), Get_Zoll_d_Height(19), gr)
        Else
            Print_Left_Center_Norm_6(LizenznehmerPR, Get_Zoll_d_Width(5), Get_Zoll_d_Height(1771), Get_Zoll_d_Height(19), gr)
        End If

        If btDemo > versNormal Then Print_Demoversion_quer(gr, 150, -100, 120)
    End Sub
    Private Sub Print_SSA_einfach(ByVal gr As Graphics)
        'DEGA-Logo
        Dim degaRect As New System.Drawing.Rectangle(xPage + Get_Zoll_e_Width(29), yPage + Get_Zoll_e_Height(13), Get_Zoll_e_Height(87), Get_Zoll_e_Height(87)) 'xpage+get_zoll_e_width(),ypage+get_zoll_e_height()+cint((get-zoll_e_height()-get_zoll_e_width(cint(
        gr.DrawImage(PB_DEGA.BackgroundImage, degaRect)

        'Logo
        If Not IsNothing(PB_e_Logo.BackgroundImage) Then
            Dim imScale As Double
            If 140 / PB_e_Logo.BackgroundImage.Width < 89 / PB_e_Logo.BackgroundImage.Height Then
                imScale = 140 / PB_e_Logo.BackgroundImage.Width
            Else
                imScale = 89 / PB_e_Logo.BackgroundImage.Height
            End If
            ''Loc: 950; 13  Siz: 140; 89
            Dim x As Integer = CInt(950 + 140 / 2 - PB_e_Logo.BackgroundImage.Width * imScale)
            Dim y As Integer = CInt(13 + 89 / 2 - PB_e_Logo.BackgroundImage.Height * imScale / 2)
            Dim w As Integer = CInt(PB_e_Logo.BackgroundImage.Width * imScale)
            Dim h As Integer = CInt(PB_e_Logo.BackgroundImage.Height * imScale)
            Dim logoRect As New System.Drawing.Rectangle(xPage + Get_Zoll_e_Width(x), yPage + Get_Zoll_e_Height(y), _
                Get_Zoll_e_Height(w), Get_Zoll_e_Height(h))

            gr.DrawImage(PB_e_Logo.BackgroundImage, logoRect)
        End If

        'Antragsteller
        With Projekt.Antragsteller
            Print_Left_Center_Norm_6(Trim(.Name), Get_Zoll_e_Width(168), Get_Zoll_e_Height(142), Get_Zoll_e_Height(20), gr)
            Print_Left_Center_Norm_6(Trim(.Zusatz), Get_Zoll_e_Width(168), Get_Zoll_e_Height(163), Get_Zoll_e_Height(20), gr)
            Print_Left_Center_Norm_6(Trim(.Strasse) & " " & Trim(.Nr), Get_Zoll_e_Width(168), Get_Zoll_e_Height(184), Get_Zoll_e_Height(20), gr)
            Dim tmpPlz As String = Trim(.PLZ) & " "
            'If tmpPlz = "0 " Then tmpPlz = ""
            Print_Left_Center_Norm_6(tmpPlz & Trim(.Ort), Get_Zoll_e_Width(168), Get_Zoll_e_Height(213), Get_Zoll_e_Height(20), gr)
        End With
        'Gebäude
        With Projekt.Gebaeude.Adresse
            Print_Left_Center_Norm_6(Trim(.Name), Get_Zoll_e_Width(611), Get_Zoll_e_Height(142), Get_Zoll_e_Height(20), gr)
            Print_Left_Center_Norm_6(.Zusatz, Get_Zoll_e_Width(611), Get_Zoll_e_Height(163), Get_Zoll_e_Height(20), gr)
            Print_Left_Center_Norm_6(Trim(.Strasse) & " " & Trim(.Nr), Get_Zoll_e_Width(611), Get_Zoll_e_Height(184), Get_Zoll_e_Height(20), gr)
            Dim tmpPlz As String = Trim(.PLZ) & " "
            'If tmpPlz = "0 " Then tmpPlz = ""
            Print_Left_Center_Norm_6(tmpPlz & Trim(.Ort), Get_Zoll_e_Width(611), Get_Zoll_e_Height(213), Get_Zoll_e_Height(20), gr)
        End With
        'Wohnungsbezeichnung
        Print_Center_Center(Projekt.Wohnung.Wohnungsbezeichnung, Get_Zoll_e_Width(939), Get_Zoll_e_Height(184), Get_Zoll_e_Width(151), Get_Zoll_e_Height(50), gr, False, 6)
        ''Labelfarben programmieren (Vorlage hat beim Ausdruck einen anderen Farbton)
        'Print_Klasse_Polygon("F", New System.Drawing.Point(Get_Zoll_e_Width(320), Get_Zoll_e_Height(455)), New System.Drawing.Point(Get_Zoll_e_Width(437), Get_Zoll_e_Height(455)), _
        '                            New System.Drawing.Point(Get_Zoll_e_Width(437), Get_Zoll_e_Height(615)), New System.Drawing.Point(Get_Zoll_e_Width(320), Get_Zoll_e_Height(615)), _
        '                            New System.Drawing.Point(Get_Zoll_e_Width(360), Get_Zoll_e_Height(535)), gr)
        'Print_Left_Center_Norm_6("0", Get_Zoll_e_Width(358), Get_Zoll_e_Height(591), Get_Zoll_e_Height(14), gr)
        ' '' ''Print_Center_Center("0", Get_Zoll_e_Width(358), Get_Zoll_e_Height(591), Get_Zoll_e_Width(12), Get_Zoll_e_Height(14), gr, False, 6)
        'Print_Klasse("E", Get_Zoll_e_Width(448), Get_Zoll_e_Height(455), Get_Zoll_e_Width(79), Get_Zoll_e_Height(160), gr, False, 17)
        'Print_Left_Center_Norm_6("10", Get_Zoll_e_Width(447), Get_Zoll_e_Height(591), Get_Zoll_e_Height(14), gr)
        ' '' ''Print_Center_Center("10", Get_Zoll_e_Width(447), Get_Zoll_e_Height(591), Get_Zoll_e_Width(16), Get_Zoll_e_Height(14), gr, False, 6)
        'Print_Klasse("D", Get_Zoll_e_Width(536), Get_Zoll_e_Height(455), Get_Zoll_e_Width(77), Get_Zoll_e_Height(160), gr, False, 17)
        'Print_Left_Center_Norm_6("20", Get_Zoll_e_Width(534), Get_Zoll_e_Height(591), Get_Zoll_e_Height(14), gr)
        ' '' ''Print_Center_Center("20", Get_Zoll_e_Width(534), Get_Zoll_e_Height(591), Get_Zoll_e_Width(15), Get_Zoll_e_Height(14), gr, False, 6)
        'Print_Klasse("C", Get_Zoll_e_Width(622), Get_Zoll_e_Height(455), Get_Zoll_e_Width(77), Get_Zoll_e_Height(160), gr, False, 17)
        'Print_Left_Center_Norm_6("25", Get_Zoll_e_Width(621), Get_Zoll_e_Height(591), Get_Zoll_e_Height(14), gr)
        ' '' ''Print_Center_Center("25", Get_Zoll_e_Width(621), Get_Zoll_e_Height(591), Get_Zoll_e_Width(15), Get_Zoll_e_Height(14), gr, False, 6)
        'Print_Klasse("B", Get_Zoll_e_Width(708), Get_Zoll_e_Height(455), Get_Zoll_e_Width(77), Get_Zoll_e_Height(160), gr, False, 17)
        'Print_Left_Center_Norm_6("40", Get_Zoll_e_Width(706), Get_Zoll_e_Height(591), Get_Zoll_e_Height(14), gr)
        ' '' ''        Print_Center_Center("40", Get_Zoll_e_Width(706), Get_Zoll_e_Height(591), Get_Zoll_e_Width(15), Get_Zoll_e_Height(14), gr, False, 6)
        'Print_Klasse("A", Get_Zoll_e_Width(794), Get_Zoll_e_Height(455), Get_Zoll_e_Width(77), Get_Zoll_e_Height(160), gr, False, 17)
        'Print_Left_Center_Norm_6("45", Get_Zoll_e_Width(793), Get_Zoll_e_Height(591), Get_Zoll_e_Height(14), gr)
        ' '' ''Print_Center_Center("45", Get_Zoll_e_Width(793), Get_Zoll_e_Height(591), Get_Zoll_e_Width(15), Get_Zoll_e_Height(14), gr, False, 6)
        'Print_Klasse_Polygon("A*", New System.Drawing.Point(Get_Zoll_e_Width(878), Get_Zoll_e_Height(455)), New System.Drawing.Point(Get_Zoll_e_Width(956), Get_Zoll_e_Height(455)), _
        '                            New System.Drawing.Point(Get_Zoll_e_Width(995), Get_Zoll_e_Height(535)), New System.Drawing.Point(Get_Zoll_e_Width(956), Get_Zoll_e_Height(615)), _
        '                            New System.Drawing.Point(Get_Zoll_e_Width(878), Get_Zoll_e_Height(615)), gr)
        'Print_Left_Center_Norm_6("55", Get_Zoll_e_Width(879), Get_Zoll_e_Height(591), Get_Zoll_e_Height(14), gr)
        ' '' ''Print_Center_Center("55", Get_Zoll_e_Width(879), Get_Zoll_e_Height(591), Get_Zoll_e_Width(15), Get_Zoll_e_Height(14), gr, False, 6)


        Print_Klasse_Polygon("F", New System.Drawing.Point(Get_Zoll_e_Width(320), Get_Zoll_e_Height(723)), New System.Drawing.Point(Get_Zoll_e_Width(439), Get_Zoll_e_Height(723)), _
                                    New System.Drawing.Point(Get_Zoll_e_Width(439), Get_Zoll_e_Height(883)), New System.Drawing.Point(Get_Zoll_e_Width(320), Get_Zoll_e_Height(883)), _
                                    New System.Drawing.Point(Get_Zoll_e_Width(360), Get_Zoll_e_Height(803)), gr)
        Print_Left_Center_Norm_6("0", Get_Zoll_e_Width(360), Get_Zoll_e_Height(858), Get_Zoll_e_Height(12), gr)
        '' ''Print_Center_Center("0", Get_Zoll_e_Width(360), Get_Zoll_e_Height(997), Get_Zoll_e_Width(8), Get_Zoll_e_Height(12), gr, False, 6)
        Print_Klasse("E", Get_Zoll_e_Width(448), Get_Zoll_e_Height(723), Get_Zoll_e_Width(80), Get_Zoll_e_Height(161), gr, False, 17)
        Print_Left_Center_Norm_6("30", Get_Zoll_e_Width(451), Get_Zoll_e_Height(858), Get_Zoll_e_Height(12), gr)
        '' ''Print_Center_Center("30", Get_Zoll_e_Width(448), Get_Zoll_e_Height(997), Get_Zoll_e_Width(15), Get_Zoll_e_Height(12), gr, False, 6)
        Print_Klasse("D", Get_Zoll_e_Width(536), Get_Zoll_e_Height(723), Get_Zoll_e_Width(79), Get_Zoll_e_Height(161), gr, False, 17)
        Print_Left_Center_Norm_6("80", Get_Zoll_e_Width(539), Get_Zoll_e_Height(858), Get_Zoll_e_Height(12), gr)
        '' ''        Print_Center_Center("80", Get_Zoll_e_Width(535), Get_Zoll_e_Height(997), Get_Zoll_e_Width(15), Get_Zoll_e_Height(12), gr, False, 6)
        Print_Klasse("C", Get_Zoll_e_Width(622), Get_Zoll_e_Height(723), Get_Zoll_e_Width(78), Get_Zoll_e_Height(161), gr, False, 17)
        Print_Left_Center_Norm_6("145", Get_Zoll_e_Width(625), Get_Zoll_e_Height(858), Get_Zoll_e_Height(12), gr)
        '' ''Print_Center_Center("145", Get_Zoll_e_Width(623), Get_Zoll_e_Height(997), Get_Zoll_e_Width(21), Get_Zoll_e_Height(12), gr, False, 6)
        Print_Klasse("B", Get_Zoll_e_Width(708), Get_Zoll_e_Height(723), Get_Zoll_e_Width(78), Get_Zoll_e_Height(161), gr, False, 17)
        Print_Left_Center_Norm_6("210", Get_Zoll_e_Width(711), Get_Zoll_e_Height(858), Get_Zoll_e_Height(12), gr)
        '' ''        Print_Center_Center("210", Get_Zoll_e_Width(708), Get_Zoll_e_Height(997), Get_Zoll_e_Width(22), Get_Zoll_e_Height(12), gr, False, 6)
        Print_Klasse("A", Get_Zoll_e_Width(794), Get_Zoll_e_Height(723), Get_Zoll_e_Width(79), Get_Zoll_e_Height(161), gr, False, 17)
        Print_Left_Center_Norm_6("270", Get_Zoll_e_Width(797), Get_Zoll_e_Height(858), Get_Zoll_e_Height(12), gr)
        ''Print_Center_Center("270", Get_Zoll_e_Width(794), Get_Zoll_e_Height(997), Get_Zoll_e_Width(22), Get_Zoll_e_Height(12), gr, False, 6)
        Print_Klasse_Polygon("A*", New System.Drawing.Point(Get_Zoll_e_Width(880), Get_Zoll_e_Height(723)), New System.Drawing.Point(Get_Zoll_e_Width(958), Get_Zoll_e_Height(723)), _
                                    New System.Drawing.Point(Get_Zoll_e_Width(999), Get_Zoll_e_Height(803)), New System.Drawing.Point(Get_Zoll_e_Width(958), Get_Zoll_e_Height(883)), _
                                    New System.Drawing.Point(Get_Zoll_e_Width(880), Get_Zoll_e_Height(883)), gr)
        Print_Left_Center_Norm_6("340", Get_Zoll_e_Width(883), Get_Zoll_e_Height(858), Get_Zoll_e_Height(12), gr)
        '' ''Print_Center_Center("340", Get_Zoll_e_Width(880), Get_Zoll_e_Height(997), Get_Zoll_e_Width(22), Get_Zoll_e_Height(12), gr, False, 6)

        'Aussteller
        With My.Settings 'Aussteller.Adresse
            '132; 1484  132; 1519   132; 1555   132; 1591
            Print_Left_Center_Norm_6(Trim(.Aussteller_Name), Get_Zoll_e_Width(132), Get_Zoll_e_Height(1444), Get_Zoll_e_Height(20), gr)
            Print_Left_Center_Norm_6(Trim(.Aussteller_Zusatz), Get_Zoll_e_Width(132), Get_Zoll_e_Height(1479), Get_Zoll_e_Height(20), gr)
            Print_Left_Center_Norm_6(Trim(.Aussteller_Strasse) & " " & Trim(.Aussteller_Nr), Get_Zoll_e_Width(132), Get_Zoll_e_Height(1515), Get_Zoll_e_Height(20), gr)
            Dim tmpPlz As String = Trim(.Aussteller_PLZ) & " "
            'If tmpPlz = "0 " Then tmpPlz = ""
            Print_Left_Center_Norm_6(tmpPlz & Trim(.Aussteller_Ort), Get_Zoll_e_Width(132), Get_Zoll_e_Height(1551), Get_Zoll_e_Height(20), gr)

            '    Logo_einlesen() '.PB_Logo.Image = New Bitmap(exedir & "\Logo.bmp")   ' Me.PB_Logo.Image
        End With
        'Pkte der Rubriken übertragen
        Dim Get_Calc_Pkte_I As String = Me.Calc_Pkte_I
        Dim Get_Calc_Pkte_II As String = Me.Calc_Pkte_II
        Print_Center_Center(Get_Calc_Pkte_I, Get_Zoll_e_Width(7), Get_Zoll_e_Height(373), Get_Zoll_e_Width(273), Get_Zoll_e_Height(56), gr, True, 24)
        Print_Center_Center(Get_Calc_Pkte_II, Get_Zoll_e_Width(7), Get_Zoll_e_Height(632), Get_Zoll_e_Width(273), Get_Zoll_e_Height(152), gr, True, 24)

        'Klassen der Rubriken übertragen
        Print_Klasse(Get_Klasse_I(), Get_Zoll_e_Width(1022), Get_Zoll_e_Height(373), Get_Zoll_e_Width(85), Get_Zoll_e_Height(80), gr, True, 24)
        'Me.Label_e_Klasse_I.BackColor = Get_Klasse_Color(Me.Label_e_Klasse_I.Text)
        Print_Klasse(Get_Klasse_II(), Get_Zoll_e_Width(1022), Get_Zoll_e_Height(613), Get_Zoll_e_Width(85), Get_Zoll_e_Height(263), gr, True, 24)
        'Me.Label_e_Klasse_II.BackColor = Get_Klasse_Color(Me.Label_e_Klasse_II.Text)
        'Pfeile positionieren
        'gr.DrawImage(My.Resources.ssa_pfeil, Get_Klasse_PfeilRect_e(Get_Klasse_I, 374, False))
        gr.DrawImage(My.Resources.ssa_pfeil, Get_Klasse_PfeilRect_e(Get_Klasse_II, 639, True))
        'Mindestpktzahl angeben
        Print_Center_Center(Get_txt_mind_I(Get_Klasse_I), Get_Zoll_e_Width(7), Get_Zoll_e_Height(429), Get_Zoll_e_Width(273), Get_Zoll_e_Height(23), gr, False, 6)
        Print_Center_Center(Get_txt_mind_II(Get_Klasse_II), Get_Zoll_e_Width(7), Get_Zoll_e_Height(821), Get_Zoll_e_Width(273), Get_Zoll_e_Height(73), gr, False, 6)
        'Bonus angeben
        Print_Center_Center(Get_txt_Bonus, Get_Zoll_e_Width(7), Get_Zoll_e_Height(784), Get_Zoll_e_Width(273), Get_Zoll_e_Height(37), gr, False, 6)

        Print_Center_Center(Get_Mes_ja(), Get_Zoll_e_Width(986), Get_Zoll_e_Height(543), Get_Zoll_e_Width(33), Get_Zoll_e_Height(18), gr, False, 6)
        Print_Center_Center(Get_Mes_nein(), Get_Zoll_e_Width(986), Get_Zoll_e_Height(565), Get_Zoll_e_Width(33), Get_Zoll_e_Height(18), gr, False, 6)
        Print_Center_Center(Get_gesKl_ja(), Get_Zoll_e_Width(986), Get_Zoll_e_Height(587), Get_Zoll_e_Width(33), Get_Zoll_e_Height(18), gr, False, 6)
        Print_Center_Center(Get_gesKl_nein(), Get_Zoll_e_Width(986), Get_Zoll_e_Height(610), Get_Zoll_e_Width(33), Get_Zoll_e_Height(18), gr, False, 6)
        'Beschreibung der Klasse
        Print_Block_Norm_5(Get_Klasse_I_Beschreibung(), Get_Zoll_e_Width(284), Get_Zoll_e_Height(371), Get_Zoll_e_Width(735), Get_Zoll_e_Height(81), gr)
        Print_Block_Norm_5(Get_Klasse_II_Beschreibung(), Get_Zoll_e_Width(284), Get_Zoll_e_Height(1015), Get_Zoll_e_Width(822), Get_Zoll_e_Height(97), gr)
        'Gebäudeinfos
        With Projekt.Gebaeude
            Print_Left_Center_Norm_6(.Gebaeudetyp, Get_Zoll_e_Width(815), Get_Zoll_e_Height(1150), Get_Zoll_e_Height(26), gr)
            Dim tmpBaujahr As String = Trim(.Baujahr.ToString)
            If tmpBaujahr = "1800" Then tmpBaujahr = ""
            Print_Left_Center_Norm_6(tmpBaujahr, Get_Zoll_e_Width(815), Get_Zoll_e_Height(1186), Get_Zoll_e_Height(26), gr)
            Print_Left_Center_Norm_6(.Wohneinheiten, Get_Zoll_e_Width(815), Get_Zoll_e_Height(1222), Get_Zoll_e_Height(26), gr)
        End With
        With Projekt.Wohnung
            Print_Left_Center_Norm_6(.Wohnungsbezeichnung, Get_Zoll_e_Width(815), Get_Zoll_e_Height(1258), Get_Zoll_e_Height(26), gr)
            Print_Left_Center_Norm_6(Get_Str_Geschoss(), Get_Zoll_e_Width(815), Get_Zoll_e_Height(1294), Get_Zoll_e_Height(26), gr)
            Print_Left_Center_Norm_6(.Raeume, Get_Zoll_e_Width(815), Get_Zoll_e_Height(1330), Get_Zoll_e_Height(26), gr)
            Dim tmpWohnflaeche As String = Trim(.Wohnflaeche.ToString)
            If tmpWohnflaeche = "0" Then tmpWohnflaeche = ""
            Print_Left_Center_Norm_6(tmpWohnflaeche, Get_Zoll_e_Width(815), Get_Zoll_e_Height(1366), Get_Zoll_e_Height(26), gr)
        End With
        With Projekt
            If Not IsNothing(.Datum) Then
                If .Datum <> "" Then
                    Print_Left_Center_Norm_6(.Datum,
                            Get_Zoll_e_Width(128), Get_Zoll_e_Height(1591), Get_Zoll_e_Height(16), gr)
                    'Print_Center_Center(CType(.Datum, Date).Day & "." & CType(.Datum, Date).Month & "." & CType(.Datum, Date).Year, _
                    '        Get_Zoll_e_Width(13), Get_Zoll_e_Height(1665), Get_Zoll_e_Width(74), Get_Zoll_e_Height(16), gr, False, 6)
                    Dim tDatum As Date = CType(.Datum, Date)
                    Print_Left_Center_Norm_6(IIf(tDatum.Day < 10, "0" & tDatum.Day, tDatum.Day).ToString & "." & IIf(tDatum.Month < 10, "0" & tDatum.Month, tDatum.Month).ToString & "." & tDatum.AddYears(10).Year, _
                            Get_Zoll_e_Width(686), Get_Zoll_e_Height(1591), Get_Zoll_e_Height(16), gr)
                End If
            End If
        End With
        Dim GSize As Size
        Dim tmpscale As Single
        Dim sigRect As System.Drawing.Rectangle
        If Not IsNothing(GetImageFromString(My.Settings.Signatur)) Then

            Dim newBtmSig As New Bitmap(GetImageFromString(My.Settings.Signatur))
            GSize = New Size(newBtmSig.Width, newBtmSig.Height)
            '858; 1573  217; 106
            '858; 1524
            If GSize.Width / 217 > GSize.Height / 106 Then
                'Anpassung an Breite
                tmpscale = CSng(232 / GSize.Width)
                sigRect = New System.Drawing.Rectangle(xPage + Get_Zoll_e_Width(858), _
                        yPage + Get_Zoll_e_Height(1484) + CInt((Get_Zoll_e_Height(106) - Get_Zoll_e_Width(CInt(GSize.Height * tmpscale))) / 2), _
                        Get_Zoll_e_Width(217), Get_Zoll_e_Width(CInt(GSize.Height * tmpscale)))
            Else 'Anpassung asn Höhe
                tmpscale = CSng(106 / GSize.Height)
                sigRect = New System.Drawing.Rectangle(xPage + Get_Zoll_e_Width(858) + CInt((Get_Zoll_e_Width(217) - Get_Zoll_e_Height(CInt(GSize.Width * tmpscale))) / 2), _
                        yPage + Get_Zoll_e_Height(1484), Get_Zoll_e_Height(CInt(GSize.Width * tmpscale)), Get_Zoll_e_Height(106))
            End If
            gr.DrawImage(newBtmSig, sigRect)
        End If
        '5; 1730
        If btDemo > versNormal Then
            Print_Left_Center_Norm_6("Demoversion", Get_Zoll_e_Width(5), Get_Zoll_e_Height(1690), Get_Zoll_e_Height(19), gr)
        Else
            Print_Left_Center_Norm_6(LizenznehmerPR, Get_Zoll_e_Width(5), Get_Zoll_e_Height(1690), Get_Zoll_e_Height(19), gr)
        End If
        If btDemo > versNormal Then Print_Demoversion_quer(gr, 150, -100, 120)
    End Sub
    Private Function Get_Klasse_PfeilRect_e(ByVal Klasse As String, ByVal yKoord As Integer, ByVal bPfeilII As Boolean) As System.Drawing.Rectangle
        Dim tmpPoint As System.Drawing.Rectangle
        tmpPoint.X = 10
        tmpPoint.Y = yKoord
        Select Case Klasse
            Case "A*"
                tmpPoint.X = 887
            Case "A"
                tmpPoint.X = 797
            Case "B"
                tmpPoint.X = 711
            Case "C"
                tmpPoint.X = 625
            Case "D"
                tmpPoint.X = 538
            Case "E"
                tmpPoint.X = 451
            Case "F"
                tmpPoint.X = 363
        End Select
        tmpPoint.X = xPage + Get_Zoll_e_Width(tmpPoint.X)
        tmpPoint.Y = yPage + Get_Zoll_e_Height(tmpPoint.Y)
        tmpPoint.Width = Get_Zoll_e_Width(70)
        If bPfeilII Then
            tmpPoint.Height = Get_Zoll_e_Height(70)
        Else
            tmpPoint.Height = Get_Zoll_e_Height(74)
        End If

        Return tmpPoint
    End Function
    Private Sub Print_Klasse(ByVal Klasse As String, ByVal xPos As Integer, ByVal yPos As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal gr As Graphics, ByVal Fett As Boolean, ByVal Groesse As Integer)
        Print_Klasse_Color(Klasse, xPos, yPos, Width, Height, gr)
        Print_Center_Center(Klasse, xPos, yPos, Width, Height, gr, Fett, Groesse)
    End Sub
    Private Sub Print_Klasse_Polygon(ByVal Klasse As String, ByVal Point1 As System.Drawing.Point, ByVal Point2 As System.Drawing.Point, ByVal Point3 As System.Drawing.Point, ByVal Point4 As System.Drawing.Point, ByVal Point5 As System.Drawing.Point, ByVal gr As Graphics)
        Dim tmpBrsh As SolidBrush = New SolidBrush(Get_Klasse_Color(Klasse))
        'Dim tmpBrsh As Brush
        'Select Case Klasse
        '    Case "A*"
        '        tmpBrsh = Brushes.DarkGreen
        '    Case "F"
        '        tmpBrsh = Brushes.Red
        '    Case Else
        '        tmpBrsh = Brushes.Transparent
        'End Select
        Dim tmpPoints(4) As System.Drawing.Point
        tmpPoints(0).X = xPage + Point1.X
        tmpPoints(0).Y = yPage + Point1.Y
        tmpPoints(1).X = xPage + Point2.X
        tmpPoints(1).Y = yPage + Point2.Y
        tmpPoints(2).X = xPage + Point3.X
        tmpPoints(2).Y = yPage + Point3.Y
        tmpPoints(3).X = xPage + Point4.X
        tmpPoints(3).Y = yPage + Point4.Y
        tmpPoints(4).X = xPage + Point5.X
        tmpPoints(4).Y = yPage + Point5.Y
        gr.FillPolygon(tmpBrsh, tmpPoints)

        Print_Center_Center(Klasse, Point1.X, Point1.Y, Point2.X - Point1.X, Point4.Y - Point1.Y, gr, False, 17)
    End Sub

    Private Sub Print_Klasse_Color(ByVal Klasse As String, ByVal xPos As Integer, ByVal yPos As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal gr As Graphics)

        Dim solBrsh As SolidBrush
        solBrsh = New SolidBrush(Get_Klasse_Color(Klasse))
        gr.FillRectangle(solBrsh, xPage + xPos, yPage + yPos, Width, Height)

        'Dim tmpBrsh As Brush
        'Select Case Klasse
        '    Case "A*"
        '        tmpBrsh = Brushes.DarkGreen
        '    Case "A"
        '        tmpBrsh = Brushes.SeaGreen
        '    Case "B"
        '        tmpBrsh = Brushes.YellowGreen
        '    Case "C"
        '        tmpBrsh = Brushes.Yellow
        '    Case "D"
        '        tmpBrsh = Brushes.DarkOrange
        '    Case "E"
        '        tmpBrsh = Brushes.OrangeRed
        '    Case "F"
        '        tmpBrsh = Brushes.Red
        '    Case Else
        '        tmpBrsh = Brushes.Transparent 'White
        'End Select
        'gr.FillRectangle(tmpBrsh, xPage + xPos, yPage + yPos, Width, Height)
    End Sub

    Private Sub Print_Header_TS_D(ByVal Text As String, ByVal xPos As Integer, ByVal yPos As Integer, ByVal xWidth As Integer, ByVal yHeight As Integer, ByVal gr As Graphics, ByVal Fett As Boolean, ByVal Groesse As Integer)
        Dim Arial As System.Drawing.Font

        If Fett Then
            Arial = New System.Drawing.Font("Arial", Groesse, FontStyle.Bold)
        Else
            Arial = New System.Drawing.Font("Arial", Groesse)
        End If

        Dim sf_Center As New StringFormat
        sf_Center.Alignment = StringAlignment.Center
        sf_Center.LineAlignment = StringAlignment.Center

        Dim tmpDrawRect As New System.Drawing.Rectangle(xPage + xPos, yPage + yPos, xWidth, yHeight)
        gr.FillRectangle(Brushes.White, tmpDrawRect)

        gr.DrawString(Text, Arial, Brushes.Black, tmpDrawRect, sf_Center)

        ''gr.DrawString(Text, Arial, Brushes.Black, CSng(xPage + xPos + xWidth / 2), _
        ''            CSng(yPage + yPos + yHeight / 2), sf_Center)

    End Sub

    Private Sub Print_Center_Center(ByVal Text As String, ByVal xPos As Integer, ByVal yPos As Integer, ByVal xWidth As Integer, ByVal yHeight As Integer, ByVal gr As Graphics, ByVal Fett As Boolean, ByVal Groesse As Integer)
        Dim Arial As System.Drawing.Font

        If Fett Then
            Arial = New System.Drawing.Font("Arial", Groesse, FontStyle.Bold)
        Else
            Arial = New System.Drawing.Font("Arial", Groesse)
        End If

        Dim sf_Center As New StringFormat
        sf_Center.Alignment = StringAlignment.Center
        sf_Center.LineAlignment = StringAlignment.Center

        Dim tmpDrawRect As New System.Drawing.Rectangle(xPage + xPos, yPage + yPos, xWidth, yHeight)

        gr.DrawString(Text, Arial, Brushes.Black, tmpDrawRect, sf_Center)

        ''gr.DrawString(Text, Arial, Brushes.Black, CSng(xPage + xPos + xWidth / 2), _
        ''            CSng(yPage + yPos + yHeight / 2), sf_Center)

    End Sub
    Private Sub Print_Left_Center(ByVal Text As String, ByVal xPos As Integer, ByVal yPos As Integer, ByVal xWidth As Integer, ByVal yHeight As Integer, ByVal gr As Graphics, ByVal Fett As Boolean, ByVal Groesse As Integer)
        Dim Arial As System.Drawing.Font

        If Fett Then
            Arial = New System.Drawing.Font("Arial", Groesse, FontStyle.Bold)
        Else
            Arial = New System.Drawing.Font("Arial", Groesse)
        End If

        Dim sf_Center As New StringFormat
        sf_Center.Alignment = StringAlignment.Center
        sf_Center.LineAlignment = StringAlignment.Near

        Dim tmpDrawRect As New System.Drawing.Rectangle(xPage + xPos, yPage + yPos, xWidth, yHeight)

        gr.DrawString(Text, Arial, Brushes.Black, tmpDrawRect, sf_Center)


    End Sub

    Private Sub Print_Left_Center_Norm_6(ByVal Text As String, ByVal xPos As Integer, ByVal yPos As Integer, ByVal yHeight As Integer, ByVal gr As Graphics)
        Dim Arial_Norm_6 As New System.Drawing.Font("Arial", 6)

        Dim sf_Center As New StringFormat
        sf_Center.Alignment = StringAlignment.Near
        sf_Center.LineAlignment = StringAlignment.Center

        ' If Not String.IsNullOrEmpty(Text) Then gr.DrawString(Text, Arial_Norm_6, Brushes.Black, xPage + xPos, CSng(yPage + yPos + yHeight / 2), sf_Center)
        If Text <> "" And Text <> LEER100 And Text <> LEER10 And Text <> LEER5 And Text <> LEER2 And Text <> LEER55 And Text <> LEER512 Then gr.DrawString(Text, Arial_Norm_6, Brushes.Black, xPage + xPos, CSng(yPage + yPos + yHeight / 2), sf_Center)
    End Sub

    Private Sub Print_Block_Norm_5(ByVal Text As String, ByVal xPos As Integer, ByVal yPos As Integer, ByVal xWidth As Integer, ByVal yHeight As Integer, ByVal gr As Graphics)
        Dim Arial_Norm_5 As New System.Drawing.Font("Arial", 5)
        Dim tmpRect As New System.Drawing.Rectangle(xPage + xPos, yPage + yPos, xWidth, yHeight)

        Dim sf_Block As New StringFormat
        sf_Block.Alignment = StringAlignment.Center 'Near
        sf_Block.LineAlignment = StringAlignment.Center

        gr.DrawString(Text, Arial_Norm_5, Brushes.Black, tmpRect, sf_Block)

    End Sub

#Region "Panel_SSA_*_Click"
    Private Sub Panel_SSA_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_SSA_1.Click
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
        Form_Help.Close()
    End Sub
    Private Sub Panel_SSA_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_SSA_2.Click
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
        Form_Help.Close()
    End Sub
    Private Sub Panel_SSA_3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_SSA_3.Click
        Form_Help.Close()
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
    End Sub
    Private Sub Panel_SSA_LS_W_Messung_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_SSA_LS_W_Messung.Click
        Form_Help.Close()
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
    End Sub
    Private Sub Panel_SSA_LS_W_Prognose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_SSA_LS_W_Prognose.Click
        Form_Help.Close()
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
    End Sub
    Private Sub Panel_SSA_LS_D_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_SSA_LS_D.Click
        Form_Help.Close()
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
    End Sub
    Private Sub Panel_SSA_LS_D_Messung_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_SSA_LS_D_Messung.Click
        Form_Help.Close()
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
    End Sub
    Private Sub Panel_SSA_LS_D_Prognose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_SSA_LS_D_Prognose.Click
        Form_Help.Close()
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
    End Sub
    Private Sub Panel_SSA_TS_D_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_SSA_TS_D.Click
        Form_Help.Close()
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
    End Sub
    Private Sub Panel_SSA_TS_D_Messung_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_SSA_TS_D_Messung.Click
        Form_Help.Close()
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
    End Sub
    Private Sub Panel_SSA_TS_D_Prognose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_SSA_TS_D_Prognose.Click
        Form_Help.Close()
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
    End Sub
    Private Sub Panel_SSA_TS_TPHf_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_SSA_TS_TPHf.Click
        Form_Help.Close()
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
    End Sub
    Private Sub Panel_SSA_TS_TPHf_Messung_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_SSA_TS_TPHf_Messung.Click
        Form_Help.Close()
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
    End Sub
    Private Sub Panel_SSA_TS_TPHf_Prognose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_SSA_TS_TPLH_Prognose.Click
        Form_Help.Close()
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
    End Sub
    Private Sub Panel_SSA_TS_BLLT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_SSA_TS_BLLT.Click

        Form_Help.Close()
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
    End Sub
    Private Sub Panel_SSA_TS_BLLT_Messung_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_SSA_TS_BLLT_Messung.Click

        Form_Help.Close()
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
    End Sub
    Private Sub Panel_SSA_TS_BLLT_Prognose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_SSA_TS_BLT_Prognose.Click

        Form_Help.Close()
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
    End Sub
    Private Sub Panel_SSA_Tuere_Untersuchung_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_SSA_Tuere_Untersuchung.Click

        Form_Help.Close()
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
    End Sub
    Private Sub Panel_SSA_Tueren_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_SSA_Tueren.Click

        Form_Help.Close()
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
    End Sub
    Private Sub Panel_SSA_Nutzer_Pegel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_SSA_Nutzer_Pegel.Click

        Form_Help.Close()
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
    End Sub
    Private Sub Panel_SSA_Koerper_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_SSA_NutzerKoerper_ProgMes.Click

        Form_Help.Close()
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
    End Sub
    Private Sub Panel_SSA_Koerper_Pegel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_SSA_Koerper_Pegel.Click

        Form_Help.Close()
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
    End Sub
#End Region

#Region "Help"
    Private Sub Show_Help(ByVal Pic As String, ByVal ctrl As Control)
        Dim obj As System.Drawing.Bitmap = CType(My.Resources.ResourceManager.GetObject(Pic), System.Drawing.Bitmap)
        Dim tmpSize As New Size(obj.Width + 40, obj.Height + 50)

        Dim y As Integer = ctrl.Location.Y
        Dim pCtrl As Control = ctrl.Parent
        Do While pCtrl.Name <> "TabControl1"
            y = y + pCtrl.Location.Y
            pCtrl = pCtrl.Parent
        Loop

        Dim loc As New System.Drawing.Point(Me.Location.X + 3, y + Me.TabControl1.Location.Y + Me.Location.Y + 20 + ctrl.Height + 3)

        Form_Help.Size = tmpSize


        '        Me.Panel_Help.Size = tmpSize
        '        Me.Panel_Help.Location = loc

        Form_Help.PB.Size = obj.Size 'tmpSize
        Form_Help.PB.Image = obj

        Form_Help.PB_2.Visible = False
        Form_Help.PB_3.Visible = False

        Form_Help.Show()
        Form_Help.BringToFront()
        Form_Help.Location = loc
        'Me.Panel_Help.Show()

    End Sub
    'Private Sub Show_Help(ByVal Pic As String, ByVal ctrl As Control)
    '    Dim obj As System.Drawing.Bitmap = CType(My.Resources.ResourceManager.GetObject(Pic), System.Drawing.Bitmap)
    '    Dim tmpSize As New Size(obj.Width + 12, obj.Height + 12)

    '    Dim y As Integer = ctrl.Location.Y
    '    Dim pCtrl As Control = ctrl.Parent
    '    Do While pCtrl.Name <> "TabControl1"
    '        y = y + pCtrl.Location.Y
    '        pCtrl = pCtrl.Parent
    '    Loop

    '    Dim loc As New System.Drawing.Point(3, y + Me.TabControl1.Location.Y + ctrl.Height + 3)

    '    Me.Panel_Help.Size = tmpSize
    '    Me.Panel_Help.Location = loc

    '    Me.PB.Size = obj.Size 'tmpSize
    '    Me.PB.Image = obj

    '    Me.PB_2.Visible = False
    '    Me.PB_3.Visible = False

    '    Me.Panel_Help.Show()

    'End Sub
    Private Sub Show_Help_2(ByVal Pic_1 As String, ByVal Pic_2 As String, ByVal ctrl As Control)
        Dim obj_1 As System.Drawing.Bitmap = CType(My.Resources.ResourceManager.GetObject(Pic_1), System.Drawing.Bitmap)
        Dim obj_2 As System.Drawing.Bitmap = CType(My.Resources.ResourceManager.GetObject(Pic_2), System.Drawing.Bitmap)

        Dim y As Integer = ctrl.Location.Y
        Dim pCtrl As Control = ctrl.Parent
        Do While pCtrl.Name <> "TabControl1"
            y = y + pCtrl.Location.Y
            pCtrl = pCtrl.Parent
        Loop

        Dim loc As New System.Drawing.Point(Me.Location.X + 3, y + Me.TabControl1.Location.Y + Me.Location.Y + 20 + ctrl.Height + 3)

        Dim tmpSize As Size
        Dim obj_width As Integer
        If obj_1.Width > obj_2.Width Then
            tmpSize = New Size(obj_1.Width + 40, obj_1.Height + obj_2.Height + 50)
            obj_width = obj_1.Width
        Else
            tmpSize = New Size(obj_2.Width + 40, obj_1.Height + obj_2.Height + 50)
            obj_width = obj_2.Width
        End If
        Form_Help.Size = tmpSize

        'Me.Panel_Help.Size = tmpSize

        'Me.PB.Size = obj_1.Size 'tmpSize
        Form_Help.PB.Size = New Size(obj_width, obj_1.Height)
        Form_Help.PB.Image = obj_1

        'Me.PB_2.Size = obj_2.Size
        Form_Help.PB_2.Size = New Size(obj_width, obj_2.Height)
        Form_Help.PB_2.Image = obj_2
        Form_Help.PB_2.Location = New System.Drawing.Point(3, 5 + obj_1.Height)
        Form_Help.PB_2.Visible = True

        Form_Help.PB_3.Visible = False

        'me.Panel_Help.Location = loc
        'Me.Panel_Help.Show()
        Form_Help.Location = loc
        Form_Help.Show()
        Form_Help.BringToFront()

    End Sub
    'Private Sub Show_Help_2(ByVal Pic_1 As String, ByVal Pic_2 As String, ByVal ctrl As Control)
    '    Dim obj_1 As System.Drawing.Bitmap = CType(My.Resources.ResourceManager.GetObject(Pic_1), System.Drawing.Bitmap)
    '    Dim obj_2 As System.Drawing.Bitmap = CType(My.Resources.ResourceManager.GetObject(Pic_2), System.Drawing.Bitmap)

    '    Dim y As Integer = ctrl.Location.Y
    '    Dim pCtrl As Control = ctrl.Parent
    '    Do While pCtrl.Name <> "TabControl1"
    '        y = y + pCtrl.Location.Y
    '        pCtrl = pCtrl.Parent
    '    Loop

    '    Dim loc As New System.Drawing.Point(3, y + Me.TabControl1.Location.Y + ctrl.Height + 3)
    '    Dim tmpSize As Size
    '    Dim obj_width As Integer
    '    If obj_1.Width > obj_2.Width Then
    '        tmpSize = New Size(obj_1.Width + 12, obj_1.Height + obj_2.Height + 14)
    '        obj_width = obj_1.Width
    '    Else
    '        tmpSize = New Size(obj_2.Width + 12, obj_1.Height + obj_2.Height + 14)
    '        obj_width = obj_2.Width
    '    End If
    '    Me.Panel_Help.Size = tmpSize

    '    'Me.PB.Size = obj_1.Size 'tmpSize
    '    Me.PB.Size = New Size(obj_width, obj_1.Height)
    '    Me.PB.Image = obj_1

    '    'Me.PB_2.Size = obj_2.Size
    '    Me.PB_2.Size = New Size(obj_width, obj_2.Height)
    '    Me.PB_2.Image = obj_2
    '    Me.PB_2.Location = New System.Drawing.Point(3, 5 + obj_1.Height)
    '    Me.PB_2.Visible = True

    '    Me.PB_3.Visible = False

    '    Me.Panel_Help.Location = loc
    '    Me.Panel_Help.Show()

    'End Sub
    Private Sub Show_Help_3(ByVal Pic_1 As String, ByVal Pic_2 As String, ByVal Pic_3 As String, ByVal ctrl As Control)
        Dim obj_1 As System.Drawing.Bitmap = CType(My.Resources.ResourceManager.GetObject(Pic_1), System.Drawing.Bitmap)
        Dim obj_2 As System.Drawing.Bitmap = CType(My.Resources.ResourceManager.GetObject(Pic_2), System.Drawing.Bitmap)
        Dim obj_3 As System.Drawing.Bitmap = CType(My.Resources.ResourceManager.GetObject(Pic_3), System.Drawing.Bitmap)

        Dim y As Integer = ctrl.Location.Y
        Dim pCtrl As Control = ctrl.Parent
        Do While pCtrl.Name <> "TabControl1"
            y = y + pCtrl.Location.Y
            pCtrl = pCtrl.Parent
        Loop

        Dim loc As New System.Drawing.Point(Me.Location.X + 3, y + Me.TabControl1.Location.Y + Me.Location.Y + 20 + ctrl.Height + 3)

        Dim tmpSize As Size
        Dim obj_width As Integer
        If obj_1.Width >= obj_2.Width And obj_1.Width >= obj_3.Width Then
            tmpSize = New Size(obj_1.Width + 40, obj_1.Height + obj_2.Height + obj_3.Height + 50)
            obj_width = obj_1.Width
        ElseIf obj_2.Width >= obj_1.Width And obj_2.Width >= obj_3.Width Then
            tmpSize = New Size(obj_2.Width + 40, obj_1.Height + obj_2.Height + obj_3.Height + 50)
            obj_width = obj_2.Width
        Else
            tmpSize = New Size(obj_3.Width + 40, obj_1.Height + obj_2.Height + obj_3.Height + 50)
            obj_width = obj_3.Width
        End If
        Form_Help.Size = tmpSize
        'Me.Panel_Help.Size = tmpSize

        Form_Help.PB.Size = New Size(obj_width, obj_1.Height) 'tmpSize
        Form_Help.PB.Image = obj_1

        Form_Help.PB_2.Size = New Size(obj_width, obj_2.Height)
        Form_Help.PB_2.Image = obj_2
        Form_Help.PB_2.Location = New System.Drawing.Point(3, 5 + obj_1.Height)
        Form_Help.PB_2.Visible = True

        Form_Help.PB_3.Size = New Size(obj_width, obj_3.Height)
        Form_Help.PB_3.Image = obj_3
        Form_Help.PB_3.Location = New System.Drawing.Point(3, 7 + obj_1.Height + obj_2.Height)
        Form_Help.PB_3.Visible = True

        'Me.Panel_Help.Location = loc
        'Me.Panel_Help.Show()
        Form_Help.Location = loc
        Form_Help.BringToFront()
        Form_Help.Show()
    End Sub
    'Private Sub Show_Help_3(ByVal Pic_1 As String, ByVal Pic_2 As String, ByVal Pic_3 As String, ByVal ctrl As Control)
    '    Dim obj_1 As System.Drawing.Bitmap = CType(My.Resources.ResourceManager.GetObject(Pic_1), System.Drawing.Bitmap)
    '    Dim obj_2 As System.Drawing.Bitmap = CType(My.Resources.ResourceManager.GetObject(Pic_2), System.Drawing.Bitmap)
    '    Dim obj_3 As System.Drawing.Bitmap = CType(My.Resources.ResourceManager.GetObject(Pic_3), System.Drawing.Bitmap)

    '    Dim y As Integer = ctrl.Location.Y
    '    Dim pCtrl As Control = ctrl.Parent
    '    Do While pCtrl.Name <> "TabControl1"
    '        y = y + pCtrl.Location.Y
    '        pCtrl = pCtrl.Parent
    '    Loop

    '    Dim loc As New System.Drawing.Point(3, y + Me.TabControl1.Location.Y + ctrl.Height + 3)

    '    Dim tmpSize As Size
    '    Dim obj_width As Integer
    '    If obj_1.Width >= obj_2.Width And obj_1.Width >= obj_3.Width Then
    '        tmpSize = New Size(obj_1.Width + 12, obj_1.Height + obj_2.Height + obj_3.Height + 16)
    '        obj_width = obj_1.Width
    '    ElseIf obj_2.Width >= obj_1.Width And obj_2.Width >= obj_3.Width Then
    '        tmpSize = New Size(obj_2.Width + 12, obj_1.Height + obj_2.Height + obj_3.Height + 16)
    '        obj_width = obj_2.Width
    '    Else
    '        tmpSize = New Size(obj_3.Width + 12, obj_1.Height + obj_2.Height + obj_3.Height + 16)
    '        obj_width = obj_3.Width
    '    End If
    '    Me.Panel_Help.Size = tmpSize

    '    Me.PB.Size = New Size(obj_width, obj_1.Height) 'tmpSize
    '    Me.PB.Image = obj_1

    '    Me.PB_2.Size = New Size(obj_width, obj_2.Height)
    '    Me.PB_2.Image = obj_2
    '    Me.PB_2.Location = New System.Drawing.Point(3, 5 + obj_1.Height)
    '    Me.PB_2.Visible = True

    '    Me.PB_3.Size = New Size(obj_width, obj_3.Height)
    '    Me.PB_3.Image = obj_3
    '    Me.PB_3.Location = New System.Drawing.Point(3, 7 + obj_1.Height + obj_2.Height)
    '    Me.PB_3.Visible = True

    '    Me.Panel_Help.Location = loc
    '    Me.Panel_Help.Show()
    'End Sub

    Private Sub Label_Help_42_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_42", CType(sender, System.Windows.Forms.Label))
    End Sub
    Private Sub Label_Help_42_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub

    Private Sub Label_Help_E_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_E.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help_2("ssa_help_KlasseE_7", "ssa_help_BeschKlE_54", CType(sender, Control))
    End Sub
    Private Sub Label_Help_E_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_E.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub

    Private Sub Label_Help_F_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_F.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help_2("ssa_help_KlasseF_8", "ssa_help_BeschKlF_55", CType(sender, Control))
    End Sub

    Private Sub Label_Help_D_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_D.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help_3("ssa_help_KlasseD_6", "ssa_help_abwKenn_22", "ssa_help_BeschKlD_53", CType(sender, Control))
    End Sub

    Private Sub Label_Help_C_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_C.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help_2("ssa_help_KlasseC_5", "ssa_help_BeschKlC_52", CType(sender, Control))
    End Sub

    Private Sub Label_Help_B_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_B.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help_3("ssa_help_4", "ssa_help_23", "ssa_help_BeschKlB_51", CType(sender, Control))
    End Sub

    Private Sub Label_Help_A_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_A.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help_3("ssa_help_3", "ssa_help_23", "ssa_help_BeschKlA_50", CType(sender, Control))
    End Sub

    Private Sub Label_Help_AA_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_AA.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help_3("ssa_help_2", "ssa_help_23", "ssa_help_BeschKlA__49", CType(sender, Control))

    End Sub

    Private Sub Label_gueltigBis_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_gueltigBis.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_46", CType(sender, Control))
    End Sub

    Private Sub Label_Help_34_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_34.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_ungSituation_34", CType(sender, Control))
    End Sub

    Private Sub Label_Help_35_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_35.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_allgAnw_35", CType(sender, Control))
    End Sub

    Private Sub Label_Help_44_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_44.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_44", CType(sender, Control))
    End Sub

    Private Sub Label_Help_47_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_47.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_47", CType(sender, Control))
    End Sub

    Private Sub Label_33_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_33.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_SA_33", CType(sender, Control))
    End Sub

    Private Sub Label_Help_45_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_45.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_HinweiseKat_45", CType(sender, Control))
    End Sub

    Private Sub Label_Help_25_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_25.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_LS_25", CType(sender, Control))
    End Sub

    Private Sub Label_Help_LS_W_MA_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_Estrich_43", CType(sender, Control))
    End Sub

    Private Sub Label_Help_LS_D_MA_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_LS_D_MA.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_Estrich_43", CType(sender, Control))
    End Sub

    Private Sub Label_Help_TS_D_MA_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_TS_D_MA.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_Estrich_43", CType(sender, Control))
    End Sub

    Private Sub Label_Help_LS_W_AnzM_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_LS_W_AnzM.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_1", CType(sender, Control))
    End Sub

    Private Sub Label_Help_LS_D_AnzM_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_LS_D_AnzM.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_1", CType(sender, Control))
    End Sub

    Private Sub Label_Help_TS_D_AnzM_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_TS_D_AnzM.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_1", CType(sender, Control))
    End Sub

    Private Sub Label_Help_TS_TPH_AnzM_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_TS_TPH_AnzM.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_1", CType(sender, Control))
    End Sub

    Private Sub Label_Help_TS_BLLT_AnzM_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_TS_BLLT_AnzM.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_1", CType(sender, Control))
    End Sub

    Private Sub Label_Help_LS_W_M_R_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_LS_W_M_R.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_14", CType(sender, Control))
    End Sub

    Private Sub Label_Help_LS_W_P_R_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_LS_W_P_R.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_14", CType(sender, Control))
    End Sub

    Private Sub Label_Help_LS_D_M_R_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_LS_D_M_R.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_14", CType(sender, Control))
    End Sub

    Private Sub Label_Help_LS_D_P_R_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_LS_D_P_R.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_14", CType(sender, Control))
    End Sub

    Private Sub Label_Help_TS_D_M_L_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_TS_D_M_L.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_15", CType(sender, Control))
    End Sub

    Private Sub Label_Help_TS_D_P_L_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_TS_D_P_L.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_15", CType(sender, Control))
    End Sub

    Private Sub Label_Help_TS_TPH_M_L_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_TS_TPH_M_L.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_15", CType(sender, Control))
    End Sub

    Private Sub Label_Help_TS_TPH_P_L_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_TS_TPH_P_L.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_15", CType(sender, Control))
    End Sub

    Private Sub Label_Help_TS_BLLT_M_L_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_TS_BLLT_M_L.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_15", CType(sender, Control))
    End Sub

    Private Sub Label_Help_TS_BLLT_P_L_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_TS_BLLT_P_L.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_15", CType(sender, Control))
    End Sub
    Private Sub Label_Help_TS_BLLT_P_L_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_TS_BLLT_P_L.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub

    Private Sub Label_Help_26_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_26.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_TS_26", CType(sender, Control))
    End Sub
    Private Sub Label_Help_26_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_26.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub

    Private Sub Label_Help_LS_Tuer_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_LS_Tuer.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_LS_25", CType(sender, Control))
    End Sub
    Private Sub Label_Help_LS_Tuer_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_LS_Tuer.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub

    Private Sub Label_Help_29_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_29.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_LSAußen_29", CType(sender, Control))
    End Sub
    Private Sub Label_Help_29_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_29.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub

    Private Sub Label_Help_27_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_27.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_Wasser_27", CType(sender, Control))
    End Sub

    Private Sub Label_Help_28_40_48_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_28_40_48.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help_3("ssa_help_NutzerKoerperAnf_40", "ssa_help_NutzerKoerper_28", "ssa_help_Empf_48", CType(sender, Control))
    End Sub

    Private Sub Label_Help_38_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_38.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_38", CType(sender, Control))
    End Sub
    Private Sub Label_Help_38_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_38.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub

    Private Sub Label_Help_30_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_30.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_fremd_30", CType(sender, Control))
    End Sub
    Private Sub Label_Help_30_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_30.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub

    Private Sub Label_Help_31_24_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_31_24.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help_2("ssa_help_erlEW_24", "ssa_help_EW_31", CType(sender, Control))
    End Sub

    Private Sub Label_Help_16_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_16.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_LWasser_16", CType(sender, Control))
    End Sub
    Private Sub Label_Help_16_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_16.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub

    Private Sub Label_Help_Nutzer_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_Nutzer.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_Nutzer_18", CType(sender, Control))
    End Sub
    Private Sub Label_Help_Nutzer_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_Nutzer.MouseLeave
        'If Me.Cursor = Cursors.Help Then Form_Help.Close() 'Me.Panel_Help.Hide()
    End Sub

    Private Sub Label_Help_Koerper_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label_Help_Koerper.MouseHover
        If Me.Cursor = Cursors.Help Then Show_Help("ssa_help_Nutzer_18", CType(sender, Control))
    End Sub
#End Region

    Private Sub TSB_Neu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSB_Neu.Click
        Projekt_neu()
    End Sub

    Private Sub TSB_Drucken_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSB_Drucken.Click
        Drucken()
    End Sub

    Private Sub TSB_Oeffnen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSB_Oeffnen.Click
        Projekt_laden()
    End Sub

    Private Sub TSB_Speichern_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSB_Speichern.Click
        If stProjekt_Pfad <> "" And stProjekt_Name <> "" Then
            Projektdaten_speichern()
        Else
            Projektdaten_speichernUnter()
        End If
    End Sub

    Private Sub TSB_Hilfe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSB_Hilfe.Click
        If Me.Cursor = Cursors.Help Then
            Me.Cursor = Cursors.Default
            Me.TSMI_DirekteHilfe.Checked = False
        Else
            Me.Cursor = Cursors.Help
            Me.TSMI_DirekteHilfe.Checked = True
        End If
    End Sub

    Private Sub TSB_Eingabe_leeren_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSB_Eingabe_leeren.Click
        Projekt_Default()
        Update_Anzeige()
    End Sub

#Region "Button_Tab"
    Private Sub Show_Tab(ByVal Pic As String)
        Dim obj As System.Drawing.Bitmap = CType(My.Resources.ResourceManager.GetObject(Pic), System.Drawing.Bitmap)
        Dim tmpSize As New Size(obj.Width + (Form_Tab.Width - Form_Tab.Panel1.Width) + 5, obj.Height + (Form_Tab.Height - Form_Tab.Panel1.Height) + 5)
        Form_Tab.Size = tmpSize
        Form_Tab.PB.Size = obj.Size 'tmpSize
        Form_Tab.PB.Image = obj

        Form_Tab.Show()
        Form_Tab.WindowState = FormWindowState.Normal
        Form_Tab.BringToFront()
    End Sub
    Private Sub Button_Tab_GC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tab_GC.Click
        Show_Tab("gc")
        'Show_Tab("GC")
    End Sub

    Private Sub Button_Tab_AL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tab_AL.Click
        Show_Tab("al")
        'Show_Tab("AL")
    End Sub

    Private Sub Button_Tab_LS_W_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tab_LS_W.Click
        Show_Tab("ls_w")
        'Show_Tab("LS_W")
    End Sub

    Private Sub Button_Tab_LS_D_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tab_LS_D.Click
        Show_Tab("ls_d")
        'Show_Tab("LS_D")
    End Sub

    Private Sub Button_Tab_TS_D_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tab_TS_D.Click
        Show_Tab("ts_d")
        'Show_Tab("TS_D")
    End Sub

    Private Sub Button_Tab_TS_TPH_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tab_TS_TPH.Click
        'Show_Tab("TS_TPH")
        Show_Tab("ts_tph")
    End Sub

    Private Sub Button_Tab_TS_BLLT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tab_TS_BLLT.Click
        Show_Tab("ts_bllt")
        'Show_Tab("TS_BLLT")
    End Sub

    Private Sub Button_Tab_Tueren_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tab_Tueren.Click
        Show_Tab("tueren")
        'Show_Tab("Tueren")
    End Sub

    Private Sub Button_Tab_LS_AB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tab_LS_AB.Click
        Show_Tab("ls_ab")
        'Show_Tab("LS_AB")
    End Sub

    Private Sub Button_Tab_Wasser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tab_Wasser.Click
        Show_Tab("wasser")
        'Show_Tab("Wasser")
    End Sub

    Private Sub Button_Tab_Nutzer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tab_Nutzer.Click
        Show_Tab("nutzerkoerper")
        'Show_Tab("NutzerKoerper")
    End Sub

    Private Sub Button_Tab_Koerper_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Show_Tab("nutzerkoerper")
        'Show_Tab("NutzerKoerper")
    End Sub

    Private Sub Button_Tab_Nachbarn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tab_Nachbarn.Click
        Show_Tab("raeume")
        'Show_Tab("Raeume")
    End Sub

    Private Sub Button_Tab_anordRaeume_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tab_anordRaeume.Click
        Show_Tab("raeume")
        'Show_Tab("Raeume")
    End Sub

    Private Sub Button_Tab_lauteRaeume_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tab_lauteRaeume.Click
        Show_Tab("raeume")
        'Show_Tab("Raeume")
    End Sub

    Private Sub Button_Tab_NHZ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tab_NHZ.Click
        Show_Tab("nhz")
    End Sub

    Private Sub Button_Tab_EW_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Tab_EW.Click
        Show_Tab("ew")
        'Show_Tab("EW")
    End Sub
#End Region

    'Private Sub TSCB_Zoom_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Select Case Me.TSCB_Zoom.Text
    '        Case "500 %"
    '            Form_Main_Zoom(5, Me.Panel2)
    '        Case "200 %"
    '            Form_Main_Zoom(2, Me.Panel2)
    '        Case "150 %"
    '            Form_Main_Zoom(1.5, Me.Panel2)
    '        Case "75 %"
    '            Form_Main_Zoom(0.75, Me.Panel2)
    '        Case "50 %"
    '            Form_Main_Zoom(0.5, Me.Panel2)
    '        Case "25 %"
    '            Form_Main_Zoom(0.25, Me.Panel2)
    '        Case "10 %"
    '            Form_Main_Zoom(0.1, Me.Panel2)
    '        Case Else '"100 %"
    '            Form_Main_Zoom(1, Me.Panel2)
    '    End Select
    'End Sub
    'Private Sub Form_Main_Zoom(ByVal Scale As Single, ByVal parent As Control)
    '    ' For Each ctrl As Control In parent.Controls
    '    If Not IsNothing(parent.Controls) Then
    '        For Each child As Control In parent.Controls
    '            Form_Main_Zoom(Scale, child)
    '        Next
    '    End If

    '    Dim nSize As Size = New Size(CInt(parent.Size.Width * Scale), CInt(parent.Size.Height * Scale))

    '    parent.Size = New Size(CInt(parent.Size.Width * Scale), CInt(parent.Size.Height * Scale))
    '    parent.Location = New System.Drawing.Point(CInt(parent.Location.X * Scale), CInt(parent.Location.Y * Scale))
    '    ' Next
    'End Sub

    'Private Sub me_paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
    '    Dim gr As Graphics = e.Graphics
    '    gr.ScaleTransform(0.5, 5)
    'End Sub

#Region "Update"
    Public Sub Update_SSA()
        'Aussteller
        With My.Settings 'Aussteller.Adresse
            Me.Label_Aussteller_Name.Text = Trim(.Aussteller_Name)
            Dim tmpPlz As String = Trim(.Aussteller_PLZ) & " "
            'If tmpPlz = "0 " Then tmpPlz = ""
            Me.Label_Aussteller_Ort.Text = tmpPlz & Trim(.Aussteller_Ort)
            Me.Label_Aussteller_Strasse.Text = Trim(.Aussteller_Strasse) & " " & Trim(.Aussteller_Nr)
            Me.Label_Aussteller_Zusatz.Text = Trim(.Aussteller_Zusatz)
        End With
        Update_SSA_Deckblatt()
        Update_SSA_detailiert()
        'Update_SSA_einfach()
        'Update_SSA_Status()
    End Sub
    Public Sub Update_SSA_Status()
        Me.Label_Status_Klasse.Text = Get_Klasse_II()
        Me.Label_Status_Klasse.BackColor = Get_Klasse_Color(Me.Label_Status_Klasse.Text)

        Me.Label_Status_SA_Klasse.Text = Get_Klasse_I()
        Me.Label_Status_SA_Klasse.BackColor = Get_Klasse_Color(Me.Label_Status_SA_Klasse.Text)

        Me.Label_Status_BSS_Klasse.Text = Get_Status_BSS_Text()
    End Sub
    Public Function Get_Status_BSS_Text() As String
        Dim Kl_II As Byte = Get_Byte_Klasse(Get_Klasse_II())

        Dim Kl_LS_W As Byte = Get_Byte_Klasse(Get_Klasse_LS_Waende())
        Dim KL_LS_D As Byte = Get_Byte_Klasse(Get_Klasse_LS_Decken())
        Dim KL_TS_D As Byte = Get_Byte_Klasse(Get_Klasse_TS_Decken())
        Dim KL_TS_TPH As Byte = Get_Byte_Klasse(Get_Klasse_TS_TPLH())
        Dim KL_TS_BLLt As Byte = Get_Byte_Klasse(Get_Klasse_TS_BLT())
        Dim KL_Tueren As Byte = Get_Byte_Klasse(Get_Klasse_Tueren())
        Dim KL_Aussenbauteile As Byte = Get_Byte_Klasse(Get_Klasse_Aussenbauteile())
        Dim KL_Wasser As Byte = Get_Byte_Klasse(Get_Klasse_Wasser())
        Dim KL_lauteRaeume As Byte = Get_Byte_Klasse(Get_Klasse_lauteRaeume())

        If Get_gesKl_ja() = "X" Then
            Dim erPkt As Integer = CInt(Calc_Pkte_II())
            If Kl_II < AA Then
                If Kl_II = A Then
                    Return 340 - erPkt & " Punkte"
                ElseIf Kl_II = B Then
                    Return 270 - erPkt & " Punkte"
                ElseIf Kl_II = C Then
                    Return 210 - erPkt & " Punkte"
                ElseIf Kl_II = D Then
                    Return 145 - erPkt & " Punkte"
                ElseIf Kl_II = E Then
                    Return 80 - erPkt & " Punkte"
                ElseIf Kl_II = F Then
                    Return 30 - erPkt & " Punkte"
                Else
                    Return "Eingabe ist nicht vollständig"
                End If
            Else
                Return "Die beste Klasse ist erreicht!"
            End If

        Else
            If Kl_LS_W < Kl_II Then
                Return "Verbesserung von:   Luftschall Wände"
            ElseIf KL_LS_D < Kl_II Then
                Return "Verbesserung von:   Luftschall Decken"
            ElseIf KL_TS_D < Kl_II Then
                Return "Verbesserung von:   Trittschall Decken"
            ElseIf KL_TS_TPH < Kl_II Then
                Return "Verbesserung von:   Trittschall Treppen, Podeste, Hausflure"
            ElseIf KL_TS_BLLt < Kl_II Then
                Return "Verbesserung von:   Trittschall Balkone, Laubengänge, Loggien, Terrassen"
            ElseIf KL_Tueren < Kl_II Then
                Return "Verbesserung von:   Luftschall Türen"
            ElseIf KL_Aussenbauteile < Kl_II Then
                Return "Verbesserung von:   Luftschall Außenbauteile"
            ElseIf KL_Wasser < Kl_II Then
                Return "Verbesserung von:   Wasserinstallation / Haustechn. Anlagen"
            ElseIf KL_lauteRaeume < Kl_II Then
                Return "Verbesserung von:   laute Räume angrenzend"
            Else
                Return "Eingabe ist nicht vollständig"
            End If
        End If

    End Function
    Public Function Get_Byte_Klasse(ByVal sKlasse As String) As Byte
        Select Case sKlasse
            Case "A*", "-"
                Return AA
            Case "A"
                Return A
            Case "B"
                Return B
            Case "C"
                Return C
            Case "D"
                Return D
            Case "E"
                Return E
            Case "F"
                Return F
            Case Else
                Return 0
        End Select
    End Function
    Public Sub Update_SSA_Deckblatt()


        'Grundriss/Signatur/Logo
        Me.PB_DB_Grundriss.BackgroundImage = btmGrundriss 'Me.PB_Grundriss.backgroundImage
        Me.PB_Geb_Signatur.BackgroundImage = GetImageFromString(My.Settings.Signatur)
        Me.PB_Geb_Logo.BackgroundImage = GetImageFromString(My.Settings.Logo)

        'Aussteller
        With My.Settings 'Aussteller.Adresse
            Me.Label_Geb_AS_Name.Text = Trim(.Aussteller_Name)
            Dim tmpPlz As String = Trim(.Aussteller_PLZ) & " "
            'If tmpPlz = "0 " Then tmpPlz = ""
            Me.Label_Geb_AS_Ort.Text = tmpPlz & Trim(.Aussteller_Ort)
            Me.Label_Geb_AS_Strasse.Text = Trim(.Aussteller_Strasse) & " " & Trim(.Aussteller_Nr)
            Me.Label_Geb_AS_Zusatz.Text = Trim(.Aussteller_Zusatz)
        End With
    End Sub
    Public Sub Update_SSA_einfach()

        'Aussteller
        With My.Settings 'Aussteller.Adresse
            Me.Label_Aussteller_e_Name.Text = Trim(.Aussteller_Name)
            Dim tmpPlz As String = Trim(.Aussteller_PLZ) & " "
            'If tmpPlz = "0 " Then tmpPlz = ""
            Me.Label_Aussteller_e_Ort.Text = tmpPlz & Trim(.Aussteller_Ort)
            Me.Label_Aussteller_e_Strasse.Text = Trim(.Aussteller_Strasse) & " " & Trim(.Aussteller_Nr)
            Me.Label_Aussteller_e_Zusatz.Text = Trim(.Aussteller_Zusatz)
        End With
        'Logo/Signatur
        Me.PB_E_Sig.BackgroundImage = GetImageFromString(My.Settings.Signatur)
        Me.PB_e_Logo.BackgroundImage = GetImageFromString(My.Settings.Logo)

        'Pkte der Rubriken übertragen
        Me.Label_Pkte_e_I.Text = Me.Label_Pkte_I.Text
        Me.Label_Pkte_e_II.Text = Me.Label_Pkte_II.Text

        'Klassen der Rubriken übertragen
        Me.Label_e_Klasse_I.Text = Get_Klasse_I() 'Me.Label_d_Klasse_I.Text
        Me.Label_e_Klasse_I.BackColor = Get_Klasse_Color(Me.Label_e_Klasse_I.Text)
        Me.Label_e_Klasse_II.Text = Get_Klasse_II()   'Me.Label_d_Klasse_II.Text
        Me.Label_e_Klasse_II.BackColor = Get_Klasse_Color(Me.Label_e_Klasse_II.Text)
        'Pfeile positionieren
        'Me.Panel_Pfeil_I.Location = Get_Klasse_PfeilPos(Me.Label_e_Klasse_I.Text, 374)
        Me.Panel_Pfeil_II.Location = Get_Klasse_PfeilPos(Me.Label_e_Klasse_II.Text, 639)
        'Mindestpktzahl angeben
        Me.Label_Pkte_mind_I.Text = Get_txt_mind_I(Me.Label_e_Klasse_I.Text)
        Me.Label_Pkte_mind_II.Text = Get_txt_mind_II(Me.Label_e_Klasse_II.Text)
        'Bonus angeben
        Me.Label_Bonus_II.Text = Get_txt_Bonus()

        Me.Label_Mes_ja.Text = Get_Mes_ja()
        Me.Label_Mes_nein.Text = Get_Mes_nein()
        Me.Label_gesKl_ja.Text = Get_gesKl_ja()
        Me.Label_gesKl_nein.Text = Get_gesKl_nein()
        'Beschreibung der Klasse
        Me.Label_Klasse_I_Beschreibung.Text = Get_Klasse_I_Beschreibung()
        Me.Label_Klasse_II_Beschreibung.Text = Get_Klasse_II_Beschreibung()
        'Gebäudeinfos
        With Projekt.Gebaeude
            Me.Label_Gebaeudetyp.Text = .Gebaeudetyp
            If .Baujahr = 1800 Then
                Me.Label_Baujahr.Text = ""
            Else
                Me.Label_Baujahr.Text = .Baujahr.ToString
            End If
            Me.Label_AnzWohneinheiten.Text = .Wohneinheiten '.ToString
        End With
        With Projekt.Wohnung
            Me.Label_Wohnungsbez.Text = Trim(.Wohnungsbezeichnung)
            Me.Label_Geschoss.Text = Get_Str_Geschoss()
            Me.Label_AnzRaeume.Text = .Raeume '.ToString
            Me.Label_Wohnflaeche.Text = .Wohnflaeche.ToString
        End With

    End Sub

    Public Sub Update_SSA_detailiert()
        Update_SSA_Datum()
        Update_SSA_Antragsteller()
        Update_SSA_Gebaeude()

        'Aussteller
        With My.Settings 'Aussteller.Adresse
            Me.Label_Aussteller_d_Name.Text = Trim(.Aussteller_Name)
            Dim tmpPlz As String = Trim(.Aussteller_PLZ) & " "
            'If tmpPlz = "0 " Then tmpPlz = ""
            Me.Label_Aussteller_d_Ort.Text = tmpPlz & Trim(.Aussteller_Ort)
            Me.Label_Aussteller_d_Strasse.Text = Trim(.Aussteller_Strasse) & " " & Trim(.Aussteller_Nr)
            Me.Label_Aussteller_d_Zusatz.Text = Trim(.Aussteller_Zusatz)
        End With

        Me.PB_d_Sig.BackgroundImage = GetImageFromString(My.Settings.Signatur)

        'Rubrik Standort und Außenlärm
        Me.Update_SSA_GC()
        Me.Update_SSA_AL()

        'Baulicher Schallschutz
        Update_SSA_LS_W()
        Update_SSA_LS_D()
        Update_SSA_TS_D()
        Update_SSA_TS_TPLH()
        Update_SSA_TS_BLT()
        Update_SSA_Tueren()
        Update_SSA_Aussenbauteile()
        Update_SSA_Wasser()
        Update_SSA_Nutzer_Koerper()
        Update_SSA_Nachbarn()
        Update_SSA_anordRaeume()
        Update_SSA_lauteRaeume()
        Update_SSA_NHZ()
        Update_SSA_EW()

        Update_SSA_Bewertung()
    End Sub
    Public Sub Update_SSA_Bewertung()
        Me.Label_Pkte_I.Text = Calc_Pkte_I()
        Me.Label_Pkte_II.Text = Calc_Pkte_II()
        Me.Label_Ges_Pkte_I.Text = Me.Label_Pkte_I.Text
        Me.Label_Ges_Pkte_II.Text = Me.Label_Pkte_II.Text

        Me.Label_d_Klasse_I.Text = Get_Klasse_I()
        Me.Label_d_Klasse_II.Text = Get_Klasse_II()

        Me.Label_d_Beurteilung_I.Text = Get_Klasse_I()
        Me.Label_d_Beurteilung_II.Text = Get_Klasse_II()

        Me.Label_Ges_Pkte_I.BackColor = Get_Klasse_Color(Me.Label_d_Beurteilung_I.Text)
        Me.Label_Ges_Pkte_II.BackColor = Get_Klasse_Color(Me.Label_d_Beurteilung_II.Text)
        Me.Label_d_Beurteilung_I.BackColor = Get_Klasse_Color(Me.Label_d_Beurteilung_I.Text)
        Me.Label_d_Beurteilung_II.BackColor = Get_Klasse_Color(Me.Label_d_Beurteilung_II.Text)

        Update_SSA_einfach()
        Update_SSA_Status()
    End Sub
    Public Sub Update_SSA_Datum()
        'gültig bis
        With Projekt
            If .Datum <> "" Then
                Dim tmpdate As Date = CDate(.Datum)
                Me.Label_d_Datum.Text = IIf(tmpdate.Day < 10, "0" & tmpdate.Day, tmpdate.Day).ToString & "." & IIf(tmpdate.Month < 10, "0" & tmpdate.Month, tmpdate.Month).ToString & "." & tmpdate.Year.ToString 'Projekt.Datum
                Me.Label_Geb_Datum.Text = IIf(tmpdate.Day < 10, "0" & tmpdate.Day, tmpdate.Day).ToString & "." & IIf(tmpdate.Month < 10, "0" & tmpdate.Month, tmpdate.Month).ToString & "." & tmpdate.Year.ToString 'CType(.Datum, Date).Day & "." & CType(.Datum, Date).Month & "." & CType(.Datum, Date).Year
                Me.Label_e_Datum.Text = IIf(tmpdate.Day < 10, "0" & tmpdate.Day, tmpdate.Day).ToString & "." & IIf(tmpdate.Month < 10, "0" & tmpdate.Month, tmpdate.Month).ToString & "." & tmpdate.Year.ToString 'Projekt.Datum

                tmpdate = tmpdate.AddYears(10)

                Me.Label_d_gueltigBis.Text = IIf(tmpdate.Day < 10, "0" & tmpdate.Day, tmpdate.Day).ToString & "." & IIf(tmpdate.Month < 10, "0" & tmpdate.Month, tmpdate.Month).ToString & "." & tmpdate.Year 'CType(.Datum, Date).Day & "." & CType(.Datum, Date).Month & "." & CType(.Datum, Date).AddYears(10).Year
                Me.Label_Geb_gueltigBis.Text = IIf(tmpdate.Day < 10, "0" & tmpdate.Day, tmpdate.Day).ToString & "." & IIf(tmpdate.Month < 10, "0" & tmpdate.Month, tmpdate.Month).ToString & "." & tmpdate.Year 'CType(.Datum, Date).Day & "." & CType(.Datum, Date).Month & "." & CType(.Datum, Date).AddYears(10).Year
                Me.Label_e_gueltigBis.Text = IIf(tmpdate.Day < 10, "0" & tmpdate.Day, tmpdate.Day).ToString & "." & IIf(tmpdate.Month < 10, "0" & tmpdate.Month, tmpdate.Month).ToString & "." & tmpdate.Year 'CType(.Datum, Date).Day & "." & CType(.Datum, Date).Month & "." & CType(.Datum, Date).AddYears(10).Year
            End If
        End With

        ''gültig bis
        'With Projekt
        '    If .Datum <> "" Then
        '        Dim tmpDate As Date = CDate(.Datum)
        '        Me.Label_Geb_Datum.Text = IIf(tmpDate.Day < 10, "0" & tmpDate.Day, tmpDate.Day).ToString & "." & IIf(tmpDate.Month < 10, "0" & tmpDate.Month, tmpDate.Month).ToString & "." & tmpDate.Year.ToString 'CType(.Datum, Date).Day & "." & CType(.Datum, Date).Month & "." & CType(.Datum, Date).Year
        '        tmpDate.AddYears(10)
        '        Me.Label_Geb_gueltigBis.Text = IIf(tmpDate.Day < 10, "0" & tmpDate.Day, tmpDate.Day).ToString & "." & IIf(tmpDate.Month < 10, "0" & tmpDate.Month, tmpDate.Month).ToString & "." & tmpDate.Year.ToString 'CType(.Datum, Date).Day & "." & CType(.Datum, Date).Month & "." & CType(.Datum, Date).AddYears(10).Year
        '    End If
        'End With

        ''gültig bis
        'With Projekt
        '    If .Datum <> "" Then
        '        Me.Label_e_gueltigBis.Text = CType(.Datum, Date).Day & "." & CType(.Datum, Date).Month & "." & CType(.Datum, Date).AddYears(10).Year
        '    End If
        'End With

        'Me.Label_e_Datum.Text = Projekt.Datum

    End Sub
    Public Sub Update_SSA_Antragsteller()
        'Antragsteller
        With Projekt.Antragsteller
            Me.Label_Antragsteller_d_Name.Text = Trim(.Name)
            Dim tmpPlz As String = Trim(.PLZ) & " "
            'If tmpPlz = "0 " Then tmpPlz = ""
            Me.Label_Antragsteller_d_Ort.Text = tmpPlz & Trim(.Ort)
            Me.Label_Antragsteller_d_Strasse.Text = Trim(.Strasse) & " " & Trim(.Nr)
            Me.Label_Antragsteller_d_Zusatz.Text = Trim(.Zusatz)
        End With
        'Antragsteller
        With Projekt.Antragsteller
            Me.Label_Antragsteller_e_Name.Text = Trim(.Name)
            Dim tmpPlz As String = Trim(.PLZ) & " "
            'If tmpPlz = "0 " Then tmpPlz = ""
            Me.Label_Antragsteller_e_Ort.Text = tmpPlz & Trim(.Ort)
            Me.Label_Antragsteller_e_Strasse.Text = Trim(.Strasse) & " " & Trim(.Nr)
            Me.Label_Antragsteller_e_Zusatz.Text = Trim(.Zusatz)
        End With

    End Sub
    Public Sub Update_SSA_Gebaeude()
        'Gebäude
        With Projekt.Gebaeude.Adresse
            Me.Label_Gebaeude_d_Name.Text = Trim(.Name)
            Dim tmpPlz As String = Trim(.PLZ) & " "
            'If tmpPlz = "0 " Then tmpPlz = ""
            Me.Label_Gebaeude_d_Ort.Text = tmpPlz & Trim(.Ort)
            Me.Label_Gebaeude_d_Strasse.Text = Trim(.Strasse) & " " & Trim(.Nr)
            Me.Label_Gebaeude_d_Zusatz.Text = Trim(.Zusatz)
        End With
        'Wohnungsbezeichnung
        Me.Label_d_Wohnungsbez.Text = Trim(Projekt.Wohnung.Wohnungsbezeichnung)


        'Gebäude
        With Projekt.Gebaeude.Adresse
            Me.Label_Gebaeude_e_Name.Text = Trim(.Name)
            Dim tmpPlz As String = Trim(.PLZ) & " "
            'If tmpPlz = "0 " Then tmpPlz = ""
            Me.Label_Gebaeude_e_Ort.Text = tmpPlz & Trim(.Ort)
            Me.Label_Gebaeude_e_Strasse.Text = Trim(.Strasse) & " " & Trim(.Nr)
            Me.Label_Gebaeude_e_Zusatz.Text = Trim(.Zusatz)
        End With
        'Wohnungsbezeichnung
        Me.Label_e_Wohnungsbez.Text = Trim(Projekt.Wohnung.Wohnungsbezeichnung)

        'Gebäude
        With Projekt.Gebaeude.Adresse
            Dim tmpPlz As String = Trim(.PLZ) & " "
            'If tmpPlz = "0 " Then tmpPlz = ""
            Me.Label_Geb_Adresse.Text = Trim(.Strasse) & Trim(.Nr) & Chr(13) & Chr(10) & tmpPlz & Trim(.Ort)
        End With
        'Gebäudeinfos
        With Projekt.Gebaeude
            Me.Label_Geb_Typ.Text = .Gebaeudetyp
            'If .Baujahr = 0 Then
            '    Me.Label_Geb_Bauj.Text = ""
            'Else
            '    Me.Label_Geb_Bauj.Text = .Baujahr.ToString
            'End If
            Me.Label_Geb_Bauj.Text = Me.TB_Gebaeude_Baujahr.Text

            Me.Label_Geb_AnzWohnein.Text = .Wohneinheiten '.ToString

        End With
        With Projekt.Wohnung
            Me.Label_Geb_Wohnungsbez.Text = Trim(.Wohnungsbezeichnung)
            Me.Label_Geb_Geschoss.Text = Get_Str_Geschoss()
            'If .Raeume = 0 Then
            'Me.Label_Geb_AnzRaeume.Text = ""
            'Else
            Me.Label_Geb_AnzRaeume.Text = .Raeume '.ToString
            'End If
            If .Wohnflaeche = 0 Then
                Me.Label_Geb_Wohnfl.Text = ""
            Else
                Me.Label_Geb_Wohnfl.Text = .Wohnflaeche.ToString
            End If
        End With
    End Sub
    Public Sub Update_SSA_GC()
        Me.Label_GC.Text = Get_Str_GC()

        Me.Label_Pkte_GC.Text = Get_Pkte_GC()
        Me.Label_Klasse_GC.Text = Get_Klasse_GC()
        Me.Label_Klasse_GC.BackColor = Get_Klasse_Color(Me.Label_Klasse_GC.Text)

        If bUpdate_Anzeige = False Then Update_SSA_Bewertung()
    End Sub
    Public Sub Update_SSA_AL()
        Me.Label_AP.Text = Get_Str_AP()
        Me.Label_aF.Text = Get_Str_aF()

        Me.Label_Pkte_Aussenlaerm.Text = Get_Pkte_Aussenlaerm()
        Me.Label_Klasse_Aussenlaerm.Text = Get_Klasse_AP()
        Me.Label_Klasse_Aussenlaerm.BackColor = Get_Klasse_Color(Me.Label_Klasse_Aussenlaerm.Text)

        If bUpdate_Anzeige = False Then Update_SSA_Bewertung()
    End Sub
    Public Sub Update_SSA_LS_W()
        Me.Label_LS_W_P.Text = Get_Str_LS_W_P()
        Me.Label_LS_W_M.Text = Get_Str_LS_W_M()
        Me.Label_LS_W_C.Text = Get_Str_LS_W_C()
        Me.Label_LS_W_R.Text = Get_Str_LS_W_R()

        Me.Label_Pkte_LS_W.Text = Get_Pkte_LS_W()
        Me.Label_Klasse_LS_W.Text = Get_Klasse_LS_Waende()
        Me.Label_Klasse_LS_W.BackColor = Get_Klasse_Color(Me.Label_Klasse_LS_W.Text)

        If Me.Label_LS_W_P.Text = "-" And Me.Label_LS_W_M.Text = "-" Then
            Me.TB_Bemerkung_LS_W.Text = "nicht vorhanden"
            'ElseIf Me.Label_LS_W_MV.Text = "KMV" And _
            '        (Me.Label_Klasse_LS_W.Text = "A*" Or Me.Label_Klasse_LS_W.Text = "A" Or Me.Label_Klasse_LS_W.Text = "B") Then
            '    Me.TB_Bemerkung_LS_W.Text = "KMV nicht geeignet"
        Else
            Me.TB_Bemerkung_LS_W.Text = ""
        End If

        If bUpdate_Anzeige = False Then Update_SSA_Bewertung()
    End Sub
    Public Sub Update_SSA_LS_D()
        Me.Label_LS_D_P.Text = Get_Str_LS_D_P()
        Me.Label_LS_D_M.Text = Get_Str_LS_D_M()
        Me.Label_LS_D_C.Text = Get_Str_LS_D_C()
        Me.Label_LS_D_R.Text = Get_Str_LS_D_R()

        Me.Label_Pkte_LS_D.Text = Get_Pkte_LS_D()
        Me.Label_Klasse_LS_D.Text = Get_Klasse_LS_Decken()
        Me.Label_Klasse_LS_D.BackColor = Get_Klasse_Color(Me.Label_Klasse_LS_D.Text)

        If Me.Label_LS_D_P.Text = "-" And Me.Label_LS_D_M.Text = "-" Then
            Me.TB_Bemerkung_LS_D.Text = "nicht vorhanden"
            'ElseIf Me.Label_LS_D_MV.Text = "KMV" And _
            '    (Me.Label_Klasse_LS_D.Text = "A*" Or Me.Label_Klasse_LS_D.Text = "A" Or Me.Label_Klasse_LS_D.Text = "B") Then
            '    Me.TB_Bemerkung_LS_D.Text = "KMV nicht geeignet"
        Else
            Me.TB_Bemerkung_LS_D.Text = ""
        End If

        If bUpdate_Anzeige = False Then Update_SSA_Bewertung()
    End Sub
    Public Sub Update_SSA_TS_D()
        Me.Label_TS_D_P.Text = Get_Str_TS_D_P()
        Me.Label_TS_D_M.Text = Get_Str_TS_D_M()
        Me.Label_TS_D_C.Text = Get_Str_TS_D_C()
        Me.Label_TS_D_L.Text = Get_Str_TS_D_L()
        'If Me.Label_TS_D_M.Text = "X" Then
        Me.Label_TS_D_fE.BackColor = System.Drawing.Color.Silver 'systemcolors.control
        Me.Label_TS_D_Be.BackColor = System.Drawing.Color.Silver 'systemcolors.control
        'Else
        'Me.Label_TS_D_MA.BackColor = Color.Transparent 'white
        'Me.Label_TS_D_MV.BackColor = Color.Transparent 'white
        'End If
        Me.Label_TS_D_fE.Text = Get_Str_TS_D_fE()
        Me.Label_TS_D_Be.Text = Get_Str_TS_D_Be()

        Me.Label_Pkte_TS_D.Text = Get_Pkte_TS_D()
        Me.Label_Klasse_TS_D.Text = Get_Klasse_TS_Decken()
        Me.Label_Klasse_TS_D.BackColor = Get_Klasse_Color(Me.Label_Klasse_TS_D.Text)

        If Me.Label_TS_D_P.Text = "-" And Me.Label_TS_D_M.Text = "-" Then
            Me.TB_Bemerkung_TS_D.Text = "nicht vorhanden"
            'ElseIf Me.Label_TS_D_MV.Text = "KMV" And _
            '    (Me.Label_Klasse_TS_D.Text = "A*" Or Me.Label_Klasse_TS_D.Text = "A" Or Me.Label_Klasse_TS_D.Text = "B") Then
            '    Me.TB_Bemerkung_TS_D.Text = "KMV nicht geeignet"
        Else
            Me.TB_Bemerkung_TS_D.Text = ""
        End If

        If bUpdate_Anzeige = False Then Update_SSA_Bewertung()
    End Sub
    Public Sub Update_SSA_TS_TPLH()
        Me.Label_TS_TPH_P.Text = Get_Str_TS_TPLH_P()
        Me.Label_TS_TPH_M.Text = Get_Str_TS_TPLH_M()
        Me.Label_TS_TPH_Be.Text = Get_Str_TS_TPLH_Be()
        Me.Label_TS_TPH_C.Text = Get_Str_TS_TPLH_C()
        Me.Label_TS_TPH_L.Text = Get_Str_TS_TPLH_L()
        'If Me.Label_TS_TPH_M.Text = "X" Then
        Me.Label_TS_TPH_Be.BackColor = System.Drawing.Color.Silver 'systemcolors.control
        'Else
        'Me.Label_TS_TPH_MV.BackColor = Color.Transparent 'white
        'End If

        Me.Label_Pkte_TS_TPH.Text = Get_Pkte_TS_TPLH()
        Me.Label_Klasse_TS_TPH.Text = Get_Klasse_TS_TPLH()
        Me.Label_Klasse_TS_TPH.BackColor = Get_Klasse_Color(Me.Label_Klasse_TS_TPH.Text)

        If Me.Label_TS_TPH_P.Text = "-" And Me.Label_TS_TPH_M.Text = "-" Then
            Me.TB_Bemerkung_TS_TPH.Text = "nicht vorhanden"
        ElseIf Me.Label_TS_TPH_Be.Text = "X" And _
            (Me.Label_Klasse_TS_TPH.Text = "A*" Or Me.Label_Klasse_TS_TPH.Text = "A" Or _
            Me.Label_Klasse_TS_TPH.Text = "D") Then
            Me.TB_Bemerkung_TS_TPH.Text = "Bodenbelag nicht anrechenbar"
        Else
            Me.TB_Bemerkung_TS_TPH.Text = ""
        End If

        If bUpdate_Anzeige = False Then Update_SSA_Bewertung()
    End Sub
    Public Sub Update_SSA_TS_BLT()
        Me.Label_TS_BLLT_P.Text = Get_Str_TS_BLT_P()
        Me.Label_TS_BLLT_M.Text = Get_Str_TS_BLT_M()
        Me.Label_TS_BLLT_Be.Text = Get_Str_TS_BLT_Be()
        Me.Label_TS_BLLT_C.Text = Get_Str_TS_BLT_C()
        Me.Label_TS_BLLT_L.Text = Get_Str_TS_BLT_L()
        'If Me.Label_TS_BLLT_M.Text = "X" Then
        Me.Label_TS_BLLT_Be.BackColor = System.Drawing.Color.Silver 'systemcolors.control
        'Else
        'Me.Label_TS_BLLT_Be.BackColor = Color.Transparent 'white
        'End If

        Me.Label_Pkte_TS_BLLT.Text = Get_Pkte_TS_BLT()
        Me.Label_Klasse_TS_BLLT.Text = Get_Klasse_TS_BLT()
        Me.Label_Klasse_TS_BLLT.BackColor = Get_Klasse_Color(Me.Label_Klasse_TS_BLLT.Text)

        If Me.Label_TS_BLLT_P.Text = "-" And Me.Label_TS_BLLT_M.Text = "-" Then
            Me.TB_Bemerkung_TS_BLLT.Text = "nicht vorhanden"
        ElseIf Me.Label_TS_BLLT_Be.Text = "X" And _
            (Me.Label_Klasse_TS_BLLT.Text = "A*" Or Me.Label_Klasse_TS_BLLT.Text = "A" Or _
            Me.Label_Klasse_TS_BLLT.Text = "D") Then
            Me.TB_Bemerkung_TS_BLLT.Text = "Bodenbelag nicht anrechenbar"
        Else
            Me.TB_Bemerkung_TS_BLLT.Text = ""
        End If

        If bUpdate_Anzeige = False Then Update_SSA_Bewertung()
    End Sub
    Public Sub Update_SSA_Tueren()
        Me.Label_Tueren_P_D.Text = Get_Str_Tueren_P_D()
        Me.Label_Tueren_M_D.Text = Get_Str_Tueren_M_D()
        Me.Label_Tueren_R_D.Text = Get_Str_Tueren_R_D()

        Me.Label_Tueren_P_A.Text = Get_Str_Tueren_P_A()
        Me.Label_Tueren_M_A.Text = Get_Str_Tueren_M_A()
        Me.Label_Tueren_R_A.Text = Get_Str_Tueren_R_A()

        Me.TB_Bemerkung_Tueren.Text = Get_Str_Tueren_B()

        Me.Label_Pkte_Tueren_D.Text = Get_Pkte_Tueren_D()
        Me.Label_Pkte_Tueren_A.Text = Get_Pkte_Tueren_A()

        Me.Label_Klasse_Tueren_A.Text = Get_Klasse_Tueren_A()
        Me.Label_Klasse_Tueren_A.BackColor = Get_Klasse_Color(Me.Label_Klasse_Tueren_A.Text)
        Me.Label_Klasse_Tueren_D.Text = Get_Klasse_Tueren_D()
        Me.Label_Klasse_Tueren_D.BackColor = Get_Klasse_Color(Me.Label_Klasse_Tueren_D.Text)

        If bUpdate_Anzeige = False Then Update_SSA_Bewertung()
    End Sub
    Public Sub Update_SSA_Aussenbauteile()
        Me.Label_Aussenbauteile_on.Text = Get_Str_Aussenbauteile_oN()
        'Me.Label_Aussenbauteile_FmD.Text = Get_Str_Aussenbauteile_FmD()
        Me.Label_Aussenbauteile_DIN.Text = Get_Str_Aussenbauteile_DIN()
        Me.Label_Aussenbauteile_DINPlus.Text = Get_Str_Aussenbauteile_DINPlus()

        Me.Label_Pkte_Aussenbauteile.Text = Get_Pkte_Aussenbauteile()
        Me.Label_Klasse_Aussenbauteile.Text = Get_Klasse_Aussenbauteile()
        Me.Label_Klasse_Aussenbauteile.BackColor = Get_Klasse_Color(Me.Label_Klasse_Aussenbauteile.Text)

        Me.TB_Bem_Aussenbauteile.Text = Projekt.Bem_Aussenbauteile

        If bUpdate_Anzeige = False Then Update_SSA_Bewertung()
    End Sub
    Public Sub Update_SSA_Wasser()
        Me.Label_Wasser_P.Text = Get_Str_Wasser_P()
        Me.Label_Wasser_M.Text = Get_Str_Wasser_M()
        Me.Label_Wasser_erfuellt.Text = Get_Str_Wasser_LcLa()
        Me.Label_Wasser_L.Text = Get_Str_Wasser_L()

        Me.Label_Pkte_Wasser.Text = Get_Pkte_Wasser()
        Me.Label_Klasse_Wasser.Text = Get_Klasse_Wasser()
        Me.Label_Klasse_Wasser.BackColor = Get_Klasse_Color(Me.Label_Klasse_Wasser.Text)

        Me.TB_Bem_WasserHausanlagen.Text = Projekt.Wasser.Bem_Wasser

        If bUpdate_Anzeige = False Then Update_SSA_Bewertung()
    End Sub
    Public Sub Update_SSA_Nutzer_Koerper()
        Me.Label_Nutzer_oN.Text = Get_Str_Nutzer_nv()
        Me.Label_Nutzer_P.Text = Get_Str_Nutzer_P()
        Me.Label_Nutzer_M.Text = Get_Str_Nutzer_M()
        Me.Label_Nutzer_L.Text = Get_Str_Nutzer_L()
        Me.TB_Bemerkung_Nutzer.Text = Get_Str_Nutzer_B()
         

        Me.Label_Koerper_oN.Text = Get_Str_Koerper_nv()
        Me.Label_Koerper_P.Text = Get_Str_Koerper_P()
        Me.Label_Koerper_M.Text = Get_Str_Koerper_M()
        Me.Label_Koerper_L.Text = Get_Str_Koerper_L()
        'Me.TB_Bemerkung_Koerper.Text = Get_Str_Koerper_B()
         
        Me.Label_Pkte_Koerper.Text = Get_Pkte_NutzerKoerper()

        'Me.Label_Klasse_Nutzer.Text = Get_Klasse_Nutzer()
        'Me.Label_Klasse_Koerper.Text = Get_Klasse_Koerper()

        Me.Label_Klasse_NutzerKoerper.Text = Get_Klasse_NutzerKoerper()
        Me.Label_Klasse_NutzerKoerper.BackColor = Get_Klasse_Color(Me.Label_Klasse_NutzerKoerper.Text)

        If bUpdate_Anzeige = False Then Update_SSA_Bewertung()
    End Sub
    Public Sub Update_SSA_Nachbarn()
        Me.Label_Nachbarn.Text = Get_Str_Nachbar()
        'If Me.Label_Nachbarn.Text <> "" Then
        '    Me.TB_Bemerkung_Nachbarn.Text = "nur Empfehlung"
        'Else
        '    Me.TB_Bemerkung_Nachbarn.Text = ""
        'End If
        Me.TB_Bem_AnordnungRaeume.Text = Projekt.Bemerkung_Nachbarn

        Me.Label_Pkte_Nachbarn.Text = Get_Pkte_Nachbarn()

        Me.Label_Klasse_Nachbarn.Text = Get_Klasse_Nachbarn()
        Me.Label_Klasse_Nachbarn.BackColor = Get_Klasse_Color(Me.Label_Klasse_Nachbarn.Text)

        If bUpdate_Anzeige = False Then Update_SSA_Bewertung()
    End Sub
    Public Sub Update_SSA_anordRaeume()
        Me.Label_anordRaeume_ug.Text = Get_Str_anordRaeume_ug()
        Me.Label_anordRaeume_g.Text = Get_Str_anordRaeume_g()

        Me.Label_Pkte_anordRaeume.Text = Get_Pkte_anordRaeume()

        Me.TB_Bem_AnordnungRaeume.Text = Projekt.Bem_anordnungRaeume

        If bUpdate_Anzeige = False Then Update_SSA_Bewertung()
    End Sub
    Public Sub Update_SSA_lauteRaeume()
        Me.Label_lauteRaeume_keine.Text = Get_Str_lauteRaeume_keine()
        Me.Label_lauteRaeume_25.Text = Get_Str_lauteRaeume_L_25_35()
        Me.Label_lauteRaeume_30.Text = Get_Str_lauteRaeume_L_30_40()
        Me.Label_lauteRaeume_35.Text = Get_Str_lauteRaeume_L_35_45()

        Me.Label_Pkte_lauteRaeume_keine.Text = Get_Pkte_lauteRaeume_keine()
        Me.Label_Pkte_lauteRaeume_25.Text = Get_Pkte_lauteRaeume_25()
        Me.Label_Pkte_lauteRaeume_30.Text = Get_Pkte_lauteRaeume_30()
        Me.Label_Pkte_lauteRaeume_35.Text = Get_Pkte_lauteRaeume_35()

        Me.Label_Klasse_lauteRaeume_keine.Text = Get_Klasse_lauteRaeume_kE()
        Me.Label_Klasse_lauteRaeume_keine.BackColor = Get_Klasse_Color(Me.Label_Klasse_lauteRaeume_keine.Text)

        Me.Label_Klasse_lauteRaeume_25.Text = Get_Klasse_lauteRaeume_25()
        Me.Label_Klasse_lauteRaeume_25.BackColor = Get_Klasse_Color(Me.Label_Klasse_lauteRaeume_25.Text)

        Me.Label_Klasse_lauteRaeume_30.Text = Get_Klasse_lauteRaeume_30()
        Me.Label_Klasse_lauteRaeume_30.BackColor = Get_Klasse_Color(Me.Label_Klasse_lauteRaeume_30.Text)

        Me.Label_Klasse_lauteRaeume_35.Text = Get_Klasse_lauteRaeume_35()
        Me.Label_Klasse_lauteRaeume_35.BackColor = Get_Klasse_Color(Me.Label_Klasse_lauteRaeume_35.Text)

        Me.TB_Bem_lauteRaeume.Text = Projekt.Bem_lauteRaeume

        If bUpdate_Anzeige = False Then Update_SSA_Bewertung()
    End Sub
    Public Sub Update_SSA_NHZ()
        Me.Label_NHZ.Text = Get_Str_NHZ()

        If Me.Label_NHZ.Text <> "" Then
            Me.TB_Bem_AV.Text = "nur Empfehlung"
        Else
            Me.TB_Bem_AV.Text = ""
        End If

        Me.Label_Pkte_NHZ.Text = Get_Pkte_NHZ()

        Me.Label_Klasse_NHZ.Text = Get_Klasse_NHZ()

        If bUpdate_Anzeige = False Then Update_SSA_Bewertung()
    End Sub
    Public Sub Update_SSA_EW()
        Me.Label_EW_1.Text = Get_Str_EW_1()
        Me.Label_EW_2.Text = Get_Str_EW_2()
        Me.Label_EW_3.Text = Get_Str_EW_3()
        Me.Label_EW_kE.Text = Get_Str_EW_kE()
        If Me.Label_EW_1.Text = "X" Or Me.Label_EW_2.Text = "X" Or Me.Label_EW_3.Text = "X" Or Me.Label_EW_kE.Text = "X" Then
            Me.TB_Bemerkung_EW.Text = "nur Empfehlung"
        Else
            Me.TB_Bemerkung_EW.Text = ""
        End If

        Me.Label_Pkte_EW.Text = Get_Pkte_EW()

        Me.Label_Klasse_EW.Text = Get_Klasse_EW()

        If bUpdate_Anzeige = False Then Update_SSA_Bewertung()
    End Sub


    Public Function Calc_Pkte_I() As String
        Dim iPkte As Integer = 0
        For Each ctrl As Control In Me.Panel_Pkte_I.Controls
            If TypeOf (ctrl) Is System.Windows.Forms.Label Then
                If IsNumeric(ctrl.Text) Then iPkte = iPkte + CInt(ctrl.Text)
            End If
        Next
        Calc_Pkte_I = IIf(iPkte = 0, "", iPkte.ToString)
    End Function
    Public Function Calc_Pkte_II() As String
        Dim iPkte As Integer = 0
        For Each ctrl As Control In Me.Panel_Pkte_II.Controls
            If TypeOf (ctrl) Is System.Windows.Forms.Label Then
                If IsNumeric(ctrl.Text) Then iPkte = iPkte + CInt(ctrl.Text)
            End If
        Next
        Calc_Pkte_II = iPkte.ToString
    End Function

#End Region


    Private Sub Button_Pruefdatum_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Pruefdatum.Click
        Dialog_Datum.ShowDialog()
        Dim sDateSelectionRange As System.Windows.Forms.SelectionRange = Dialog_Datum.MonthCalendar1.SelectionRange
        Me.TB_Datum.Text = Microsoft.VisualBasic.Left(sDateSelectionRange.Start.Date.ToString, 10)
    End Sub

    Private Sub TSB_Hand_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Cursor = Cursors.Hand
    End Sub

    Private Sub Button_DB_Antragsteller_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_DB_Antragsteller.Click
        Dialog_DB_Antragsteller.ShowDialog()
    End Sub

    Private Sub TSMI_DEGA_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSMI_DEGA.Click
        Dim sExecutingPath As String = Reflection.Assembly.GetEntryAssembly.Location 'Pfad zur ausführbaren datei dieses Programmes
        Dim sFolder As String = IO.Path.GetDirectoryName(sExecutingPath)

        Try
            Diagnostics.Process.Start(sFolder & "\DEGA_Empfehlung_103.pdf")
        Catch ex As IO.FileNotFoundException
            MsgBox("Die DEGA-Empfehlung kann nicht gefunden werden. Das Dokument " & Chr(34) & "DEGA_Empfehlung_103.pdf" & Chr(34) & " muss unter dem Pfad " _
                    & Chr(34) & sFolder & "\DEGA_Empfehlung_103.pdf" & Chr(34) & " gespeichert sein. Das Dokument befindet sich auf der Installations-CD.", _
                    MsgBoxStyle.OkOnly, "DEGA-Empfehlung 103")
        Catch ex As Exception
            MsgBox("Die DEGA-Empfehlung kann nicht geöffnet oder nicht gefunden werden. Es muss ein Programm zum Anzeigen von PDF-Dokumenten installiert sein." & Chr(13) & Chr(10) & _
                    "Außerdem muss das Dokument " & Chr(34) & "DEGA_Empfehlung_103.pdf" & Chr(34) & " unter dem Pfad " & Chr(34) & sFolder & "\DEGA_Empfehlung_103.pdf" & Chr(34) & _
                    " gespeichert sein. Das Dokument befindet sich auf der Installations-CD.", MsgBoxStyle.OkOnly, "DEGA-Empfehlung 103")
        End Try

    End Sub

    Private Sub Panel6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_Deckblatt.Click
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
        Form_Help.Close()

        PB_Geb_Logo.Select()
    End Sub
    Private Sub Panel8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_dSSA.Click
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
        Form_Help.Close()

        Label_Antragsteller_d_Name.Select()
    End Sub

    Private Sub Panel32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_eSSA.Click
        Me.Cursor = Cursors.Default
        Me.TSMI_DirekteHilfe.Checked = False
        Form_Help.Close()

        PB_e_Logo.Select()
    End Sub


    Private Sub TSB_DEGA_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSB_DEGA.Click
        Dim sExecutingPath As String = Reflection.Assembly.GetEntryAssembly.Location 'Pfad zur ausführbaren datei dieses Programmes
        Dim sFolder As String = IO.Path.GetDirectoryName(sExecutingPath)

        Try
            Diagnostics.Process.Start(sFolder & "\DEGA_Empfehlung_103.pdf")
        Catch ex As IO.FileNotFoundException
            MsgBox("Die DEGA-Empfehlung kann nicht gefunden werden. Das Dokument " & Chr(34) & "DEGA_Empfehlung_103.pdf" & Chr(34) & " muss unter dem Pfad " _
                    & Chr(34) & sFolder & "\DEGA_Empfehlung_103.pdf" & Chr(34) & " gespeichert sein. Das Dokument befindet sich auf der Installations-CD.", _
                    MsgBoxStyle.OkOnly, "DEGA-Empfehlung 103")
        Catch ex As Exception
            MsgBox("Die DEGA-Empfehlung kann nicht geöffnet oder nicht gefunden werden. Es muss ein Programm zum Anzeigen von PDF-Dokumenten installiert sein." & Chr(13) & Chr(10) & _
                    "Außerdem muss das Dokument " & Chr(34) & "DEGA_Empfehlung_103.pdf" & Chr(34) & " unter dem Pfad " & Chr(34) & sFolder & "\DEGA_Empfehlung_103.pdf" & Chr(34) & _
                    " gespeichert sein. Das Dokument befindet sich auf der Installations-CD.", MsgBoxStyle.OkOnly, "DEGA-Empfehlung 103")
        End Try
    End Sub


    Private Sub TSMI_Handbuch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSMI_Handbuch.Click
        Dim sExecutingPath As String = Reflection.Assembly.GetEntryAssembly.Location 'Pfad zur ausführbaren datei dieses Programmes
        Dim sFolder As String = IO.Path.GetDirectoryName(sExecutingPath)

        Try
            Diagnostics.Process.Start(sFolder & "\Handbuch.pdf")
        Catch ex As IO.FileNotFoundException
            MsgBox("Das Handbuch kann nicht gefunden werden. Das Dokument " & Chr(34) & "Handbuch.pdf" & Chr(34) & " muss unter dem Pfad " _
                    & Chr(34) & sFolder & "\Handbuch.pdf" & Chr(34) & " gespeichert sein. Das Dokument befindet sich auf der Installations-CD.", _
                    MsgBoxStyle.OkOnly, "Handbuch")
        Catch ex As Exception
            MsgBox("Das Handbuch kann nicht geöffnet oder nicht gefunden werden. Es muss ein Programm zum Anzeigen von PDF-Dokumenten installiert sein." & Chr(13) & Chr(10) & _
                    "Außerdem muss das Dokument " & Chr(34) & "Handbuch.pdf" & Chr(34) & " unter dem Pfad " & Chr(34) & sFolder & "\Handbuch.pdf" & Chr(34) & _
                    " gespeichert sein. Das Dokument befindet sich auf der Installations-CD.", MsgBoxStyle.OkOnly, "Handbuch")
        End Try

    End Sub

    Private Sub Panel_Deckblatt_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel_Deckblatt.Paint
        If btDemo > versNormal Then Print_Demoversion_quer(e.Graphics, 500, 0, 152)
        'If btDemo > versNormal Then
        '    Dim gr As Graphics = e.Graphics
        '    gr.RotateTransform(55)
        '    gr.DrawString("Demoversion", New System.Drawing.Font("arial", 152, FontStyle.Bold), New SolidBrush(Color.FromArgb(150, 255, 127, 127)), 500, 0)    'Brushes.Red
        'End If
    End Sub
    Private Sub Panel_dSSA_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel_dSSA.Paint
        If btDemo > versNormal Then Print_Demoversion_quer(e.Graphics, 500, 0, 152)

        'If btDemo > versNormal Then
        '    Dim gr As Graphics = e.Graphics
        '    gr.RotateTransform(55)
        '    gr.DrawString("Demoversion", New System.Drawing.Font("arial", 152, FontStyle.Bold), New SolidBrush(Color.FromArgb(150, 255, 127, 127)), 500, 0)    'Brushes.Red
        'End If
    End Sub

    Private Sub Panel_eSSA_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel_eSSA.Paint
        If btDemo > versNormal Then Print_Demoversion_quer(e.Graphics, 500, 0, 152)
        'Dim gr As Graphics = e.Graphics
        'gr.RotateTransform(55)
        'gr.DrawString("Demoversion", New System.Drawing.Font("arial", 152, FontStyle.Bold), New SolidBrush(Color.FromArgb(150, 255, 127, 127)), 500, 0)    'Brushes.Red
        'End If
    End Sub
    Private Sub Print_Demoversion_quer(ByVal gr As Graphics, ByVal xKoord As Integer, ByVal yKoord As Integer, ByVal schriftGroesse As Integer)

        gr.RotateTransform(55)
        gr.DrawString("Demoversion", New System.Drawing.Font("arial", schriftGroesse, FontStyle.Bold), New SolidBrush(Color.FromArgb(150, 255, 127, 127)), xKoord, yKoord)    'Brushes.Red

    End Sub


    Private Sub NUD_TS_TPLH_Mes_L_1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NUD_TS_TPLH_Mes_L_1.ValueChanged

    End Sub

    Private Sub Label_e_gueltigBis_Click(sender As Object, e As EventArgs) Handles Label_e_gueltigBis.Click

    End Sub
End Class
