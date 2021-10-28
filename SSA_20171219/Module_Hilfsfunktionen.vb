Module Module_Hilfsfunktionen
    Public Sub Loesche_Projektdaten()

        stProjekt_Name = ""
        stProjekt_Pfad = ""

        Form_Main.Projekt_Default()

        'Projekt.Name = ""
        'Projekt.Pfad = ""
        'Projekt.Grundriss = ""

        'Projekt.bStatus_Grundriss = False
        'Projekt.bStatus_Projekt = False


        'With Projekt
        '    .Antragsteller.Name = ""
        '    .Antragsteller.Zusatz = ""
        '    .Antragsteller.Strasse = ""
        '    .Antragsteller.Nr = ""
        '    .Antragsteller.PLZ = ""
        '    .Antragsteller.Ort = ""

        '    .Gebaeude.Adresse.Name = ""
        '    .Gebaeude.Adresse.Zusatz = ""
        '    .Gebaeude.Adresse.Strasse = ""
        '    .Gebaeude.Adresse.Nr = ""
        '    .Gebaeude.Adresse.PLZ = ""
        '    .Gebaeude.Adresse.Ort = ""
        '    .Gebaeude.Gebaeudetyp = ""
        '    .Gebaeude.Baujahr = 0
        '    .Gebaeude.Wohneinheiten = ""
        '    .Gebaeude.Kosten = 0

        '    .Wohnung.Wohnungsbezeichnung = ""
        '    .Wohnung.Geschoss.Nr = 0
        '    .Wohnung.Geschoss.Typ = 0
        '    .Wohnung.Geschoss.Bezeichnung = ""
        '    .Wohnung.Raeume = ""
        '    .Wohnung.Wohnflaeche = 1
        'End With
        'With Projekt.Standort
        '    'Standort
        '    .Gebietscharakter = 0
        '    .Bem_Gebietscharakter = ""
        '    'Außenlärm
        '    .Aussenlaermpegel = 0
        '    .Bem_Aussenlaermpegel = ""
        '    'Freibereich
        '    .abgewandFreibereich = 0
        'End With
        ''LS Wände
        'With Projekt.LS_Wand
        '    .Untersuchung = 0
        '    .Prognose.Pegel = 0
        '    .Prognose.C = ""
        '    '.Bemerkung_LS = ""
        'End With
        'With Projekt.LS_Wand.Messung

        '    .Messung_1.C = ""
        '    .Messung_1.Pegel = 0

        '    .Messung_2.C = ""
        '    .Messung_2.Pegel = 0

        '    .Messung_3.C = ""
        '    .Messung_3.Pegel = 0

        '    .Messung_4.C = ""
        '    .Messung_4.Pegel = 0

        '    .Messung_5.C = ""
        '    .Messung_5.Pegel = 0

        '    .Messung_6.C = ""
        '    .Messung_6.Pegel = 0
        'End With

        ''LS Decken
        'With Projekt.LS_Decke
        '    .Untersuchung = 0
        '    .Prognose.Pegel = 0
        '    .Prognose.C = ""
        '    '.Bemerkung_LS = ""
        'End With
        'With Projekt.LS_Decke.Messung

        '    .Messung_1.C = ""
        '    .Messung_1.Pegel = 0

        '    .Messung_2.C = ""
        '    .Messung_2.Pegel = 0

        '    .Messung_3.C = ""
        '    .Messung_3.Pegel = 0

        '    .Messung_4.C = ""
        '    .Messung_4.Pegel = 0

        '    .Messung_5.C = ""
        '    .Messung_5.Pegel = 0

        '    .Messung_6.C = ""
        '    .Messung_6.Pegel = 0
        'End With

        ''TS Decken
        'With Projekt.TS_Decke
        '    .Untersuchung = 0
        '    .Prognose.Bodenbelag = 0
        '    .Prognose.Pegel.Pegel = 0
        '    .Prognose.Pegel.C = ""
        '    '.Bemerkung_TS_Decke = ""
        'End With
        'With Projekt.TS_Decke.Messung
        '    .Messung_1.Bodenbelag = 0
        '    .Messung_1.Pegel.C = ""
        '    .Messung_1.Pegel.Pegel = 0

        '    .Messung_2.Bodenbelag = 0
        '    .Messung_2.Pegel.C = ""
        '    .Messung_2.Pegel.Pegel = 0

        '    .Messung_3.Bodenbelag = 0
        '    .Messung_3.Pegel.C = ""
        '    .Messung_3.Pegel.Pegel = 0

        '    .Messung_4.Bodenbelag = 0
        '    .Messung_4.Pegel.C = ""
        '    .Messung_4.Pegel.Pegel = 0

        '    .Messung_5.Bodenbelag = 0
        '    .Messung_5.Pegel.C = ""
        '    .Messung_5.Pegel.Pegel = 0

        '    .Messung_6.Bodenbelag = 0
        '    .Messung_6.Pegel.C = ""
        '    .Messung_6.Pegel.Pegel = 0
        'End With

        ''TS TPH
        'With Projekt.TS_TPlH
        '    .Untersuchung = 0
        '    .Prognose.Treppe.Bodenbelag = 0
        '    .Prognose.Treppe.Pegel.Pegel = 0
        '    .Prognose.Treppe.Pegel.C = ""
        '    .Prognose.Podest.Bodenbelag = 0
        '    .Prognose.Podest.Pegel.Pegel = 0
        '    .Prognose.Podest.Pegel.C = ""
        '    .Prognose.Laube.Bodenbelag = 0
        '    .Prognose.Laube.Pegel.Pegel = 0
        '    .Prognose.Laube.Pegel.C = ""
        '    .Prognose.Hausflur.Bodenbelag = 0
        '    .Prognose.Hausflur.Pegel.Pegel = 0
        '    .Prognose.Hausflur.Pegel.C = ""
        '    '.Bemerkung_TS_TPH = ""
        'End With
        'With Projekt.TS_TPlH.Messung
        '    .Anzahl = 0

        '    .Messung_1.Bodenbelag = 0
        '    .Messung_1.Pegel.C = ""
        '    .Messung_1.Pegel.Pegel = 0

        '    .Messung_2.Bodenbelag = 0
        '    .Messung_2.Pegel.C = ""
        '    .Messung_2.Pegel.Pegel = 0

        '    .Messung_3.Bodenbelag = 0
        '    .Messung_3.Pegel.C = ""
        '    .Messung_3.Pegel.Pegel = 0

        '    .Messung_4.Bodenbelag = 0
        '    .Messung_4.Pegel.C = ""
        '    .Messung_4.Pegel.Pegel = 0

        '    .Messung_5.Bodenbelag = 0
        '    .Messung_5.Pegel.C = ""
        '    .Messung_5.Pegel.Pegel = 0

        '    .Messung_6.Bodenbelag = 0
        '    .Messung_6.Pegel.C = ""
        '    .Messung_6.Pegel.Pegel = 0
        'End With
        ''TS BLLT
        'With Projekt.TS_BLT
        '    .Untersuchung = 0
        '    '.Bemerkung_TS_BLLT=""
        '    .Prognose.Balkon.Bodenbelag = 0
        '    .Prognose.Balkon.Pegel.Pegel = 0
        '    .Prognose.Balkon.Pegel.C = ""
        '    .Prognose.Loggia.Bodenbelag = 0
        '    .Prognose.Loggia.Pegel.Pegel = 0
        '    .Prognose.Loggia.Pegel.C = ""
        '    .Prognose.Terrasse.Bodenbelag = 0
        '    .Prognose.Terrasse.Pegel.Pegel = 0
        '    .Prognose.Terrasse.Pegel.C = ""
        '    '.Bemerkung_TS_BLLT = ""
        'End With
        'With Projekt.TS_BLT.Messung
        '    .Anzahl = 0

        '    .Messung_1.Bodenbelag = 0
        '    .Messung_1.Pegel.C = ""
        '    .Messung_1.Pegel.Pegel = 0

        '    .Messung_2.Bodenbelag = 0
        '    .Messung_2.Pegel.C = ""
        '    .Messung_2.Pegel.Pegel = 0

        '    .Messung_3.Bodenbelag = 0
        '    .Messung_3.Pegel.C = ""
        '    .Messung_3.Pegel.Pegel = 0

        '    .Messung_4.Bodenbelag = 0
        '    .Messung_4.Pegel.C = ""
        '    .Messung_4.Pegel.Pegel = 0

        '    .Messung_5.Bodenbelag = 0
        '    .Messung_5.Pegel.C = ""
        '    .Messung_5.Pegel.Pegel = 0

        '    .Messung_6.Bodenbelag = 0
        '    .Messung_6.Pegel.C = ""
        '    .Messung_6.Pegel.Pegel = 0
        'End With

        ''Tueren
        'With Projekt.Tueren
        '    .Ort = 0
        '    .Untersuchung = 0
        '    .L = 0
        '    '.Bemerkung_Tueren = ""
        'End With

        ''LS Außenbauteile
        'Projekt.Aussenbauteile = 0
        'Projekt.Bem_Aussenbauteile = ""
        ''Wasser + haustechn. Anlagen
        'With Projekt.Wasser
        '    .Untersuchung = 0
        '    .L_Intervall = 0
        '    .LcLa_erfuellt = 0
        '    .Bem_Wasser = ""
        'End With
        ''Nutzer, Körper
        'With Projekt.NutzerKoerper
        '    .Untersuchung = 0
        '    .MessungPrognose = 0
        '    .L_Intervall = 0
        '    .BemerkungKoerper = ""
        '    .BemerkungNutzer = ""
        '    '.Bemerkung = ""
        'End With
        ' ''Körper
        ''With Projekt.Koerper
        ''    .Untersuchung = 0
        ''    .L_Intervall = 0
        ''    '.Bemerkung = ""
        ''End With
        ''Nachbarn
        'Projekt.Nachbarn = 0
        ''Projekt.Bemerkung_Nachbarn = ""
        ''Anordnung lauter Räume
        'Projekt.anordnungRaeume = 0
        'Projekt.Bem_anordnungRaeume = ""
        ''laute Räume angrenzend
        'Projekt.lauteRaeume = 0
        'Projekt.Bem_lauteRaeume = ""
        ''eigener Wohnbereich
        'Projekt.eigenerWohnbereich = 0
        ''Projekt.Bemerkung_eigenerWohnbereich = ""

        'If Not IsNothing(Form_Main.PB_Grundriss.BackgroundImage) Then
        '    ' Form_Main.PB_Grundriss.BackgroundImage.Dispose()
        '    Form_Main.PB_Grundriss.BackgroundImage = Nothing
        'End If
    End Sub
End Module
