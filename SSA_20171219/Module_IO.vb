Option Strict Off

'Imports Excel
Imports microsoft.Office.Interop.Excel

Module Module_IO
    'Public btmLogo As Bitmap
    ' Public btmSignatur As Bitmap
    Public btmGrundriss As Bitmap

    Public Function GetStringFromArraylist(ByVal arList As ArrayList) As String
        If Not IsNothing(arList) Then
            Dim tStr As String = ""
            For i As Integer = 0 To arList.Count - 1
                Dim tDB As AdresseData = CType(arList(i), AdresseData)
                tStr = tStr & tDB.Name & "+" & tDB.Nr & "+" & tDB.Ort & "+" & tDB.PLZ & "+" & tDB.Strasse & "+" & tDB.Zusatz & ";"
            Next
            Return Left(tStr, tStr.Length - 1)
        Else
            Return ""
        End If
    End Function
    Public Function GetArraylistFromString(ByVal thisStr As String) As ArrayList
        If thisStr = "" Then
            Return Nothing
        Else
            Dim tArList As New ArrayList
            Dim aAdr() As String = thisStr.Split(";")

            For iEl As Integer = 0 To aAdr.Length - 1
                Dim tAdr As AdresseData
                tAdr.Name = aAdr(iEl).Split("+")(0)
                tAdr.Nr = aAdr(iEl).Split("+")(1)
                tAdr.Ort = aAdr(iEl).Split("+")(2)
                tAdr.PLZ = aAdr(iEl).Split("+")(3)
                tAdr.Strasse = aAdr(iEl).Split("+")(4)
                tAdr.Zusatz = aAdr(iEl).Split("+")(5)

                tArList.Add(tAdr)
            Next
            Return tArList
        End If
    End Function

    Public Function GetStringFromImage(ByVal image As Image) As String
        If Image IsNot Nothing Then
            Dim ic As New ImageConverter
            Dim buffer As Byte() = DirectCast(ic.ConvertTo(Image, GetType(Byte())), Byte())
            Return Convert.ToBase64String(buffer, Base64FormattingOptions.InsertLineBreaks)
        Else
            Return Nothing
        End If
    End Function
    Public Function GetImageFromString(ByVal base64String As String) As Image
        If String.IsNullOrEmpty(base64String) Then Return Nothing
        Dim buffer As Byte() = Convert.FromBase64String(base64String)
        If buffer IsNot Nothing Then
            Dim ic As New ImageConverter()
            Return TryCast(ic.ConvertFrom(buffer), Image)
        Else
            Return Nothing
        End If
    End Function

    Public Sub Projektdaten_speichernUnter()
        'Abfrage zu neuem Projekt: Projektname und Dateiverzeichnis in neuem Fenster
        Dim WForm As New Form_Projekt_einrichten

        If Not IsNothing(stProjekt_Pfad) Then
            If stProjekt_Pfad <> "" Then
                Dim tmpPfad As String = Microsoft.VisualBasic.Left(stProjekt_Pfad, stProjekt_Pfad.Length - stProjekt_Name.Length - 1)
                If Right(tmpPfad, 1) <> "\" Then tmpPfad = tmpPfad & "\"

                WForm.TB_Projektverzeichnis.Text = tmpPfad

                Dim tmpName As String = stProjekt_Name

                Dim dio As IO.DirectoryInfo = New IO.DirectoryInfo(tmpPfad & tmpName)
                Dim i As Integer = 0
                Do Until dio.Exists = False
                    i = i + 1
                    dio = New IO.DirectoryInfo(tmpPfad & tmpName & "(" & i & ")")
                Loop

                WForm.TB_Projektname.Text = tmpName & "(" & i & ")"

            End If
        End If

        WForm.Text = "Speichern Unter ..."
        Dim resultDlg As DialogResult = WForm.ShowDialog
        If resultDlg = System.Windows.Forms.DialogResult.OK And stProjekt_Name <> "" And stProjekt_Pfad <> "" Then
            'Es ist ein neues Projekt erzeugt
            Projekt.bStatus_Projekt = True
            Form_Main.Text = "Schallschutzausweis 7Label 2018- " + stProjekt_Name '+ " - PMKG"    'Projekt.Name + " - PMKG"
            Projektdaten_speichern()
            'Projekt.bStatus_Lageplan = True
        End If
    End Sub
    Public Sub Projektdaten_speichern()
        Dim stFilename As String
        Dim RLen As Short
        Dim iDSLaenge As Integer

        Form_Main.Cursor = Cursors.WaitCursor

        'Projektdaten in Datei speichern
        '===============================
        Dim fil As New IO.FileInfo(stProjekt_Pfad & "\" & stProjekt_Name & ".PRJ")
        If fil.Exists Then fil.Delete()

        Dim ProjektRecord As ProjektData
        ProjektRecord = Nothing

        iDSLaenge = 0

        RLen = CShort(Len(CType(ProjektRecord, ProjektData)))
        stFilename = stProjekt_Pfad & "\" & stProjekt_Name & ".PRJ"

        FileOpen(1, stFilename, OpenMode.Random, , , RLen)

        ProjektRecord = Projekt
     

        iDSLaenge = CInt(iDSLaenge + 1)
        FilePut(1, ProjektRecord, 1) 'iDSLaenge)

        FileClose(1)

        'PB_Grundriss.backgroundImage.Save(stProjekt_Pfad & "\Daten\Grundriss.jpg", System.Drawing.Imaging.ImageFormat.Jpeg)
        If Not IsNothing(Form_Main.PB_Grundriss.BackgroundImage) Then
            Grundriss_speichern()
        Else
            ''Ermitteln, ob die Datei "Grundriss" im Projektordner schon existiert
            Dim filGR As New IO.FileInfo(stProjekt_Pfad & "\Daten\Grundriss.jpg")
            If filGR.Exists Then filGR.Delete()
        End If

        db_Antragsteller_aufnehmen()

        Form_Main.Cursor = Cursors.Default
    End Sub

    Public Sub db_Antragsteller_aufnehmen()
        'Prüfen, ob der Antragsteller schon vorhanden
        Dim bExist As Boolean = False

        With Projekt.Antragsteller
            If .Name = "" And .Zusatz = "" And .Strasse = "" And .Nr = "" And .PLZ = "" And .Ort = "" Then
                bExist = True
            Else
                If Not IsNothing(DB_Antragsteller) Then
                    For i As Integer = 0 To DB_Antragsteller.Count - 1
                        If CType(DB_Antragsteller(i), AdresseData).Name = .Name And CType(DB_Antragsteller(i), AdresseData).Zusatz = .Zusatz And CType(DB_Antragsteller(i), AdresseData).Strasse = .Strasse _
                            And CType(DB_Antragsteller(i), AdresseData).Nr = .Nr And CType(DB_Antragsteller(i), AdresseData).PLZ = .PLZ And CType(DB_Antragsteller(i), AdresseData).Ort = .Ort Then bExist = True
                    Next
                End If
            End If
        End With

        If bExist = False Then
          
            If DB_Antragsteller Is Nothing Then
                DB_Antragsteller = New Collections.ArrayList
                DB_Antragsteller.Add(Projekt.Antragsteller)
            Else

                DB_Antragsteller.Insert(0, Projekt.Antragsteller) 'Projekt.Antragsteller)
                If DB_Antragsteller.Count = 31 Then DB_Antragsteller.RemoveAt(30)
            End If
          
        End If

        My.Settings.Save()

    End Sub

    Public Sub Projektdaten_einlesen()

        Try
            Dim stFilename As String
            Dim RLen As Short

            'Projektdaten aus Datei einlesen
            '===============================

            Dim ProjektRecord As ProjektData
            ProjektRecord = Nothing
            Dim ProjektbRecord As System.ValueType = ProjektRecord

            RLen = CShort(Len(CType(ProjektRecord, ProjektData)))
            stFilename = stProjekt_Pfad & "\" & stProjekt_Name & ".PRJ"
            FileOpen(1, stFilename, OpenMode.Random, , , RLen)

            Do While Not EOF(1)
                Try
                    FileGet(1, ProjektbRecord)

                    ProjektRecord = CType(ProjektbRecord, ProjektData)
                    Projekt = ProjektRecord

                    Projekt.Name = Trim(CStr(ProjektRecord.Name))
                    Projekt.Pfad = Trim(CStr(ProjektRecord.Pfad))
                    Projekt.Grundriss = Trim(CStr(ProjektRecord.Grundriss))

                    Projekt.bStatus_Grundriss = ProjektRecord.bStatus_Grundriss
                    Projekt.bStatus_Projekt = True

                    Projekt.Datum = Trim(ProjektRecord.Datum)

                    Projekt.Antragsteller.Name = Trim(ProjektRecord.Antragsteller.Name)
                    Projekt.Antragsteller.Zusatz = Trim(ProjektRecord.Antragsteller.Zusatz)
                    Projekt.Antragsteller.Strasse = Trim(ProjektRecord.Antragsteller.Strasse)
                    Projekt.Antragsteller.Nr = Trim(ProjektRecord.Antragsteller.Nr)
                    Projekt.Antragsteller.PLZ = Trim(ProjektRecord.Antragsteller.PLZ)
                    Projekt.Antragsteller.Ort = Trim(ProjektRecord.Antragsteller.Ort)

                    Projekt.Gebaeude.Baujahr = ProjektRecord.Gebaeude.Baujahr
                    Projekt.Gebaeude.Gebaeudetyp = Trim(ProjektRecord.Gebaeude.Gebaeudetyp)
                    Projekt.Gebaeude.Kosten = ProjektRecord.Gebaeude.Kosten
                    Projekt.Gebaeude.Wohneinheiten = Trim(ProjektRecord.Gebaeude.Wohneinheiten)
                    Projekt.Gebaeude.Adresse.Name = Trim(ProjektRecord.Gebaeude.Adresse.Name)
                    Projekt.Gebaeude.Adresse.Zusatz = Trim(ProjektRecord.Gebaeude.Adresse.Zusatz)
                    Projekt.Gebaeude.Adresse.Strasse = Trim(ProjektRecord.Gebaeude.Adresse.Strasse)
                    Projekt.Gebaeude.Adresse.Nr = Trim(ProjektRecord.Gebaeude.Adresse.Nr)
                    Projekt.Gebaeude.Adresse.PLZ = Trim(ProjektRecord.Gebaeude.Adresse.PLZ)
                    Projekt.Gebaeude.Adresse.Ort = Trim(ProjektRecord.Gebaeude.Adresse.Ort)

                    Projekt.Wohnung = ProjektRecord.Wohnung
                    Projekt.Standort = ProjektRecord.Standort

                    Projekt.LS_Wand = ProjektRecord.LS_Wand
                    Projekt.LS_Wand.Prognose.C = Trim(ProjektRecord.LS_Wand.Prognose.C)
                    Projekt.LS_Wand.Messung.Anzahl = Trim(ProjektRecord.LS_Wand.Messung.Anzahl)
                    Projekt.LS_Wand.Messung.Messung_1.C = Trim(Projekt.LS_Wand.Messung.Messung_1.C)
                    Projekt.LS_Wand.Messung.Messung_2.C = Trim(Projekt.LS_Wand.Messung.Messung_2.C)
                    Projekt.LS_Wand.Messung.Messung_3.C = Trim(Projekt.LS_Wand.Messung.Messung_3.C)
                    Projekt.LS_Wand.Messung.Messung_4.C = Trim(Projekt.LS_Wand.Messung.Messung_4.C)
                    Projekt.LS_Wand.Messung.Messung_5.C = Trim(Projekt.LS_Wand.Messung.Messung_5.C)
                    Projekt.LS_Wand.Messung.Messung_6.C = Trim(Projekt.LS_Wand.Messung.Messung_6.C)

                    Projekt.LS_Decke = ProjektRecord.LS_Decke
                    Projekt.LS_Decke.Prognose.C = Trim(ProjektRecord.LS_Decke.Prognose.C)
                    Projekt.LS_Decke.Messung.Anzahl = Trim(ProjektRecord.LS_Decke.Messung.Anzahl)
                    Projekt.LS_Decke.Messung.Messung_1.C = Trim(Projekt.LS_Decke.Messung.Messung_1.C)
                    Projekt.LS_Decke.Messung.Messung_2.C = Trim(Projekt.LS_Decke.Messung.Messung_2.C)
                    Projekt.LS_Decke.Messung.Messung_3.C = Trim(Projekt.LS_Decke.Messung.Messung_3.C)
                    Projekt.LS_Decke.Messung.Messung_4.C = Trim(Projekt.LS_Decke.Messung.Messung_4.C)
                    Projekt.LS_Decke.Messung.Messung_5.C = Trim(Projekt.LS_Decke.Messung.Messung_5.C)
                    Projekt.LS_Decke.Messung.Messung_6.C = Trim(Projekt.LS_Decke.Messung.Messung_6.C)

                    Projekt.TS_Decke = ProjektRecord.TS_Decke
                    Projekt.TS_Decke.Prognose.Bodenbelag = Trim(ProjektRecord.TS_Decke.Prognose.Bodenbelag)
                    Projekt.TS_Decke.Prognose.Pegel.C = Trim(ProjektRecord.TS_Decke.Prognose.Pegel.C)
                    Projekt.TS_Decke.Messung.Anzahl = Trim(ProjektRecord.TS_Decke.Messung.Anzahl)
                    Projekt.TS_Decke.Messung.Messung_1.Pegel.C = Trim(Projekt.TS_Decke.Messung.Messung_1.Pegel.C)
                    Projekt.TS_Decke.Messung.Messung_2.Pegel.C = Trim(Projekt.TS_Decke.Messung.Messung_2.Pegel.C)
                    Projekt.TS_Decke.Messung.Messung_3.Pegel.C = Trim(Projekt.TS_Decke.Messung.Messung_3.Pegel.C)
                    Projekt.TS_Decke.Messung.Messung_4.Pegel.C = Trim(Projekt.TS_Decke.Messung.Messung_4.Pegel.C)
                    Projekt.TS_Decke.Messung.Messung_5.Pegel.C = Trim(Projekt.TS_Decke.Messung.Messung_5.Pegel.C)
                    Projekt.TS_Decke.Messung.Messung_6.Pegel.C = Trim(Projekt.TS_Decke.Messung.Messung_6.Pegel.C)

                    Projekt.TS_TPLH = ProjektRecord.TS_TPLH
                    Projekt.TS_TPLH.Prognose.Treppe.Pegel.C = Trim(ProjektRecord.TS_TPLH.Prognose.Treppe.Pegel.C)
                    Projekt.TS_TPLH.Prognose.Podest.Pegel.C = Trim(ProjektRecord.TS_TPLH.Prognose.Podest.Pegel.C)
                    Projekt.TS_TPLH.Prognose.Laube.Pegel.C = Trim(ProjektRecord.TS_TPLH.Prognose.Laube.Pegel.C)
                    Projekt.TS_TPLH.Prognose.Hausflur.Pegel.C = Trim(ProjektRecord.TS_TPLH.Prognose.Hausflur.Pegel.C)
                    Projekt.TS_TPlH.Messung.Messung_1.Pegel.C = Trim(Projekt.TS_TPlH.Messung.Messung_1.Pegel.C)
                    Projekt.TS_TPlH.Messung.Messung_2.Pegel.C = Trim(Projekt.TS_TPlH.Messung.Messung_2.Pegel.C)
                    Projekt.TS_TPlH.Messung.Messung_3.Pegel.C = Trim(Projekt.TS_TPlH.Messung.Messung_3.Pegel.C)
                    Projekt.TS_TPlH.Messung.Messung_4.Pegel.C = Trim(Projekt.TS_TPlH.Messung.Messung_4.Pegel.C)
                    Projekt.TS_TPlH.Messung.Messung_5.Pegel.C = Trim(Projekt.TS_TPlH.Messung.Messung_5.Pegel.C)
                    Projekt.TS_TPlH.Messung.Messung_6.Pegel.C = Trim(Projekt.TS_TPlH.Messung.Messung_6.Pegel.C)

                    Projekt.TS_BLT = ProjektRecord.TS_BLT
                    Projekt.TS_BLT.Prognose.Balkon.Pegel.C = Trim(ProjektRecord.TS_BLT.Prognose.Balkon.Pegel.C)
                    Projekt.TS_BLT.Prognose.Loggia.Pegel.C = Trim(ProjektRecord.TS_BLT.Prognose.Loggia.Pegel.C)
                    Projekt.TS_BLT.Prognose.Terrasse.Pegel.C = Trim(ProjektRecord.TS_BLT.Prognose.Terrasse.Pegel.C)
                    Projekt.TS_BLT.Messung.Messung_1.Pegel.C = Trim(Projekt.TS_BLT.Messung.Messung_1.Pegel.C)
                    Projekt.TS_BLT.Messung.Messung_2.Pegel.C = Trim(Projekt.TS_BLT.Messung.Messung_2.Pegel.C)
                    Projekt.TS_BLT.Messung.Messung_3.Pegel.C = Trim(Projekt.TS_BLT.Messung.Messung_3.Pegel.C)
                    Projekt.TS_BLT.Messung.Messung_4.Pegel.C = Trim(Projekt.TS_BLT.Messung.Messung_4.Pegel.C)
                    Projekt.TS_BLT.Messung.Messung_5.Pegel.C = Trim(Projekt.TS_BLT.Messung.Messung_5.Pegel.C)
                    Projekt.TS_BLT.Messung.Messung_6.Pegel.C = Trim(Projekt.TS_BLT.Messung.Messung_6.Pegel.C)

                    Projekt.Tueren = ProjektRecord.Tueren
                    Projekt.Aussenbauteile = ProjektRecord.Aussenbauteile
                    Projekt.Wasser = ProjektRecord.Wasser
                    Projekt.NutzerKoerper = ProjektRecord.NutzerKoerper
                    Projekt.NutzerKoerper = ProjektRecord.NutzerKoerper
                    Projekt.Nachbarn = ProjektRecord.Nachbarn
                    Projekt.anordnungRaeume = ProjektRecord.anordnungRaeume
                    Projekt.lauteRaeume = ProjektRecord.lauteRaeume
                    Projekt.NHZ = ProjektRecord.NHZ
                    Projekt.eigenerWohnbereich = ProjektRecord.eigenerWohnbereich


                    Projekt.Bem_anordnungRaeume = Trim(Projekt.Bem_anordnungRaeume)
                    Projekt.Bem_Aussenbauteile = Trim(Projekt.Bem_Aussenbauteile)
                    Projekt.Bem_lauteRaeume = Trim(Projekt.Bem_lauteRaeume)
                    Projekt.Bemerkung_eigenerWohnbereich = Trim(Projekt.Bemerkung_eigenerWohnbereich)
                    Projekt.Bemerkung_Nachbarn = Trim(Projekt.Bemerkung_Nachbarn)

                    Projekt.NutzerKoerper.BemerkungKoerper = Trim(Projekt.NutzerKoerper.BemerkungKoerper)
                    Projekt.LS_Decke.Bemerkung_LS = Trim(Projekt.LS_Decke.Bemerkung_LS)
                    Projekt.LS_Wand.Bemerkung_LS = Trim(Projekt.LS_Wand.Bemerkung_LS)
                    Projekt.NutzerKoerper.BemerkungNutzer = Trim(Projekt.NutzerKoerper.BemerkungNutzer)
                    Projekt.Standort.Bem_Aussenlaermpegel = Trim(Projekt.Standort.Bem_Aussenlaermpegel)
                    Projekt.Standort.Bem_Gebietscharakter = Trim(Projekt.Standort.Bem_Gebietscharakter)
                    Projekt.TS_BLT.Bemerkung_TS_BLLT = Trim(Projekt.TS_BLT.Bemerkung_TS_BLLT)
                    Projekt.TS_Decke.Bemerkung_TS_Decke = Trim(Projekt.TS_Decke.Bemerkung_TS_Decke)
                    Projekt.TS_TPLH.Bemerkung_TS_TPH = Trim(Projekt.TS_TPLH.Bemerkung_TS_TPH)
                    Projekt.Tueren.Bemerkung_Tueren = Trim(Projekt.Tueren.Bemerkung_Tueren)
                    Projekt.Wasser.Bem_Wasser = Trim(Projekt.Wasser.Bem_Wasser)

                Catch ex As Exception
                    FileClose(1)
                    MsgBox("Fehler beim Einlesen der Projektdaten. Das Projekt wurde vermutlich von einer älteren Version erzeugt!", MsgBoxStyle.OkOnly, "Fehlermeldung")
                    'Projektdaten_bisV20140818_einlesen()
                End Try
            Loop

            FileClose(1)


            'PB_Grundriss.backgroundImage.Save(stProjekt_Pfad & "\Daten\Grundriss.jpg", System.Drawing.Imaging.ImageFormat.Jpeg)
            Grundriss_einlesen()

        Catch ex As Exception
            Form_Main.Cursor = Cursors.Default
            MsgBox("Fehler beim Einlesen der Projektdaten. Das Projekt wurde vermutlich von einer älteren Version erzeugt!", MsgBoxStyle.OkOnly, "Fehlermeldung")
        End Try

    End Sub
    Public Sub Projektdaten_XLS_einlesen(ByVal xlsPfad As String)
        Loesche_Projektdaten()

        Dim sPf() As String = Split(xlsPfad, Chr(92))
        stProjekt_Name = Split(sPf(sPf.Length - 1), Chr(46))(0)
        stProjekt_Pfad = ""

        Projekt.Name = stProjekt_Name
        Projekt.Pfad = ""

        Projekt.bStatus_Projekt = True

        Dim xlApp As Application
        Dim xlMappe As Workbook
        Dim xlBlatt As Worksheet

        xlApp = CType(CreateObject("Excel.Application"), Application)
        'xlsPfad = Left(xlsPfad, xlsPfad.Length - 4)
        xlMappe = CType(xlApp.Workbooks.Open(xlsPfad, UpdateLinks:=True, ReadOnly:=True), Workbook)   'sPathXLDB & "Muster_Export.xls"

        Dim bExist As Boolean = False
        For Each ws As Worksheet In xlMappe.Worksheets
            If ws.Name = "Eingabeblatt" Then bExist = True
        Next ws
        If bExist Then
            xlBlatt = CType(xlMappe.Worksheets("Eingabeblatt"), Worksheet)

            With Projekt
                If IsDate(xlBlatt.Cells(6, 11).value) Then .Datum = xlBlatt.Cells(6, 11).value
            End With
            With Projekt.Antragsteller
                If Not IsNothing(xlBlatt.Cells(13, 6).value) Then .Name = xlBlatt.Cells(13, 6).value.ToString
                If Not IsNothing(xlBlatt.Cells(14, 6).value) Then .Zusatz = xlBlatt.Cells(14, 6).value.ToString
                If Not IsNothing(xlBlatt.Cells(15, 6).value) Then .Strasse = xlBlatt.Cells(15, 6).value.ToString
                If Not IsNothing(xlBlatt.Cells(16, 6).value) Then If IsNumeric(xlBlatt.Cells(16, 6).value) Then .PLZ = xlBlatt.Cells(16, 6).value.ToString
                If Not IsNothing(xlBlatt.Cells(16, 8).value) Then .Ort = xlBlatt.Cells(16, 8).value.ToString
            End With
            With Projekt.Gebaeude.Adresse
                If Not IsNothing(xlBlatt.Cells(20, 6).value) Then .Name = xlBlatt.Cells(20, 6).value.ToString
                If Not IsNothing(xlBlatt.Cells(21, 6).value) Then .Zusatz = xlBlatt.Cells(21, 6).value.ToString
                If Not IsNothing(xlBlatt.Cells(22, 6).value) Then .Strasse = xlBlatt.Cells(22, 6).value.ToString
                If Not IsNothing(xlBlatt.Cells(23, 6).value) Then If IsNumeric(xlBlatt.Cells(23, 6).value) Then .PLZ = xlBlatt.Cells(23, 6).value.ToString
                If Not IsNothing(xlBlatt.Cells(34, 8).value) Then .Ort = xlBlatt.Cells(34, 8).value.ToString
            End With
            With Projekt.Gebaeude
                If IsNumeric(xlBlatt.Cells(25, 6).value) Then .Baujahr = CInt(xlBlatt.Cells(25, 6).value)
                If Not IsNothing(xlBlatt.Cells(24, 6).value) Then .Gebaeudetyp = xlBlatt.Cells(24, 6).value.ToString
                If Not IsNothing(xlBlatt.Cells(26, 6).value) Then .Wohneinheiten = xlBlatt.Cells(26, 6).value.ToString
                If IsNumeric(xlBlatt.Cells(35, 6).value) Then .Kosten = CSng(xlBlatt.Cells(35, 6).value)
            End With
            With Projekt.Wohnung
                If Not IsNothing(xlBlatt.Cells(32, 9).value) Then .Geschoss.Bezeichnung = xlBlatt.Cells(32, 9).value.ToString
                If IsNumeric(xlBlatt.Cells(32, 8).value) Then .Geschoss.Nr = CInt(xlBlatt.Cells(32, 8).value)
                If Not IsNothing(xlBlatt.Cells(32, 6).value) Then .Geschoss.Typ = Get_Byte_Geschosstyp(xlBlatt.Cells(32, 6).value.ToString)
                If Not IsNothing(xlBlatt.Cells(33, 6).value) Then .Raeume = xlBlatt.Cells(33, 6).value.ToString
                If Not IsNothing(xlBlatt.Cells(34, 6).value) Then If IsNumeric(xlBlatt.Cells(34, 6).value) Then .Wohnflaeche = xlBlatt.Cells(34, 6).value.ToString
                If Not IsNothing(xlBlatt.Cells(31, 6).value) Then .Wohnungsbezeichnung = xlBlatt.Cells(31, 6).value.ToString
            End With
            With Projekt.Standort
                .Gebietscharakter = Get_Byte_GC(xlBlatt.Cells(41, 7).value)
                .Aussenlaermpegel = Get_Byte_AP(xlBlatt.Cells(43, 7).value)
                .abgewandFreibereich = Get_Byte_aF(xlBlatt.Cells(45, 7).value)
            End With

            'Ls
            Projekt.LS_Wand.Untersuchung = Get_Byte_Untersuchung(xlBlatt.Cells(49, 7).value)
            With Projekt.LS_Wand.Messung
                .Anzahl = 4
            End With
            With Projekt.LS_Wand.Messung.Messung_1
                If Not IsNothing(xlBlatt.Cells(52, 8).value) Then If IsNumeric(xlBlatt.Cells(52, 8).value) Then .Pegel = xlBlatt.Cells(52, 8).value.ToString
                If Not IsNothing(xlBlatt.Cells(53, 8).value) Then .C = xlBlatt.Cells(53, 8).value.ToString
            End With
            With Projekt.LS_Wand.Messung.Messung_2
                If Not IsNothing(xlBlatt.Cells(52, 9).value) Then If IsNumeric(xlBlatt.Cells(52, 9).value) Then .Pegel = xlBlatt.Cells(52, 9).value.ToString
                If Not IsNothing(xlBlatt.Cells(53, 9).value) Then .C = xlBlatt.Cells(53, 9).value.ToString
            End With
            With Projekt.LS_Wand.Messung.Messung_3
                If Not IsNothing(xlBlatt.Cells(52, 10).value) Then If IsNumeric(xlBlatt.Cells(52, 10).value) Then .Pegel = xlBlatt.Cells(52, 10).value.ToString
                If Not IsNothing(xlBlatt.Cells(53, 10).value) Then .C = xlBlatt.Cells(53, 10).value.ToString
            End With
            With Projekt.LS_Wand.Messung.Messung_4
                If Not IsNothing(xlBlatt.Cells(52, 11).value) Then If IsNumeric(xlBlatt.Cells(52, 11).value) Then .Pegel = xlBlatt.Cells(52, 11).value.ToString
                If Not IsNothing(xlBlatt.Cells(53, 11).value) Then .C = xlBlatt.Cells(53, 11).value.ToString
            End With
            With Projekt.LS_Wand.Messung.Messung_5
                If Not IsNothing(xlBlatt.Cells(52, 12).value) Then If IsNumeric(xlBlatt.Cells(52, 12).value) Then .Pegel = xlBlatt.Cells(52, 12).value.ToString
                If Not IsNothing(xlBlatt.Cells(53, 12).value) Then .C = xlBlatt.Cells(53, 12).value.ToString
            End With
            With Projekt.LS_Wand.Prognose
                .Pegel = CSng(xlBlatt.Cells(54, 8).value)
                If Not IsNothing(xlBlatt.Cells(55, 8).value) Then .C = xlBlatt.Cells(55, 8).value.ToString
            End With

            Projekt.LS_Decke.Untersuchung = Get_Byte_Untersuchung(xlBlatt.Cells(57, 7).value)
            With Projekt.LS_Decke.Messung
                .Anzahl = 4
            End With
            With Projekt.LS_Decke.Messung.Messung_1
                If Not IsNothing(xlBlatt.Cells(60, 8).value) Then If IsNumeric(xlBlatt.Cells(60, 8).value) Then .Pegel = xlBlatt.Cells(60, 8).value.ToString
                If Not IsNothing(xlBlatt.Cells(61, 8).value) Then .C = xlBlatt.Cells(61, 8).value.ToString
            End With
            With Projekt.LS_Decke.Messung.Messung_2
                If Not IsNothing(xlBlatt.Cells(60, 9).value) Then If IsNumeric(xlBlatt.Cells(60, 9).value) Then .Pegel = xlBlatt.Cells(60, 9).value.ToString
                If Not IsNothing(xlBlatt.Cells(61, 9).value) Then .C = xlBlatt.Cells(61, 9).value.ToString
            End With
            With Projekt.LS_Decke.Messung.Messung_3
                If Not IsNothing(xlBlatt.Cells(60, 10).value) Then If IsNumeric(xlBlatt.Cells(60, 10).value) Then .Pegel = xlBlatt.Cells(60, 10).value.ToString
                If Not IsNothing(xlBlatt.Cells(61, 10).value) Then .C = xlBlatt.Cells(61, 10).value.ToString
            End With
            With Projekt.LS_Decke.Messung.Messung_4
                If Not IsNothing(xlBlatt.Cells(60, 11).value) Then If IsNumeric(xlBlatt.Cells(60, 11).value) Then .Pegel = xlBlatt.Cells(60, 11).value.ToString
                If Not IsNothing(xlBlatt.Cells(61, 11).value) Then .C = xlBlatt.Cells(61, 11).value.ToString
            End With
            With Projekt.LS_Decke.Messung.Messung_5
                If Not IsNothing(xlBlatt.Cells(60, 12).value) Then If IsNumeric(xlBlatt.Cells(60, 12).value) Then .Pegel = xlBlatt.Cells(60, 12).value.ToString
                If Not IsNothing(xlBlatt.Cells(61, 12).value) Then .C = xlBlatt.Cells(61, 12).value.ToString
            End With
            With Projekt.LS_Decke.Prognose
                .Pegel = CSng(xlBlatt.Cells(62, 8).value)
                If Not IsNothing(xlBlatt.Cells(63, 8).value) Then .C = xlBlatt.Cells(63, 8).value.ToString
            End With

            Projekt.TS_Decke.Untersuchung = Get_Byte_Untersuchung(xlBlatt.Cells(66, 7).value)
            With Projekt.TS_Decke.Messung
                .Anzahl = 4
            End With
            With Projekt.TS_Decke.Messung.Messung_1
                If Not IsNothing(xlBlatt.Cells(69, 8).value) Then If IsNumeric(xlBlatt.Cells(69, 8).value) Then .Pegel.Pegel = xlBlatt.Cells(69, 8).value.ToString
                If Not IsNothing(xlBlatt.Cells(70, 8).value) Then .Pegel.C = xlBlatt.Cells(70, 8).value.ToString
            End With
            With Projekt.TS_Decke.Messung.Messung_2
                If Not IsNothing(xlBlatt.Cells(69, 9).value) Then If IsNumeric(xlBlatt.Cells(69, 9).value) Then .Pegel.Pegel = xlBlatt.Cells(69, 9).value.ToString
                If Not IsNothing(xlBlatt.Cells(70, 9).value) Then .Pegel.C = xlBlatt.Cells(70, 9).value.ToString
            End With
            With Projekt.TS_Decke.Messung.Messung_3
                If Not IsNothing(xlBlatt.Cells(69, 10).value) Then If IsNumeric(xlBlatt.Cells(69, 10).value) Then .Pegel.Pegel = xlBlatt.Cells(69, 10).value.ToString
                If Not IsNothing(xlBlatt.Cells(70, 10).value) Then .Pegel.C = xlBlatt.Cells(70, 10).value.ToString
            End With
            With Projekt.TS_Decke.Messung.Messung_4
                If Not IsNothing(xlBlatt.Cells(69, 11).value) Then If IsNumeric(xlBlatt.Cells(69, 11).value) Then .Pegel.Pegel = xlBlatt.Cells(69, 11).value.ToString
                If Not IsNothing(xlBlatt.Cells(70, 11).value) Then .Pegel.C = xlBlatt.Cells(70, 11).value.ToString
            End With
            With Projekt.TS_Decke.Messung.Messung_5
                If Not IsNothing(xlBlatt.Cells(69, 12).value) Then If IsNumeric(xlBlatt.Cells(69, 12).value) Then .Pegel.Pegel = xlBlatt.Cells(69, 12).value.ToString
                If Not IsNothing(xlBlatt.Cells(70, 12).value) Then .Pegel.C = xlBlatt.Cells(70, 12).value.ToString
            End With
            With Projekt.TS_Decke.Prognose
                .Pegel.Pegel = CSng(xlBlatt.Cells(71, 8).value)
                If Not IsNothing(xlBlatt.Cells(72, 8).value) Then .Pegel.C = xlBlatt.Cells(72, 8).value.ToString
            End With

            Projekt.TS_TPLH.Untersuchung = Get_Byte_Untersuchung(xlBlatt.Cells(74, 7).value)
            Projekt.TS_TPLH.Messung.Anzahl = 4
            With Projekt.TS_TPLH.Messung.Messung_1
                .Bauteil = Get_Byte_Bauteil_TPH(xlBlatt.Cells(75, 8).value)
                If Not IsNothing(xlBlatt.Cells(77, 8).value) Then If IsNumeric(xlBlatt.Cells(77, 8).value) Then .Pegel.Pegel = xlBlatt.Cells(77, 8).value.ToString
                If Not IsNothing(xlBlatt.Cells(78, 8).value) Then .Pegel.C = xlBlatt.Cells(78, 8).value.ToString
            End With
            With Projekt.TS_TPLH.Messung.Messung_2
                .Bauteil = Get_Byte_Bauteil_TPH(xlBlatt.Cells(75, 9).value)
                If Not IsNothing(xlBlatt.Cells(77, 9).value) Then If IsNumeric(xlBlatt.Cells(77, 9).value) Then .Pegel.Pegel = xlBlatt.Cells(77, 9).value.ToString
                If Not IsNothing(xlBlatt.Cells(78, 9).value) Then .Pegel.C = xlBlatt.Cells(78, 9).value.ToString
            End With
            With Projekt.TS_TPLH.Messung.Messung_3
                .Bauteil = Get_Byte_Bauteil_TPH(xlBlatt.Cells(75, 10).value)
                If Not IsNothing(xlBlatt.Cells(77, 10).value) Then If IsNumeric(xlBlatt.Cells(77, 10).value) Then .Pegel.Pegel = xlBlatt.Cells(77, 10).value.ToString
                If Not IsNothing(xlBlatt.Cells(78, 10).value) Then .Pegel.C = xlBlatt.Cells(78, 10).value.ToString
            End With
            With Projekt.TS_TPLH.Messung.Messung_4
                .Bauteil = Get_Byte_Bauteil_TPH(xlBlatt.Cells(75, 11).value)
                If Not IsNothing(xlBlatt.Cells(77, 11).value) Then If IsNumeric(xlBlatt.Cells(77, 11).value) Then .Pegel.Pegel = xlBlatt.Cells(77, 11).value.ToString
                If Not IsNothing(xlBlatt.Cells(78, 11).value) Then .Pegel.C = xlBlatt.Cells(78, 11).value.ToString
            End With
            With Projekt.TS_TPLH.Messung.Messung_5
                .Bauteil = Get_Byte_Bauteil_TPH(xlBlatt.Cells(75, 12).value)
                If Not IsNothing(xlBlatt.Cells(77, 12).value) Then If IsNumeric(xlBlatt.Cells(77, 12).value) Then .Pegel.Pegel = xlBlatt.Cells(77, 12).value.ToString
                If Not IsNothing(xlBlatt.Cells(78, 12).value) Then .Pegel.C = xlBlatt.Cells(78, 12).value.ToString
            End With
            With Projekt.TS_TPLH.Prognose.Treppe
                .Pegel.Pegel = CSng(xlBlatt.Cells(80, 8).value)
                If Not IsNothing(xlBlatt.Cells(81, 8).value) Then .Pegel.C = xlBlatt.Cells(81, 8).value.ToString
            End With
            With Projekt.TS_TPLH.Prognose.Podest
                .Pegel.Pegel = CSng(xlBlatt.Cells(80, 9).value)
                If Not IsNothing(xlBlatt.Cells(81, 9).value) Then .Pegel.C = xlBlatt.Cells(81, 9).value.ToString
            End With
            With Projekt.TS_TPLH.Prognose.Hausflur
                .Pegel.Pegel = CSng(xlBlatt.Cells(80, 10).value)
                If Not IsNothing(xlBlatt.Cells(81, 10).value) Then .Pegel.C = xlBlatt.Cells(81, 10).value.ToString
            End With

            Projekt.TS_BLT.Untersuchung = Get_Byte_Untersuchung(xlBlatt.Cells(83, 7).value)
            Projekt.TS_BLT.Messung.Anzahl = 4
            With Projekt.TS_BLT.Messung.Messung_1
                .Bauteil = Get_Byte_Bauteil_BLLT(xlBlatt.Cells(84, 8).value)
                If Not IsNothing(xlBlatt.Cells(86, 8).value) Then If IsNumeric(xlBlatt.Cells(86, 8).value) Then .Pegel.Pegel = xlBlatt.Cells(86, 8).value.ToString
                If Not IsNothing(xlBlatt.Cells(87, 8).value) Then .Pegel.C = xlBlatt.Cells(87, 8).value.ToString
            End With
            With Projekt.TS_BLT.Messung.Messung_2
                .Bauteil = Get_Byte_Bauteil_BLLT(xlBlatt.Cells(84, 9).value)
                If Not IsNothing(xlBlatt.Cells(86, 9).value) Then If IsNumeric(xlBlatt.Cells(86, 9).value) Then .Pegel.Pegel = xlBlatt.Cells(86, 9).value.ToString
                If Not IsNothing(xlBlatt.Cells(87, 9).value) Then .Pegel.C = xlBlatt.Cells(87, 9).value.ToString
            End With
            With Projekt.TS_BLT.Messung.Messung_3
                .Bauteil = Get_Byte_Bauteil_BLLT(xlBlatt.Cells(84, 10).value)
                If Not IsNothing(xlBlatt.Cells(86, 10).value) Then If IsNumeric(xlBlatt.Cells(86, 10).value) Then .Pegel.Pegel = xlBlatt.Cells(86, 10).value.ToString
                If Not IsNothing(xlBlatt.Cells(87, 10).value) Then .Pegel.C = xlBlatt.Cells(87, 10).value.ToString
            End With
            With Projekt.TS_BLT.Messung.Messung_4
                .Bauteil = Get_Byte_Bauteil_BLLT(xlBlatt.Cells(84, 11).value)
                If Not IsNothing(xlBlatt.Cells(86, 11).value) Then If IsNumeric(xlBlatt.Cells(86, 11).value) Then .Pegel.Pegel = xlBlatt.Cells(86, 11).value.ToString
                If Not IsNothing(xlBlatt.Cells(87, 11).value) Then .Pegel.C = xlBlatt.Cells(87, 11).value.ToString
            End With
            With Projekt.TS_BLT.Messung.Messung_5
                .Bauteil = Get_Byte_Bauteil_BLLT(xlBlatt.Cells(84, 12).value)
                If Not IsNothing(xlBlatt.Cells(86, 12).value) Then If IsNumeric(xlBlatt.Cells(86, 12).value) Then .Pegel.Pegel = xlBlatt.Cells(86, 12).value.ToString
                If Not IsNothing(xlBlatt.Cells(87, 12).value) Then .Pegel.C = xlBlatt.Cells(87, 12).value.ToString
            End With
            With Projekt.TS_BLT.Prognose.Balkon
                .Pegel.Pegel = CSng(xlBlatt.Cells(89, 8).value)
                If Not IsNothing(xlBlatt.Cells(90, 8).value) Then .Pegel.C = xlBlatt.Cells(90, 8).value.ToString
            End With
            With Projekt.TS_TPLH.Prognose.Laube
                .Pegel.Pegel = CSng(xlBlatt.Cells(89, 9).value)
                If Not IsNothing(xlBlatt.Cells(90, 9).value) Then .Pegel.C = xlBlatt.Cells(90, 9).value.ToString
            End With
            With Projekt.TS_BLT.Prognose.Loggia
                .Pegel.Pegel = CSng(xlBlatt.Cells(89, 10).value)
                If Not IsNothing(xlBlatt.Cells(90, 10).value) Then .Pegel.C = xlBlatt.Cells(90, 10).value.ToString
            End With
            With Projekt.TS_BLT.Prognose.Terrasse
                .Pegel.Pegel = CSng(xlBlatt.Cells(89, 11).value)
                If Not IsNothing(xlBlatt.Cells(90, 11).value) Then .Pegel.C = xlBlatt.Cells(90, 11).value.ToString
            End With

            With Projekt.Tueren
                .Ort = Get_Byte_Tueren_Ort(xlBlatt.Cells(93, 7).value)
                .Untersuchung = Get_Byte_Tueren_Untersuchung(xlBlatt.Cells(94, 7).value)
                .L = CSng(xlBlatt.Cells(95, 8).value)
            End With
            Projekt.Aussenbauteile = Get_Byte_Aussenbauteile(xlBlatt.Cells(97, 7).value)
            With Projekt.Wasser
                .Untersuchung = Get_Byte_Untersuchung(xlBlatt.Cells(99, 7).value)
                .L_Intervall = Get_Byte_L_Intervall_Wasser(xlBlatt.Cells(100, 8).value)
                .LcLa_erfuellt = Get_Byte_Wasser_LcLa(xlBlatt.Cells(101, 8).value)
            End With
            'With Projekt.Nutzer
            '    .Untersuchung = Get_Byte_Untersuchung(xlBlatt.Cells(103, 7).value)
            '    .L_Intervall = Get_Byte_L_Intervall_Nutzer(xlBlatt.Cells(104, 8).value)
            'End With
            'With Projekt.Koerper
            '    .Untersuchung = Get_Byte_Untersuchung(xlBlatt.Cells(105, 7).value)
            '    .L_Intervall = Get_Byte_L_Intervall_Koerper(xlBlatt.Cells(106, 8).value)
            'End With
            Projekt.Nachbarn = Get_Byte_Nachbarn(xlBlatt.Cells(108, 7).value)
            Projekt.anordnungRaeume = Get_Byte_AnordRaeume(xlBlatt.Cells(110, 7).value)
            Projekt.lauteRaeume = Get_Byte_lauteRaeume(xlBlatt.Cells(112, 7).value)
            Projekt.eigenerWohnbereich = Get_Byte_EW(xlBlatt.Cells(114, 7).value)
        Else
            MsgBox("Das Eingabeblatt konnte nicht gefunden werden.", MsgBoxStyle.OkOnly, "Fehlermeldung")
        End If
        'Excel1.ActiveWorkbook.Close()
        xlApp.ActiveWorkbook.Close(False)
        'xlMappe.Close()

        'Excel1.Visible = False
    End Sub

    Public Sub Grundriss_speichern()

        'Form_Main.PB_Grundriss.backgroundImage.Save(stProjekt_Pfad & "\Daten\Grundriss.jpg", System.Drawing.Imaging.ImageFormat.Jpeg)
        ' stProjekt_Pfad = "G:\7\70\707\7071\SSA"
        Try
            If Not IsNothing(btmGrundriss) Then btmGrundriss.Save(stProjekt_Pfad & "\Daten\Grundriss.jpg", Imaging.ImageFormat.Jpeg) ', Imaging.ImageFormat.Jpeg)
        Catch ex As Exception
            Dim BExc As System.Exception = ex.GetBaseException
            Dim str As String = BExc.ToString
            Dim s As String = BExc.Message

        End Try


        Projekt.bStatus_Grundriss = True

    End Sub

    Public Sub Grundriss_einlesen()
        Dim fil As New IO.FileInfo(stProjekt_Pfad & "\Daten\Grundriss.jpg")
        If fil.Exists Then
            btmGrundriss = Nothing
            btmGrundriss = System.Drawing.Image.FromFile(stProjekt_Pfad & "\Daten\Grundriss.jpg")
            'Form_Main.PB_Grundriss.BackgroundImage = btmGrundriss 'System.Drawing.Image.FromFile(stProjekt_Pfad & "\Daten\Grundriss.jpg")
            'Form_Main.PB_DB_Grundriss.BackgroundImage = btmGrundriss
        End If
    End Sub
 
End Module
