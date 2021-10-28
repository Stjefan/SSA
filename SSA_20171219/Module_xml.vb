Module Module_xml


    '******************* upload - begin
    'Upload file using input type=file
    'UploadFile("www.schallschutzausweis.tu-bs.de/upload.php", str) '"c:\thefile.xml")

    Sub UploadFile(ByVal DestURL As String, ByVal FileName As String, _
      Optional ByVal FieldName As String = "thefile")
        Dim sFormData As String, d As String

        'Boundary of fields.
        'Be sure this string is Not In the source file
        Const Boundary As String = "---------------------------0123456789012"

        ''Get source file As a string.
        sFormData = FileName

        ''Build source form with file contents
        d = "--" + Boundary + vbCrLf
        d = d + "Content-Disposition: form-data; name=""" + FieldName + """;"
        d = d + " filename=' '" + vbCrLf
        d = d + "Content-Type: application/upload" + vbCrLf + vbCrLf
        d = d + sFormData
        d = d + vbCrLf + "--" + Boundary + "--" + vbCrLf

        'Post the data To the destination URL
        IEPostStringRequest(DestURL, d, Boundary)
    End Sub

    'sends URL encoded form data To the URL using IE
    Sub IEPostStringRequest(ByVal URL As String, ByVal FormData As String, ByVal Boundary As String)
        Dim WebBrowser As WebBrowser = Form_Webbrowser.WebBrowser1

        Form_Webbrowser.Show()

        'Send the form data To URL As POST request
        Dim instance As New System.Text.UnicodeEncoding
        Dim s As String = FormData
        Dim charCount As Integer = FormData.Length
        Dim bytes As Byte()
        ReDim bytes(Len(FormData) - 1)
       
        bytes = instance.GetBytes(s)

        Dim bFormData As Byte() = Nothing
        ReDim bFormData(CInt(Len(FormData) / 2) - 1)

        Dim bytesTest As Byte()
        ReDim bytesTest(CInt(bytes.Length / 2) - 1)

        Dim iBytes As Integer = 1
        Dim txt As String = ""
        For i As Integer = 0 To bytesTest.Length - 1
            bytesTest(i) = bytes(i * 2)
        Next

        WebBrowser.Navigate(URL, "", bytesTest, "Content-Type: multipart/form-data; boundary=" + Boundary + vbCrLf)
        WebBrowser.Navigate(URL, False)
    End Sub

    'read binary file As a string value
    Function GetFile(ByVal FileName As String) As String
        GetFile = My.Computer.FileSystem.ReadAllText("c:\thefile.xml")
    End Function
    '******************* upload - end

    Public Function Built_xml() As String

        Dim tmpStr As String
        Dim str As String

        With Projekt
            'xsi:schemaLocation="http://www.schallschutzausweis.tu-bs.de/gebaeude/1.0
            'http://www.schallschutzausweis.tu-bs.de/schema/ssaDega.xsd">

            Str = "<?xml version='1.0' encoding='utf-8'?>" '& vbCrLf 'iso-8859-1
            'str = str & "<gebaeude xmlns='http://www.schallschutzausweis.tu-bs.de/gebaeude/1.0' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'" & _
            '    " xsi:schemaLocation='http://www.schallschutzausweis.tu-bs.de/gebaeude/1.0 http://www.schallschutzausweis.tu-bs.de/schema/ssaDega.xsd'>"
            str = str & "<gebaeude xmlns='http://www.schallschutzausweis.de/gebaeude/1.0' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'" & _
                            " xsi:schemaLocation='http://www.schallschutzausweis.de/gebaeude/1.0 http://www.dega-schallschutzausweis.de/download/ssaDega.xsd'>"

            tmpStr = Get_str_PLZ
            If tmpStr = "" Then
                Built_xml = "Fehler: Die Daten können nicht übertragen werden. Bitte geben Sie eine gültige PLZ für das Gebäude ein."
                Exit Function
            End If
            str = str & tmpStr

            str = str & Get_str_Kosten

            tmpStr = Get_str_Baujahr
            If tmpStr = "" Then
                Built_xml = "Fehler: Die Daten können nicht übertragen werden. Bitte geben Sie ein gültiges Baujahr ein."
                Exit Function
            End If
            str = str & tmpStr

            tmpStr = Get_str_Gueltigkeit
            If tmpStr = "" Then
                Built_xml = "Fehler: Die Daten können nicht übertragen werden. Bitte geben Sie ein gültiges Datum ein."
                Exit Function
            End If
            str = str & tmpStr

            tmpStr = Get_str_Gebietscharakter
            If tmpStr = "" Then
                Built_xml = "Fehler: Die Daten können nicht übertragen werden. Bitte geben Sie gültige Daten für den Gebietscharakter ein."
                Exit Function
            End If
            str = str & tmpStr

            tmpStr = Get_str_Gebaeudetyp
            If tmpStr = "" Then
                Built_xml = "Fehler: Die Daten können nicht übertragen werden. Bitte geben Sie einen gültigen Gebäudetyp ein."
                Exit Function
            End If
            str = str & tmpStr

            tmpStr = Get_str_anzWohneinheiten
            If tmpStr = "" Then
                Built_xml = "Fehler: Die Daten können nicht übertragen werden. Bitte geben Sie einen gültigen Wert für die Anzahl der Wohneinheiten ein."
                Exit Function
            End If
            str = str & tmpStr

            str = str & "<wohneinheit>"

            tmpStr = Get_str_anzRaeume
            If tmpStr = "" Then
                Built_xml = "Fehler: Die Daten können nicht übertragen werden. Bitte geben Sie einen gültigen Wert für die Anzahl der Räume ein."
                Exit Function
            End If
            str = str & tmpStr

            tmpStr = Get_str_Stockwerk
            If tmpStr = "" Then
                Built_xml = "Fehler: Die Daten können nicht übertragen werden. Bitte geben Sie einen gültigen Wert für das Stockwerk ein."
                Exit Function
            End If
            str = str & tmpStr

            tmpStr = Get_str_Wohnflaeche
            If tmpStr = "" Then
                Built_xml = "Fehler: Die Daten können nicht übertragen werden. Bitte geben Sie einen gültigen Wert für die Wohnfläche ein."
                Exit Function
            End If
            str = str & tmpStr

            tmpStr = Get_str_Aussenlaermpegel
            If tmpStr = "" Then
                Built_xml = "Fehler: Die Daten können nicht übertragen werden. Bitte geben Sie gültige Werte für den Außenlärmsituation ein."
                Exit Function
            End If
            str = str & tmpStr

            '######## Treppen, Podeste, Hausflure
            str = str & Get_str_Treppen
            str = str & Get_str_Podest
            str = str & Get_str_Hausflur

            '#######    Balkone, Loggien, Laubengänge, Terrassen
            str = str & Get_str_Balkone
            str = str & Get_str_Laubengang
            str = str & Get_str_Loggie
            str = str & Get_str_Terrasse

            '#######    Wohnungseingangstüren
            str = str & Get_str_Tueren

            '#########  LuftschallAussenbauteile
            str = str & Get_str_LSAussenbauteile

            '#########  Wasserinstallation
            str = str & Get_str_Wasserinstallation

            'If Get_Pkte_Koerper() = "0" Then
            '#########  nutzerGeraeusche
            tmpStr = Get_str_Nutzergeraeusche()
            If tmpStr = "" Then
                Built_xml = "Fehler: Die Daten können nicht übertragen werden. Bitte geben Sie gültige Werte für die Nutzergeräusche ein."
                Exit Function
            End If
            str = str & tmpStr
            'Else
            '#########  koerperschallGeraeusche
            tmpStr = Get_str_Koerperschallentkopplung()
            If tmpStr = "" Then
                Built_xml = "Fehler: Die Daten können nicht übertragen werden. Bitte geben Sie gültige Werte für die Körperschallentkopplung ein."
                Exit Function
            End If
            str = str & tmpStr
            'End If

            '#########  grundrisssituation
            tmpStr = Get_str_Grundrisssituation()
            If tmpStr = "" Then
                Built_xml = "Fehler: Die Daten können nicht übertragen werden. Bitte geben Sie gültige Werte für 'fremde Nutzer direkt angrenzend', 'Anordnung lauter Räume schalltechnisch' und 'laute Räume gem. DIN 4109:1989-11 angrenzend' ein."
                Exit Function
            End If
            str = str & tmpStr

            '#########  eigenerWohnbereich
            tmpStr = Get_str_EW()
            If tmpStr = "" Then
                Built_xml = "Fehler: Die Daten können nicht übertragen werden. Bitte geben Sie gültige Werte für den eigenen Wohnbereich ein."
                Exit Function
            End If
            str = str & tmpStr

            Dim tmpRaum As String
            '#########  Raum - LS Wände
            tmpRaum = Get_str_LS_Wand()
            'str = str & Get_Ls_Waende

            '#########  Raum - LS / TS Decken
            tmpRaum = tmpRaum & Get_str_TS_Decken()
            tmpRaum = tmpRaum & Get_str_LS_Decke()
            'str = str & Get_str_TS_Decken
            'str = str & Get_str_LS_Decken
            If tmpRaum = "" Then tmpRaum = "<raum></raum>"

            str = str & tmpRaum

            str = str & "</wohneinheit>"

            str = str & "</gebaeude>" & vbCrLf & vbCrLf

        End With

        Built_xml = str
        Clipboard.SetText(Built_xml)
    End Function

    Public Function Get_str_PLZ() As String
        Dim str As String

        With Projekt

            str = "<plz>"

            If Len(.Gebaeude.Adresse.PLZ) = 5 Then '-> Wenn PLZ vorhanden, dann korrekt (keine weitere Überprüfung): And .Gebaeude.Adresse.PLZ > 9999 And .Gebaeude.Adresse.PLZ < 100000 Then
                str = str & .Gebaeude.Adresse.PLZ
            Else
                Get_str_PLZ = ""
                Exit Function
            End If
            str = str & "</plz>"

        End With

        Get_str_PLZ = str

    End Function

    Public Function Get_str_anzWohneinheiten() As String
        Dim str As String

        With Projekt

            str = "<anzWohneinheiten>"
            If IsNumeric(.Gebaeude.Wohneinheiten) Then
                If CDec(.Gebaeude.Wohneinheiten) >= 0 And CDec(.Gebaeude.Wohneinheiten) <= 100 Then
                    str = str & .Gebaeude.Wohneinheiten.ToString
                Else
                    Get_str_anzWohneinheiten = "" 'str & "0"
                    Exit Function
                End If
            Else
                Get_str_anzWohneinheiten = ""
                Exit Function
            End If
            str = str & "</anzWohneinheiten>"

        End With

        Get_str_anzWohneinheiten = str

    End Function

    Public Function Get_str_Wohnflaeche() As String
        Dim str As String

        With Projekt

            str = "<wohnflaeche>"
            If .Wohnung.Wohnflaeche > 1 Then
                str = str & CDbl(.Wohnung.Wohnflaeche).ToString
                str = str.Replace(",", ".")
                If InStr(str, ".") = 0 Then str = str & ".0"
            Else
                Get_str_Wohnflaeche = ""
                Exit Function
            End If
            str = str & "</wohnflaeche>"

        End With

        Get_str_Wohnflaeche = str
    End Function

    Public Function Get_str_anzRaeume() As String
        Dim str As String

        With Projekt.Wohnung

            str = "<anzRaeume>"
            If IsNumeric(.Raeume) Then
                If CDec(.Raeume) >= 0 And CDec(.Raeume) <= 100 Then
                    str = str & .Raeume
                Else
                    Get_str_anzRaeume = ""
                    Exit Function
                End If
            Else
                Get_str_anzRaeume = ""
                Exit Function
            End If
            str = str & "</anzRaeume>"

        End With

        Get_str_anzRaeume = str
    End Function

    Public Function Get_str_TS_Decken() As String
        Dim str As String = ""
        Dim iMes As Integer
        Dim lnw As Single
        Dim C As String = ""

        With Projekt.TS_Decke

            If .Untersuchung = Messung Then
                Dim bGueltig As Boolean
                bGueltig = False

                For iMes = 0 To 5
                    'Dim sMA As String
                    'sMA = ""
                    'If .Messung.Messanteil = groesser50 Then sMA = "+50"
                    'Dim MV As String = ""
                    If iMes = 0 Then
                        'MV = Get_Str_MV(.Messung.Messung_1.Messverfahren)
                        lnw = .Messung.Messung_1.Pegel.Pegel
                        C = .Messung.Messung_1.Pegel.C
                    ElseIf iMes = 1 Then
                        'MV = Get_Str_MV(.Messung.Messung_2.Messverfahren)
                        lnw = .Messung.Messung_1.Pegel.Pegel
                        C = .Messung.Messung_1.Pegel.C
                    ElseIf iMes = 2 Then
                        'MV = Get_Str_MV(.Messung.Messung_3.Messverfahren)
                        lnw = .Messung.Messung_1.Pegel.Pegel
                        C = .Messung.Messung_1.Pegel.C
                    ElseIf iMes = 3 Then
                        'MV = Get_Str_MV(.Messung.Messung_4.Messverfahren)
                        lnw = .Messung.Messung_1.Pegel.Pegel
                        C = .Messung.Messung_1.Pegel.C
                    ElseIf iMes = 4 Then
                        'MV = Get_Str_MV(.Messung.Messung_5.Messverfahren)
                        lnw = .Messung.Messung_1.Pegel.Pegel
                        C = .Messung.Messung_1.Pegel.C
                    ElseIf iMes = 5 Then
                        'MV = Get_Str_MV(.Messung.Messung_6.Messverfahren)
                        lnw = .Messung.Messung_1.Pegel.Pegel
                        C = .Messung.Messung_1.Pegel.C
                    End If
                    

                    'If MV = "KMV" And lnw > 0 And lnw <= 100 _
                    '    And .Messung.Messanteil > 0 Then

                    '    bGueltig = True
                    '    str = Enter_TS_Decken(CInt(lnw), C) ', "Kurzmessverfahren" & sMA

                    'ElseIf MV = "NMV" Then
                    If lnw > 0 And lnw <= 100 And (IsNumeric(C) Or C = "") Then '_
                        'And .Messung.Messanteil > 0 Then

                        bGueltig = True
                        str = Enter_TS_Decken(CInt(lnw), C) ', "Normmessverfahren" & sMA
                    End If
                    'End If
                Next

                If bGueltig = False Then
                    Get_str_TS_Decken = ""
                Else
                    Get_str_TS_Decken = str
                End If

            ElseIf .Untersuchung = Prognose Then
                lnw = .Prognose.Pegel.Pegel '(.Cells(row_TS_Decken_Prognose_L, 8).Value)
                C = .Prognose.Pegel.C '.Cells(row_TS_Decken_Prognose_C, 8).Value

                If lnw > 0 And lnw <= 100 And _
                        (IsNumeric(C) Or C = "") Then

                    Get_str_TS_Decken = Enter_TS_Decken(CInt(lnw), C) ', "rechnerisch"
                Else
                    Get_str_TS_Decken = ""
                    Exit Function
                End If
            Else 'If .Untersuchung = nv Or .Cells(row_TS_Decken_Untersuchung, 7).Value = "" Then
                Get_str_TS_Decken = ""
                Exit Function
            End If
        End With
    End Function
    Private Function Enter_TS_Decken(ByVal Pegel As Integer, ByVal sC As String) As String 'ByVal MV As String,
        Dim str As String

        str = "<raum>"

        str = str & "<Decke>"

        str = str & "<trittschall>"

        str = str & "<trittschallpegel>"
        str = str & "<ln-w>"
        str = str & Pegel
        str = str & "</ln-w>"

        str = str & "<nachweis>"
        str = str & "<verfahren>"
        str = str & ""
        str = str & "</verfahren>"

        str = str & "<ln-wPlusC>"
        If IsNumeric(sC) Then ' <> "" Then
            Dim c As Integer = CInt(sC)
            If (Pegel <= 28 And Pegel + c <= 28) Or (Pegel > 28 And Pegel <= 34 And Pegel + c <= 34) Or (Pegel > 34 And Pegel <= 40 And Pegel + c <= 40) Or _
                (Pegel > 40 And Pegel <= 46 And Pegel + c <= 46) Or (Pegel > 46 And Pegel <= 53 And Pegel + c <= 53) Or (Pegel > 53 And Pegel <= 60 And Pegel + c <= 60) Then
                str = str & "ja"
            Else
                str = str & "nein"
            End If
        Else
            str = str & "nein"
        End If
        str = str & "</ln-wPlusC>"
        str = str & "</nachweis>"
        str = str & "</trittschallpegel>"

        str = str & "</trittschall>"

        str = str & "</Decke>"

        str = str & "</raum>"

        Enter_TS_Decken = str
    End Function

    Public Function Get_str_LS_Wand() As String
        Dim str As String = ""
        Dim iMes As Integer
        Dim lnw As Single
        Dim C As String = ""

        With Projekt.LS_Wand

            If .Untersuchung = Messung Then
                Dim bGueltig As Boolean
                bGueltig = False

                For iMes = 0 To 5
                    'Dim sMA As String
                    'sMA = ""
                    'If .Messung.Messanteil = groesser50 Then sMA = "+50"
                    'Dim MV As String = ""
                    If iMes = 0 Then
                        'MV = Get_Str_MV(.Messung.Messung_1.Messverfahren)
                        lnw = .Messung.Messung_1.Pegel
                        C = .Messung.Messung_1.C
                    ElseIf iMes = 1 Then
                        'MV = Get_Str_MV(.Messung.Messung_1.Messverfahren)
                        lnw = .Messung.Messung_1.Pegel
                        C = .Messung.Messung_1.C
                    ElseIf iMes = 2 Then
                        'MV = Get_Str_MV(.Messung.Messung_1.Messverfahren)
                        lnw = .Messung.Messung_1.Pegel
                        C = .Messung.Messung_1.C
                    ElseIf iMes = 3 Then
                        'MV = Get_Str_MV(.Messung.Messung_1.Messverfahren)
                        lnw = .Messung.Messung_1.Pegel
                        C = .Messung.Messung_1.C
                    ElseIf iMes = 4 Then
                        'MV = Get_Str_MV(.Messung.Messung_1.Messverfahren)
                        lnw = .Messung.Messung_1.Pegel
                        C = .Messung.Messung_1.C
                    ElseIf iMes = 5 Then
                        'MV = Get_Str_MV(.Messung.Messung_1.Messverfahren)
                        lnw = .Messung.Messung_1.Pegel
                        C = .Messung.Messung_1.C
                    End If

                    'If MV = "KMV" And lnw > 0 And lnw <= 100 _
                    '    And .Messung.Messanteil > 0 Then

                    '    bGueltig = True
                    '    str = Enter_LS_Wand(CInt(lnw), "Kurzmessverfahren" & sMA, C)

                    'ElseIf MV = "NMV" Then
                    If lnw > 0 And lnw <= 100 And _
                         (IsNumeric(C) Or C = "") Then _
                        'And .Messung.Messanteil > 0 Then

                        bGueltig = True
                        str = Enter_LS_Wand(CInt(lnw), "Messung", C) '"Normmessverfahren" & sMA,
                    End If
                    'End If
                Next

                If bGueltig = False Then
                    Get_str_LS_Wand = ""
                Else
                    Get_str_LS_Wand = str
                End If

            ElseIf .Untersuchung = Prognose Then
                lnw = .Prognose.Pegel
                C = .Prognose.C

                If lnw > 0 And lnw <= 100 And _
                        (IsNumeric(C) Or C = "") Then

                    Get_str_LS_Wand = Enter_LS_Wand(CInt(lnw), "rechnerisch", C)

                Else
                    Get_str_LS_Wand = ""
                    Exit Function
                End If
            Else 'If .Cells(row_LS_wandn_Untersuchung, 7).Value = "nicht vorhanden" Or .Cells(row_LS_wandn_Untersuchung, 7).Value = "" Then
                Get_str_LS_Wand = ""
                Exit Function
            End If
        End With
    End Function
    Private Function Enter_LS_Wand(ByVal Pegel As Integer, ByVal MV As String, ByVal sC As String) As String '

        Dim str As String


        str = "<raum>"

        str = str & "<wand>"

        str = str & "<luftschall>"

        str = str & "<daemmmass>"
        str = str & "<R-w>"
        str = str & Pegel
        str = str & "</R-w>"

        str = str & "<nachweis>"
        str = str & "<verfahren>"
        str = str & MV
        str = str & "</verfahren>"

        str = str & "<R-wPlusC>"
        If IsNumeric(sC) Then ' <> "" Then
            Dim c As Single = CDec(sC)
            If (Pegel >= 72 And Pegel + c >= 72) Or (Pegel >= 67 And Pegel < 72 And Pegel + c >= 67) Or (Pegel >= 62 And Pegel < 67 And Pegel + c >= 62) Or _
                    (Pegel >= 57 And Pegel < 62 And Pegel + c >= 57) Or (Pegel >= 53 And Pegel < 57 And Pegel + c >= 53) Or _
                    (Pegel >= 50 And Pegel < 53 And Pegel + c >= 50) Then
                str = str & "ja"
            Else
                str = str & "nein"
            End If
        Else
            str = str & "nein"
        End If
        str = str & "</R-wPlusC>"
        str = str & "</nachweis>"

        str = str & "</daemmmass>"

        str = str & "</luftschall>"

        str = str & "</wand>"

        str = str & "</raum>"

        Enter_LS_Wand = str
    End Function

    Private Function Get_str_LS_Decke() As String
        Dim str As String = ""
        Dim iMes As Integer
        Dim lnw As Single
        Dim C As String = ""

        With Projekt.LS_Decke

            If .Untersuchung = Messung Then
                Dim bGueltig As Boolean
                bGueltig = False

                For iMes = 0 To 5
                    'Dim sMA As String
                    'sMA = ""
                    'If .Messung.Messanteil = groesser50 Then sMA = "+50"
                    'Dim MV As String = ""
                    If iMes = 0 Then
                        'MV = Get_Str_MV(.Messung.Messung_1.Messverfahren)
                        lnw = .Messung.Messung_1.Pegel
                        C = .Messung.Messung_1.C
                    ElseIf iMes = 1 Then
                        'MV = Get_Str_MV(.Messung.Messung_1.Messverfahren)
                        lnw = .Messung.Messung_1.Pegel
                        C = .Messung.Messung_1.C
                    ElseIf iMes = 2 Then
                        'MV = Get_Str_MV(.Messung.Messung_1.Messverfahren)
                        lnw = .Messung.Messung_1.Pegel
                        C = .Messung.Messung_1.C
                    ElseIf iMes = 3 Then
                        'MV = Get_Str_MV(.Messung.Messung_1.Messverfahren)
                        lnw = .Messung.Messung_1.Pegel
                        C = .Messung.Messung_1.C
                    ElseIf iMes = 4 Then
                        'MV = Get_Str_MV(.Messung.Messung_1.Messverfahren)
                        lnw = .Messung.Messung_1.Pegel
                        C = .Messung.Messung_1.C
                    ElseIf iMes = 5 Then
                        'MV = Get_Str_MV(.Messung.Messung_1.Messverfahren)
                        lnw = .Messung.Messung_1.Pegel
                        C = .Messung.Messung_1.C
                    End If

                    'If MV = "KMV" And lnw > 0 And lnw <= 100 _
                    '    And .Messung.Messanteil > 0 Then

                    '    bGueltig = True
                    '    str = Enter_LS_Decken(CInt(lnw), "Kurzmessverfahren" & sMA, C)

                    'ElseIf MV = "NMV" Then
                    If lnw > 0 And lnw <= 100 And _
                         (IsNumeric(C) Or C = "") Then _
                        'And .Messung.Messanteil > 0 Then

                        bGueltig = True
                        str = Enter_LS_Decken(CInt(lnw), "Messung", C) ' "Normmessverfahren" & sMA,
                    End If
                    'End If
                Next

                If bGueltig = False Then
                    Get_str_LS_Decke = ""
                Else
                    Get_str_LS_Decke = str
                End If

            ElseIf .Untersuchung = Prognose Then
                lnw = .Prognose.Pegel
                C = .Prognose.C

                If lnw > 0 And lnw <= 100 And _
                        (IsNumeric(C) Or C = "") Then

                    Get_str_LS_Decke = Enter_LS_Decken(CInt(lnw), "rechnerisch", C)

                Else
                    Get_str_LS_Decke = ""
                    Exit Function
                End If
            Else 'If .Cells(row_LS_Decken_Untersuchung, 7).Value = "nicht vorhanden" Or .Cells(row_LS_Decken_Untersuchung, 7).Value = "" Then
                Get_str_LS_Decke = ""
                Exit Function
            End If
        End With


        '##############################
    End Function
    Private Function Enter_LS_Decken(ByVal Pegel As Integer, ByVal MV As String, ByVal sC As String) As String '
        Dim str As String

        str = "<raum>"

        str = str & "<Decke>"

        str = str & "<luftschall>"

        str = str & "<daemmmass>"

        str = str & "<R-w>"
        str = str & Pegel
        str = str & "</R-w>"

        str = str & "<nachweis>"
        str = str & "<verfahren>"
        str = str & MV
        str = str & "</verfahren>"

        str = str & "<R-wPlusC>"
        If IsNumeric(sC) Then ' <> "" Then
            Dim c As Integer = CInt(sC)
            If (Pegel >= 72 And Pegel + c >= 72) Or (Pegel >= 67 And Pegel < 72 And Pegel + c >= 67) Or (Pegel >= 62 And Pegel < 67 And Pegel + c >= 62) Or _
                    (Pegel >= 57 And Pegel < 62 And Pegel + c >= 57) Or (Pegel >= 53 And Pegel < 57 And Pegel + c >= 54) Or _
                    (Pegel >= 50 And Pegel < 54 And Pegel + c >= 50) Then
                str = str & "ja"
            Else
                str = str & "nein"
            End If
        Else
            str = str & "nein"
        End If
        str = str & "</R-wPlusC>"
        str = str & "</nachweis>"

        str = str & "</daemmmass>"

        str = str & "</luftschall>"

        str = str & "</Decke>"

        str = str & "</raum>"

        Enter_LS_Decken = str
    End Function

    Public Function Get_str_Gueltigkeit() As String
        Dim str As String

        With Projekt

            str = "<gueltigkeit>"
            If IsDate(.Datum) Then
                Dim sDate As Date = CType(.Datum, Date)
                Dim sMonth As String = sDate.Month.ToString
                Dim sDay As String = sDate.Day.ToString
                If sMonth.Length = 1 Then sMonth = "0" & sMonth
                If sDay.Length = 1 Then sDay = "0" & sDay
                str = str & sDate.AddYears(10).Year & "-" & sMonth & "-" & sDay  '"1982-05-16"    'str = str & .Cells(8, 11).Value
            Else
                Get_str_Gueltigkeit = "" 'str & "1000-01-01"
                Exit Function
            End If
            str = str & "</gueltigkeit>"

        End With

        Get_str_Gueltigkeit = str
    End Function

    Public Function Get_str_Baujahr() As String
        Dim str As String

        With Projekt.Gebaeude

            str = "<baujahr>"
            If Len(.Baujahr) = 4 And .Baujahr >= 1800 And .Baujahr <= 2200 Then
                str = str & .Baujahr.ToString
            Else
                Get_str_Baujahr = "" 'str & "1800"
                Exit Function
            End If
            str = str & "</baujahr>"

        End With

        Get_str_Baujahr = str

    End Function

    Public Function Get_str_Kosten() As String
        Dim str As String = ""

        With Projekt.Gebaeude

            If .Kosten > 0 Then
                str = "<kosten>"
                Dim sK As String
                sK = .Kosten.ToString
                sK = Replace(sK, ",", ".")
                str = str & sK '.Cells(35, 6).Value
                str = str & "</kosten>"
            End If

        End With

        Get_str_Kosten = str

    End Function
    Public Function Get_str_Gebietscharakter() As String
        Dim str As String
        With Projekt.Standort
            str = "<gebietscharakter>"
            If .Gebietscharakter = GC_WR Then
                str = str & "WR"
            ElseIf .Gebietscharakter = GC_WA Then
                str = str & "WA"
            ElseIf .Gebietscharakter = GC_MIWB Then
                str = str & "MI/WB"
            ElseIf .Gebietscharakter = GC_GE Then
                str = str & "GE"
            ElseIf .Gebietscharakter = GC_GI Then
                str = str & "GI"
            Else
                Get_str_Gebietscharakter = ""
                Exit Function
            End If
            str = str & "</gebietscharakter>"
        End With
        Get_str_Gebietscharakter = str
    End Function

    Public Function Get_str_Gebaeudetyp() As String
        Dim str As String

        With Projekt.Gebaeude

            str = "<gebaeudetyp>"
            If .Gebaeudetyp = "Einfamilienhaus" Or .Gebaeudetyp = "Doppelhaus" Or .Gebaeudetyp = "Reihenhaus" Or .Gebaeudetyp = "Mehrfamilienhaus" Then
                str = str & .Gebaeudetyp
            ElseIf .Gebaeudetyp = "Wohn- und Geschäftshaus" Then
                str = str & "Wohn- und Geschaeftshaus"
            Else
                Get_str_Gebaeudetyp = ""
                Exit Function
            End If
            str = str & "</gebaeudetyp>"

        End With

        Get_str_Gebaeudetyp = str
    End Function

    Public Function Get_str_Stockwerk() As String
        Dim str As String

        With Projekt.Wohnung.Geschoss

            str = "<stockwerk>"
            If .Typ = Sonstiges Then
                str = str & "7777"
            ElseIf .Typ = DG Then
                str = str & "9999"
            ElseIf .Typ = EG Then
                str = str & "0"
            ElseIf .Typ = UG Then
                If .Nr > 0 And .Nr <= 128 Then
                    str = str & "-" & .Nr.ToString
                Else
                    Get_str_Stockwerk = ""
                    Exit Function
                End If
            ElseIf .Typ = OG Then
                If .Nr > 0 And .Nr <= 127 Then
                    str = str & .Nr
                Else
                    Get_str_Stockwerk = "" 'str & "7777"
                    Exit Function
                End If
            Else
                Get_str_Stockwerk = "" 'str & "7777"
                Exit Function
            End If
            str = str & "</stockwerk>"

        End With

        Get_str_Stockwerk = str
    End Function

    Public Function Get_str_Aussenlaermpegel() As String
        Dim str As String

        With Projekt.Standort

            str = "<aussenlaermpegel>"

            str = str & "<laermpegelbereich>"
            If .Aussenlaermpegel = AP_bis55 Then
                str = str & "55"
            ElseIf .Aussenlaermpegel = AP_56_60 Then
                str = str & "56"
            ElseIf .Aussenlaermpegel = AP_61_65 Then
                str = str & "61"
            ElseIf .Aussenlaermpegel = AP_66_70 Then
                str = str & "66"
            ElseIf .Aussenlaermpegel = AP_71_75 Then
                str = str & "71"
            ElseIf .Aussenlaermpegel = AP_76 Then
                str = str & "76"
            Else
                Get_str_Aussenlaermpegel = "" 'str & "0"
                Exit Function
            End If
            str = str & "</laermpegelbereich>"

            str = str & "<orientierung>"
            If .abgewandFreibereich = aF_ja Then
                str = str & "Freibereich abgewandt"
            ElseIf .abgewandFreibereich = aF_nein Then
                str = str & "Orientierung beliebig"
            Else
                Get_str_Aussenlaermpegel = ""
                Exit Function
            End If
            str = str & "</orientierung>"

            str = str & "</aussenlaermpegel>"

        End With

        Get_str_Aussenlaermpegel = str
    End Function
    
    Public Function Get_str_Treppen() As String
        Dim str As String = ""
        Dim iMes As Integer
        Dim lnw As Single
        Dim C As String = ""
        'Dim MV As String = ""

        With Projekt.TS_TPLH

            If .Untersuchung = Messung Then
                Dim bGueltig As Boolean
                bGueltig = False

                For iMes = 0 To 5
                    Dim bTreppe As Boolean = False
                    If iMes = 0 Then
                        If .Messung.Messung_1.Bauteil = Treppe Then bTreppe = True
                        'MV = Get_Str_MV(.Messung.Messung_1.Messverfahren)
                        lnw = .Messung.Messung_1.Pegel.Pegel
                        C = .Messung.Messung_1.Pegel.C
                    ElseIf iMes = 1 Then
                        If .Messung.Messung_2.Bauteil = Treppe Then bTreppe = True
                        'MV = Get_Str_MV(.Messung.Messung_2.Messverfahren)
                        lnw = .Messung.Messung_2.Pegel.Pegel
                        C = .Messung.Messung_2.Pegel.C
                    ElseIf iMes = 2 Then
                        If .Messung.Messung_3.Bauteil = Treppe Then bTreppe = True
                        'MV = Get_Str_MV(.Messung.Messung_3.Messverfahren)
                        lnw = .Messung.Messung_3.Pegel.Pegel
                        C = .Messung.Messung_3.Pegel.C
                    ElseIf iMes = 3 Then
                        If .Messung.Messung_4.Bauteil = Treppe Then bTreppe = True
                        'MV = Get_Str_MV(.Messung.Messung_4.Messverfahren)
                        lnw = .Messung.Messung_4.Pegel.Pegel
                        C = .Messung.Messung_4.Pegel.C
                    ElseIf iMes = 4 Then
                        If .Messung.Messung_5.Bauteil = Treppe Then bTreppe = True
                        'MV = Get_Str_MV(.Messung.Messung_5.Messverfahren)
                        lnw = .Messung.Messung_5.Pegel.Pegel
                        C = .Messung.Messung_5.Pegel.C
                    ElseIf iMes = 5 Then
                        If .Messung.Messung_6.Bauteil = Treppe Then bTreppe = True
                        'MV = Get_Str_MV(.Messung.Messung_6.Messverfahren)
                        lnw = .Messung.Messung_6.Pegel.Pegel
                        C = .Messung.Messung_6.Pegel.C
                    End If

                    If bTreppe Then
                        'If MV = "KMV" And lnw > 0 And lnw <= 100 Then

                        '    bGueltig = True
                        '    str = Enter_Treppen(CInt(lnw), "Kurzmessverfahren", C)

                        'ElseIf MV = "NMV" Then
                        If lnw > 0 And lnw <= 100 And _
                             (IsNumeric(C) Or C = "") Then

                            bGueltig = True
                            str = Enter_Treppen(CInt(lnw), "Messung", C) ', "Normmessverfahren"
                        End If
                        'End If
                    End If
                Next

                If bGueltig = False Then
                    Get_str_Treppen = ""
                Else
                    Get_str_Treppen = str
                End If

            ElseIf .Untersuchung = Prognose Then

                lnw = .Prognose.Treppe.Pegel.Pegel
                C = .Prognose.Treppe.Pegel.C

                If lnw > 0 And lnw <= 100 And (IsNumeric(C) Or C = "") Then
                    Get_str_Treppen = Enter_Treppen(CInt(lnw), "rechnerisch", C) '
                Else
                    Get_str_Treppen = ""
                    Exit Function
                End If
            Else 'If .Cells(74, 7).Value = "nicht vorhanden" Or .Cells(74, 7).Value = "" Then
                Get_str_Treppen = ""
                Exit Function
            End If

        End With

    End Function
    Private Function Enter_Treppen(ByVal Pegel As Integer, ByVal MV As String, ByVal sC As String) As String '
        Dim str As String
        str = "<treppe>"

        str = str & "<trittschallpegel>"
        str = str & "<ln-w>"
        str = str & Pegel
        str = str & "</ln-w>"

        str = str & "<nachweis>"

        str = str & "<verfahren>"
        str = str & MV
        str = str & "</verfahren>"

        str = str & "<ln-wPlusC>"
        If IsNumeric(sC) Then ' <> "" Then
            Dim C As Single = CDec(sC)
            If (Pegel <= 28 And Pegel + C <= 28) Or (Pegel <= 34 And Pegel + C <= 34) Or (Pegel <= 40 And Pegel + C <= 40) Or _
                (Pegel <= 46 And Pegel + C <= 46) Or (Pegel <= 53 And Pegel + C <= 53) Or (Pegel <= 60 And Pegel + C <= 60) Then
                str = str & "ja"
            Else
                str = str & "nein"
            End If
        Else
            str = str & "nein"
        End If
        str = str & "</ln-wPlusC>"

        str = str & "</nachweis>"
        str = str & "</trittschallpegel>"

        str = str & "</treppe>"

        Enter_Treppen = str
    End Function

    Public Function Get_str_Podest() As String
        Dim str As String = ""
        Dim iMes As Integer
        Dim lnw As Single
        Dim C As String = ""
        'Dim MV As String = ""

        With Projekt.TS_TPLH

            If .Untersuchung = Messung Then
                Dim bGueltig As Boolean
                bGueltig = False

                For iMes = 0 To 5
                    Dim bPodest As Boolean = False
                    If iMes = 0 Then
                        If .Messung.Messung_1.Bauteil = Podest Then bPodest = True
                        'MV = Get_Str_MV(.Messung.Messung_1.Messverfahren)
                        lnw = .Messung.Messung_1.Pegel.Pegel
                        C = .Messung.Messung_1.Pegel.C
                    ElseIf iMes = 1 Then
                        If .Messung.Messung_2.Bauteil = Podest Then bPodest = True
                        'MV = Get_Str_MV(.Messung.Messung_2.Messverfahren)
                        lnw = .Messung.Messung_2.Pegel.Pegel
                        C = .Messung.Messung_2.Pegel.C
                    ElseIf iMes = 2 Then
                        If .Messung.Messung_3.Bauteil = Podest Then bPodest = True
                        'MV = Get_Str_MV(.Messung.Messung_3.Messverfahren)
                        lnw = .Messung.Messung_3.Pegel.Pegel
                        C = .Messung.Messung_3.Pegel.C
                    ElseIf iMes = 3 Then
                        If .Messung.Messung_4.Bauteil = Podest Then bPodest = True
                        'MV = Get_Str_MV(.Messung.Messung_4.Messverfahren)
                        lnw = .Messung.Messung_4.Pegel.Pegel
                        C = .Messung.Messung_4.Pegel.C
                    ElseIf iMes = 4 Then
                        If .Messung.Messung_5.Bauteil = Podest Then bPodest = True
                        'MV = Get_Str_MV(.Messung.Messung_5.Messverfahren)
                        lnw = .Messung.Messung_5.Pegel.Pegel
                        C = .Messung.Messung_5.Pegel.C
                    ElseIf iMes = 5 Then
                        If .Messung.Messung_6.Bauteil = Podest Then bPodest = True
                        'MV = Get_Str_MV(.Messung.Messung_6.Messverfahren)
                        lnw = .Messung.Messung_6.Pegel.Pegel
                        C = .Messung.Messung_6.Pegel.C
                    End If

                    If bPodest Then
                        'If MV = "KMV" And lnw > 0 And lnw <= 100 Then

                        '    bGueltig = True
                        '    str = str & Enter_Podest(CInt(lnw), "Kurzmessverfahren", C)

                        'ElseIf MV = "NMV" Then
                        If lnw > 0 And lnw <= 100 And _
                             (IsNumeric(C) Or C = "") Then

                            bGueltig = True
                            str = str & Enter_Podest(CInt(lnw), "Messung", C) ', "Normmessverfahren"
                        End If
                        'End If
                    End If
                Next

                If bGueltig = False Then
                    Get_str_Podest = ""
                Else
                    Get_str_Podest = str
                End If

            ElseIf .Untersuchung = Prognose Then

                'If .Prognose.Bauteil = Podest Then
                lnw = .Prognose.Podest.Pegel.Pegel
                C = .Prognose.Podest.Pegel.C

                If lnw > 0 And lnw <= 100 And _
                        (IsNumeric(C) Or C = "") Then
                    Get_str_Podest = Enter_Podest(CInt(lnw), "rechnerisch", C)
                Else
                    Get_str_Podest = ""
                    Exit Function
                End If
                'Else
                '    Get_str_Podest = ""
                '    Exit Function
                'End If
            Else 'If .Cells(74, 7).Value = "nicht vorhanden" Or .Cells(74, 7).Value = "" Then
                Get_str_Podest = ""
                Exit Function
            End If

        End With


    End Function

    Private Function Enter_Podest(ByVal Pegel As Integer, ByVal MV As String, ByVal sC As String) As String


        Dim str As String

        str = "<podest>"

        str = str & "<trittschallpegel>"

        str = str & "<ln-w>"
        str = str & Pegel
        str = str & "</ln-w>"

        str = str & "<nachweis>"

        str = str & "<verfahren>"
        str = str & MV
        str = str & "</verfahren>"

        str = str & "<ln-wPlusC>"
        If IsNumeric(sC) Then ' <> "" Then
            Dim c As Single = CDec(sC)
            If (Pegel <= 28 And Pegel + c <= 28) Or (Pegel <= 34 And Pegel + c <= 34) Or (Pegel <= 40 And Pegel + c <= 40) Or _
                (Pegel <= 46 And Pegel + c <= 46) Or (Pegel <= 53 And Pegel + c <= 53) Or (Pegel <= 60 And Pegel + c <= 60) Then
                str = str & "ja"
            Else
                str = str & "nein"
            End If
        Else
            str = str & "nein"
        End If
        str = str & "</ln-wPlusC>"

        str = str & "</nachweis>"

        str = str & "</trittschallpegel>"

        str = str & "</podest>"

        Enter_Podest = str
    End Function

    Public Function Get_str_Hausflur() As String

        Dim str As String = ""
        Dim iMes As Integer
        Dim lnw As Single
        Dim C As String = ""

        With Projekt.TS_TPLH

            If .Untersuchung = Messung Then
                Dim bGueltig As Boolean
                bGueltig = False

                For iMes = 0 To 5
                    Dim bHausflur As Boolean = False
                    If iMes = 0 Then
                        If .Messung.Messung_1.Bauteil = Hausflur Then bHausflur = True
                        lnw = .Messung.Messung_1.Pegel.Pegel
                        C = .Messung.Messung_1.Pegel.C
                    ElseIf iMes = 1 Then
                        If .Messung.Messung_2.Bauteil = Hausflur Then bHausflur = True
                        lnw = .Messung.Messung_2.Pegel.Pegel
                        C = .Messung.Messung_2.Pegel.C
                    ElseIf iMes = 2 Then
                        If .Messung.Messung_3.Bauteil = Hausflur Then bHausflur = True
                        lnw = .Messung.Messung_3.Pegel.Pegel
                        C = .Messung.Messung_3.Pegel.C
                    ElseIf iMes = 3 Then
                        If .Messung.Messung_4.Bauteil = Hausflur Then bHausflur = True
                        lnw = .Messung.Messung_4.Pegel.Pegel
                        C = .Messung.Messung_4.Pegel.C
                    ElseIf iMes = 4 Then
                        If .Messung.Messung_5.Bauteil = Hausflur Then bHausflur = True
                        lnw = .Messung.Messung_5.Pegel.Pegel
                        C = .Messung.Messung_5.Pegel.C
                    ElseIf iMes = 5 Then
                        If .Messung.Messung_6.Bauteil = Hausflur Then bHausflur = True
                        lnw = .Messung.Messung_6.Pegel.Pegel
                        C = .Messung.Messung_6.Pegel.C
                    End If

                    If bHausflur Then
                        'If MV = "KMV" And lnw > 0 And lnw <= 100 Then

                        '    bGueltig = True
                        '    str = str & Enter_Hausflur(CInt(lnw), "Kurzmessverfahren", C)

                        'ElseIf MV = "NMV" Then
                        If lnw > 0 And lnw <= 100 And _
                             (IsNumeric(C) Or C = "") Then

                            bGueltig = True
                            str = str & Enter_Hausflur(CInt(lnw), "Normmessverfahren", C)
                            'End If
                        End If
                    End If
                Next

                If bGueltig = False Then
                    Get_str_Hausflur = ""
                Else
                    Get_str_Hausflur = str
                End If

            ElseIf .Untersuchung = Prognose Then

                'If .Prognose.Bauteil = Hausflur Then
                lnw = .Prognose.Hausflur.Pegel.Pegel
                C = .Prognose.Hausflur.Pegel.C

                If lnw > 0 And lnw <= 100 And _
                        (IsNumeric(C) Or C = "") Then
                    Get_str_Hausflur = Enter_Hausflur(CInt(lnw), "rechnerisch", C)
                Else
                    Get_str_Hausflur = ""
                    Exit Function
                End If
                'Else
                '    Get_str_Hausflur = ""
                '    Exit Function
                'End If
            Else 'If .Cells(74, 7).Value = "nicht vorhanden" Or .Cells(74, 7).Value = "" Then
                Get_str_Hausflur = ""
                Exit Function
            End If

        End With


    End Function

    Private Function Enter_Hausflur(ByVal lnw As Integer, ByVal MV As String, ByVal sC As String) As String

        Dim str As String

        str = "<hausflur>"

        str = str & "<trittschallpegel>"


        str = str & "<ln-w>"
        str = str & lnw
        str = str & "</ln-w>"

        str = str & "<nachweis>"

        str = str & "<verfahren>"
        str = str & MV
        str = str & "</verfahren>"

        str = str & "<ln-wPlusC>"
        If IsNumeric(sC) Then ' <> "" Then
            Dim c As Single = CDec(sC)
            If (lnw <= 28 And lnw + c <= 28) Or (lnw <= 34 And lnw + c <= 34) Or (lnw <= 40 And lnw + c <= 40) Or _
                (lnw <= 46 And lnw + c <= 46) Or (lnw <= 53 And lnw + c <= 53) Or (lnw <= 60 And lnw + c <= 60) Then
                str = str & "ja"
            Else
                str = str & "nein"
            End If
        Else
            str = str & "nein"
        End If
        str = str & "</ln-wPlusC>"

        str = str & "</nachweis>"


        str = str & "</trittschallpegel>"

        str = str & "</hausflur>"

        Enter_Hausflur = str
    End Function

    Public Function Get_str_Balkone() As String

        Dim str As String = ""
        Dim iMes As Integer
        Dim lnw As Single
        Dim C As String = ""
        'Dim MV As String = ""

        With Projekt.TS_BLT

            If .Untersuchung = Messung Then
                Dim bGueltig As Boolean
                bGueltig = False

                For iMes = 0 To 5
                    Dim bBalkon As Boolean = False
                    If iMes = 0 Then
                        If .Messung.Messung_1.Bauteil = Balkon Then bBalkon = True
                        'MV = Get_Str_MV(.Messung.Messung_1.Messverfahren)
                        lnw = .Messung.Messung_1.Pegel.Pegel
                        C = .Messung.Messung_1.Pegel.C
                    ElseIf iMes = 1 Then
                        If .Messung.Messung_2.Bauteil = Balkon Then bBalkon = True
                        'MV = Get_Str_MV(.Messung.Messung_2.Messverfahren)
                        lnw = .Messung.Messung_2.Pegel.Pegel
                        C = .Messung.Messung_2.Pegel.C
                    ElseIf iMes = 2 Then
                        If .Messung.Messung_3.Bauteil = Balkon Then bBalkon = True
                        'MV = Get_Str_MV(.Messung.Messung_3.Messverfahren)
                        lnw = .Messung.Messung_3.Pegel.Pegel
                        C = .Messung.Messung_3.Pegel.C
                    ElseIf iMes = 3 Then
                        If .Messung.Messung_4.Bauteil = Balkon Then bBalkon = True
                        'MV = Get_Str_MV(.Messung.Messung_4.Messverfahren)
                        lnw = .Messung.Messung_4.Pegel.Pegel
                        C = .Messung.Messung_4.Pegel.C
                    ElseIf iMes = 4 Then
                        If .Messung.Messung_5.Bauteil = Balkon Then bBalkon = True
                        'MV = Get_Str_MV(.Messung.Messung_5.Messverfahren)
                        lnw = .Messung.Messung_5.Pegel.Pegel
                        C = .Messung.Messung_5.Pegel.C
                    ElseIf iMes = 5 Then
                        If .Messung.Messung_6.Bauteil = Balkon Then bBalkon = True
                        'MV = Get_Str_MV(.Messung.Messung_6.Messverfahren)
                        lnw = .Messung.Messung_6.Pegel.Pegel
                        C = .Messung.Messung_6.Pegel.C
                    End If

                    If bBalkon Then
                        'If MV = "KMV" And lnw > 0 And lnw <= 100 Then

                        '    bGueltig = True
                        '    str = str & Enter_Balkone(CInt(lnw), "Kurzmessverfahren", C)

                        'ElseIf MV = "NMV" Then
                        If lnw > 0 And lnw <= 100 And _
                             (IsNumeric(C) Or C = "") Then

                            bGueltig = True
                            str = str & Enter_Balkone(CInt(lnw), "Normmessverfahren", C)
                        End If
                        'End If
                    End If
                Next

                If bGueltig = False Then
                    Get_str_Balkone = ""
                Else
                    Get_str_Balkone = str
                End If

            ElseIf .Untersuchung = Prognose Then

                'If .Prognose.Bauteil = Balkon Then
                lnw = .Prognose.Balkon.Pegel.Pegel
                C = .Prognose.Balkon.Pegel.C

                If lnw > 0 And lnw <= 100 And _
                        (IsNumeric(C) Or C = "") Then
                    Get_str_Balkone = Enter_Balkone(CInt(lnw), "rechnerisch", C)
                Else
                    Get_str_Balkone = ""
                    Exit Function
                End If
                'Else
                '    Get_str_Balkone = ""
                '    Exit Function
                'End If
            Else 'If .Cells(74, 7).Value = "nicht vorhanden" Or .Cells(74, 7).Value = "" Then
                Get_str_Balkone = ""
                Exit Function
            End If

        End With

    End Function
    Private Function Enter_Balkone(ByVal lnw As Integer, ByVal MV As String, ByVal sC As String) As String
        Dim str As String

        str = "<balkon>"

        str = str & "<trittschallpegel>"


        str = str & "<ln-w>"
        str = str & lnw
        str = str & "</ln-w>"

        str = str & "<nachweis>"

        str = str & "<verfahren>"
        str = str & MV
        str = str & "</verfahren>"

        str = str & "<ln-wPlusC>"
        If IsNumeric(sC) Then ' <> "" Then
            Dim C As Single = CDec(sC)

            If (lnw <= 28 And lnw + C <= 28) Or (lnw <= 34 And lnw + C <= 34) Or (lnw <= 40 And lnw + C <= 40) Or _
                (lnw <= 46 And lnw + C <= 46) Or (lnw <= 53 And lnw + C <= 53) Or (lnw <= 60 And lnw + C <= 60) Then
                str = str & "ja"
            Else
                str = str & "nein"
            End If
        Else
            str = str & "nein"
        End If
        str = str & "</ln-wPlusC>"

        str = str & "</nachweis>"

        str = str & "</trittschallpegel>"

        str = str & "</balkon>"


        Enter_Balkone = str

    End Function

    Public Function Get_str_Laubengang() As String
        Dim str As String = ""
        Dim iMes As Integer
        Dim lnw As Single
        Dim C As String = ""

        With Projekt.TS_TPLH

            If .Untersuchung = Messung Then
                Dim bGueltig As Boolean
                bGueltig = False

                For iMes = 0 To 5
                    Dim bLaube As Boolean = False
                    If iMes = 0 Then
                        If .Messung.Messung_1.Bauteil = Laube Then bLaube = True
                        lnw = .Messung.Messung_1.Pegel.Pegel
                        C = .Messung.Messung_1.Pegel.C
                    ElseIf iMes = 1 Then
                        If .Messung.Messung_2.Bauteil = Laube Then bLaube = True
                        lnw = .Messung.Messung_2.Pegel.Pegel
                        C = .Messung.Messung_2.Pegel.C
                    ElseIf iMes = 2 Then
                        If .Messung.Messung_3.Bauteil = Laube Then bLaube = True
                        lnw = .Messung.Messung_3.Pegel.Pegel
                        C = .Messung.Messung_3.Pegel.C
                    ElseIf iMes = 3 Then
                        If .Messung.Messung_4.Bauteil = Laube Then bLaube = True
                        lnw = .Messung.Messung_4.Pegel.Pegel
                        C = .Messung.Messung_4.Pegel.C
                    ElseIf iMes = 4 Then
                        If .Messung.Messung_5.Bauteil = Laube Then bLaube = True
                        lnw = .Messung.Messung_5.Pegel.Pegel
                        C = .Messung.Messung_5.Pegel.C
                    ElseIf iMes = 5 Then
                        If .Messung.Messung_6.Bauteil = Laube Then bLaube = True
                        lnw = .Messung.Messung_6.Pegel.Pegel
                        C = .Messung.Messung_6.Pegel.C
                    End If

                    If bLaube Then
                        'If MV = "KMV" And lnw > 0 And lnw <= 100 Then

                        '    bGueltig = True
                        '    str = str & Enter_Laubengang(CInt(lnw), "Kurzmessverfahren", C)

                        'ElseIf MV = "NMV" Then
                        If lnw > 0 And lnw <= 100 And _
                             (IsNumeric(C) Or C = "") Then

                            bGueltig = True
                            str = str & Enter_Laubengang(CInt(lnw), "Normmessverfahren", C)
                        End If
                        'End If
                    End If
                Next

                If bGueltig = False Then
                    Get_str_Laubengang = ""
                Else
                    Get_str_Laubengang = str
                End If

            ElseIf .Untersuchung = Prognose Then

                'If .Prognose.Bauteil = Laube Then
                lnw = .Prognose.Laube.Pegel.Pegel
                C = .Prognose.Laube.Pegel.C

                If lnw > 0 And lnw <= 100 And _
                        (IsNumeric(C) Or C = "") Then
                    Get_str_Laubengang = Enter_Laubengang(CInt(lnw), "rechnerisch", C)
                Else
                    Get_str_Laubengang = ""
                    Exit Function
                End If
                'Else
                '    Get_str_Laubengang = ""
                '    Exit Function
                'End If
            Else 'If .Cells(74, 7).Value = "nicht vorhanden" Or .Cells(74, 7).Value = "" Then
                Get_str_Laubengang = ""
                Exit Function
            End If

        End With


    End Function

    Private Function Enter_Laubengang(ByVal lnw As Integer, ByVal MV As String, ByVal sC As String) As String

        Dim str As String

        str = "<laubengang>"

        str = str & "<trittschallpegel>"

        str = str & "<ln-w>"
        str = str & lnw
        str = str & "</ln-w>"

        str = str & "<nachweis>"

        str = str & "<verfahren>"
        str = str & MV
        str = str & "</verfahren>"

        str = str & "<ln-wPlusC>"
        If IsNumeric(sC) Then ' <> "" Then
            Dim c As Single = CDec(sC)
            If (lnw <= 28 And lnw + c <= 28) Or (lnw <= 34 And lnw + c <= 34) Or (lnw <= 40 And lnw + c <= 40) Or _
                (lnw <= 46 And lnw + c <= 46) Or (lnw <= 53 And lnw + c <= 53) Or (lnw <= 60 And lnw + c <= 60) Then
                str = str & "ja"
            Else
                str = str & "nein"
            End If
        Else
            str = str & "nein"
        End If
        str = str & "</ln-wPlusC>"

        str = str & "</nachweis>"

        str = str & "</trittschallpegel>"

        str = str & "</laubengang>"

        Enter_Laubengang = str

    End Function

    Public Function Get_str_Loggie() As String
        Dim str As String = ""
        Dim iMes As Integer
        Dim lnw As Single
        Dim C As String = ""

        With Projekt.TS_BLT

            If .Untersuchung = Messung Then
                Dim bGueltig As Boolean
                bGueltig = False

                For iMes = 0 To 5
                    Dim bLoggie As Boolean = False
                    If iMes = 0 Then
                        If .Messung.Messung_1.Bauteil = Loggie Then bLoggie = True
                        'MV = Get_Str_MV(.Messung.Messung_1.Messverfahren)
                        lnw = .Messung.Messung_1.Pegel.Pegel
                        C = .Messung.Messung_1.Pegel.C
                    ElseIf iMes = 1 Then
                        If .Messung.Messung_2.Bauteil = Loggie Then bLoggie = True
                        'MV = Get_Str_MV(.Messung.Messung_2.Messverfahren)
                        lnw = .Messung.Messung_2.Pegel.Pegel
                        C = .Messung.Messung_2.Pegel.C
                    ElseIf iMes = 2 Then
                        If .Messung.Messung_3.Bauteil = Loggie Then bLoggie = True
                        'MV = Get_Str_MV(.Messung.Messung_3.Messverfahren)
                        lnw = .Messung.Messung_3.Pegel.Pegel
                        C = .Messung.Messung_3.Pegel.C
                    ElseIf iMes = 3 Then
                        If .Messung.Messung_4.Bauteil = Loggie Then bLoggie = True
                        'MV = Get_Str_MV(.Messung.Messung_4.Messverfahren)
                        lnw = .Messung.Messung_4.Pegel.Pegel
                        C = .Messung.Messung_4.Pegel.C
                    ElseIf iMes = 4 Then
                        If .Messung.Messung_5.Bauteil = Loggie Then bLoggie = True
                        'MV = Get_Str_MV(.Messung.Messung_5.Messverfahren)
                        lnw = .Messung.Messung_5.Pegel.Pegel
                        C = .Messung.Messung_5.Pegel.C
                    ElseIf iMes = 5 Then
                        If .Messung.Messung_6.Bauteil = Loggie Then bLoggie = True
                        'MV = Get_Str_MV(.Messung.Messung_6.Messverfahren)
                        lnw = .Messung.Messung_6.Pegel.Pegel
                        C = .Messung.Messung_6.Pegel.C
                    End If

                    If bLoggie Then
                        'If MV = "KMV" And lnw > 0 And lnw <= 100 Then

                        '    bGueltig = True
                        '    str = str & Enter_Loggie(CInt(lnw), "Kurzmessverfahren", C)

                        'ElseIf MV = "NMV" Then
                        If lnw > 0 And lnw <= 100 And _
                            (IsNumeric(C) Or C = "") Then

                            bGueltig = True
                            str = str & Enter_Loggie(CInt(lnw), "Normmessverfahren", C)
                        End If
                        'End If
                    End If
                Next

                If bGueltig = False Then
                    Get_str_Loggie = ""
                Else
                    Get_str_Loggie = str
                End If

            ElseIf .Untersuchung = Prognose Then

                'If .Prognose.Bauteil = Loggie Then
                lnw = .Prognose.Loggia.Pegel.Pegel
                C = .Prognose.Loggia.Pegel.C

                If lnw > 0 And lnw <= 100 And _
                        (IsNumeric(C) Or C = "") Then
                    Get_str_Loggie = Enter_Loggie(CInt(lnw), "rechnerisch", C)
                Else
                    Get_str_Loggie = ""
                    Exit Function
                End If
                'Else
                '    Get_str_Loggie = ""
                '    Exit Function
                'End If
            Else 'If .Cells(74, 7).Value = "nicht vorhanden" Or .Cells(74, 7).Value = "" Then
                Get_str_Loggie = ""
                Exit Function
            End If

        End With


    End Function
    Private Function Enter_Loggie(ByVal lnw As Integer, ByVal MV As String, ByVal sC As String) As String
        Dim str As String


        str = "<loggie>"

        str = str & "<trittschallpegel>"

        str = str & "<ln-w>"
        str = str & lnw
        str = str & "</ln-w>"

        str = str & "<nachweis>"

        str = str & "<verfahren>"
        str = str & MV
        str = str & "</verfahren>"

        str = str & "<ln-wPlusC>"
        If IsNumeric(sC) Then ' <> "" Then
            Dim c As Single = CDec(sC)
            If (lnw <= 28 And lnw + c <= 28) Or (lnw <= 34 And lnw + c <= 34) Or (lnw <= 40 And lnw + c <= 40) Or _
                (lnw <= 46 And lnw + c <= 46) Or (lnw <= 53 And lnw + c <= 53) Or (lnw <= 60 And lnw + c <= 60) Then
                str = str & "ja"
            Else
                str = str & "nein"
            End If
        Else
            str = str & "nein"
        End If
        str = str & "</ln-wPlusC>"

        str = str & "</nachweis>"

        str = str & "</trittschallpegel>"

        str = str & "</loggie>"

        Enter_Loggie = str

    End Function

    Public Function Get_str_Terrasse() As String
        Dim str As String = ""
        Dim iMes As Integer
        Dim lnw As Single
        Dim C As String = ""
        Dim MV As String = ""

        With Projekt.TS_BLT

            If .Untersuchung = Messung Then
                Dim bGueltig As Boolean
                bGueltig = False

                For iMes = 0 To 5
                    Dim bTerrasse As Boolean = False
                    If iMes = 0 Then
                        If .Messung.Messung_1.Bauteil = Terrasse Then bTerrasse = True
                        'MV = Get_Str_MV(.Messung.Messung_1.Messverfahren)
                        lnw = .Messung.Messung_1.Pegel.Pegel
                        C = .Messung.Messung_1.Pegel.C
                    ElseIf iMes = 1 Then
                        If .Messung.Messung_2.Bauteil = Terrasse Then bTerrasse = True
                        'MV = Get_Str_MV(.Messung.Messung_2.Messverfahren)
                        lnw = .Messung.Messung_2.Pegel.Pegel
                        C = .Messung.Messung_2.Pegel.C
                    ElseIf iMes = 2 Then
                        If .Messung.Messung_3.Bauteil = Terrasse Then bTerrasse = True
                        'MV = Get_Str_MV(.Messung.Messung_3.Messverfahren)
                        lnw = .Messung.Messung_3.Pegel.Pegel
                        C = .Messung.Messung_3.Pegel.C
                    ElseIf iMes = 3 Then
                        If .Messung.Messung_4.Bauteil = Terrasse Then bTerrasse = True
                        'MV = Get_Str_MV(.Messung.Messung_4.Messverfahren)
                        lnw = .Messung.Messung_4.Pegel.Pegel
                        C = .Messung.Messung_4.Pegel.C
                    ElseIf iMes = 4 Then
                        If .Messung.Messung_5.Bauteil = Terrasse Then bTerrasse = True
                        'MV = Get_Str_MV(.Messung.Messung_5.Messverfahren)
                        lnw = .Messung.Messung_5.Pegel.Pegel
                        C = .Messung.Messung_5.Pegel.C
                    ElseIf iMes = 5 Then
                        If .Messung.Messung_6.Bauteil = Terrasse Then bTerrasse = True
                        'MV = Get_Str_MV(.Messung.Messung_6.Messverfahren)
                        lnw = .Messung.Messung_6.Pegel.Pegel
                        C = .Messung.Messung_6.Pegel.C
                    End If

                    If bTerrasse Then
                        If MV = "KMV" And lnw > 0 And lnw <= 100 Then

                            bGueltig = True
                            str = str & Enter_Terrasse(CInt(lnw), "Kurzmessverfahren", C)

                        ElseIf MV = "NMV" Then
                            If lnw > 0 And lnw <= 100 And _
                                (IsNumeric(C) Or C = "") Then

                                bGueltig = True
                                str = str & Enter_Terrasse(CInt(lnw), "Normmessverfahren", C)
                            End If
                        End If
                    End If
                Next

                If bGueltig = False Then
                    Get_str_Terrasse = ""
                Else
                    Get_str_Terrasse = str
                End If

            ElseIf .Untersuchung = Prognose Then

                'If .Prognose.Bauteil = Terrasse Then
                lnw = .Prognose.Terrasse.Pegel.Pegel
                C = .Prognose.Terrasse.Pegel.C

                If lnw > 0 And lnw <= 100 And _
                        (IsNumeric(C) Or C = "") Then
                    Get_str_Terrasse = Enter_Terrasse(CInt(lnw), "rechnerisch", C)
                Else
                    Get_str_Terrasse = ""
                    Exit Function
                End If
                'Else
                '    Get_str_Terrasse = ""
                '    Exit Function
                'End If
            Else 'If .Cells(74, 7).Value = "nicht vorhanden" Or .Cells(74, 7).Value = "" Then
                Get_str_Terrasse = ""
                Exit Function
            End If

        End With

    End Function

    Private Function Enter_Terrasse(ByVal lnw As Integer, ByVal MV As String, ByVal sC As String) As String
        Dim str As String

        str = "<terasse>"

        str = str & "<trittschallpegel>"

        str = str & "<ln-w>"
        str = str & lnw
        str = str & "</ln-w>"

        str = str & "<nachweis>"

        str = str & "<verfahren>"
        str = str & MV
        str = str & "</verfahren>"

        str = str & "<ln-wPlusC>"
        If IsNumeric(sC) Then ' <> "" Then
            Dim c As Single = CDec(sC)
            If (lnw <= 28 And lnw + c <= 28) Or (lnw <= 34 And lnw + c <= 34) Or (lnw <= 40 And lnw + c <= 40) Or _
                (lnw <= 46 And lnw + c <= 46) Or (lnw <= 53 And lnw + c <= 53) Or (lnw <= 60 And lnw + c <= 60) Then
                str = str & "ja"
            Else
                str = str & "nein"
            End If
        Else
            str = str & "nein"
        End If
        str = str & "</ln-wPlusC>"

        str = str & "</nachweis>"

        str = str & "</trittschallpegel>"

        str = str & "</terasse>"

        Enter_Terrasse = str

    End Function


    Public Function Get_str_Tueren() As String
        Dim str As String
        str = ""

        With Projekt.Tueren

            If .Ort = Diele Then
                If .L > 0 And .L <= 100 And (.Untersuchung = Pruefzeugnis Or .Untersuchung = Messung) Then

                    str = str & "<whgEingangstuerFlurDiele>"
                    str = str & "<Rw>"
                    str = str & .L.ToString
                    str = str & "</Rw>"


                    str = str & "<nachweisTuer>"
                    If .Untersuchung = Pruefzeugnis Then
                        str = str & "pruefzeugnis"
                    ElseIf .Untersuchung = Messung Then
                        str = str & "messung"
                    End If
                    str = str & "</nachweisTuer>"

                    str = str & "</whgEingangstuerFlurDiele>"
                End If

            ElseIf .Ort = Aufenthalt Then

                If .L > 0 And .L <= 100 And _
                        (.Untersuchung = Pruefzeugnis Or .Untersuchung = Messung) Then

                    str = str & "<whgEingangstuerInAufenthaltsraum>"
                    str = str & "<Rw>"
                    str = str & .L.ToString
                    str = str & "</Rw>"

                    str = str & "<nachweisTuer>"
                    If .Untersuchung = Pruefzeugnis Then
                        str = str & "pruefzeugnis"
                    ElseIf .Untersuchung = Messung Then
                        str = str & "messung"
                    End If
                    str = str & "</nachweisTuer>"

                    str = str & "</whgEingangstuerInAufenthaltsraum>"

                End If
            End If
        End With
        Get_str_Tueren = str
    End Function

    Public Function Get_str_LSAussenbauteile() As String
        Dim str As String

        With Projekt

            If .Aussenbauteile = ohneNachweis Then
                str = "<luftschallaussenbauteile>ohne Nachweis</luftschallaussenbauteile>"
            ElseIf .Aussenbauteile = FensterMitDichtung Then
                str = "<luftschallaussenbauteile>Fenster mit Dichtung</luftschallaussenbauteile>"
            ElseIf .Aussenbauteile = DINerfuellt Then
                str = "<luftschallaussenbauteile>DIN 4109</luftschallaussenbauteile>"
            ElseIf .Aussenbauteile = DINPlusErfuellt Then
                str = "<luftschallaussenbauteile>DIN 4109 +5 dB</luftschallaussenbauteile>"
            Else
                str = ""
            End If

        End With

        Get_str_LSAussenbauteile = str
    End Function

    Public Function Get_str_Wasserinstallation() As String
        Dim str As String
        str = ""

        With Projekt.Wasser

            Dim sU As Byte
            sU = .Untersuchung
            Dim Laf As Byte = .L_Intervall
            Dim sLcLa As Byte = .LcLa_erfuellt

            If Laf > 0 And Laf <= HT_groesser35 And (sU = Messung Or sU = Prognose) And sLcLa > 0 And sLcLa <= keineAngabe Then
                str = str & "<wasserinstallation>"

                str = str & "<lafmaxn>"

                If Laf = HT_kleiner20 Then
                    str = str & "20"
                ElseIf Laf = HT_20_24 Then
                    str = str & "24"
                ElseIf Laf = HT_24_27 Then
                    str = str & "27"
                ElseIf Laf = HT_27_30 Then
                    str = str & "30"
                ElseIf Laf = HT_30_35 Then
                    str = str & "35"
                ElseIf Laf = HT_groesser35 Then
                    str = str & "36"
                Else    'Bei keiner Angabe
                    Get_str_Wasserinstallation = "" 'str & "0"
                    Exit Function
                End If

                str = str & "</lafmaxn>"


                str = str & "<nachweisInstall>"

                If sU = Messung Then
                    str = str & "Messung"
                ElseIf sU = Prognose Then
                    str = str & "Prognose"
                End If

                str = str & "</nachweisInstall>"


                '*** NEU nach neuem Schema
                str = str & "<lcla>"

                If sLcLa = erfuellt Then
                    str = str & "ja"
                Else
                    str = str & "nein"
                End If

                str = str & "</lcla>"


                str = str & "</wasserinstallation>"

            End If
        End With

        Get_str_Wasserinstallation = str
    End Function

    Public Function Get_str_Nutzergeraeusche() As String
        Dim str As String

        With Projekt.NutzerKoerper

            str = "<nutzerGraeusche>"
            'str = "<nutzerKopplung>"
            If .Untersuchung = Nutzer Then
                Dim sKrit As Byte
                Dim Laf As Byte
                sKrit = .MessungPrognose '.Cells(row_Nutzer, 7).Value
                Laf = .L_Intervall '.Cells(row_Nutzer_Laf, 8).Value

                'str = str & "<L>"
                If sKrit = Prognose Or sKrit = Messung Then
                    If Laf = kleiner20 Then
                        str = str & "<L>20</L>"
                    ElseIf Laf = L_20_25 Then
                        str = str & "<L>25</L>"
                    ElseIf Laf = L_25_30 Then
                        str = str & "<L>30</L>"
                    ElseIf Laf = L_30_35 Then
                        str = str & "<L>35</L>"
                    ElseIf Laf = L_35_40 Then
                        str = str & "<L>40</L>"
                    ElseIf Laf = L_40_45 Then
                        str = str & "<L>45</L>"
                    ElseIf Laf = groesser45 Then
                        str = str & "<L>46</L>"
                    Else    'Bei keiner Angabe
                        Get_str_Nutzergeraeusche = "" 'str & "0"
                        Exit Function
                    End If
                ElseIf sKrit <> nv Then
                    Get_str_Nutzergeraeusche = "" ' str & "0"
                    Exit Function
                End If
                'str = str & "</L>"

                str = str & "<kriterium>"
                If sKrit = Prognose Then
                    str = str & "Prognose"
                ElseIf sKrit = Messung Then
                    str = str & "Messung"
                ElseIf sKrit = nv Then
                    str = str & "Kein Nachweis"
                Else
                    Get_str_Nutzergeraeusche = ""
                    Exit Function
                End If
                str = str & "</kriterium>"
            End If

            str = str & "</nutzerGraeusche>"

        End With

        Get_str_Nutzergeraeusche = str
    End Function

    Public Function Get_str_Koerperschallentkopplung() As String
        Dim str As String

        With Projekt.NutzerKoerper
            Dim sKrit As Byte
            Dim Laf As Byte

            str = "<koerperschallentkopplung>"
            'str = "<nutzerKopplung>"
            If .Untersuchung = Koerper Then
                sKrit = .MessungPrognose
                Laf = .L_Intervall

                'str = str & "<L>"
                If sKrit = Prognose Or sKrit = Messung Then
                    If Laf = kleiner38 Then
                        str = str & "<L>38</L>"
                    ElseIf Laf = L_38_43 Then
                        str = str & "<L>43</L>"
                    ElseIf Laf = L_43_48 Then
                        str = str & "<L>48</L>"
                    ElseIf Laf = L_48_53 Then
                        str = str & "<L>53</L>"
                    ElseIf Laf = L_53_58 Then
                        str = str & "<L>58</L>"
                    ElseIf Laf = L_58_63 Then
                        str = str & "<L>63</L>"
                    ElseIf Laf = groesser63 Then
                        str = str & "<L>64</L>"
                    Else
                        Get_str_Koerperschallentkopplung = "" 'str & "0"
                        Exit Function
                    End If
                ElseIf sKrit <> nv Then
                    Get_str_Koerperschallentkopplung = "" 'str & "0"
                    Exit Function
                End If


                str = str & "<kriterium>"
                If sKrit = Prognose Then
                    str = str & "Prognose"
                ElseIf sKrit = Messung Then
                    str = str & "Messung"
                ElseIf sKrit = nv Then
                    str = str & "Kein Nachweis"
                Else
                    Get_str_Koerperschallentkopplung = ""
                    Exit Function
                End If
                str = str & "</kriterium>"
            End If
            str = str & "</koerperschallentkopplung>"

        End With

        Get_str_Koerperschallentkopplung = str
    End Function

    Public Function Get_str_Grundrisssituation() As String
        Dim str As String
        Dim slr As Integer
        Dim slmax As Integer
        With Projekt

            str = "<grundrisssituation>"
            'str = "<grundriss>"

            str = str & "<angrenzendeNutzer>"
            If .Nachbarn = 1 Then
                str = str & "1"
            ElseIf .Nachbarn = 2 Then
                str = str & "2"
            ElseIf .Nachbarn = 3 Then
                str = str & "3"
            ElseIf .Nachbarn = 4 Then
                str = str & "4"
            ElseIf .Nachbarn = 5 Then
                str = str & "5"
            Else
                Get_str_Grundrisssituation = "" 'str & "0"
                Exit Function
            End If
            str = str & "</angrenzendeNutzer>"

            str = str & "<anordnungLauterRaeume>"
            If .anordnungRaeume = guenstig Then
                str = str & "guenstig"
            ElseIf .anordnungRaeume = unguenstig Then
                str = str & "unguenstig"
            Else
                Get_str_Grundrisssituation = ""
                Exit Function
            End If
            str = str & "</anordnungLauterRaeume>"

            str = str & "<angrenzendLauteRaeume>"

            str = str & "<keine>"
            If .lauteRaeume = keineAngrenzend Then
                str = str & "ja"
            Else
                str = str & "nein"
            End If
            str = str & "</keine>"

            If .lauteRaeume = keineAngrenzend Then
                sLr = 0
                sLmax = 0
            ElseIf .lauteRaeume = L_25_35 Then
                sLr = 15
                sLmax = 25
            ElseIf .lauteRaeume = L_30_40 Then
                sLr = 20
                sLmax = 30
            ElseIf .lauteRaeume = L_35_45 Then
                sLr = 25
                sLmax = 35
            Else
                Get_str_Grundrisssituation = ""
                Exit Function
            End If

            str = str & "<lr>"
            str = str & sLr
            str = str & "</lr>"

            str = str & "<lmax>"
            str = str & sLmax
            str = str & "</lmax>"

            str = str & "</angrenzendLauteRaeume>"

            str = str & "</grundrisssituation>"

        End With

        Get_str_Grundrisssituation = str
    End Function

    Public Function Get_str_EW() As String
        Dim str As String

        With Projekt

            str = "<eigenerWohnbereich>"
            If .eigenerWohnbereich = keineEmpfehlung Then
                str = str & "keine Empfehlung vereinbart"
            ElseIf .eigenerWohnbereich = EW1 Then
                str = str & "EW1 erfuellt"
            ElseIf .eigenerWohnbereich = EW2 Then
                str = str & "EW2 erfuellt"
            Else
                Get_str_EW = ""
                Exit Function
            End If
            str = str & "</eigenerWohnbereich>"

        End With

        Get_str_EW = str
    End Function
End Module
