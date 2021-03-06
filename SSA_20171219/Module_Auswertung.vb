Module Module_Auswertung

#Region "Get_Byte_*"
    Public Function Get_Byte_Geschosstyp(ByVal sTyp As String) As Byte
        Select Case sTyp
            Case "Untergeschoß"
                Return UG
            Case "Erdgeschoß"
                Return EG
            Case "Obergeschoß"
                Return OG
            Case "Dachgeschoß"
                Return DG
            Case "Sonstiges"
                Return Sonstiges
            Case Else
                Return 0
        End Select
    End Function
    Public Function Get_Byte_GC(ByVal sGC As String) As Byte
        Select Case sGC
            Case "reines Wohngebiet"
                Return GC_WR
            Case "allgemeines Wohngebiet"
                Return GC_WA
            Case "Mischgebiet / besonderes Wohngebiet"
                Return GC_MIWB
            Case "Gewerbegebiet"
                Return GC_GE
            Case "Industriegebiet"
                Return GC_GI
            Case Else
                Return 0
        End Select
    End Function
    Public Function Get_Byte_AP(ByVal sAP As String) As Byte
        Select Case sAP
            Case "bis 55"
                Return AP_bis55
            Case "56 - 60"
                Return AP_56_60
            Case "61 - 65"
                Return AP_61_65
            Case "66 - 70"
                Return AP_66_70
            Case "71 - 75"
                Return AP_71_75
            Case ">= 76"
                Return AP_76
            Case Else
                Return 0
        End Select
    End Function
    Public Function Get_Byte_aF(ByVal sAF As String) As Byte
        Select Case sAF
            Case "ja"
                Return aF_ja
            Case "nein"
                Return aF_nein
            Case Else
                Return 0
        End Select
    End Function
    '' ''Public Function Get_Byte_Messanteil(ByVal sMA As String) As Byte
    '' ''    Select Case sMA
    '' ''        Case "< 50 %"
    '' ''            Return kleiner50
    '' ''        Case ">= 50 %"
    '' ''            Return groesser50
    '' ''        Case Else
    '' ''            Return 0
    '' ''    End Select
    '' ''End Function
    '' ''Public Function Get_Byte_Messverfahren(ByVal sMV As String) As Byte
    '' ''    Select Case sMV
    '' ''        Case "KMV"
    '' ''            Return KMV
    '' ''        Case "NMV"
    '' ''            Return NMV
    '' ''        Case Else
    '' ''            Return 0
    '' ''    End Select
    '' ''End Function
    Public Function Get_Byte_Untersuchung(ByVal sUnters As String) As Byte
        Select Case sUnters
            Case "Prognose"
                Return Prognose
            Case "Messung"
                Return Messung
            Case "nicht vorhanden", "ohne Nachweis"
                Return nv
            Case Else
                Return 0
        End Select
    End Function
    Public Function Get_Byte_Bauteil_TPH(ByVal sTPH As String) As Byte
        Select Case sTPH
            Case "Treppe"
                Return Treppe
            Case "Podest"
                Return Podest
            Case "Hausflur"
                Return Hausflur
            Case Else
                Return 0
        End Select
    End Function
    Public Function Get_Byte_Bauteil_BLLT(ByVal sBLLT As String) As Byte
        Select Case sBLLT
            Case "Balkon"
                Return Balkon
            Case "Laubengang"
                Return Laube
            Case "Loggia"
                Return Loggie
            Case "Terrasse"
                Return Terrasse
            Case Else
                Return 0
        End Select
    End Function
    Public Function Get_Byte_Tueren_Ort(ByVal sTuerenOrt As String) As Byte
        Select Case sTuerenOrt
            Case "in Flur oder Dielen"
                Return Diele
            Case "in Aufenthaltsräumen"
                Return Aufenthalt
            Case "nicht vorhanden"
                Return nv
            Case Else
                Return 0
        End Select
    End Function
    Public Function Get_Byte_Tueren_Untersuchung(ByVal sUnters As String) As Byte
        Select Case sUnters
            Case "Prüfzeugnis (Rechenwert)"
                Return Pruefzeugnis
            Case "Messung am Bau"
                Return Messung
            Case Else
                Return 0
        End Select
    End Function
    Public Function Get_Byte_Aussenbauteile(ByVal sAB As String) As Byte
        Select Case sAB
            Case "ohne Nachweis"
                Return ohneNachweis
            Case "ohne Nachweis, Fenster mit Dichtungen"
                Return FensterMitDichtung
            Case "Anforderung nach DIN 4109 erfüllt"
                Return DINerfuellt
            Case "Anforderung nach DIN 4109 + 5 dB erfüllt"
                Return DINPlusErfuellt
            Case Else
                Return 0
        End Select
    End Function
    'Wasser
    Public Function Get_Byte_L_Intervall_Wasser(ByVal sIntervall As String) As Byte
        Select Case sIntervall
            Case "≤ 20"
                Return HT_kleiner20
            Case "20 < L ≤ 24"
                Return HT_20_24
            Case "24 < L ≤ 27"
                Return HT_24_27
            Case "27 < L ≤ 30"
                Return HT_27_30
            Case "30 < L ≤ 35"
                Return HT_30_35
            Case "> 35"
                Return HT_groesser35
            Case Else
                Return 0
        End Select
    End Function
    Public Function Get_Byte_Wasser_LcLa(ByVal sLcLa As String) As Byte
        Select Case sLcLa
            Case "erfüllt"
                Return erfuellt
            Case "nicht erfüllt", ""
                Return keineAngabe
            Case Else
                Return 0
        End Select
    End Function
    'Nutzer
    Public Function Get_Byte_L_Intervall_Nutzer(ByVal sIntervall As String) As Byte
        Select Case sIntervall
            Case "≤ 20"
                Return kleiner20
            Case "20 < L ≤ 25"
                Return L_20_25
            Case "25 < L ≤ 30"
                Return L_25_30
            Case "30 < L ≤ 35"
                Return L_30_35
            Case "35 < L ≤ 40"
                Return L_35_40
            Case "40 < L ≤ 45"
                Return L_40_45
            Case "> 45"
                Return groesser45
            Case Else
                Return 0
        End Select
    End Function
    'Körper
    Public Function Get_Byte_L_Intervall_Koerper(ByVal sIntervall As String) As Byte
        Select Case sIntervall
            Case "≤ 38"
                Return kleiner38
            Case "38 < L ≤ 43"
                Return L_38_43
            Case "43 < L ≤ 48"
                Return L_43_48
            Case "48 < L ≤ 53"
                Return L_48_53
            Case "53 < L ≤ 58"
                Return L_53_58
            Case "58 < L ≤ 63"
                Return L_58_63
            Case ">63"
                Return groesser63
            Case Else
                Return 0
        End Select
    End Function
    Public Function Get_Byte_Nachbarn(ByVal sNachbarn As String) As Byte
        Select Case sNachbarn
            Case "0 - 1"
                Return 1
            Case "2"
                Return 2
            Case "3"
                Return 3
            Case "4"
                Return 4
            Case ">= 5"
                Return 5
            Case Else
                Return 0
        End Select
    End Function
    Public Function Get_Byte_AnordRaeume(ByVal sAnord As String) As Byte
        Select Case sAnord
            Case "günstig"
                Return guenstig
            Case "ungünstig"
                Return unguenstig
            Case Else
                Return 0
        End Select
    End Function
    Public Function Get_Byte_lauteRaeume(ByVal sLauteRaeume As String) As Byte
        Select Case sLauteRaeume
            Case "keine lauten angrenzenden Gewerberäume"
                Return keineAngrenzend
            Case "Lr  t / n:      25 / 15 dB (A)    und     Lmax  t / n:    35 / 25 dB (A)"
                Return L_25_35
            Case "Lr  t / n:      30 / 20 dB (A)    und     Lmax  t / n:    40 / 30 dB (A)"
                Return L_30_40
            Case "Lr  t / n:      35 / 25 dB (A)    und     Lmax  t / n:    45 / 35 dB (A)"
                Return L_35_45
            Case Else
                Return 0
        End Select
    End Function
    Public Function Get_Byte_EW(ByVal sEW As String) As Byte
        Select Case sEW
            Case "keine Empfehlung vereinbart"
                Return keineEmpfehlung
            Case "Klasse EW1 erfüllt"
                Return EW1
            Case "Klasse EW2 erfüllt"
                Return EW2
            Case Else
                Return 0

        End Select
    End Function
#End Region

#Region "Get_Str_*"
    ''Public Function Get_Str_MV(ByVal MV As Byte) As String
    ''    Select Case MV
    ''        Case KMV
    ''            Return "KMV"
    ''        Case NMV
    ''            Return "NMV"
    ''        Case Else
    ''            Return ""
    ''    End Select
    ''End Function
    ''Public Function Get_Str_MA(ByVal MA As Byte) As String
    ''    Select Case MA
    ''        Case kleiner50
    ''            Return "< 50"
    ''        Case groesser50
    ''            Return ">= 50"
    ''        Case Else
    ''            Return ""
    ''    End Select
    ''End Function

#Region "Standort und Außenlärm"
    Public Function Get_Str_GC() As String
        Select Case Projekt.Standort.Gebietscharakter
            Case GC_WR
                Return "reines Wohngebiet"
            Case GC_WA
                Return "allgemeines Wohngebiet"
            Case GC_MIWB
                Return "Mischgebiet / besonderes Wohngebiet"
            Case GC_GE
                Return "Gewerbegebiet"
            Case GC_GI
                Return "Industriegebiet"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_AP() As String
        Select Case Projekt.Standort.Aussenlaermpegel
            Case AP_bis55
                Return "bis 55"
            Case AP_56_60
                Return "56 - 60"
            Case AP_61_65
                Return "61 - 65"
            Case AP_66_70
                Return "66 - 70"
            Case AP_71_75
                Return "71 - 75"
            Case AP_76
                Return ">= 76"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_aF() As String
        Select Case Projekt.Standort.abgewandFreibereich
            Case aF_ja
                Return "ja"
            Case aF_nein
                Return "nein"
            Case Else
                Return ""
        End Select
    End Function
#End Region
#Region "LS_W"
    Public Function Get_Str_LS_W_P() As String
        Select Case Projekt.LS_Wand.Untersuchung
            Case Prognose
                Return "X"
            Case Messung, nv
                Return "-"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_LS_W_M() As String
        Select Case Projekt.LS_Wand.Untersuchung
            Case Prognose, nv
                Return "-"
            Case Messung
                Return "X"
            Case Else
                Return ""
        End Select
    End Function

    Public Function Get_Str_LS_W_C() As String
        If Projekt.LS_Wand.Untersuchung = Messung Then
            Dim iMes As Integer = Get_Mes_LS_Waende()
            With Projekt.LS_Wand.Messung
                If iMes = 1 Then
                    Return .Messung_1.C '.ToString
                ElseIf iMes = 2 Then
                    Return .Messung_2.C '.ToString
                ElseIf iMes = 3 Then
                    Return .Messung_3.C '.ToString
                ElseIf iMes = 4 Then
                    Return .Messung_4.C '.ToString
                ElseIf iMes = 5 Then
                    Return .Messung_5.C '.ToString
                ElseIf iMes = 6 Then
                    Return .Messung_6.C '.ToString
                Else
                    Return ""
                End If
            End With
        ElseIf Projekt.LS_Wand.Untersuchung = Prognose Then
            Return Projekt.LS_Wand.Prognose.C
        Else
            Return ""
        End If
    End Function

    Public Function Get_Str_LS_W_R() As String
        If Projekt.LS_Wand.Untersuchung = Messung Then
            Dim iMes As Integer = Get_Mes_LS_Waende()
            With Projekt.LS_Wand.Messung
                If iMes = 1 Then
                    Return .Messung_1.Pegel.ToString
                ElseIf iMes = 2 Then
                    Return .Messung_2.Pegel.ToString
                ElseIf iMes = 3 Then
                    Return .Messung_3.Pegel.ToString
                ElseIf iMes = 4 Then
                    Return .Messung_4.Pegel.ToString
                ElseIf iMes = 5 Then
                    Return .Messung_5.Pegel.ToString
                ElseIf iMes = 6 Then
                    Return .Messung_6.Pegel.ToString
                Else
                    Return ""
                End If
            End With
        ElseIf Projekt.LS_Wand.Untersuchung = Prognose Then
            Return Projekt.LS_Wand.Prognose.Pegel.ToString
        Else
            Return ""
        End If
    End Function
    Public Function Get_Str_LS_W_B() As String
        If Projekt.LS_Wand.Untersuchung = nv Then
            Return "nicht vorhanden"
        Else
            Return Form_Main.TB_Bemerkung_LS_W.Text
        End If
    End Function
#End Region
#Region "LS_D"
    Public Function Get_Str_LS_D_P() As String
        Select Case Projekt.LS_Decke.Untersuchung
            Case Prognose
                Return "X"
            Case Messung, nv
                Return "-"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_LS_D_M() As String
        Select Case Projekt.LS_Decke.Untersuchung
            Case Prognose, nv
                Return "-"
            Case Messung
                Return "X"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_LS_D_C() As String
        If Projekt.LS_Decke.Untersuchung = Messung Then
            Dim iMes As Integer = Get_Mes_LS_Decken()
            With Projekt.LS_Decke.Messung
                If iMes = 1 Then
                    Return .Messung_1.C.ToString
                ElseIf iMes = 2 Then
                    Return .Messung_2.C.ToString
                ElseIf iMes = 3 Then
                    Return .Messung_3.C.ToString
                ElseIf iMes = 4 Then
                    Return .Messung_4.C.ToString
                ElseIf iMes = 5 Then
                    Return .Messung_5.C.ToString
                ElseIf iMes = 6 Then
                    Return .Messung_6.C.ToString
                Else
                    Return ""
                End If
            End With
        ElseIf Projekt.LS_Decke.Untersuchung = Prognose Then
            Return Projekt.LS_Decke.Prognose.C
        Else
            Return ""
        End If
    End Function
    Public Function Get_Str_LS_D_R() As String
        If Projekt.LS_Decke.Untersuchung = Messung Then
            Dim iMes As Integer = Get_Mes_LS_Decken()
            With Projekt.LS_Decke.Messung
                If iMes = 1 Then
                    Return .Messung_1.Pegel.ToString
                ElseIf iMes = 2 Then
                    Return .Messung_2.Pegel.ToString
                ElseIf iMes = 3 Then
                    Return .Messung_3.Pegel.ToString
                ElseIf iMes = 4 Then
                    Return .Messung_4.Pegel.ToString
                ElseIf iMes = 5 Then
                    Return .Messung_5.Pegel.ToString
                ElseIf iMes = 6 Then
                    Return .Messung_6.Pegel.ToString
                Else
                    Return ""
                End If
            End With
        ElseIf Projekt.LS_Decke.Untersuchung = Prognose Then
            Return Projekt.LS_Decke.Prognose.Pegel.ToString
        Else
            Return ""
        End If
    End Function
    Public Function Get_Str_LS_D_B() As String
        If Projekt.LS_Decke.Untersuchung = nv Then
            Return "nicht vorhanden"
        Else
            Return Form_Main.TB_Bemerkung_LS_D.Text
        End If
    End Function
#End Region
#Region "TS_D"
    Public Function Get_Str_TS_D_P() As String
        Select Case Projekt.TS_Decke.Untersuchung
            Case Prognose
                Return "X"
            Case Messung, nv
                Return "-"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_TS_D_M() As String
        Select Case Projekt.TS_Decke.Untersuchung
            Case Prognose, nv
                Return "-"
            Case Messung
                Return "X"
            Case Else
                Return ""
        End Select
    End Function

    Public Function Get_Header_TS_D_fE() As String

        If Projekt.TS_Decke.fEstrich = fE_kleiner50 Then
            Return "fᵣ < 50 Hz"
        Else
            Return "fᵣ ≥ 50 Hz"
        End If
    End Function

    Public Function Get_Str_TS_D_fE() As String
        ' Hotfix
        If Projekt.TS_Decke.Bemerkung_TS_Decke = "nicht vorhanden" Then
            Return ""
        End If
        If Projekt.TS_Decke.fEstrich = fE_kleiner50 Then
            Return "X"
        ElseIf Projekt.TS_Decke.fEstrich = fE_groesser50 Then
            Return "X"
        Else
            Return "-"
        End If
    End Function
    Public Function Get_Str_TS_D_Be() As String
        Dim Bodenbelag As Byte = 0
        If Projekt.TS_Decke.Untersuchung = Messung Then
            Dim iMes As Integer = Get_Mes_TS_Decken()
            With Projekt.TS_Decke.Messung
                If iMes = 1 Then
                    Bodenbelag = .Messung_1.Bodenbelag
                ElseIf iMes = 2 Then
                    Bodenbelag = .Messung_2.Bodenbelag
                ElseIf iMes = 3 Then
                    Bodenbelag = .Messung_3.Bodenbelag
                ElseIf iMes = 4 Then
                    Bodenbelag = .Messung_4.Bodenbelag
                ElseIf iMes = 5 Then
                    Bodenbelag = .Messung_5.Bodenbelag
                ElseIf iMes = 6 Then
                    Bodenbelag = .Messung_6.Bodenbelag
                Else
                    Bodenbelag = 0
                End If
            End With
        ElseIf Projekt.TS_Decke.Untersuchung = Prognose Then
            Bodenbelag = Projekt.TS_Decke.Prognose.Bodenbelag
        Else
            Bodenbelag = 0
        End If
        If Bodenbelag = Be_mit Then
            Return "X"
        ElseIf Bodenbelag = Be_ohne Then
            Return "-"
        Else
            Return ""
        End If
    End Function
   Public Function Get_Str_TS_D_C() As String
        If Projekt.TS_Decke.Untersuchung = Messung Then
            Dim iMes As Integer = Get_Mes_TS_Decken()
            With Projekt.TS_Decke.Messung
                If iMes = 1 Then
                    Return .Messung_1.Pegel.C.ToString
                ElseIf iMes = 2 Then
                    Return .Messung_2.Pegel.C.ToString
                ElseIf iMes = 3 Then
                    Return .Messung_3.Pegel.C.ToString
                ElseIf iMes = 4 Then
                    Return .Messung_4.Pegel.C.ToString
                ElseIf iMes = 5 Then
                    Return .Messung_5.Pegel.C.ToString
                ElseIf iMes = 6 Then
                    Return .Messung_6.Pegel.C.ToString
                Else
                    Return ""
                End If
            End With
        ElseIf Projekt.TS_Decke.Untersuchung = Prognose Then
            Return Projekt.TS_Decke.Prognose.Pegel.C
        Else
            Return ""
        End If
    End Function
    Public Function Get_Str_TS_D_L() As String
        If Projekt.TS_Decke.Untersuchung = Messung Then
            Dim iMes As Integer = Get_Mes_TS_Decken()
            With Projekt.TS_Decke.Messung
                If iMes = 1 Then
                    Return .Messung_1.Pegel.Pegel.ToString
                ElseIf iMes = 2 Then
                    Return .Messung_2.Pegel.Pegel.ToString
                ElseIf iMes = 3 Then
                    Return .Messung_3.Pegel.Pegel.ToString
                ElseIf iMes = 4 Then
                    Return .Messung_4.Pegel.Pegel.ToString
                ElseIf iMes = 5 Then
                    Return .Messung_5.Pegel.Pegel.ToString
                ElseIf iMes = 6 Then
                    Return .Messung_6.Pegel.Pegel.ToString
                Else
                    Return ""
                End If
            End With
        ElseIf Projekt.TS_Decke.Untersuchung = Prognose Then
            Return Projekt.TS_Decke.Prognose.Pegel.Pegel.ToString()
        Else
            Return ""
        End If
    End Function
    Public Function Get_Str_TS_D_B() As String
        If Projekt.TS_Decke.Untersuchung = nv Then
            Return "nicht vorhanden"
        Else
            Return Form_Main.TB_Bemerkung_TS_D.Text
        End If
    End Function
#End Region
#Region "TS_TPH"
    Public Function Get_Str_TS_TPLH_P() As String
        Select Case Projekt.TS_TPlH.Untersuchung
            Case Prognose
                Return "X"
            Case Messung, nv
                Return "-"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_TS_TPLH_M() As String
        Select Case Projekt.TS_TPlH.Untersuchung
            Case Prognose, nv
                Return "-"
            Case Messung
                Return "X"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_TS_TPLH_Be() As String
        Dim btBe As Byte = 0
        If Projekt.TS_TPlH.Untersuchung = Messung Then
            Dim iMes As Integer = Get_Mes_TS_TPLH()
            With Projekt.TS_TPlH.Messung
                If iMes = 1 Then
                    btBe = .Messung_1.Bodenbelag
                ElseIf iMes = 2 Then
                    btBe = .Messung_2.Bodenbelag
                ElseIf iMes = 3 Then
                    btBe = .Messung_3.Bodenbelag
                ElseIf iMes = 4 Then
                    btBe = .Messung_4.Bodenbelag
                ElseIf iMes = 5 Then
                    btBe = .Messung_5.Bodenbelag
                ElseIf iMes = 6 Then
                    btBe = .Messung_6.Bodenbelag
                    '               Else
                    '                    Return ""
                End If
            End With
        ElseIf Projekt.TS_TPlH.Untersuchung = Prognose Then
            Dim btBauteil As Byte = Get_Prog_TS_TPLH()
            If btBauteil = Treppe Then
                btBe = Projekt.TS_TPlH.Prognose.Treppe.Bodenbelag
            ElseIf btBauteil = Podest Then
                btBe = Projekt.TS_TPlH.Prognose.Podest.Bodenbelag
            ElseIf btBauteil = Laube Then
                btBe = Projekt.TS_TPlH.Prognose.Laube.Bodenbelag
            ElseIf btBauteil = Hausflur Then
                btBe = Projekt.TS_TPlH.Prognose.Hausflur.Bodenbelag
                'Else
                'Return ""
            End If
        Else
            'Return ""
        End If
        If btBe = Be_mit Then
            Return "X"
        ElseIf btBe = Be_ohne Then
            Return "-"
        Else
            Return ""
        End If
    End Function
    Public Function Get_Str_TS_TPLH_C() As String
        If Projekt.TS_TPlH.Untersuchung = Messung Then
            Dim iMes As Integer = Get_Mes_TS_TPLH()
            With Projekt.TS_TPlH.Messung
                If iMes = 1 Then
                    Return .Messung_1.Pegel.C.ToString
                ElseIf iMes = 2 Then
                    Return .Messung_2.Pegel.C.ToString
                ElseIf iMes = 3 Then
                    Return .Messung_3.Pegel.C.ToString
                ElseIf iMes = 4 Then
                    Return .Messung_4.Pegel.C.ToString
                ElseIf iMes = 5 Then
                    Return .Messung_5.Pegel.C.ToString
                ElseIf iMes = 6 Then
                    Return .Messung_6.Pegel.C.ToString
                Else
                    Return ""
                End If
            End With
        ElseIf Projekt.TS_TPlH.Untersuchung = Prognose Then
            Dim btBauteil As Byte = Get_Prog_TS_TPLH()
            If btBauteil = Treppe Then
                Return Projekt.TS_TPlH.Prognose.Treppe.Pegel.C
            ElseIf btBauteil = Podest Then
                Return Projekt.TS_TPlH.Prognose.Podest.Pegel.C
            ElseIf btBauteil = Laube Then
                Return Projekt.TS_TPlH.Prognose.Laube.Pegel.C
            ElseIf btBauteil = Hausflur Then
                Return Projekt.TS_TPlH.Prognose.Hausflur.Pegel.C
            Else
                Return ""
            End If
        Else
            Return ""
        End If
    End Function
    Public Function Get_Str_TS_TPLH_L() As String
        If Projekt.TS_TPlH.Untersuchung = Messung Then
            Dim iMes As Integer = Get_Mes_TS_TPLH()
            With Projekt.TS_TPlH.Messung
                If iMes = 1 Then
                    Return .Messung_1.Pegel.Pegel.ToString
                ElseIf iMes = 2 Then
                    Return .Messung_2.Pegel.Pegel.ToString
                ElseIf iMes = 3 Then
                    Return .Messung_3.Pegel.Pegel.ToString
                ElseIf iMes = 4 Then
                    Return .Messung_4.Pegel.Pegel.ToString
                ElseIf iMes = 5 Then
                    Return .Messung_5.Pegel.Pegel.ToString
                ElseIf iMes = 6 Then
                    Return .Messung_6.Pegel.Pegel.ToString
                Else
                    Return ""
                End If
            End With
        ElseIf Projekt.TS_TPlH.Untersuchung = Prognose Then
            Dim btBauteil As Byte = Get_Prog_TS_TPLH()
            If btBauteil = Treppe Then
                Return Projekt.TS_TPlH.Prognose.Treppe.Pegel.Pegel.ToString
            ElseIf btBauteil = Podest Then
                Return Projekt.TS_TPlH.Prognose.Podest.Pegel.Pegel.ToString
            ElseIf btBauteil = Laube Then
                Return Projekt.TS_TPlH.Prognose.Laube.Pegel.Pegel.ToString
            ElseIf btBauteil = Hausflur Then
                Return Projekt.TS_TPlH.Prognose.Hausflur.Pegel.Pegel.ToString
            Else
                Return ""
            End If
        Else
            Return ""
        End If
    End Function
    Public Function Get_Str_TS_TPH_B() As String
        If Projekt.TS_TPlH.Untersuchung = nv Then
            Return Form_Main.TB_Bemerkung_TS_TPH.Text
        Else
            Return Form_Main.TB_Bemerkung_TS_TPH.Text
        End If
    End Function
#End Region
#Region "TS_BLLT"
    Public Function Get_Str_TS_BLT_P() As String
        Select Case Projekt.TS_BLT.Untersuchung
            Case Prognose
                Return "X"
            Case Messung, nv
                Return "-"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_TS_BLT_M() As String
        Select Case Projekt.TS_BLT.Untersuchung
            Case Prognose, nv
                Return "-"
            Case Messung
                Return "X"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_TS_BLT_Be() As String
        Dim btBE As Byte = 0
        If Projekt.TS_BLT.Untersuchung = Messung Then
            Dim iMes As Integer = Get_Mes_TS_BLT()
            With Projekt.TS_BLT.Messung
                If iMes = 1 Then
                    btBE = .Messung_1.Bodenbelag
                ElseIf iMes = 2 Then
                    btBE = .Messung_2.Bodenbelag
                ElseIf iMes = 3 Then
                    btBE = .Messung_3.Bodenbelag
                ElseIf iMes = 4 Then
                    btBE = .Messung_4.Bodenbelag
                ElseIf iMes = 5 Then
                    btBE = .Messung_5.Bodenbelag
                ElseIf iMes = 6 Then
                    btBE = .Messung_6.Bodenbelag
                    'Else
                    'Return ""
                End If
            End With
        ElseIf Projekt.TS_BLT.Untersuchung = Prognose Then
            Dim btBauteil As Byte = Get_Prog_TS_BLT()
            If btBauteil = Balkon Then
                btBE = Projekt.TS_BLT.Prognose.Balkon.Bodenbelag
            ElseIf btBauteil = Loggie Then
                btBE = Projekt.TS_BLT.Prognose.Loggia.Bodenbelag
            ElseIf btBauteil = Terrasse Then
                btBE = Projekt.TS_BLT.Prognose.Terrasse.Bodenbelag
                'Else
                '    Return ""
            End If
            'Else
            'Return ""
        End If
        If btBE = Be_mit Then
            Return "X"
        ElseIf btBE = Be_ohne Then
            Return "-"
        Else
            Return ""
        End If
    End Function
    Public Function Get_Str_TS_BLT_C() As String
        If Projekt.TS_BLT.Untersuchung = Messung Then
            Dim iMes As Integer = Get_Mes_TS_BLT()
            With Projekt.TS_BLT.Messung
                If iMes = 1 Then
                    Return .Messung_1.Pegel.C.ToString
                ElseIf iMes = 2 Then
                    Return .Messung_2.Pegel.C.ToString
                ElseIf iMes = 3 Then
                    Return .Messung_3.Pegel.C.ToString
                ElseIf iMes = 4 Then
                    Return .Messung_4.Pegel.C.ToString
                ElseIf iMes = 5 Then
                    Return .Messung_5.Pegel.C.ToString
                ElseIf iMes = 6 Then
                    Return .Messung_6.Pegel.C.ToString
                Else
                    Return ""
                End If
            End With
        ElseIf Projekt.TS_BLT.Untersuchung = Prognose Then
            Dim btBauteil As Byte = Get_Prog_TS_BLT()
            If btBauteil = Balkon Then
                Return Projekt.TS_BLT.Prognose.Balkon.Pegel.C
            ElseIf btBauteil = Loggie Then
                Return Projekt.TS_BLT.Prognose.Loggia.Pegel.C
            ElseIf btBauteil = Terrasse Then
                Return Projekt.TS_BLT.Prognose.Terrasse.Pegel.C
            Else
                Return ""
            End If
        Else
            Return ""
        End If
    End Function
    Public Function Get_Str_TS_BLT_L() As String
        If Projekt.TS_BLT.Untersuchung = Messung Then
            Dim iMes As Integer = Get_Mes_TS_BLT()
            With Projekt.TS_BLT.Messung
                If iMes = 1 Then
                    Return .Messung_1.Pegel.Pegel.ToString
                ElseIf iMes = 2 Then
                    Return .Messung_2.Pegel.Pegel.ToString
                ElseIf iMes = 3 Then
                    Return .Messung_3.Pegel.Pegel.ToString
                ElseIf iMes = 4 Then
                    Return .Messung_4.Pegel.Pegel.ToString
                ElseIf iMes = 5 Then
                    Return .Messung_5.Pegel.Pegel.ToString
                ElseIf iMes = 6 Then
                    Return .Messung_6.Pegel.Pegel.ToString
                Else
                    Return ""
                End If
            End With
        ElseIf Projekt.TS_BLT.Untersuchung = Prognose Then
            Dim btBauteil As Byte = Get_Prog_TS_BLT()
            If btBauteil = Balkon Then
                Return Projekt.TS_BLT.Prognose.Balkon.Pegel.Pegel.ToString
            ElseIf btBauteil = Loggie Then
                Return Projekt.TS_BLT.Prognose.Loggia.Pegel.Pegel.ToString
            ElseIf btBauteil = Terrasse Then
                Return Projekt.TS_BLT.Prognose.Terrasse.Pegel.Pegel.ToString
            Else
                Return ""
            End If
        Else
            Return ""
        End If
    End Function
    Public Function Get_Str_TS_BLLT_B() As String
        If Projekt.TS_BLT.Untersuchung = nv Then
            Return "nicht vorhanden"
        Else
            Return Form_Main.TB_Bemerkung_TS_BLLT.Text
        End If
    End Function
#End Region
#Region "Tueren"
    Public Function Get_Str_Tueren_P_D() As String
        If Projekt.Tueren.Untersuchung = Pruefzeugnis And Projekt.Tueren.Ort = Diele Then
            Return "X"
        ElseIf (Projekt.Tueren.Untersuchung > 0 And Projekt.Tueren.Ort > 0) Or Projekt.Tueren.Ort = nv Then
            Return "-"
        Else
            Return ""
        End If
    End Function
    Public Function Get_Str_Tueren_M_D() As String
        If Projekt.Tueren.Untersuchung = Messung And Projekt.Tueren.Ort = Diele Then
            Return "X"
        ElseIf (Projekt.Tueren.Untersuchung > 0 And Projekt.Tueren.Ort > 0) Or Projekt.Tueren.Ort = nv Then
            Return "-"
        Else
            Return ""
        End If
    End Function
    Public Function Get_Str_Tueren_R_D() As String
        With Projekt.Tueren
            If .Ort = Diele And (.Untersuchung = Pruefzeugnis Or .Untersuchung = Messung) Then
                Return .L.ToString
            ElseIf (.Ort = Aufenthalt And (.Untersuchung = Pruefzeugnis Or .Untersuchung = Messung)) Or .Ort = nv Then
                Return "-"
            Else
                Return ""
            End If
        End With
    End Function

    Public Function Get_Str_Tueren_P_A() As String
        If Projekt.Tueren.Untersuchung = Pruefzeugnis And Projekt.Tueren.Ort = Aufenthalt Then
            Return "X"
        ElseIf (Projekt.Tueren.Untersuchung > 0 And Projekt.Tueren.Ort > 0) Or Projekt.Tueren.Ort = nv Then
            Return "-"
        Else
            Return ""
        End If
    End Function
    Public Function Get_Str_Tueren_M_A() As String
        If Projekt.Tueren.Untersuchung = Messung And Projekt.Tueren.Ort = Aufenthalt Then
            Return "X"
        ElseIf (Projekt.Tueren.Untersuchung > 0 And Projekt.Tueren.Ort > 0) Or Projekt.Tueren.Ort = nv Then
            Return "-"
        Else
            Return ""
        End If
    End Function
    Public Function Get_Str_Tueren_R_A() As String
        With Projekt.Tueren
            If .Ort = Aufenthalt And (.Untersuchung = Pruefzeugnis Or .Untersuchung = Messung) Then
                Return .L.ToString
            ElseIf (.Ort = Diele And (.Untersuchung = Pruefzeugnis Or .Untersuchung = Messung)) Or .Ort = nv Then
                Return "-"
            Else
                Return ""
            End If
        End With
    End Function
    Public Function Get_Str_Tueren_B() As String
        If Projekt.Tueren.Ort <> Aufenthalt And Projekt.Tueren.Ort <> Diele Then
            Return "nicht vorhanden"
        Else
            Return "" 'Form_Main.TB_Bemerkung_Tueren.Text
        End If
    End Function
#End Region
#Region "Außenbauteile"
    Public Function Get_Str_Aussenbauteile_oN() As String
        Select Case Projekt.Aussenbauteile
            Case ohneNachweis
                Return "X"
            Case FensterMitDichtung, DINerfuellt, DINPlusErfuellt
                Return "-"
            Case Else
                Return ""
        End Select
    End Function
    'Public Function Get_Str_Aussenbauteile_FmD() As String
    '    Select Case Projekt.Aussenbauteile
    '        Case FensterMitDichtung
    '            Return "X"
    '        Case ohneNachweis, DINerfuellt, DINPlusErfuellt
    '            Return "-"
    '        Case Else
    '            Return ""
    '    End Select
    'End Function
    Public Function Get_Str_Aussenbauteile_DIN() As String
        Select Case Projekt.Aussenbauteile
            Case DINerfuellt
                Return "X"
            Case ohneNachweis, FensterMitDichtung, DINPlusErfuellt
                Return "-"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_Aussenbauteile_DINPlus() As String
        Select Case Projekt.Aussenbauteile
            Case DINPlusErfuellt
                Return "X"
            Case ohneNachweis, FensterMitDichtung, DINerfuellt
                Return "-"
            Case Else
                Return ""
        End Select
    End Function
#End Region
#Region "Wasser"
    Public Function Get_Str_Wasser_P() As String
        Select Case Projekt.Wasser.Untersuchung
            Case Prognose
                Return "X"
            Case Messung
                Return "-"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_Wasser_M() As String
        Select Case Projekt.Wasser.Untersuchung
            Case Prognose
                Return "-"
            Case Messung
                Return "X"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_Wasser_LcLa() As String
        Select Case Projekt.Wasser.LcLa_erfuellt
            Case erfuellt
                Return "X"
            Case keineAngabe
                Return "-"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_Wasser_L() As String
        With Projekt.Wasser
            Select Case .L_Intervall
                Case HT_kleiner20
                    Return "≤ 20"
                Case HT_20_24
                    Return "20 < L ≤ 24"
                Case HT_24_27
                    Return "24 < L ≤ 27"
                Case HT_27_30
                    Return "27 < L ≤ 30"
                Case HT_30_35
                    Return "30 < L ≤ 35"
                Case HT_groesser35
                    Return "> 35"
                Case Else
                    Return ""
            End Select
        End With
    End Function
    'Public Function Get_Str_Wasser_Laf() As String
    '    If Projekt.Wasser.Untersuchung = Messung Or Projekt.Wasser.Untersuchung = Prognose Then
    '        Return Projekt.Wasser.Laf.ToString
    '    Else
    '        Return ""
    '    End If
    'End Function
#End Region
#Region "Nutzer - Körper"
    Public Function Get_Str_Nutzer_nv() As String
        Select Case Projekt.NutzerKoerper.Untersuchung
            Case nv
                Return "X"
            Case Koerper, Nutzer 'Prognose, Messung
                Return "-"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_Koerper_nv() As String
        Select Case Projekt.NutzerKoerper.Untersuchung
            Case nv
                Return "X"
            Case Koerper, Nutzer
                Return "-"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_Nutzer_P() As String
        With Projekt.NutzerKoerper
            If .Untersuchung = Nutzer Then
                Select Case .MessungPrognose
                    Case Prognose
                        Return "X"
                    Case Messung
                        Return "-"
                    Case Else
                        Return ""
                End Select
            ElseIf .Untersuchung = nv Or .Untersuchung = Koerper Then
                Return "-"
            Else
                Return ""
            End If
        end with
    End Function
    Public Function Get_Str_Koerper_P() As String
        With Projekt.NutzerKoerper
            If .Untersuchung = Koerper Then
                Select Case .MessungPrognose
                    Case Prognose
                        Return "X"
                    Case Messung
                        Return "-"
                    Case Else
                        Return ""
                End Select
            ElseIf .Untersuchung = nv Or .Untersuchung = Nutzer Then
                Return "-"
            Else
                Return ""
            End If
        End With
    End Function
    Public Function Get_Str_Nutzer_M() As String
        With Projekt.NutzerKoerper
            If .Untersuchung = Nutzer Then
                Select Case .MessungPrognose
                    Case Messung
                        Return "X"
                    Case Prognose
                        Return "-"
                    Case Else
                        Return ""
                End Select
            ElseIf .Untersuchung = nv Or .Untersuchung = Koerper Then
                Return "-"
            Else
                Return ""
            End If
        End With
    End Function
    Public Function Get_Str_Koerper_M() As String
        With Projekt.NutzerKoerper
            If .Untersuchung = Koerper Then
                Select Case .MessungPrognose
                    Case Messung
                        Return "X"
                    Case Prognose
                        Return "-"
                    Case Else
                        Return ""
                End Select
            ElseIf .Untersuchung = nv Or .Untersuchung = Nutzer Then
                Return "-"
            Else
                Return ""
            End If
        End With
    End Function
    Public Function Get_Str_Nutzer_L() As String
        With Projekt.NutzerKoerper
            If .Untersuchung = Nutzer Then
                Select Case .L_Intervall
                    Case kleiner20
                        Return "≤ 20"
                    Case L_20_25
                        Return "20 < L ≤ 25"
                    Case L_25_30
                        Return "25 < L ≤ 30"
                    Case L_30_35
                        Return "30 < L ≤ 35"
                    Case L_35_40
                        Return "35 < L ≤ 40"
                    Case L_40_45
                        Return "40 < L ≤ 45"
                    Case groesser45
                        Return "> 45"
                    Case Else
                        Return ""
                End Select
            ElseIf .Untersuchung = nv Or .Untersuchung = Koerper Then
                Return "-"
            Else
                Return ""
            End If
        End With
    End Function
    Public Function Get_Str_Koerper_L() As String
        With Projekt.NutzerKoerper
            If .Untersuchung = Koerper Then
                Select Case .L_Intervall
                    Case kleiner38
                        Return "≤ 38"
                    Case L_38_43
                        Return "38 < L ≤ 43"
                    Case L_43_48
                        Return "43 < L ≤ 48"
                    Case L_48_53
                        Return "48 < L ≤ 53"
                    Case L_53_58
                        Return "53 < L ≤ 58"
                    Case L_58_63
                        Return "58 < L ≤ 63"
                    Case groesser63
                        Return "> 63"
                    Case Else
                        Return ""
                End Select
            ElseIf .Untersuchung = nv Or .Untersuchung = Nutzer Then
                Return "-"
            Else
                Return ""
            End If
        End With
    End Function
    Public Function Get_Str_Nutzer_B() As String
        If Projekt.NutzerKoerper.Untersuchung = nv Then
            Return "nicht vorhanden" '"nicht vorhanden"
            'ElseIf Projekt.Nutzer.Untersuchung = Prognose Or Projekt.Nutzer.Untersuchung = Messung Then
            '    Return "nur Empfehlung"
        Else
            Return ""
        End If
    End Function
    'Public Function Get_Str_Koerper_B() As String
    '    If Projekt.Koerper.Untersuchung = nv Then
    '        Return "ohne Nachweis" '"nicht vorhanden"
    '    ElseIf Projekt.Koerper.Untersuchung = Prognose Or Projekt.Nutzer.Untersuchung = Messung Then
    '        Return "nur Empfehlung"
    '    Else
    '        Return ""
    '    End If
    'End Function
#End Region
#Region "Nachbarn"
    Public Function Get_Str_Nachbar() As String
        Select Case Projekt.Nachbarn
            Case 1
                Return "0 - 1"
            Case 2, 3, 4
                Return Projekt.Nachbarn.ToString
            Case 5
                Return "≥ 5"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_Nachbar_B() As String
        If Projekt.Nachbarn > 0 Then
            Return "nur Empfehlung"
        Else
            Return ""
        End If
    End Function
#End Region
#Region "Anord Räume"
    Public Function Get_Str_anordRaeume_ug() As String
        Select Case Projekt.anordnungRaeume
            Case unguenstig
                Return "X"
            Case guenstig
                Return "-"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_anordRaeume_g() As String
        Select Case Projekt.anordnungRaeume
            Case guenstig
                Return "X"
            Case unguenstig
                Return "-"
            Case Else
                Return ""
        End Select
    End Function
#End Region
#Region "lauteRäume"
    Public Function Get_Str_lauteRaeume_keine() As String
        Select Case Projekt.lauteRaeume
            Case keineAngrenzend
                Return "X"
            Case L_25_35, L_30_40, L_35_45
                Return "-"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_lauteRaeume_L_25_35() As String
        Select Case Projekt.lauteRaeume
            Case L_25_35
                Return "X"
            Case keineAngrenzend, L_30_40, L_35_45
                Return "-"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_lauteRaeume_L_30_40() As String
        Select Case Projekt.lauteRaeume
            Case L_30_40
                Return "X"
            Case keineAngrenzend, L_25_35, L_35_45
                Return "-"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_lauteRaeume_L_35_45() As String
        Select Case Projekt.lauteRaeume
            Case L_35_45
                Return "X"
            Case keineAngrenzend, L_25_35, L_30_40
                Return "-"
            Case Else
                Return ""
        End Select
    End Function
#End Region
#Region "NHZ"
    Public Function Get_Str_Nhz_B() As String
        If Projekt.NHZ > 0 Then
            Return "nur Empfehlung"
        Else
            Return ""
        End If
    End Function
    Public Function Get_Str_NHZ_010() As String
        Select Case Projekt.NHZ
            Case NHZ_010
                Return "X"
            Case NHZ_020
                Return "-"
            Case NHZ_keine
                Return "-"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_NHZ_020() As String
        Select Case Projekt.NHZ
            Case NHZ_020
                Return "X"
            Case NHZ_010
                Return "-"
            Case NHZ_keine
                Return "-"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_NHZ_keine() As String
        Select Case Projekt.NHZ
            Case NHZ_020
                Return "-"
            Case NHZ_010
                Return "-"
            Case NHZ_keine
                Return "X"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_NHZ() As String
        Select Case Projekt.NHZ
            Case NHZ_010
                Return "A/V ≥ 0,10"
            Case NHZ_020
                Return "A/V ≥ 0,20 oder kein gemeinsames Treppenhaus"
            Case NHZ_keine
                Return "Keine Maßnahmen"
            Case Else
                Return ""
        End Select
    End Function
#End Region
#Region "EW"
    Public Function Get_Str_EW_kE() As String
        Select Case Projekt.eigenerWohnbereich
            Case keineEmpfehlung
                Return "X"
            Case EW1, EW2, EW3
                Return "-"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_EW_1() As String
        Select Case Projekt.eigenerWohnbereich
            Case EW1
                Return "X"
            Case keineEmpfehlung, EW2, EW3
                Return "-"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_EW_2() As String
        Select Case Projekt.eigenerWohnbereich
            Case EW2
                Return "X"
            Case keineEmpfehlung, EW1, EW3
                Return "-"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_EW_3() As String
        Select Case Projekt.eigenerWohnbereich
            Case EW3
                Return "X"
            Case keineEmpfehlung, EW1, EW2
                Return "-"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Str_EW_B() As String
        If Projekt.eigenerWohnbereich > 0 Then
            Return "nur Empfehlung"
        Else
            Return ""
        End If
    End Function
#End Region
#End Region
#Region "Get_Pkte_*"
    Public Function Get_Bonus_Untersuchung(ByVal Untersuchung As Byte) As Integer
        Select Case Untersuchung
            Case Messung
                Return 8
            Case Else
                Return 0
        End Select
    End Function
    '    Public Function Get_Bonus_MV(ByVal MV As Byte, ByVal Klasse As String) As Integer

    'If MV = KMV And Klasse <> "A*" And Klasse <> "A" And Klasse <> "B" Then
    '    Return 2
    'ElseIf MV = NMV Then
    '    Return 6
    'Else
    '    Return 0
    'End If
    '    End Function
    'Public Function Get_Bonus_MVPlus(ByVal MA As Byte, ByVal MV As Byte, ByVal Klasse As String) As Integer

    '    If MA = kleiner50 And MV = KMV And Klasse <> "A*" And Klasse <> "A" And Klasse <> "B" Then
    '        Return 2
    '    ElseIf MA = groesser50 And MV = KMV And Klasse <> "A*" And Klasse <> "A" And Klasse <> "B" Then
    '        Return 4
    '    ElseIf MA = kleiner50 And MV = NMV Then
    '        Return 6
    '    ElseIf MA = groesser50 And MV = NMV Then
    '        Return 8
    '    Else
    '        Return 0
    '    End If
    'End Function

#Region "Get_Pkte_* Standort/Außenlärm"
    Public Function Get_Pkte_GC() As String
        Select Case Projekt.Standort.Gebietscharakter
            Case GC_WR
                Return "30"
            Case GC_WA
                Return "20"
            Case GC_MIWB
                Return "10"
            Case GC_GE
                Return "5"
            Case GC_GI
                Return "0"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Pkte_Aussenlaerm() As String
        With Projekt.Standort
            If .Aussenlaermpegel = AP_bis55 And .abgewandFreibereich = aF_ja Then
                Return "27"
            ElseIf .Aussenlaermpegel = AP_bis55 And .abgewandFreibereich = aF_nein Then
                Return "25"
            ElseIf .Aussenlaermpegel = AP_56_60 And .abgewandFreibereich = aF_ja Then
                Return "22"
            ElseIf .Aussenlaermpegel = AP_56_60 And .abgewandFreibereich = aF_nein Then
                Return "20"
            ElseIf .Aussenlaermpegel = AP_61_65 And .abgewandFreibereich = aF_ja Then
                Return "17"
            ElseIf .Aussenlaermpegel = AP_61_65 And .abgewandFreibereich = aF_nein Then
                Return "15"
            ElseIf .Aussenlaermpegel = AP_66_70 And .abgewandFreibereich = aF_ja Then
                Return "12"
            ElseIf .Aussenlaermpegel = AP_66_70 And .abgewandFreibereich = aF_nein Then
                Return "10"
            ElseIf .Aussenlaermpegel = AP_71_75 And .abgewandFreibereich = aF_ja Then
                Return "7"
            ElseIf .Aussenlaermpegel = AP_71_75 And .abgewandFreibereich = aF_nein Then
                Return "5"
            ElseIf .Aussenlaermpegel = AP_76 And .abgewandFreibereich = aF_ja Then
                Return "2"
            ElseIf .Aussenlaermpegel = AP_76 And .abgewandFreibereich = aF_nein Then
                Return "0"
            Else
                Return ""
            End If
        End With
    End Function
#End Region
#Region "LS_W"
    Public Function Get_Pkte_LS_W() As String
        Dim R As Single
        With Projekt.LS_Wand
            If .Untersuchung = Messung Then
                Dim iMes As Integer = Get_Mes_LS_Waende()
                If iMes = 1 Then
                    R = .Messung.Messung_1.Pegel
                ElseIf iMes = 2 Then
                    R = .Messung.Messung_2.Pegel
                ElseIf iMes = 3 Then
                    R = .Messung.Messung_3.Pegel
                ElseIf iMes = 4 Then
                    R = .Messung.Messung_4.Pegel
                ElseIf iMes = 5 Then
                    R = .Messung.Messung_5.Pegel
                ElseIf iMes = 6 Then
                    R = .Messung.Messung_6.Pegel
                End If
            ElseIf .Untersuchung = Prognose Then
                R = .Prognose.Pegel
            ElseIf .Untersuchung = nv Then
                Return "50"
            End If

            If R >= 72 Then
                Return CStr(50 + Get_Bonus_LS_W())
            ElseIf R >= 67 Then
                Return CStr(40 + Get_Bonus_LS_W())
            ElseIf R >= 62 Then
                Return CStr(30 + Get_Bonus_LS_W())
            ElseIf R >= 56 Then
                Return CStr(20 + Get_Bonus_LS_W())
            ElseIf R >= 53 Then
                Return CStr(10 + Get_Bonus_LS_W())
            ElseIf R >= 50 Then
                Return CStr(5 + Get_Bonus_LS_W())
            Else
                Return "0"
            End If
        End With
    End Function
    Public Function Get_Bonus_LS_W() As Integer
        Dim iBonus As Integer = 0
        Dim R As Single = 0
        Dim C As String = ""
        With Projekt.LS_Wand
            If .Untersuchung = Messung Then
                Dim iMes As Integer = Get_Mes_LS_Waende()
                iBonus = 8 'iBonus + Get_Bonus_MVPlus(.Messung.Messanteil, .Messung.Messung_1.Messverfahren, Get_Klasse_LS_Waende)
                If iMes = 1 Then
                    R = .Messung.Messung_1.Pegel
                    C = .Messung.Messung_1.C
                ElseIf iMes = 2 Then
                    R = .Messung.Messung_2.Pegel
                    C = .Messung.Messung_2.C
                ElseIf iMes = 3 Then
                    R = .Messung.Messung_3.Pegel
                    C = .Messung.Messung_3.C
                ElseIf iMes = 4 Then
                    R = .Messung.Messung_4.Pegel
                    C = .Messung.Messung_4.C
                ElseIf iMes = 5 Then
                    R = .Messung.Messung_5.Pegel
                    C = .Messung.Messung_5.C
                ElseIf iMes = 6 Then
                    R = .Messung.Messung_6.Pegel
                    C = .Messung.Messung_6.C
                End If
            ElseIf .Untersuchung = Prognose Then
                R = .Prognose.Pegel
                C = .Prognose.C
            Else
                Return 0
            End If
            iBonus = iBonus + Get_Bonus_LS_W_R_C(R, C)
        End With

        If R < 53 Then iBonus = 0 'Festlegung mit ku 22.08.2017: keine Bonuspunkte in Klasse F und E

        Return iBonus
    End Function
    Public Function Get_Bonus_LS_W_R_C(ByVal R As Single, ByVal sC As String) As Integer
        If IsNumeric(sC) Then ' <> "" Then
            Dim c As Single = CDec(sC)
            If R >= 72 And R + c >= 72 Then
                Return 4
            ElseIf R < 72 And R >= 67 And R + c >= 67 Then
                Return 4
            ElseIf R < 67 And R >= 62 And R + c >= 62 Then
                Return 4
            ElseIf R < 62 And R >= 56 And R + c >= 56 Then
                Return 4
            ElseIf R < 56 And R >= 53 And R + c >= 53 Then
                Return 4
            ElseIf R < 53 And R >= 50 And R + c >= 50 Then
                Return 4
            Else
                Return 0
            End If
        End If
    End Function
#End Region
#Region "LS_D"
    Public Function Get_Pkte_LS_D() As String
        Dim R As Single
        With Projekt.LS_Decke
            If .Untersuchung = Messung Then
                Dim iMes As Integer = Get_Mes_LS_Decken()
                If iMes = 1 Then
                    R = .Messung.Messung_1.Pegel
                ElseIf iMes = 2 Then
                    R = .Messung.Messung_2.Pegel
                ElseIf iMes = 3 Then
                    R = .Messung.Messung_3.Pegel
                ElseIf iMes = 4 Then
                    R = .Messung.Messung_4.Pegel
                ElseIf iMes = 5 Then
                    R = .Messung.Messung_5.Pegel
                ElseIf iMes = 6 Then
                    R = .Messung.Messung_6.Pegel
                End If
            ElseIf .Untersuchung = Prognose Then
                R = .Prognose.Pegel
            ElseIf .Untersuchung = nv Then
                Return "50"
            End If

            If R >= 72 Then
                Return CStr(50 + Get_Bonus_LS_D())
            ElseIf R >= 67 Then
                Return CStr(40 + Get_Bonus_LS_D())
            ElseIf R >= 62 Then
                Return CStr(30 + Get_Bonus_LS_D())
            ElseIf R >= 57 Then
                Return CStr(20 + Get_Bonus_LS_D())
            ElseIf R >= 54 Then
                Return CStr(10 + Get_Bonus_LS_D())
            ElseIf R >= 50 Then
                Return CStr(5 + Get_Bonus_LS_D())
            Else
                Return "0"
            End If
        End With
    End Function
    Public Function Get_Bonus_LS_D() As Integer
        Dim iBonus As Integer = 0
        Dim R As Single = 0
        Dim C As String = ""
        With Projekt.LS_Decke
            If .Untersuchung = Messung Then
                Dim iMes As Integer = Get_Mes_LS_Decken()
                iBonus = 8 'iBonus + Get_Bonus_MVPlus(.Messung.Messanteil, .Messung.Messung_1.Messverfahren, Get_Klasse_LS_Decken)
                If iMes = 1 Then
                    R = .Messung.Messung_1.Pegel
                    C = .Messung.Messung_1.C
                ElseIf iMes = 2 Then
                    R = .Messung.Messung_2.Pegel
                    C = .Messung.Messung_2.C
                ElseIf iMes = 3 Then
                    R = .Messung.Messung_3.Pegel
                    C = .Messung.Messung_3.C
                ElseIf iMes = 4 Then
                    R = .Messung.Messung_4.Pegel
                    C = .Messung.Messung_4.C
                ElseIf iMes = 5 Then
                    R = .Messung.Messung_5.Pegel
                    C = .Messung.Messung_5.C
                ElseIf iMes = 6 Then
                    R = .Messung.Messung_6.Pegel
                    C = .Messung.Messung_6.C
                End If
            ElseIf .Untersuchung = Prognose Then
                R = .Prognose.Pegel
                C = .Prognose.C
            Else
                Return 0
            End If
            iBonus = iBonus + Get_Bonus_LS_D_R_C(R, C)
        End With

        If R < 54 Then iBonus = 0 'Festlegung mit ku 22.08.2017: keine Bonuspunkte in Klasse F und E

        Return iBonus
    End Function
    Public Function Get_Bonus_LS_D_R_C(ByVal R As Single, ByVal sC As String) As Integer
        If IsNumeric(sC) Then ' <> "" Then
            Dim c As Single = CDec(sC)
            If R >= 72 And R + c >= 72 Then
                Return 4
            ElseIf R < 72 And R >= 67 And R + c >= 67 Then
                Return 4
            ElseIf R < 67 And R >= 62 And R + c >= 62 Then
                Return 4
            ElseIf R < 62 And R >= 57 And R + c >= 57 Then
                Return 4
            ElseIf R < 57 And R >= 54 And R + c >= 54 Then
                Return 4
            ElseIf R < 54 And R >= 50 And R + c >= 50 Then
                Return 4
            Else
                Return 0
            End If
        End If
    End Function
#End Region
#Region "TS_D"
    Public Function Get_Pkte_TS_D() As String
        Dim L As Single
        With Projekt.TS_Decke
            If .Untersuchung = Messung Then
                Dim iMes As Integer = Get_Mes_TS_Decken()
                If iMes = 1 Then
                    L = .Messung.Messung_1.Pegel.Pegel
                ElseIf iMes = 2 Then
                    L = .Messung.Messung_2.Pegel.Pegel
                ElseIf iMes = 3 Then
                    L = .Messung.Messung_3.Pegel.Pegel
                ElseIf iMes = 4 Then
                    L = .Messung.Messung_4.Pegel.Pegel
                ElseIf iMes = 5 Then
                    L = .Messung.Messung_5.Pegel.Pegel
                ElseIf iMes = 6 Then
                    L = .Messung.Messung_6.Pegel.Pegel
                End If
            ElseIf .Untersuchung = Prognose Then
                L = .Prognose.Pegel.Pegel
            ElseIf .Untersuchung = nv Then
                Return "50"
            Else
                Return "0"
            End If

            If L <= 30 Then
                Return CStr(50 + Get_Bonus_TS_D())
            ElseIf L <= 35 Then
                Return CStr(40 + Get_Bonus_TS_D())
            ElseIf L <= 40 Then
                Return CStr(30 + Get_Bonus_TS_D())
            ElseIf L <= 45 Then
                Return CStr(20 + Get_Bonus_TS_D())
            ElseIf L <= 50 Then
                Return CStr(10 + Get_Bonus_TS_D())
            ElseIf L <= 60 Then
                Return CStr(5) '+ Get_Bonus_TS_D())
            Else
                Return "0"
            End If
        End With
    End Function
    Public Function Get_Bonus_TS_D() As Integer
        Dim iBonus As Integer = 0
        Dim L As Single = 0
        Dim C As String = ""
        With Projekt.TS_Decke
            Dim fE As Byte = .fEstrich
            If .Untersuchung = Messung Then
                Dim iMes As Integer = Get_Mes_TS_Decken()
                iBonus = 8 'iBonus + Get_Bonus_MVPlus(.Messung.Messanteil, .Messung.Messung_1.Messverfahren, Get_Klasse_TS_Decken)
                If iMes = 1 Then
                    L = .Messung.Messung_1.Pegel.Pegel
                    C = .Messung.Messung_1.Pegel.C
                ElseIf iMes = 2 Then
                    L = .Messung.Messung_2.Pegel.Pegel
                    C = .Messung.Messung_2.Pegel.C
                ElseIf iMes = 3 Then
                    L = .Messung.Messung_3.Pegel.Pegel
                    C = .Messung.Messung_3.Pegel.C
                ElseIf iMes = 4 Then
                    L = .Messung.Messung_4.Pegel.Pegel
                    C = .Messung.Messung_4.Pegel.C
                ElseIf iMes = 5 Then
                    L = .Messung.Messung_5.Pegel.Pegel
                    C = .Messung.Messung_5.Pegel.C
                ElseIf iMes = 6 Then
                    L = .Messung.Messung_6.Pegel.Pegel
                    C = .Messung.Messung_6.Pegel.C
                End If
            ElseIf .Untersuchung = Prognose Then
                L = .Prognose.Pegel.Pegel
                C = .Prognose.Pegel.C
            Else
                Return 0
            End If
            iBonus = iBonus + Get_Bonus_TS_D_L_C(L, C, fE)
        End With
        If L > 50 Then iBonus = 0 'Festlegung mit ku 22.08.2017: keine Bonuspunkte in Klasse F und E
        Return iBonus
    End Function
    Public Function Get_Bonus_TS_D_L_C(ByVal L As Single, ByVal sC As String, ByVal fEstrich As Byte) As Integer
        If fEstrich = fE_kleiner50 Then
            Return 8
        ElseIf IsNumeric(sC) Then ' <> "" Then
            Dim c As Single = CDec(sC)
            If L <= 30 And L + c <= 30 Then
                Return 8
            ElseIf L > 30 And L <= 35 And L + c <= 35 Then
                Return 8
            ElseIf L > 35 And L <= 40 And L + c <= 40 Then
                Return 8
            ElseIf L > 40 And L <= 45 And L + c <= 45 Then
                Return 8
            ElseIf L > 45 And L <= 50 And L + c <= 50 Then
                Return 8
            ElseIf L > 50 And L <= 60 And L + c <= 60 Then
                Return 8
            ElseIf L > 60 Then
                Return 8
            Else
                Return 0
            End If
        Else
            Return 0
        End If
    End Function
#End Region
#Region "TS_TPH"
    Public Function Get_Pkte_TS_TPLH() As String
        Dim L As Single
        Dim bt As Byte = 0

        With Projekt.TS_TPlH
            If .Untersuchung = Messung Then
                Dim iMes As Integer = Get_Mes_TS_TPLH()
                If iMes = 1 Then
                    L = .Messung.Messung_1.Pegel.Pegel
                    bt = .Messung.Messung_1.Bauteil
                ElseIf iMes = 2 Then
                    L = .Messung.Messung_2.Pegel.Pegel
                    bt = .Messung.Messung_2.Bauteil
                ElseIf iMes = 3 Then
                    L = .Messung.Messung_3.Pegel.Pegel
                    bt = .Messung.Messung_3.Bauteil
                ElseIf iMes = 4 Then
                    L = .Messung.Messung_4.Pegel.Pegel
                    bt = .Messung.Messung_4.Bauteil
                ElseIf iMes = 5 Then
                    L = .Messung.Messung_5.Pegel.Pegel
                    bt = .Messung.Messung_5.Bauteil
                ElseIf iMes = 6 Then
                    L = .Messung.Messung_6.Pegel.Pegel
                    bt = .Messung.Messung_6.Bauteil
                End If
            ElseIf .Untersuchung = Prognose Then
                Dim btBauteil As Byte = Get_Prog_TS_TPLH()
                If btBauteil = Treppe Then
                    L = .Prognose.Treppe.Pegel.Pegel
                ElseIf btBauteil = Podest Then
                    L = .Prognose.Podest.Pegel.Pegel
                ElseIf btBauteil = Laube Then
                    L = .Prognose.Laube.Pegel.Pegel
                ElseIf btBauteil = Hausflur Then
                    L = .Prognose.Hausflur.Pegel.Pegel
                End If
                bt = btBauteil
            ElseIf .Untersuchung = nv Then
                Return "50"
            Else
                Return "0"
            End If

            If L <= 33 Then
                Return CStr(50 + Get_Bonus_TS_TPLH())
            ElseIf L <= 38 Then
                Return CStr(40 + Get_Bonus_TS_TPLH())
            ElseIf L <= 43 Then
                Return CStr(30 + Get_Bonus_TS_TPLH())
            ElseIf L <= 48 Then
                Return CStr(20 + Get_Bonus_TS_TPLH())
            ElseIf (L <= 53 And bt <> Hausflur) Or (L <= 50 And bt = Hausflur) Then
                'ElseIf L <= 53 Then
                Return CStr(10 + Get_Bonus_TS_TPLH())
            ElseIf L <= 63 Then
                Return CStr(5 + Get_Bonus_TS_TPLH())
            Else
                Return "0"
            End If
        End With
    End Function
    Public Function Get_Bonus_TS_TPLH() As Integer
        Dim iBonus As Integer = 0
        Dim L As Single = 0
        Dim C As String = ""
        Dim bt As Byte = 0
        With Projekt.TS_TPlH
            If .Untersuchung = Messung Then
                Dim iMes As Integer = Get_Mes_TS_TPLH()
                iBonus = 8 'iBonus + Get_Bonus_MV(.Messung.Messung_1.Messverfahren, Get_Klasse_TS_TPH)
                If iMes = 1 Then
                    L = .Messung.Messung_1.Pegel.Pegel
                    C = .Messung.Messung_1.Pegel.C
                    bt = .Messung.Messung_1.Bauteil
                ElseIf iMes = 2 Then
                    L = .Messung.Messung_2.Pegel.Pegel
                    C = .Messung.Messung_2.Pegel.C
                    bt = .Messung.Messung_2.Bauteil
                ElseIf iMes = 3 Then
                    L = .Messung.Messung_3.Pegel.Pegel
                    C = .Messung.Messung_3.Pegel.C
                    bt = .Messung.Messung_3.Bauteil
                ElseIf iMes = 4 Then
                    L = .Messung.Messung_4.Pegel.Pegel
                    C = .Messung.Messung_4.Pegel.C
                    bt = .Messung.Messung_4.Bauteil
                ElseIf iMes = 5 Then
                    L = .Messung.Messung_5.Pegel.Pegel
                    C = .Messung.Messung_5.Pegel.C
                    bt = .Messung.Messung_5.Bauteil
                ElseIf iMes = 6 Then
                    L = .Messung.Messung_6.Pegel.Pegel
                    C = .Messung.Messung_6.Pegel.C
                    bt = .Messung.Messung_6.Bauteil
                End If
            ElseIf .Untersuchung = Prognose Then
                Dim btBauteil As Byte = Get_Prog_TS_TPLH()
                If btBauteil = Treppe Then
                    L = .Prognose.Treppe.Pegel.Pegel
                    C = .Prognose.Treppe.Pegel.C
                    bt = Treppe
                ElseIf btBauteil = Podest Then
                    L = .Prognose.Podest.Pegel.Pegel
                    C = .Prognose.Podest.Pegel.C
                    bt = Podest
                ElseIf btBauteil = Laube Then
                    L = .Prognose.Laube.Pegel.Pegel
                    C = .Prognose.Laube.Pegel.C
                    bt = Laube
                ElseIf btBauteil = Hausflur Then
                    L = .Prognose.Hausflur.Pegel.Pegel
                    C = .Prognose.Hausflur.Pegel.C
                    bt = Hausflur
                End If
            Else
                Return 0
            End If
            iBonus = iBonus + Get_Bonus_TS_TPLH_L_C(L, C, bt)
        End With
        If (L > 53 And bt <> Hausflur) Or (L > 50 And bt = Hausflur) Then iBonus = 0 'Festlegung mit ku 22.08.2017: keine Bonuspunkte in Klasse F und E
        Return iBonus
    End Function
    Public Function Get_Bonus_TS_TPLH_L_C(ByVal L As Single, ByVal sC As String, ByVal bBauteil As Byte) As Integer
        If IsNumeric(sC) Then ' <> "" Then
            Dim c As Single = CDec(sC)
            If L <= 33 And L + c <= 33 Then
                Return 4
            ElseIf L > 33 And L <= 38 And L + c <= 38 Then
                Return 4
            ElseIf L > 38 And L <= 43 And L + c <= 43 Then
                Return 4
            ElseIf L > 43 And L <= 48 And L + c <= 48 Then
                Return 4
            ElseIf (L > 48 And L <= 53 And L + c <= 53 And bBauteil <> Hausflur) Or (L > 48 And L <= 50 And L + c <= 50 And bBauteil = Hausflur) Then
                Return 4
                'ElseIf (L > 53 And L <= 63 And L + c <= 63 And bBauteil <> Hausflur) Or (L > 50 And L <= 63 And L + c <= 63 And bBauteil = Hausflur) Then
                '    Return 4
            Else 'Klasse E und F
                Return 0
            End If
        End If
    End Function
#End Region
#Region "TS_BLLT"
    Public Function Get_Pkte_TS_BLT() As String
        Dim L As Single
        Dim bt As Byte = 0
        With Projekt.TS_BLT
            If .Untersuchung = Messung Then
                Dim iMes As Integer = Get_Mes_TS_BLT()
                If iMes = 1 Then
                    L = .Messung.Messung_1.Pegel.Pegel
                    bt = .Messung.Messung_1.Bauteil
                ElseIf iMes = 2 Then
                    L = .Messung.Messung_2.Pegel.Pegel
                    bt = .Messung.Messung_2.Bauteil
                ElseIf iMes = 3 Then
                    L = .Messung.Messung_3.Pegel.Pegel
                    bt = .Messung.Messung_3.Bauteil
                ElseIf iMes = 4 Then
                    L = .Messung.Messung_4.Pegel.Pegel
                    bt = .Messung.Messung_4.Bauteil
                ElseIf iMes = 5 Then
                    L = .Messung.Messung_5.Pegel.Pegel
                    bt = .Messung.Messung_5.Bauteil
                ElseIf iMes = 6 Then
                    L = .Messung.Messung_6.Pegel.Pegel
                    bt = .Messung.Messung_6.Bauteil
                End If
            ElseIf .Untersuchung = Prognose Then
                Dim btBauteil As Byte = Get_Prog_TS_BLT()
                If btBauteil = Balkon Then
                    L = .Prognose.Balkon.Pegel.Pegel
                ElseIf btBauteil = Loggie Then
                    L = .Prognose.Loggia.Pegel.Pegel
                ElseIf btBauteil = Terrasse Then
                    L = .Prognose.Terrasse.Pegel.Pegel
                End If
                bt = btBauteil
            ElseIf .Untersuchung = nv Then
                Return "25"
            Else
                Return "0"
            End If

            If L <= 33 Then
                Return CStr(25 + Get_Bonus_TS_BLT())
            ElseIf L <= 38 Then
                Return CStr(20 + Get_Bonus_TS_BLT())
            ElseIf L <= 43 Then
                Return CStr(15 + Get_Bonus_TS_BLT())
            ElseIf L <= 48 Then
                Return CStr(10 + Get_Bonus_TS_BLT())
            ElseIf (L <= 50 And bt <> Balkon) Or (L <= 58 And bt = Balkon) Then
                Return CStr(5 + Get_Bonus_TS_BLT())
            ElseIf L <= 63 Then
                Return CStr(0 + Get_Bonus_TS_BLT())
            Else
                Return "0"
            End If
        End With
    End Function
    Public Function Get_Bonus_TS_BLT() As Integer
        Dim iBonus As Integer = 0
        Dim L As Single = 0
        Dim C As String = ""
        Dim bt As Byte = 0
        With Projekt.TS_BLT
            If .Untersuchung = Messung Then
                Dim iMes As Integer = Get_Mes_TS_BLT()
                iBonus = 4 'iBonus + Get_Bonus_MV(.Messung.Messung_1.Messverfahren, Get_Klasse_TS_BLLT)
                If iMes = 1 Then
                    L = .Messung.Messung_1.Pegel.Pegel
                    C = .Messung.Messung_1.Pegel.C
                    bt = .Messung.Messung_1.Bauteil
                ElseIf iMes = 2 Then
                    L = .Messung.Messung_2.Pegel.Pegel
                    C = .Messung.Messung_2.Pegel.C
                    bt = .Messung.Messung_2.Bauteil
                ElseIf iMes = 3 Then
                    L = .Messung.Messung_3.Pegel.Pegel
                    C = .Messung.Messung_3.Pegel.C
                    bt = .Messung.Messung_3.Bauteil
                ElseIf iMes = 4 Then
                    L = .Messung.Messung_4.Pegel.Pegel
                    C = .Messung.Messung_4.Pegel.C
                    bt = .Messung.Messung_4.Bauteil
                ElseIf iMes = 5 Then
                    L = .Messung.Messung_5.Pegel.Pegel
                    C = .Messung.Messung_5.Pegel.C
                    bt = .Messung.Messung_5.Bauteil
                ElseIf iMes = 6 Then
                    L = .Messung.Messung_6.Pegel.Pegel
                    C = .Messung.Messung_6.Pegel.C
                    bt = .Messung.Messung_6.Bauteil
                End If
            ElseIf .Untersuchung = Prognose Then
                Dim btBauteil As Byte = Get_Prog_TS_BLT()
                If btBauteil = Balkon Then
                    L = .Prognose.Balkon.Pegel.Pegel
                    C = .Prognose.Balkon.Pegel.C
                    bt = Balkon
                ElseIf btBauteil = Loggie Then
                    L = .Prognose.Loggia.Pegel.Pegel
                    C = .Prognose.Loggia.Pegel.C
                    bt = Loggie
                ElseIf btBauteil = Terrasse Then
                    L = .Prognose.Terrasse.Pegel.Pegel
                    C = .Prognose.Terrasse.Pegel.C
                    bt = Terrasse
                End If
            Else
                Return 0
            End If
            iBonus = iBonus + Get_Bonus_TS_BLT_L_C(L, C, bt)
        End With
        If (L > 50 And bt <> Balkon) Or (L > 58 And bt = Balkon) Then iBonus = 0 'Festlegung mit ku 22.08.2017: keine Bonuspunkte in Klasse F und E
        Return iBonus
    End Function
    Public Function Get_Bonus_TS_BLT_L_C(ByVal L As Single, ByVal sC As String, ByVal bBauteil As Byte) As Integer
        If IsNumeric(sC) Then ' <> "" Then
            Dim c As Single = CDec(sC)
            If L <= 33 And L + c <= 33 Then
                Return 4
            ElseIf L > 33 And L <= 38 And L + c <= 38 Then
                Return 4
            ElseIf L > 38 And L <= 43 And L + c <= 43 Then
                Return 4
            ElseIf L > 43 And L <= 48 And L + c <= 48 Then
                Return 4
            ElseIf (L > 48 And L <= 50 And L + c <= 50 And bBauteil <> Balkon) Or (L > 48 And L <= 58 And L + c <= 58 And bBauteil = Balkon) Then
                Return 4
            Else 'kein Bonus für Klasse E und F
                Return 0
            End If
        End If
    End Function
#End Region

#Region "Tueren"
    Public Function Get_Pkte_Tueren_D() As String
        With Projekt.Tueren
            If .Ort = nv Then
                Return "30"
            ElseIf .Ort = Diele Then
                If .L >= 40 Then
                    Return CStr(30 + Get_Bonus_Tueren())
                ElseIf .L >= 37 Then
                    Return CStr(20 + Get_Bonus_Tueren())
                ElseIf .L >= 32 Then
                    Return CStr(10 + Get_Bonus_Tueren())
                ElseIf .L >= 27 Then
                    Return CStr(5 + Get_Bonus_Tueren())
                ElseIf .L >= 22 Then
                    Return CStr(Get_Bonus_Tueren())
                Else
                    Return CStr(Get_Bonus_Tueren()) '"0"
                End If
            Else
                Return "" '"0"
            End If
        End With
    End Function
    Public Function Get_Pkte_Tueren_A() As String
        With Projekt.Tueren
            If .Ort = Aufenthalt Then
                'If .L >= 48 Then
                '    Return CStr(30 + Get_Bonus_Tueren())
                'ElseIf .L >= 45 Then
                '    Return CStr(20 + Get_Bonus_Tueren())
                'Else
                If .L >= 42 Then
                    Return CStr(10 + Get_Bonus_Tueren())
                ElseIf .L >= 37 Then
                    Return CStr(5 + Get_Bonus_Tueren())
                ElseIf .L >= 32 Then
                    Return CStr(Get_Bonus_Tueren())
                Else
                    Return CStr(Get_Bonus_Tueren()) '"0"
                End If
            Else
                Return "" '"0"
            End If
        End With
    End Function
    Public Function Get_Bonus_Tueren() As Integer
        Dim iBonus As Integer = 0
        If Projekt.Tueren.Untersuchung = Messung Then
            iBonus = 4
        Else
            iBonus = 0
        End If
        Select Case Projekt.Tueren.Ort
            Case Diele
                If Projekt.Tueren.L < 27 Then iBonus = 0
            Case Aufenthalt
                If Projekt.Tueren.L < 37 Then iBonus = 0
            Case Else
                iBonus = 0
        End Select

        Return iBonus
    End Function
#End Region
#Region "Aussenbauteile"
    Public Function Get_Pkte_Aussenbauteile() As String
        Select Case Projekt.Aussenbauteile
            'Case FensterMitDichtung
            '    Return "5"
            Case DINerfuellt
                Return "10"
            Case DINPlusErfuellt
                Return "15"
            Case Else
                Return "0"
        End Select
    End Function
#End Region
#Region "Wasser"
    Public Function Get_Pkte_Wasser() As String
        With Projekt.Wasser
            If .Untersuchung = Prognose Or .Untersuchung = Messung Then
                If .L_Intervall = HT_kleiner20 Then ' <= 20 Then
                    Return CStr(30 + Get_Bonus_Wasser())
                ElseIf .L_Intervall = HT_20_24 Then ' <= 25 Then
                    Return CStr(20 + Get_Bonus_Wasser())
                ElseIf .L_Intervall = HT_24_27 Then ' <= 30 Then
                    Return CStr(10 + Get_Bonus_Wasser())
                ElseIf .L_Intervall = HT_27_30 Then ' <= 35 Then
                    Return CStr(5 + Get_Bonus_Wasser())
                Else
                    Return "0" 'CStr(Get_Bonus_Wasser()) '"0"
                End If
            Else
                Return "0"
            End If
        End With
    End Function
    Public Function Get_Bonus_Wasser() As Integer
        With Projekt.Wasser
            If .L_Intervall = HT_30_35 Or .L_Intervall = HT_groesser35 Then
                Return 0
            ElseIf .Untersuchung = Messung And .LcLa_erfuellt = erfuellt Then
                Return 6
            ElseIf .Untersuchung = Messung And .LcLa_erfuellt = keineAngabe Then
                Return 4
            ElseIf .Untersuchung = Prognose And .LcLa_erfuellt = erfuellt Then
                Return 2
            Else
                Return 0
            End If
        End With
    End Function
#End Region
#Region "Nutzergeräusche und Körperschallentkopplung"
    Public Function Get_Pkte_NutzerKoerper() As String
        Get_Pkte_NutzerKoerper = ""
        With Projekt.NutzerKoerper
            If .Untersuchung = nv Then
                Return "0"
            ElseIf .Untersuchung = Nutzer Then
                Select Case .L_Intervall
                    Case kleiner20
                        Get_Pkte_NutzerKoerper = CStr(25 + Get_Bonus_NutzerKoerper())
                    Case L_20_25
                        Get_Pkte_NutzerKoerper = CStr(20 + Get_Bonus_NutzerKoerper())
                    Case L_25_30
                        Get_Pkte_NutzerKoerper = CStr(15 + Get_Bonus_NutzerKoerper())
                    Case L_30_35
                        Get_Pkte_NutzerKoerper = CStr(10 + Get_Bonus_NutzerKoerper())
                    Case L_35_40
                        Get_Pkte_NutzerKoerper = CStr(5 + Get_Bonus_NutzerKoerper())
                    Case L_40_45
                        Get_Pkte_NutzerKoerper = CStr(0) ' + Get_Bonus_NutzerKoerper())
                    Case groesser45
                        Get_Pkte_NutzerKoerper = CStr(0) ' + Get_Bonus_NutzerKoerper())
                End Select
            ElseIf .Untersuchung = Koerper Then
                Select Case .L_Intervall
                    Case kleiner38
                        Get_Pkte_NutzerKoerper = CStr(25 + Get_Bonus_NutzerKoerper())
                    Case L_38_43
                        Get_Pkte_NutzerKoerper = CStr(20 + Get_Bonus_NutzerKoerper())
                    Case L_43_48
                        Get_Pkte_NutzerKoerper = CStr(15 + Get_Bonus_NutzerKoerper())
                    Case L_48_53
                        Get_Pkte_NutzerKoerper = CStr(10 + Get_Bonus_NutzerKoerper())
                    Case L_53_58
                        Get_Pkte_NutzerKoerper = CStr(5 + Get_Bonus_NutzerKoerper())
                    Case L_58_63
                        Get_Pkte_NutzerKoerper = CStr(0) ' + Get_Bonus_NutzerKoerper())
                    Case groesser63
                        Get_Pkte_NutzerKoerper = CStr(0) ' + Get_Bonus_NutzerKoerper())
                End Select
            End If

        End With

    End Function
   
    Public Function Get_Bonus_NutzerKoerper() As Integer
        With Projekt.NutzerKoerper
            If .Untersuchung = Nutzer Then
                If .MessungPrognose = Messung And .L_Intervall <> groesser45 And .L_Intervall <> L_40_45 Then
                    Return 4
                Else
                    Return 0
                End If
            ElseIf .Untersuchung = Koerper Then
                If .MessungPrognose = Messung And .L_Intervall <> groesser63 And .L_Intervall <> L_58_63 Then
                    Return 4
                Else
                    Return 0
                End If
            End If

        End With
    End Function
    'Public Function Get_Bonus_Koerper() As String
    '    With Projekt.Koerper
    '        If .Untersuchung = Messung And .L_Intervall <> groesser63 And .L_Intervall <> L_58_63 Then
    '            Return "4"
    '        Else
    '            Return "0"
    '        End If
    '    End With
    'End Function
    'Public Function Get_Pkte_NutzerKoerper() As String
    '    With Projekt.NutzerKoerper
    '        If .Untersuchung = nv Then
    '            Get_Pkte_NutzerKoerper = "0"
    '        ElseIf .Untersuchung = Prognose And ( _
    '                (Projekt.Koerper.Untersuchung = Messung And _
    '                    (Projekt.Koerper.L_Intervall = groesser63 Or Projekt.Koerper.L_Intervall = L_58_63)) Or _
    '                Projekt.Koerper.Untersuchung <> Messung) Then
    '            Get_Pkte_Nutzer = "2"
    '        ElseIf .Untersuchung = Messung Then
    '            Get_Pkte_Nutzer = "6"
    '        Else
    '            Get_Pkte_Nutzer = "0"
    '        End If
    '        'If .L_Intervall = 0 Then Get_Pkte_Nutzer = ""
    '    End With
    'End Function
    'Public Function Get_Pkte_Koerper() As String
    '    With Projekt.Koerper
    '        If .Untersuchung = nv Or .L_Intervall = 0 Or .L_Intervall = groesser63 Or .L_Intervall = L_58_63 Then 'Or .L_Intervall = groesser63 
    '            Get_Pkte_Koerper = "0"
    '        ElseIf .Untersuchung = Prognose And _
    '            ((Projekt.Nutzer.Untersuchung <> Prognose And Projekt.Nutzer.Untersuchung <> Messung) Or _
    '            (Projekt.Nutzer.Untersuchung = Prognose And _
    '                (Projekt.Nutzer.L_Intervall = groesser45 Or Projekt.Nutzer.L_Intervall = L_40_45)) Or _
    '            (Projekt.Nutzer.Untersuchung = Messung And _
    '                (Projekt.Nutzer.L_Intervall = groesser45 Or Projekt.Nutzer.L_Intervall = L_40_45))) Then
    '            Get_Pkte_Koerper = "2"
    '        ElseIf .Untersuchung = Messung And _
    '            (Projekt.Nutzer.Untersuchung <> Messung Or _
    '            (Projekt.Nutzer.Untersuchung = Messung And _
    '                (Projekt.Nutzer.L_Intervall = groesser45 Or Projekt.Nutzer.L_Intervall = L_40_45))) Then
    '            Get_Pkte_Koerper = "6"
    '        Else
    '            Get_Pkte_Koerper = "0"
    '        End If
    '        'If .L_Intervall = 0 Then Get_Pkte_Koerper = ""
    '    End With
    'End Function
#End Region
#Region "Baulicher Schallschutz"
    Public Function Get_Pkte_Nachbarn() As String
        Select Case Projekt.Nachbarn
            Case 1
                Return "20"
            Case 2
                Return "15"
            Case 3
                Return "10"
            Case 4
                Return "5"
            Case Else
                Return "0"
        End Select
    End Function
    Public Function Get_Pkte_anordRaeume() As String
        If Projekt.anordnungRaeume = guenstig Then
            Return "5"
        Else
            Return "0"
        End If
    End Function
    Public Function Get_Pkte_lauteRaeume_keine() As String
        Select Case Projekt.lauteRaeume
            Case keineAngrenzend
                Return "15"
            Case Else
                Return "---"
        End Select
    End Function
    Public Function Get_Pkte_lauteRaeume_25() As String
        Select Case Projekt.lauteRaeume
            Case L_25_35
                Return "10"
            Case Else
                Return "---"
        End Select
    End Function
    Public Function Get_Pkte_lauteRaeume_30() As String
        Select Case Projekt.lauteRaeume
            Case L_30_40
                Return "5"
            Case Else
                Return "---"
        End Select
    End Function
    Public Function Get_Pkte_lauteRaeume_35() As String
        Select Case Projekt.lauteRaeume
            Case L_35_45
                Return "0"
            Case Else
                Return "---"
        End Select
    End Function
#End Region
#Region "NHZ"
    Public Function Get_Pkte_NHZ() As String
        Select Case Projekt.NHZ
            Case NHZ_010
                Return "5"
            Case NHZ_020
                Return "8"
            Case Else
                Return "0"
        End Select
    End Function
#End Region
#Region "EW"
    Public Function Get_Pkte_EW() As String
        Select Case Projekt.eigenerWohnbereich
            Case EW1
                Return "5"
            Case EW2
                Return "10"
            Case EW3
                Return "15"
            Case Else
                Return "0"
        End Select
    End Function
#End Region
#End Region

#Region "Get_Klasse_*"
    Public Function Get_Klasse_I() As String
        '=WENN(ODER(P19="";P21="");"";WENN(ODER(P19="F";P21="F");"F";WENN(ODER(P19="E";P21="E");"E";WENN(ODER(P19="D";P21="D");"D";
        'WENN(ODER(P19="C";P21="C");"C";WENN(ODER(P19="B";P21="B");"B";WENN(ODER(P19="A";P21="A");"A";"A*")))))))

        Dim P_Pkte_I As String = Form_Main.Calc_Pkte_I 'Me.Label_Pkte_I.Text
        Dim iKlassePkte As Integer = 0
        If P_Pkte_I <> "" Then
            Select Case CInt(P_Pkte_I)
                Case Is < 10
                    iKlassePkte = 7
                Case Is < 20
                    iKlassePkte = 6
                Case Is < 25
                    iKlassePkte = 5
                Case Is < 40
                    iKlassePkte = 4
                Case Is < 45
                    iKlassePkte = 3
                Case Is < 55
                    iKlassePkte = 2
                Case Else
                    iKlassePkte = 1
            End Select
        End If
        Dim iKlasseSchlechteste As Integer = 0

        Dim Kl_GC As String = Get_Klasse_GC() 'Me.Label_Klasse_GC.Text
        Dim Kl_AP As String = Get_Klasse_AP() 'Me.Label_Klasse_Aussenlaerm.Text

        If Kl_GC = "" Or Kl_AP = "" Then
            iKlasseSchlechteste = 0
        ElseIf (Kl_GC = "F" Or Kl_AP = "F") Then
            iKlasseSchlechteste = 7
        ElseIf (Kl_GC = "E" Or Kl_AP = "E") Then
            iKlasseSchlechteste = 6
        ElseIf (Kl_GC = "D" Or Kl_AP = "D") Then
            iKlasseSchlechteste = 5
        ElseIf (Kl_GC = "C" Or Kl_AP = "C") Then
            iKlasseSchlechteste = 4
        ElseIf (Kl_GC = "B" Or Kl_AP = "B") Then
            iKlasseSchlechteste = 3
        ElseIf (Kl_GC = "A" Or Kl_AP = "A") Then
            iKlasseSchlechteste = 2
        Else
            iKlasseSchlechteste = 1
        End If

        Dim iKlasseGesamt As Integer = 0
        If iKlasseSchlechteste = 0 Then
            iKlasseGesamt = 0
        ElseIf iKlassePkte > iKlasseSchlechteste Then
            iKlasseGesamt = iKlassePkte
        ElseIf iKlassePkte < iKlasseSchlechteste Then
            iKlasseGesamt = iKlasseSchlechteste - 1
        Else
            iKlasseGesamt = iKlassePkte
        End If

        Select Case iKlasseGesamt
            Case 0
                Return ""
            Case 1
                Return "A*"
            Case 2
                Return "A"
            Case 3
                Return "B"
            Case 4
                Return "C"
            Case 5
                Return "D"
            Case 6
                Return "E"
            Case 7
                Return "F"

        End Select

        'If Kl_GC = "" Or Kl_AP = "" Then
        '    Return ""
        'ElseIf (Kl_GC = "F" Or Kl_AP = "F") Then
        '    Return "F"
        'ElseIf (Kl_GC = "E" Or Kl_AP = "E") Then
        '    Return "E"
        'ElseIf (Kl_GC = "D" Or Kl_AP = "D") Then
        '    Return "D"
        'ElseIf (Kl_GC = "C" Or Kl_AP = "C") Then
        '    Return "C"
        'ElseIf (Kl_GC = "B" Or Kl_AP = "B") Then
        '    Return "B"
        'ElseIf (Kl_GC = "A" Or Kl_AP = "A") Then
        '    Return "A"
        'Else
        '    Return "A*"
        'End If
    End Function
    Public Function Get_Klasse_II() As String

        Dim Kl_LS_W As String = Get_Klasse_LS_Waende() 'Me.Label_Klasse_LS_W.Text
        Dim Kl_LS_D As String = Get_Klasse_LS_Decken() 'Me.Label_Klasse_LS_D.Text
        Dim Kl_TS_D As String = Get_Klasse_TS_Decken() 'Me.Label_Klasse_TS_D.Text
        Dim Kl_TS_TPH As String = Get_Klasse_TS_TPLH() 'Me.Label_Klasse_TS_TPH.Text
        Dim Kl_TS_BLLT As String = Get_Klasse_TS_BLT() 'Me.Label_Klasse_TS_BLLT.Text
        Dim Kl_Tueren_D As String = Get_Klasse_Tueren_D() 'Me.Label_Klasse_Tueren_D.Text
        Dim Kl_Tueren_A As String = Get_Klasse_Tueren_A() 'Me.Label_Klasse_Tueren_A.Text
        Dim Kl_Aussenbauteile As String = Get_Klasse_Aussenbauteile() 'Me.Label_Klasse_Aussenbauteile.Text
        Dim Kl_Wasser As String = Get_Klasse_Wasser() 'Me.Label_Klasse_Wasser.Text
        Dim Kl_lauteRaeume_keine As String = Get_Klasse_lauteRaeume_kE() 'Me.Label_Klasse_lauteRaeume_keine.Text
        Dim Kl_lauteRaeume_25 As String = Get_Klasse_lauteRaeume_25() 'Me.Label_Klasse_lauteRaeume_25.Text
        Dim Kl_lauteRaeume_30 As String = Get_Klasse_lauteRaeume_30() 'Me.Label_Klasse_lauteRaeume_30.Text
        Dim Kl_lauteRaeume_35 As String = Get_Klasse_lauteRaeume_35() 'Me.Label_Klasse_lauteRaeume_35.Text

        Dim Kl_Anzahl_Nachbarn As String = Get_Klasse_Nachbarn()

        Dim P_LS_W As String = Get_Pkte_LS_W() 'Me.Label_Pkte_LS_W.Text
        Dim P_LS_D As String = Get_Pkte_LS_D() 'Me.Label_Pkte_LS_D.Text
        Dim P_TS_D As String = Get_Pkte_TS_D() 'Me.Label_Pkte_TS_D.Text
        Dim P_TS_TPH As String = Get_Pkte_TS_TPLH() 'Me.Label_Pkte_TS_TPH.Text
        Dim P_TS_BLLT As String = Get_Pkte_TS_BLT() 'Me.Label_Pkte_TS_BLLT.Text
        Dim P_Pkte_II As String = Form_Main.Calc_Pkte_II 'Me.Label_Pkte_II.Text


        Dim iKlassePkte As Integer = 0
        Select Case CInt(P_Pkte_II)
            Case Is < 30
                iKlassePkte = 7
            Case Is < 80
                iKlassePkte = 6
            Case Is < 145
                iKlassePkte = 5
            Case Is < 210
                iKlassePkte = 4
            Case Is < 270
                iKlassePkte = 3
            Case Is < 340
                iKlassePkte = 2
            Case Else
                iKlassePkte = 1
        End Select

        Dim iKlasseSchlechteste As Integer = 0
        Dim use_version = 2018
        If use_version = 2012 Then
            If (Kl_LS_W = "" And P_LS_W = "0") Or (Kl_LS_D = "" And P_LS_D = "0") Or (Kl_TS_D = "" And P_TS_D = "0") Or
                        (Kl_TS_TPH = "" And P_TS_TPH = "0") Or (Kl_TS_BLLT = "" And P_TS_TPH = "0") Or (Kl_Tueren_D = "" And Kl_Tueren_A = "") Or
                        Kl_Aussenbauteile = "" Or Kl_Wasser = "" Or
                        (Kl_lauteRaeume_keine = "" And Kl_lauteRaeume_25 = "" And Kl_lauteRaeume_30 = "" And Kl_lauteRaeume_35 = "") Then
                iKlasseSchlechteste = 0
            ElseIf Kl_LS_W = "F" Or Kl_LS_D = "F" Or Kl_TS_D = "F" Or Kl_TS_TPH = "F" Or Kl_TS_BLLT = "F" Or Kl_Tueren_D = "F" Or Kl_Tueren_A = "F" _
                    Or Kl_Aussenbauteile = "F" Or Kl_Wasser = "F" Or
                    Kl_lauteRaeume_keine = "F" Or Kl_lauteRaeume_25 = "F" Or Kl_lauteRaeume_30 = "F" Or Kl_lauteRaeume_35 = "F" Then
                iKlasseSchlechteste = 7
            ElseIf Kl_LS_W = "E" Or Kl_LS_D = "E" Or Kl_TS_D = "E" Or Kl_TS_TPH = "E" Or Kl_TS_BLLT = "E" Or Kl_Tueren_D = "E" Or Kl_Tueren_A = "E" _
                    Or Kl_Aussenbauteile = "E" Or Kl_Wasser = "E" Or
                    Kl_lauteRaeume_keine = "E" Or Kl_lauteRaeume_25 = "E" Or Kl_lauteRaeume_30 = "E" Or Kl_lauteRaeume_35 = "E" Then
                iKlasseSchlechteste = 6
            ElseIf Kl_LS_W = "D" Or Kl_LS_D = "D" Or Kl_TS_D = "D" Or Kl_TS_TPH = "D" Or Kl_TS_BLLT = "D" Or Kl_Tueren_D = "D" Or Kl_Tueren_A = "D" _
                    Or Kl_Aussenbauteile = "D" Or Kl_Wasser = "D" Or
                    Kl_lauteRaeume_keine = "D" Or Kl_lauteRaeume_25 = "D" Or Kl_lauteRaeume_30 = "D" Or Kl_lauteRaeume_35 = "D" Then
                iKlasseSchlechteste = 5
            ElseIf Kl_LS_W = "C" Or Kl_LS_D = "C" Or Kl_TS_D = "C" Or Kl_TS_TPH = "C" Or Kl_TS_BLLT = "C" Or Kl_Tueren_D = "C" Or Kl_Tueren_A = "C" _
                    Or Kl_Aussenbauteile = "C" Or Kl_Wasser = "C" Or
                    Kl_lauteRaeume_keine = "C" Or Kl_lauteRaeume_25 = "C" Or Kl_lauteRaeume_30 = "C" Or Kl_lauteRaeume_35 = "C" Then
                iKlasseSchlechteste = 4
            ElseIf Kl_LS_W = "B" Or Kl_LS_D = "B" Or Kl_TS_D = "B" Or Kl_TS_TPH = "B" Or Kl_TS_BLLT = "B" Or Kl_Tueren_D = "B" Or Kl_Tueren_A = "B" _
                    Or Kl_Aussenbauteile = "B" Or Kl_Wasser = "B" Or
                    Kl_lauteRaeume_keine = "B" Or Kl_lauteRaeume_25 = "B" Or Kl_lauteRaeume_30 = "B" Or Kl_lauteRaeume_35 = "B" Then
                iKlasseSchlechteste = 3
            ElseIf Kl_LS_W = "A" Or Kl_LS_D = "A" Or Kl_TS_D = "A" Or Kl_TS_TPH = "A" Or Kl_TS_BLLT = "A" Or Kl_Tueren_D = "A" Or Kl_Tueren_A = "A" _
                    Or Kl_Aussenbauteile = "A" Or Kl_Wasser = "A" Or
                    Kl_lauteRaeume_keine = "A" Or Kl_lauteRaeume_25 = "A" Or Kl_lauteRaeume_30 = "A" Or Kl_lauteRaeume_35 = "A" Then
                iKlasseSchlechteste = 2
            Else
                iKlasseSchlechteste = 1
            End If
        ElseIf use_version = 2018 Then
            If (Kl_LS_W = "" And P_LS_W = "0") Or (Kl_LS_D = "" And P_LS_D = "0") Or (Kl_TS_D = "" And P_TS_D = "0") Or
                        (Kl_TS_TPH = "" And P_TS_TPH = "0") Or (Kl_TS_BLLT = "" And P_TS_TPH = "0") Or (Kl_Tueren_D = "" And Kl_Tueren_A = "") Or
                        Kl_Aussenbauteile = "" Or Kl_Wasser = "" Or Kl_Anzahl_Nachbarn = "" Or
                        (Kl_lauteRaeume_keine = "" And Kl_lauteRaeume_25 = "" And Kl_lauteRaeume_30 = "" And Kl_lauteRaeume_35 = "") Then
                iKlasseSchlechteste = 0
            ElseIf Kl_LS_W = "F" Or Kl_LS_D = "F" Or Kl_TS_D = "F" Or Kl_TS_TPH = "F" Or Kl_TS_BLLT = "F" Or Kl_Tueren_D = "F" Or Kl_Tueren_A = "F" _
                    Or Kl_Aussenbauteile = "F" Or Kl_Wasser = "F" Or Kl_Anzahl_Nachbarn = "F" Or
                    Kl_lauteRaeume_keine = "F" Or Kl_lauteRaeume_25 = "F" Or Kl_lauteRaeume_30 = "F" Or Kl_lauteRaeume_35 = "F" Then
                iKlasseSchlechteste = 7
            ElseIf Kl_LS_W = "E" Or Kl_LS_D = "E" Or Kl_TS_D = "E" Or Kl_TS_TPH = "E" Or Kl_TS_BLLT = "E" Or Kl_Tueren_D = "E" Or Kl_Tueren_A = "E" _
                    Or Kl_Aussenbauteile = "E" Or Kl_Wasser = "E" Or Kl_Anzahl_Nachbarn = "E" Or
                    Kl_lauteRaeume_keine = "E" Or Kl_lauteRaeume_25 = "E" Or Kl_lauteRaeume_30 = "E" Or Kl_lauteRaeume_35 = "E" Then
                iKlasseSchlechteste = 6
            ElseIf Kl_LS_W = "D" Or Kl_LS_D = "D" Or Kl_TS_D = "D" Or Kl_TS_TPH = "D" Or Kl_TS_BLLT = "D" Or Kl_Tueren_D = "D" Or Kl_Tueren_A = "D" _
                    Or Kl_Aussenbauteile = "D" Or Kl_Wasser = "D" Or Kl_Anzahl_Nachbarn = "D" Or
                    Kl_lauteRaeume_keine = "D" Or Kl_lauteRaeume_25 = "D" Or Kl_lauteRaeume_30 = "D" Or Kl_lauteRaeume_35 = "D" Then
                iKlasseSchlechteste = 5
            ElseIf Kl_LS_W = "C" Or Kl_LS_D = "C" Or Kl_TS_D = "C" Or Kl_TS_TPH = "C" Or Kl_TS_BLLT = "C" Or Kl_Tueren_D = "C" Or Kl_Tueren_A = "C" _
                    Or Kl_Aussenbauteile = "C" Or Kl_Wasser = "C" Or Kl_Anzahl_Nachbarn = "C" Or
                    Kl_lauteRaeume_keine = "C" Or Kl_lauteRaeume_25 = "C" Or Kl_lauteRaeume_30 = "C" Or Kl_lauteRaeume_35 = "C" Then
                iKlasseSchlechteste = 4
            ElseIf Kl_LS_W = "B" Or Kl_LS_D = "B" Or Kl_TS_D = "B" Or Kl_TS_TPH = "B" Or Kl_TS_BLLT = "B" Or Kl_Tueren_D = "B" Or Kl_Tueren_A = "B" _
                    Or Kl_Aussenbauteile = "B" Or Kl_Wasser = "B" Or Kl_Anzahl_Nachbarn = "B" Or
                    Kl_lauteRaeume_keine = "B" Or Kl_lauteRaeume_25 = "B" Or Kl_lauteRaeume_30 = "B" Or Kl_lauteRaeume_35 = "B" Then
                iKlasseSchlechteste = 3
            ElseIf Kl_LS_W = "A" Or Kl_LS_D = "A" Or Kl_TS_D = "A" Or Kl_TS_TPH = "A" Or Kl_TS_BLLT = "A" Or Kl_Tueren_D = "A" Or Kl_Tueren_A = "A" _
                    Or Kl_Aussenbauteile = "A" Or Kl_Wasser = "A" Or Kl_Anzahl_Nachbarn = "A" Or
                    Kl_lauteRaeume_keine = "A" Or Kl_lauteRaeume_25 = "A" Or Kl_lauteRaeume_30 = "A" Or Kl_lauteRaeume_35 = "A" Then
                iKlasseSchlechteste = 2
            Else
                iKlasseSchlechteste = 1
            End If
        End If


        Dim iKlasseGesamt As Integer = 0
        If iKlasseSchlechteste = 0 Then
            iKlasseGesamt = 0
        ElseIf iKlassePkte > iKlasseSchlechteste Then
            iKlasseGesamt = iKlassePkte
        ElseIf iKlassePkte < iKlasseSchlechteste Then
            iKlasseGesamt = iKlasseSchlechteste - 1
        Else
            iKlasseGesamt = iKlassePkte
        End If

        Select Case iKlasseGesamt
            Case 0
                Return ""
            Case 1
                Return "A*"
            Case 2
                Return "A"
            Case 3
                Return "B"
            Case 4
                Return "C"
            Case 5
                Return "D"
            Case 6
                Return "E"
            Case 7
                Return "F"

        End Select

        '=WENN(ODER(UND(P27="";N27=0);UND(P28="";N28=0);UND(P30="";N30=0);UND(P31="";N31=0);UND(P32="";N32=0);
        'UND(P34="";P35="");P38="";P40="";UND(P46="";P47="";P49="";P51=""));"";
        'WENN(ODER(P27="F";P28="F";P30="F";P31="F";P32="F";P34="F";P35="F";P38="F";P40="F";P46="F";P47="F";P49="F";P51="F");WENN(N55<30;"F";"E");
        'WENN(ODER(P27="E";P28="E";P30="E";P31="E";P32="E";P34="E";P35="E";P38="E";P40="E";P46="E";P47="E";P49="E";P51="E");WENN(N55<80;"E";"D");
        'WENN(ODER(P27="D";P28="D";P30="D";P31="D";P32="D";P34="D";P35="D";P38="D";P40="D";P46="D";P47="D";P49="D";P51="D");WENN(N55<145;"D";"C");
        'WENN(ODER(P27="C";P28="C";P30="C";P31="C";P32="C";P34="C";P35="C";P38="C";P40="C";P46="C";P47="C";P49="C";P51="C");WENN(N55<210;"C";"B");
        'WENN(ODER(P27="B";P28="B";P30="B";P31="B";P32="B";P34="B";P35="B";P38="B";P40="B";P46="B";P47="B";P49="B";P51="B");WENN(N55<270;"B";"A");
        'WENN(ODER(P27="A";P28="A";P30="A";P31="A";P32="A";P34="A";P35="A";P38="A";P40="A";P46="A";P47="A";P49="A";P51="A");WENN(N55<340;"A";"A*"))))))))

        'If (Kl_LS_W = "" And P_LS_W = "0") Or (Kl_LS_D = "" And P_LS_D = "0") Or (Kl_TS_D = "" And P_TS_D = "0") Or _
        '        (Kl_TS_TPH = "" And P_TS_TPH = "0") Or (Kl_TS_BLLT = "" And P_TS_TPH = "0") Or (Kl_Tueren_D = "" And Kl_Tueren_A = "") Or _
        '        Kl_Aussenbauteile = "" Or Kl_Wasser = "" Or _
        '        (Kl_lauteRaeume_keine = "" And Kl_lauteRaeume_25 = "" And Kl_lauteRaeume_30 = "" And Kl_lauteRaeume_35 = "") Then
        '    Return ""
        'ElseIf Kl_LS_W = "F" Or Kl_LS_D = "F" Or Kl_TS_D = "F" Or Kl_TS_TPH = "F" Or Kl_TS_BLLT = "F" Or Kl_Tueren_D = "F" Or Kl_Tueren_A = "F" _
        '        Or Kl_Aussenbauteile = "F" Or Kl_Wasser = "F" Or _
        '        Kl_lauteRaeume_keine = "F" Or Kl_lauteRaeume_25 = "F" Or Kl_lauteRaeume_30 = "F" Or Kl_lauteRaeume_35 = "F" Then
        '    If CInt(P_Pkte_II) < 30 Then
        '        Return "F"
        '    Else
        '        Return "E"
        '    End If
        'ElseIf Kl_LS_W = "E" Or Kl_LS_D = "E" Or Kl_TS_D = "E" Or Kl_TS_TPH = "E" Or Kl_TS_BLLT = "E" Or Kl_Tueren_D = "E" Or Kl_Tueren_A = "E" _
        '        Or Kl_Aussenbauteile = "E" Or Kl_Wasser = "E" Or _
        '        Kl_lauteRaeume_keine = "E" Or Kl_lauteRaeume_25 = "E" Or Kl_lauteRaeume_30 = "E" Or Kl_lauteRaeume_35 = "E" Then
        '    If CInt(P_Pkte_II) < 80 Then
        '        Return "E"
        '    Else
        '        Return "D"
        '    End If
        'ElseIf Kl_LS_W = "D" Or Kl_LS_D = "D" Or Kl_TS_D = "D" Or Kl_TS_TPH = "D" Or Kl_TS_BLLT = "D" Or Kl_Tueren_D = "D" Or Kl_Tueren_A = "D" _
        '        Or Kl_Aussenbauteile = "D" Or Kl_Wasser = "D" Or _
        '        Kl_lauteRaeume_keine = "D" Or Kl_lauteRaeume_25 = "D" Or Kl_lauteRaeume_30 = "D" Or Kl_lauteRaeume_35 = "D" Then
        '    If CInt(P_Pkte_II) < 145 Then
        '        Return "D"
        '    Else
        '        Return "C"
        '    End If
        'ElseIf Kl_LS_W = "C" Or Kl_LS_D = "C" Or Kl_TS_D = "C" Or Kl_TS_TPH = "C" Or Kl_TS_BLLT = "C" Or Kl_Tueren_D = "C" Or Kl_Tueren_A = "C" _
        '        Or Kl_Aussenbauteile = "C" Or Kl_Wasser = "C" Or _
        '        Kl_lauteRaeume_keine = "C" Or Kl_lauteRaeume_25 = "C" Or Kl_lauteRaeume_30 = "C" Or Kl_lauteRaeume_35 = "C" Then
        '    If CInt(P_Pkte_II) < 210 Then
        '        Return "C"
        '    Else
        '        Return "B"
        '    End If
        'ElseIf Kl_LS_W = "B" Or Kl_LS_D = "B" Or Kl_TS_D = "B" Or Kl_TS_TPH = "B" Or Kl_TS_BLLT = "B" Or Kl_Tueren_D = "B" Or Kl_Tueren_A = "B" _
        '        Or Kl_Aussenbauteile = "B" Or Kl_Wasser = "B" Or _
        '        Kl_lauteRaeume_keine = "B" Or Kl_lauteRaeume_25 = "B" Or Kl_lauteRaeume_30 = "B" Or Kl_lauteRaeume_35 = "B" Then
        '    If CInt(P_Pkte_II) < 270 Then
        '        Return "B"
        '    Else
        '        Return "A"
        '    End If
        'ElseIf Kl_LS_W = "A" Or Kl_LS_D = "A" Or Kl_TS_D = "A" Or Kl_TS_TPH = "A" Or Kl_TS_BLLT = "A" Or Kl_Tueren_D = "A" Or Kl_Tueren_A = "A" _
        '        Or Kl_Aussenbauteile = "A" Or Kl_Wasser = "A" Or _
        '        Kl_lauteRaeume_keine = "A" Or Kl_lauteRaeume_25 = "A" Or Kl_lauteRaeume_30 = "A" Or Kl_lauteRaeume_35 = "A" Then
        '    If CInt(P_Pkte_II) < 340 Then
        '        Return "A"
        '    Else
        '        Return "A*"
        '    End If
        'Else
        '    Return ""
        'End If
    End Function
    Public Function Get_Klasse_Color(ByVal Klasse As String) As Color

        Select Case Klasse
            Case "A*"
                Get_Klasse_Color = Color.FromArgb(255, 0, 153, 153) '.DarkGreen
            Case "A"
                Get_Klasse_Color = Color.FromArgb(255, 0, 158, 0) 'SeaGreen
            Case "B"
                Get_Klasse_Color = Color.FromArgb(255, 85, 182, 0) 'YellowGreen
            Case "C"
                Get_Klasse_Color = Color.FromArgb(255, 170, 206, 0) 'Yellow
            Case "D"
                Get_Klasse_Color = Color.FromArgb(255, 255, 255, 0) 'DarkOrange
            Case "E"
                Get_Klasse_Color = Color.FromArgb(255, 255, 116, 0) 'OrangeRed
            Case "F"
                Get_Klasse_Color = Color.FromArgb(255, 255, 0, 0) 'Red
            Case Else
                Get_Klasse_Color = Color.Transparent 'White
        End Select
    End Function
    Public Function Get_Klasse_GC() As String
        Get_Klasse_GC = ""
        Select Case Projekt.Standort.Gebietscharakter
            Case GC_WR
                Get_Klasse_GC = "A*"
            Case GC_WA
                Get_Klasse_GC = "A"
            Case GC_MIWB
                Get_Klasse_GC = "C"
            Case GC_GE
                Get_Klasse_GC = "E"
            Case GC_GI
                Get_Klasse_GC = "F"
        End Select
    End Function
    Public Function Get_Klasse_AP() As String
        Get_Klasse_AP = ""
        Select Case Projekt.Standort.Aussenlaermpegel
            Case AP_bis55
                Get_Klasse_AP = "A"
            Case AP_56_60
                Get_Klasse_AP = "B"
            Case AP_61_65
                Get_Klasse_AP = "C"
            Case AP_66_70
                Get_Klasse_AP = "D"
            Case AP_71_75
                Get_Klasse_AP = "E"
            Case AP_76
                Get_Klasse_AP = "F"
        End Select
    End Function
    Public Function Get_Klasse_LS_Waende() As String
        Get_Klasse_LS_Waende = ""
        With Projekt.LS_Wand
            If .Untersuchung = nv Then
                Return "-" '"A*"
            ElseIf .Untersuchung = Prognose Then
                If .Prognose.Pegel > 0 Then
                    Dim Rw As Single = .Prognose.Pegel
                    If Rw >= 72 Then
                        Get_Klasse_LS_Waende = "A*"
                    ElseIf Rw >= 67 Then
                        Get_Klasse_LS_Waende = "A"
                    ElseIf Rw >= 62 Then
                        Get_Klasse_LS_Waende = "B"
                    ElseIf Rw >= 56 Then
                        Get_Klasse_LS_Waende = "C"
                    ElseIf Rw >= 53 Then
                        Get_Klasse_LS_Waende = "D"
                    ElseIf Rw >= 50 Then
                        Get_Klasse_LS_Waende = "E"
                    Else
                        Get_Klasse_LS_Waende = "F"
                    End If
                End If
            ElseIf .Untersuchung = Messung Then
                Dim iMes As Integer = Get_Mes_LS_Waende()
                If iMes > 0 Then
                    Dim Rw As Single
                    If iMes = 1 Then
                        Rw = .Messung.Messung_1.Pegel
                    ElseIf iMes = 2 Then
                        Rw = .Messung.Messung_2.Pegel
                    ElseIf iMes = 3 Then
                        Rw = .Messung.Messung_3.Pegel
                    ElseIf iMes = 4 Then
                        Rw = .Messung.Messung_4.Pegel
                    ElseIf iMes = 5 Then
                        Rw = .Messung.Messung_5.Pegel
                    ElseIf iMes = 6 Then
                        Rw = .Messung.Messung_6.Pegel
                    End If
                    If Rw >= 72 Then
                        Get_Klasse_LS_Waende = "A*"
                    ElseIf Rw >= 67 Then
                        Get_Klasse_LS_Waende = "A"
                    ElseIf Rw >= 62 Then
                        Get_Klasse_LS_Waende = "B"
                    ElseIf Rw >= 56 Then
                        Get_Klasse_LS_Waende = "C"
                    ElseIf Rw >= 53 Then
                        Get_Klasse_LS_Waende = "D"
                    ElseIf Rw >= 50 Then
                        Get_Klasse_LS_Waende = "E"
                    Else
                        Get_Klasse_LS_Waende = "F"
                    End If
                End If
            End If
        End With
    End Function
    Public Function Get_Klasse_LS_Decken() As String
        Get_Klasse_LS_Decken = ""
        With Projekt.LS_Decke
            If .Untersuchung = nv Then
                Return "-" '"A*"
            ElseIf .Untersuchung = Prognose Then
                If .Prognose.Pegel > 0 Then
                    Dim rw As Single = .Prognose.Pegel
                    If rw >= 72 Then
                        Get_Klasse_LS_Decken = "A*"
                    ElseIf rw >= 67 Then
                        Get_Klasse_LS_Decken = "A"
                    ElseIf rw >= 62 Then
                        Get_Klasse_LS_Decken = "B"
                    ElseIf rw >= 57 Then
                        Get_Klasse_LS_Decken = "C"
                    ElseIf rw >= 54 Then
                        Get_Klasse_LS_Decken = "D"
                    ElseIf rw >= 50 Then
                        Get_Klasse_LS_Decken = "E"
                    Else
                        Get_Klasse_LS_Decken = "F"
                    End If
                End If
            ElseIf .Untersuchung = Messung Then
                Dim iMes As Integer = Get_Mes_LS_Decken()
                If iMes > 0 Then
                    Dim rw As Single
                    If iMes = 1 Then
                        rw = .Messung.Messung_1.Pegel
                    ElseIf iMes = 2 Then
                        rw = .Messung.Messung_2.Pegel
                    ElseIf iMes = 3 Then
                        rw = .Messung.Messung_3.Pegel
                    ElseIf iMes = 4 Then
                        rw = .Messung.Messung_4.Pegel
                    ElseIf iMes = 5 Then
                        rw = .Messung.Messung_5.Pegel
                    ElseIf iMes = 6 Then
                        rw = .Messung.Messung_6.Pegel
                    End If
                    If rw >= 72 Then
                        Get_Klasse_LS_Decken = "A*"
                    ElseIf rw >= 67 Then
                        Get_Klasse_LS_Decken = "A"
                    ElseIf rw >= 62 Then
                        Get_Klasse_LS_Decken = "B"
                    ElseIf rw >= 57 Then
                        Get_Klasse_LS_Decken = "C"
                    ElseIf rw >= 54 Then
                        Get_Klasse_LS_Decken = "D"
                    ElseIf rw >= 50 Then
                        Get_Klasse_LS_Decken = "E"
                    Else
                        Get_Klasse_LS_Decken = "F"
                    End If
                End If
            End If
        End With
    End Function
    Public Function Get_Klasse_TS_Decken() As String
        Get_Klasse_TS_Decken = ""
        With Projekt.TS_Decke
            If .Untersuchung = nv Then
                Return "-" '"A*"
            ElseIf .Untersuchung = Prognose Then
                Dim rw As Single = .Prognose.Pegel.Pegel
                If rw <= 30 Then
                    Get_Klasse_TS_Decken = "A*"
                ElseIf rw <= 35 Then
                    Get_Klasse_TS_Decken = "A"
                ElseIf rw <= 40 Then
                    Get_Klasse_TS_Decken = "B"
                ElseIf rw <= 45 Then
                    Get_Klasse_TS_Decken = "C"
                ElseIf rw <= 50 Then
                    Get_Klasse_TS_Decken = "D"
                ElseIf rw <= 60 Then
                    Get_Klasse_TS_Decken = "E"
                Else
                    Get_Klasse_TS_Decken = "F"
                End If
            ElseIf .Untersuchung = Messung Then
                Dim iMes As Integer = Get_Mes_TS_Decken()
                If iMes > 0 Then
                    Dim rw As Single
                    If iMes = 1 Then
                        rw = .Messung.Messung_1.Pegel.Pegel
                    ElseIf iMes = 2 Then
                        rw = .Messung.Messung_2.Pegel.Pegel
                    ElseIf iMes = 3 Then
                        rw = .Messung.Messung_3.Pegel.Pegel
                    ElseIf iMes = 4 Then
                        rw = .Messung.Messung_4.Pegel.Pegel
                    ElseIf iMes = 5 Then
                        rw = .Messung.Messung_5.Pegel.Pegel
                    ElseIf iMes = 6 Then
                        rw = .Messung.Messung_6.Pegel.Pegel
                    End If
                    If rw <= 30 Then
                        Get_Klasse_TS_Decken = "A*"
                    ElseIf rw <= 35 Then
                        Get_Klasse_TS_Decken = "A"
                    ElseIf rw <= 40 Then
                        Get_Klasse_TS_Decken = "B"
                    ElseIf rw <= 45 Then
                        Get_Klasse_TS_Decken = "C"
                    ElseIf rw <= 50 Then
                        Get_Klasse_TS_Decken = "D"
                    ElseIf rw <= 60 Then
                        Get_Klasse_TS_Decken = "E"
                    Else
                        Get_Klasse_TS_Decken = "F"
                    End If
                End If
            End If
        End With
    End Function
    Public Function Get_Klasse_TS_TPLH() As String
        Get_Klasse_TS_TPLH = ""
        With Projekt.TS_TPlH
            If .Untersuchung = nv Then
                Return "-" '"A*"
            Else
                Dim rw As Single = -1
                Dim bt As Byte = 0

                If .Untersuchung = Prognose Then
                    Dim btBauteil As Byte = Get_Prog_TS_TPLH()
                    If btBauteil > 0 Then
                        'Dim rw As Single
                        If btBauteil = Treppe Then
                            rw = .Prognose.Treppe.Pegel.Pegel
                            bt = Treppe
                        ElseIf btBauteil = Podest Then
                            rw = .Prognose.Podest.Pegel.Pegel
                            bt = Podest
                        ElseIf btBauteil = Laube Then
                            rw = .Prognose.Laube.Pegel.Pegel
                            bt = Laube
                        ElseIf btBauteil = Hausflur Then
                            rw = .Prognose.Hausflur.Pegel.Pegel
                            bt = Hausflur
                        End If
                    End If
                ElseIf .Untersuchung = Messung Then
                    Dim iMes As Integer = Get_Mes_TS_TPLH()
                    If iMes > 0 Then
                        If iMes = 1 Then
                            rw = .Messung.Messung_1.Pegel.Pegel
                            bt = .Messung.Messung_1.Bauteil
                        ElseIf iMes = 2 Then
                            rw = .Messung.Messung_2.Pegel.Pegel
                            bt = .Messung.Messung_2.Bauteil
                        ElseIf iMes = 3 Then
                            rw = .Messung.Messung_3.Pegel.Pegel
                            bt = .Messung.Messung_3.Bauteil
                        ElseIf iMes = 4 Then
                            rw = .Messung.Messung_4.Pegel.Pegel
                            bt = .Messung.Messung_4.Bauteil
                        ElseIf iMes = 5 Then
                            rw = .Messung.Messung_5.Pegel.Pegel
                            bt = .Messung.Messung_5.Bauteil
                        ElseIf iMes = 6 Then
                            rw = .Messung.Messung_6.Pegel.Pegel
                            bt = .Messung.Messung_6.Bauteil
                        End If
                        'If rw <= 28 Then
                        '    Get_Klasse_TS_TPH = "A*"
                        'ElseIf rw <= 34 Then
                        '    Get_Klasse_TS_TPH = "A"
                        'ElseIf rw <= 40 Then
                        '    Get_Klasse_TS_TPH = "B"
                        'ElseIf rw <= 46 Then
                        '    Get_Klasse_TS_TPH = "C"
                        'ElseIf rw <= 53 Then
                        '    Get_Klasse_TS_TPH = "D"
                        'ElseIf rw <= 60 Then
                        '    Get_Klasse_TS_TPH = "E"
                        'Else
                        '    Get_Klasse_TS_TPH = "F"
                        'End If
                    End If
                End If
                If rw > -1 Then
                    If rw <= 33 Then
                        Get_Klasse_TS_TPLH = "A*"
                    ElseIf rw <= 38 Then
                        Get_Klasse_TS_TPLH = "A"
                    ElseIf rw <= 43 Then
                        Get_Klasse_TS_TPLH = "B"
                    ElseIf rw <= 48 Then
                        Get_Klasse_TS_TPLH = "C"
                    ElseIf (rw <= 53 And bt <> Hausflur) Or (rw <= 50 And bt = Hausflur) Then
                        Get_Klasse_TS_TPLH = "D"
                    ElseIf rw <= 63 Then
                        Get_Klasse_TS_TPLH = "E"
                    Else
                        Get_Klasse_TS_TPLH = "F"
                    End If
                End If
            End If
        End With
    End Function
    Public Function Get_Klasse_TS_BLT() As String
        Get_Klasse_TS_BLT = ""
        With Projekt.TS_BLT
            If .Untersuchung = nv Then
                Return "-" '"A*"
            Else
                Dim rw As Single = -1
                Dim bt As Byte = 0

                If .Untersuchung = Prognose Then
                    Dim btBauteil As Byte = Get_Prog_TS_BLT()
                    If btBauteil > 0 Then
                        If btBauteil = Balkon Then
                            rw = .Prognose.Balkon.Pegel.Pegel
                            bt = Balkon
                        ElseIf btBauteil = Loggie Then
                            rw = .Prognose.Loggia.Pegel.Pegel
                            bt = Loggie
                        ElseIf btBauteil = Terrasse Then
                            rw = .Prognose.Terrasse.Pegel.Pegel
                            bt = Terrasse
                        End If
                    End If
                ElseIf .Untersuchung = Messung Then
                    Dim iMes As Integer = Get_Mes_TS_BLT()

                    If iMes > 0 Then
                        'Dim rw As Single
                        If iMes = 1 Then
                            rw = .Messung.Messung_1.Pegel.Pegel
                            bt = .Messung.Messung_1.Bauteil
                        ElseIf iMes = 2 Then
                            rw = .Messung.Messung_2.Pegel.Pegel
                            bt = .Messung.Messung_2.Bauteil
                        ElseIf iMes = 3 Then
                            rw = .Messung.Messung_3.Pegel.Pegel
                            bt = .Messung.Messung_3.Bauteil
                        ElseIf iMes = 4 Then
                            rw = .Messung.Messung_4.Pegel.Pegel
                            bt = .Messung.Messung_4.Bauteil
                        ElseIf iMes = 5 Then
                            rw = .Messung.Messung_5.Pegel.Pegel
                            bt = .Messung.Messung_5.Bauteil
                        ElseIf iMes = 6 Then
                            rw = .Messung.Messung_6.Pegel.Pegel
                            bt = .Messung.Messung_6.Bauteil
                        End If
                    End If
                End If
                If rw > -1 Then
                    If rw <= 33 Then
                        Get_Klasse_TS_BLT = "A*"
                    ElseIf rw <= 38 Then
                        Get_Klasse_TS_BLT = "A"
                    ElseIf rw <= 43 Then
                        Get_Klasse_TS_BLT = "B"
                    ElseIf rw <= 48 Then
                        Get_Klasse_TS_BLT = "C"
                    ElseIf (rw <= 50 And bt <> Balkon) Or (rw <= 58 And bt = Balkon) Then
                        Get_Klasse_TS_BLT = "D"
                    ElseIf rw <= 63 Then
                        Get_Klasse_TS_BLT = "E"
                    Else
                        Get_Klasse_TS_BLT = "F"
                    End If
                End If
            End If
        End With
    End Function

    Public Function Get_Klasse_Tueren() As String
        Get_Klasse_Tueren = ""
        'F34    Diele Prüfzeugnis
        'G34    Diele Messung
        'F35    Flur Prüfzeugnis
        'G35    Flur Messung
        '=WENN(UND(F34="-";G34="-";F35="-";G35="-");"A";
        '    WENN(ODER(F34="";G34="";J34="";UND(F34="-";G34="-"));"";
        '        WENN(J34<>"-";WENN(J34>=40;"A";WENN(J34>=37;"B";WENN(J34>=32;"C";WENN(J34>=27;"D";WENN(J34>=22;"E";"F")))));"---")))

        With Projekt.Tueren
            If .Ort = nv Then
                Get_Klasse_Tueren = "-" '"A"
            ElseIf .Ort = Diele Then
                If .Untersuchung = Pruefzeugnis Or .Untersuchung = Messung Then
                    If .L >= 40 Then
                        Get_Klasse_Tueren = "A"
                    ElseIf .L >= 37 Then
                        Get_Klasse_Tueren = "B"
                    ElseIf .L >= 32 Then
                        Get_Klasse_Tueren = "C"
                    ElseIf .L >= 27 Then
                        Get_Klasse_Tueren = "D"
                    ElseIf .L >= 22 Then
                        Get_Klasse_Tueren = "E"
                    Else
                        Get_Klasse_Tueren = "F"
                    End If
                End If
            ElseIf .Ort = Aufenthalt Then
                If .Untersuchung = Pruefzeugnis Or .Untersuchung = Messung Then
                    'If .L >= 48 Then
                    '    Get_Klasse_Tueren = "A"
                    'ElseIf .L >= 45 Then
                    '    Get_Klasse_Tueren = "B"
                    'Else
                    If .L >= 42 Then
                        Get_Klasse_Tueren = "C"
                    ElseIf .L >= 37 Then
                        Get_Klasse_Tueren = "D"
                    ElseIf .L >= 32 Then
                        Get_Klasse_Tueren = "E"
                    Else
                        Get_Klasse_Tueren = "F"
                    End If
                End If
            End If
        End With
    End Function
    Public Function Get_Klasse_Tueren_D() As String
        Get_Klasse_Tueren_D = ""
        With Projekt.Tueren
            If .Ort = Diele Then
                If .Untersuchung = Pruefzeugnis Or .Untersuchung = Messung Then
                    If .L >= 40 Then
                        Get_Klasse_Tueren_D = "A"
                    ElseIf .L >= 37 Then
                        Get_Klasse_Tueren_D = "B"
                    ElseIf .L >= 32 Then
                        Get_Klasse_Tueren_D = "C"
                    ElseIf .L >= 27 Then
                        Get_Klasse_Tueren_D = "D"
                    ElseIf .L >= 22 Then
                        Get_Klasse_Tueren_D = "E"
                    Else
                        Get_Klasse_Tueren_D = "F"
                    End If
                End If
            ElseIf .Ort = nv Then
                Get_Klasse_Tueren_D = "-" '"A"
            End If
        End With
    End Function
    Public Function Get_Klasse_Tueren_A() As String
        Get_Klasse_Tueren_A = ""
        With Projekt.Tueren
            If .Ort = Aufenthalt Then
                If .Untersuchung = Pruefzeugnis Or .Untersuchung = Messung Then
                    'If .L >= 48 Then
                    '    Get_Klasse_Tueren_A = "A"
                    'ElseIf .L >= 45 Then
                    '    Get_Klasse_Tueren_A = "B"
                    'Else
                    If .L >= 42 Then
                        Get_Klasse_Tueren_A = "C"
                    ElseIf .L >= 37 Then
                        Get_Klasse_Tueren_A = "D"
                    ElseIf .L >= 32 Then
                        Get_Klasse_Tueren_A = "E"
                    Else
                        Get_Klasse_Tueren_A = "F"
                    End If
                End If
            End If
        End With
    End Function
    Public Function Get_Klasse_Aussenbauteile() As String
        Get_Klasse_Aussenbauteile = ""
        Select Case Projekt.Aussenbauteile
            Case ohneNachweis
                Get_Klasse_Aussenbauteile = "E"
                'Case FensterMitDichtung
                '    Get_Klasse_Aussenbauteile = "E"
            Case DINerfuellt
                Get_Klasse_Aussenbauteile = "A"
            Case DINPlusErfuellt
                Get_Klasse_Aussenbauteile = "A*"
        End Select
    End Function
    Public Function Get_Klasse_Wasser() As String
        Get_Klasse_Wasser = ""
        With Projekt.Wasser
            If (.Untersuchung = Prognose Or .Untersuchung = Messung) And (.LcLa_erfuellt = erfuellt Or .LcLa_erfuellt = keineAngabe) Then
                If .L_Intervall = kleiner20 Then ' <= 20 Then
                    Get_Klasse_Wasser = "A*"
                ElseIf .L_Intervall = HT_20_24 Then ' <= 25 Then
                    Get_Klasse_Wasser = "B"
                ElseIf .L_Intervall = HT_24_27 Then ' <= 30 Then
                    Get_Klasse_Wasser = "C"
                ElseIf .L_Intervall = HT_27_30 Then ' <= 35 Then
                    Get_Klasse_Wasser = "D"
                ElseIf .L_Intervall = HT_30_35 Then
                    Get_Klasse_Wasser = "E"
                ElseIf .L_Intervall = HT_groesser35 Then
                    Get_Klasse_Wasser = "F"
                End If
            End If
        End With
    End Function
    Public Function Get_Klasse_NutzerKoerper() As String
        Get_Klasse_NutzerKoerper = ""
        If Projekt.NutzerKoerper.Untersuchung = Nutzer Then
            Select Case Projekt.NutzerKoerper.L_Intervall
                Case kleiner20
                    Get_Klasse_NutzerKoerper = "A*"
                Case L_20_25
                    Get_Klasse_NutzerKoerper = "A"
                Case L_25_30
                    Get_Klasse_NutzerKoerper = "B"
                Case L_30_35
                    Get_Klasse_NutzerKoerper = "C"
                Case L_35_40
                    Get_Klasse_NutzerKoerper = "D"
                Case L_40_45
                    Get_Klasse_NutzerKoerper = "E"
                Case groesser45
                    Get_Klasse_NutzerKoerper = "F"
            End Select
        ElseIf Projekt.NutzerKoerper.Untersuchung = Koerper Then
            Select Case Projekt.NutzerKoerper.L_Intervall
                Case kleiner38
                    Get_Klasse_NutzerKoerper = "A*"
                Case L_38_43
                    Get_Klasse_NutzerKoerper = "A"
                Case L_43_48
                    Get_Klasse_NutzerKoerper = "B"
                Case L_48_53
                    Get_Klasse_NutzerKoerper = "C"
                Case L_53_58
                    Get_Klasse_NutzerKoerper = "D"
                Case L_58_63
                    Get_Klasse_NutzerKoerper = "E"
                Case groesser63
                    Get_Klasse_NutzerKoerper = "F"
            End Select
        Else
            Get_Klasse_NutzerKoerper = ""
        End If
    End Function
    'Public Function Get_Klasse_Koerper() As String
    '    Get_Klasse_Koerper = ""
    '    Select Case Projekt.Koerper.L_Intervall
    '        Case kleiner38
    '            Get_Klasse_Koerper = "A*"
    '        Case L_38_43
    '            Get_Klasse_Koerper = "A"
    '        Case L_43_48
    '            Get_Klasse_Koerper = "B"
    '        Case L_48_53
    '            Get_Klasse_Koerper = "C"
    '        Case L_53_58
    '            Get_Klasse_Koerper = "D"
    '        Case L_58_63
    '            Get_Klasse_Koerper = "E"
    '        Case groesser63
    '            Get_Klasse_Koerper = "F"
    '    End Select
    'End Function
    Public Function Get_Klasse_Nachbarn() As String
        Get_Klasse_Nachbarn = ""
        Select Case Projekt.Nachbarn
            Case 1, 2
                Get_Klasse_Nachbarn = "A*"
            Case 3
                Get_Klasse_Nachbarn = "A"
            Case 4
                Get_Klasse_Nachbarn = "B"
            Case 5
                Get_Klasse_Nachbarn = "C" '"---"
        End Select
    End Function
    Public Function Get_Klasse_lauteRaeume() As String
        Get_Klasse_lauteRaeume = ""
        Select Case Projekt.lauteRaeume
            Case keineAngrenzend
                Return "A*"
            Case L_25_35
                Return "B"
            Case L_30_40
                Return "D"
            Case L_35_45
                Return "E"
        End Select
    End Function
    Public Function Get_Klasse_lauteRaeume_kE() As String
        Get_Klasse_lauteRaeume_kE = ""
        Select Case Projekt.lauteRaeume
            Case keineAngrenzend
                Get_Klasse_lauteRaeume_kE = "A*"
            Case L_25_35, L_30_40, L_35_45
                Get_Klasse_lauteRaeume_kE = "---"
        End Select
    End Function
    Public Function Get_Klasse_lauteRaeume_25() As String
        Get_Klasse_lauteRaeume_25 = ""
        Select Case Projekt.lauteRaeume
            Case L_25_35
                Get_Klasse_lauteRaeume_25 = "B"
            Case keineAngrenzend, L_30_40, L_35_45
                Get_Klasse_lauteRaeume_25 = "---"
        End Select
    End Function
    Public Function Get_Klasse_lauteRaeume_30() As String
        Get_Klasse_lauteRaeume_30 = ""
        Select Case Projekt.lauteRaeume
            Case L_30_40
                Get_Klasse_lauteRaeume_30 = "D"
            Case keineAngrenzend, L_25_35, L_35_45
                Get_Klasse_lauteRaeume_30 = "---"
        End Select
    End Function
    Public Function Get_Klasse_lauteRaeume_35() As String
        Get_Klasse_lauteRaeume_35 = ""
        Select Case Projekt.lauteRaeume
            Case L_35_45
                Get_Klasse_lauteRaeume_35 = "E"
            Case keineAngrenzend, L_25_35, L_30_40
                Get_Klasse_lauteRaeume_35 = "---"
        End Select
    End Function
    Public Function Get_Klasse_NHZ() As String
        Get_Klasse_NHZ = ""
        Select Case Projekt.NHZ
            Case NHZ_020
                Get_Klasse_NHZ = "A*"
            Case NHZ_010
                Get_Klasse_NHZ = "B"
            Case NHZ_keine
                Get_Klasse_NHZ = "-"
        End Select
        Return Get_Klasse_NHZ
    End Function
    Public Function Get_Klasse_EW() As String
        Get_Klasse_EW = ""
        Select Case Projekt.eigenerWohnbereich
            Case keineEmpfehlung
                Get_Klasse_EW = "A"
            Case EW1
                Get_Klasse_EW = "A"
            Case EW2
                Get_Klasse_EW = "A*"
            Case EW3
                Get_Klasse_EW = "A*"
        End Select
    End Function
#End Region
#Region "Get_Mes_*"
    Private Function Get_Mes_LS_Waende() As Integer
        '1. kleinster Pegel
        Dim I As Integer
        Dim iMes As Integer
        iMes = -1
        Dim Rw As Double
        'Dim sMV As String = ""

        With Form_Main
            For I = 1 To 6
                'Dim kmvButton As System.Windows.Forms.Button = CType(.Panel_SSA_LS_W_Messung.Controls.Find("Button_LS_W_MV_" & I & "_KMV", True)(0), System.Windows.Forms.Button)
                'Dim nmvButton As System.Windows.Forms.Button = CType(.Panel_SSA_LS_W_Messung.Controls.Find("Button_LS_W_MV_" & I & "_NMV", True)(0), System.Windows.Forms.Button)
                Dim pegelNUD As NumericUpDown = CType(.Panel_SSA_LS_W_Messung.Controls.Find("NUD_LS_W_Mes_R_" & I, True)(0), NumericUpDown)
                Dim cNUD As NumericUpDown = CType(.Panel_SSA_LS_W_Messung.Controls.Find("NUD_LS_W_Mes_C_" & I, True)(0), NumericUpDown)
                If pegelNUD.Value > 0 Then '(kmvButton.BackColor = colRed Or nmvButton.BackColor = colRed) And pegelNUD.Value > 0 Then
                    If pegelNUD.Value < Rw Or iMes = -1 Then
                        iMes = I
                        Rw = pegelNUD.Value
                       
                    End If
                End If
            Next
        End With
        Get_Mes_LS_Waende = iMes
    End Function
    Private Function Get_Mes_LS_Decken() As Integer
        '1. kleinster Pegel
        Dim I As Integer
        Dim iMes As Integer
        iMes = -1
        Dim Rw As Double
        'Dim sMV As String = ""

        With Form_Main
            For I = 1 To 6
                'Dim kmvButton As System.Windows.Forms.Button = CType(.Panel_SSA_LS_D_Messung.Controls.Find("Button_LS_D_MV_" & I & "_KMV", True)(0), System.Windows.Forms.Button)
                'Dim nmvButton As System.Windows.Forms.Button = CType(.Panel_SSA_LS_D_Messung.Controls.Find("Button_LS_D_MV_" & I & "_NMV", True)(0), System.Windows.Forms.Button)
                Dim pegelNUD As NumericUpDown = CType(.Panel_SSA_LS_D_Messung.Controls.Find("NUD_LS_D_Mes_R_" & I, True)(0), NumericUpDown)
                Dim cNUD As NumericUpDown = CType(.Panel_SSA_LS_D_Messung.Controls.Find("NUD_LS_D_Mes_C_" & I, True)(0), NumericUpDown)
                If pegelNUD.Value > 0 Then ' (kmvButton.BackColor = colRed Or nmvButton.BackColor = colRed) And pegelNUD.Value > 0 Then
                    If pegelNUD.Value < Rw Or iMes = -1 Then
                        iMes = I
                        Rw = pegelNUD.Value
                        
                    End If
                End If
            Next
        End With
        Get_Mes_LS_Decken = iMes
    End Function
    Private Function Get_Mes_TS_Decken() As Integer
        '1. kleinster Pegel
        Dim I As Integer
        Dim iMes As Integer
        iMes = -1
        Dim Rw As Double
        Dim sMV As String = ""

        With Form_Main
            For I = 1 To 6
                'Dim kmvButton As System.Windows.Forms.Button = CType(.Panel_SSA_TS_D_Messung.Controls.Find("Button_TS_D_MV_" & I & "_KMV", True)(0), System.Windows.Forms.Button)
                'Dim nmvButton As System.Windows.Forms.Button = CType(.Panel_SSA_TS_D_Messung.Controls.Find("Button_TS_D_MV_" & I & "_NMV", True)(0), System.Windows.Forms.Button)
                Dim pegelNUD As NumericUpDown = CType(.Panel_SSA_TS_D_Messung.Controls.Find("NUD_TS_D_Mes_L_" & I, True)(0), NumericUpDown)
                Dim cNUD As NumericUpDown = CType(.Panel_SSA_TS_D_Messung.Controls.Find("NUD_TS_D_Mes_C_" & I, True)(0), NumericUpDown)
                If pegelNUD.Value > 0 Then '(kmvButton.BackColor = colRed Or nmvButton.BackColor = colRed) And pegelNUD.Value > 0 Then
                    If pegelNUD.Value > Rw Or iMes = -1 Then
                        iMes = I
                        Rw = pegelNUD.Value
                        
                    End If
                End If
            Next
        End With
        Get_Mes_TS_Decken = iMes
    End Function
    Private Function Get_Mes_TS_TPLH() As Integer
        '1. kleinster Pegel
        Dim I As Integer
        Dim iMes As Integer
        iMes = -1
        Dim L As Double
        Dim c As String = ""
        Dim sMV As String = ""
        Dim iPkte As Integer = 0
        '## über Variablen
        With Projekt.TS_TPLH
            Dim tPkte As Integer
            For I = 1 To 6
                Dim bt As Byte = 0
                If .Untersuchung = Messung Then
                    'Dim iMes As Integer = Get_Mes_TS_TPLH()
                    If I = 1 Then
                        L = .Messung.Messung_1.Pegel.Pegel
                        c = .Messung.Messung_1.Pegel.C
                        bt = .Messung.Messung_1.Bauteil
                    ElseIf I = 2 Then
                        L = .Messung.Messung_2.Pegel.Pegel
                        c = .Messung.Messung_2.Pegel.C
                        bt = .Messung.Messung_2.Bauteil
                    ElseIf I = 3 Then
                        L = .Messung.Messung_3.Pegel.Pegel
                        c = .Messung.Messung_3.Pegel.C
                        bt = .Messung.Messung_3.Bauteil
                    ElseIf I = 4 Then
                        L = .Messung.Messung_4.Pegel.Pegel
                        c = .Messung.Messung_4.Pegel.C
                        bt = .Messung.Messung_4.Bauteil
                    ElseIf I = 5 Then
                        L = .Messung.Messung_5.Pegel.Pegel
                        c = .Messung.Messung_5.Pegel.C
                        bt = .Messung.Messung_5.Bauteil
                    ElseIf I = 6 Then
                        L = .Messung.Messung_6.Pegel.Pegel
                        c = .Messung.Messung_6.Pegel.C
                        bt = .Messung.Messung_6.Bauteil
                    End If
                End If
                If L <= 33 Then
                    tPkte = 50 + Get_Bonus_TS_TPLH_L_C(L, c, bt)  '<- auch mit/ohne Bonus kann nicht ein besseres/schlechteres Ergebnis erziehlt werden
                ElseIf L <= 38 Then
                    tPkte = 40 + Get_Bonus_TS_TPLH_L_C(L, c, bt)
                ElseIf L <= 43 Then
                    tPkte = 30 + Get_Bonus_TS_TPLH_L_C(L, c, bt)
                ElseIf L <= 48 Then
                    tPkte = 20 + Get_Bonus_TS_TPLH_L_C(L, B, bt)
                ElseIf (L <= 53 And bt <> Hausflur) Or (L <= 50 And bt = Hausflur) Then
                    'ElseIf L <= 53 Then
                    tPkte = 10 + Get_Bonus_TS_TPLH_L_C(L, B, bt)
                ElseIf L <= 63 Then
                    tPkte = 5 '+ Get_Bonus_TS_TPLH_L_C(L, B, bt)
                Else
                    tPkte = 0
                End If
                If tPkte < iPkte Or iMes = -1 Then
                    iMes = I
                    iPkte = tPkte
                End If
            Next
        End With

       
        Get_Mes_TS_TPLH = iMes
    End Function
    Private Function Get_Prog_TS_TPLH() As Byte
        '1. kleinster Pegel
        Dim I As Integer
        Dim ibt As Integer
        ibt = 0
        Dim L As Double
        Dim tL As Double = 0
        Dim c As String = ""
        Dim sMV As String = ""
        Dim iPkte As Integer = 0
        '## über Variablen
        With Projekt.TS_TPLH
            Dim tPkte As Integer
            For I = 1 To 4
                Dim bt As Byte = 0
                If .Untersuchung = Prognose Then
                    'Dim iMes As Integer = Get_Mes_TS_TPLH()
                    If I = 1 Then
                        L = .Prognose.Treppe.Pegel.Pegel
                        c = .Prognose.Treppe.Pegel.C
                        bt = Treppe
                    ElseIf I = 2 Then
                        L = .Prognose.Podest.Pegel.Pegel
                        c = .Prognose.Podest.Pegel.C
                        bt = Podest
                    ElseIf I = 3 Then
                        L = .Prognose.Laube.Pegel.Pegel
                        c = .Prognose.Laube.Pegel.C
                        bt = Laube
                    ElseIf I = 4 Then
                        L = .Prognose.Hausflur.Pegel.Pegel
                        c = .Prognose.Hausflur.Pegel.C
                        bt = Hausflur
                    End If
                End If
                If L > 0 Then
                    If L <= 33 Then
                        tPkte = 50 + Get_Bonus_TS_TPLH_L_C(L, c, bt)  '<- auch mit/ohne Bonus kann nicht ein besseres/schlechteres Ergebnis erziehlt werden
                    ElseIf L <= 38 Then
                        tPkte = 40 + Get_Bonus_TS_TPLH_L_C(L, c, bt)
                    ElseIf L <= 43 Then
                        tPkte = 30 + Get_Bonus_TS_TPLH_L_C(L, c, bt)
                    ElseIf L <= 48 Then
                        tPkte = 20 + Get_Bonus_TS_TPLH_L_C(L, c, bt)
                    ElseIf (L <= 53 And bt <> Hausflur) Or (L <= 50 And bt = Hausflur) Then
                        tPkte = 10 + Get_Bonus_TS_TPLH_L_C(L, c, bt)
                    ElseIf L <= 63 Then
                        tPkte = 5 '+ Get_Bonus_TS_TPLH_L_C(L, c, bt)
                    Else
                        tPkte = 0
                    End If
                    If tPkte < iPkte Or ibt = 0 Then
                        ibt = bt
                        iPkte = tPkte
                        tL = L
                    ElseIf tPkte = iPkte And L > tL Then
                        ibt = bt
                        iPkte = tPkte
                        tL = L
                    End If
                End If
            Next
        End With
        Return ibt
        
    End Function
    'Private Function is_Bonus_TPLH(ByVal rw As Integer, ByVal c As Integer, ByVal bt As Byte) As Boolean
    '    If rw <= 33 And rw + c <= 33 Then
    '        is_Bonus_TPLH = True
    '    ElseIf rw <= 38 And rw > 33 And rw + c <= 38 Then
    '        is_Bonus_TPLH = True
    '    ElseIf rw <= 43 And rw > 38 And rw + c <= 43 Then
    '        is_Bonus_TPLH = True
    '    ElseIf rw <= 48 And rw > 43 And rw + c <= 48 Then
    '        is_Bonus_TPLH = True
    '    ElseIf bt <> Hausflur And (rw <= 53 And rw > 48 And rw + c <= 53) Then
    '        is_Bonus_TPLH = True
    '    ElseIf bt = Hausflur And (rw <= 50 And rw > 48 And rw + c <= 50) Then
    '        is_Bonus_TPLH = True
    '    ElseIf bt <> Hausflur And (rw <= 63 And rw > 53 And rw + c <= 63) Then
    '        is_Bonus_TPLH = True
    '    ElseIf bt = Hausflur And (rw <= 63 And rw > 50 And rw + c <= 63) Then
    '        is_Bonus_TPLH = True
    '    End If
    'End Function
    Private Function Get_Prog_TS_BLT() As Byte
        '1. kleinster Pegel
        Dim I As Integer
        Dim ibt As Byte
        ibt = 0
        Dim L As Double
        Dim tL As Double = 0
        Dim c As String = ""
        Dim sMV As String = ""
        Dim iPkte As Integer = 0
        '## über Variablen
        With Projekt.TS_BLT
            Dim tPkte As Integer
            For I = 1 To 3
                Dim bt As Byte = 0
                If .Untersuchung = Prognose Then
                    If I = 1 Then
                        L = .Prognose.Balkon.Pegel.Pegel
                        c = .Prognose.Balkon.Pegel.C
                        bt = Balkon
                    ElseIf I = 2 Then
                        L = .Prognose.Loggia.Pegel.Pegel
                        c = .Prognose.Loggia.Pegel.C
                        bt = Loggie
                    ElseIf I = 3 Then
                        L = .Prognose.Terrasse.Pegel.Pegel
                        c = .Prognose.Terrasse.Pegel.C
                        bt = Terrasse

                    End If
                End If
                If L > 0 Then
                    If L <= 33 Then
                        tPkte = 30 + Get_Bonus_TS_BLT_L_C(L, c, bt)  '<- auch mit/ohne Bonus kann nicht ein besseres/schlechteres Ergebnis erziehlt werden
                    ElseIf L <= 38 Then
                        tPkte = 25 + Get_Bonus_TS_BLT_L_C(L, c, bt)
                    ElseIf L <= 43 Then
                        tPkte = 20 + Get_Bonus_TS_BLT_L_C(L, c, bt)
                    ElseIf L <= 48 Then
                        tPkte = 15 + Get_Bonus_TS_BLT_L_C(L, c, bt)
                    ElseIf (L <= 58 And bt = Balkon) Or (L <= 50 And bt <> Balkon) Then
                        tPkte = 10 + Get_Bonus_TS_BLT_L_C(L, c, bt)
                    ElseIf L <= 63 Then
                        tPkte = 5
                    Else
                        tPkte = 0
                    End If
                    If tPkte < iPkte Or ibt = 0 Then
                        ibt = bt
                        iPkte = tPkte
                        tL = L
                    ElseIf tPkte = iPkte And L > tL Then
                        ibt = bt
                        iPkte = tPkte
                        tL = L
                    End If
                End If
            Next
        End With
        Return ibt
    End Function
  
    Private Function Get_Mes_TS_BLT() As Integer
        '1. kleinster Pegel
        Dim I As Integer
        Dim iMes As Integer
        iMes = -1
        Dim L As Double
        Dim c As String = ""
        Dim sMV As String = ""
        Dim iPkte As Integer = 0
        '## über Variablen
        With Projekt.TS_BLT
            Dim tPkte As Integer
            For I = 1 To 6
                Dim bt As Byte = 0
                If .Untersuchung = Messung Then
                    'Dim iMes As Integer = Get_Mes_TS_BLT()
                    If I = 1 Then
                        L = .Messung.Messung_1.Pegel.Pegel
                        c = .Messung.Messung_1.Pegel.C
                        bt = .Messung.Messung_1.Bauteil
                    ElseIf I = 2 Then
                        L = .Messung.Messung_2.Pegel.Pegel
                        c = .Messung.Messung_2.Pegel.C
                        bt = .Messung.Messung_2.Bauteil
                    ElseIf I = 3 Then
                        L = .Messung.Messung_3.Pegel.Pegel
                        c = .Messung.Messung_3.Pegel.C
                        bt = .Messung.Messung_3.Bauteil
                    ElseIf I = 4 Then
                        L = .Messung.Messung_4.Pegel.Pegel
                        c = .Messung.Messung_4.Pegel.C
                        bt = .Messung.Messung_4.Bauteil
                    ElseIf I = 5 Then
                        L = .Messung.Messung_5.Pegel.Pegel
                        c = .Messung.Messung_5.Pegel.C
                        bt = .Messung.Messung_5.Bauteil
                    ElseIf I = 6 Then
                        L = .Messung.Messung_6.Pegel.Pegel
                        c = .Messung.Messung_6.Pegel.C
                        bt = .Messung.Messung_6.Bauteil
                    End If
                End If
                If L <= 33 Then
                    tPkte = 25 + Get_Bonus_TS_BLT_L_C(L, c, bt)  '<- auch mit/ohne Bonus kann nicht ein besseres/schlechteres Ergebnis erziehlt werden
                ElseIf L <= 38 Then
                    tPkte = 20 + Get_Bonus_TS_BLT_L_C(L, c, bt)
                ElseIf L <= 43 Then
                    tPkte = 15 + Get_Bonus_TS_BLT_L_C(L, c, bt)
                ElseIf L <= 48 Then
                    tPkte = 10 + Get_Bonus_TS_BLT_L_C(L, B, bt)
                ElseIf (L <= 50 And bt <> Balkon) Or (L <= 58 And bt = Balkon) Then
                    'ElseIf L <= 53 Then
                    tPkte = 5 + Get_Bonus_TS_BLT_L_C(L, B, bt)
                Else
                    tPkte = 0
                End If
                If tPkte < iPkte Or iMes = -1 Then
                    iMes = I
                    iPkte = tPkte
                End If
            Next
        End With

        Get_Mes_TS_BLT = iMes
    End Function
#End Region
    Public Function Get_Klasse_PfeilPos(ByVal Klasse As String, ByVal yKoord As Integer) As System.Drawing.Point
        Dim tmpPoint As System.Drawing.Point
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
        Return tmpPoint
    End Function
    Public Function Get_txt_mind_II(ByVal strKlasse As String) As String
        Dim strPkte As String = ""

        Select Case strKlasse
            Case "A*"
                strPkte = "340"
            Case "A"
                strPkte = "270"
            Case "B"
                strPkte = "210"
            Case "C"
                strPkte = "145"
            Case "D"
                strPkte = "80"
            Case "E"
                strPkte = "30"
            Case "F"
                strPkte = "0"
        End Select

        Return "von mind. " & strPkte & " in Stufe " & strKlasse
    End Function
    Public Function Get_txt_mind_I(ByVal strKlasse As String) As String
        Dim strPkte As String = ""

        Select Case strKlasse
            Case "A*"
                strPkte = "55"
            Case "A"
                strPkte = "45"
            Case "B"
                strPkte = "40"
            Case "C"
                strPkte = "25"
            Case "D"
                strPkte = "20"
            Case "E"
                strPkte = "10"
            Case "F"
                strPkte = "0"
        End Select

        Return "von mind. " & strPkte & " in Stufe " & strKlasse

    End Function
    Public Function Get_txt_Bonus() As String
        '=N54+N45+N44+N43+N42+WENN(P40="A";N40-30;WENN(P40="C";N40-20;WENN(P40="D";N40-10;WENN(P40="E";N40-5;0))))+WENN(P35="A";N35-30;WENN(P35="B";N35-20;WENN(P35="C";N35-10;WENN(P35="D";N35-5;N35))))+WENN(P34="A";N34-30;WENN(P34="B";N34-20;WENN(P34="C";N34-10;WENN(P34="D";N34-5;N34))))+WENN(P32="A*";N32-25;WENN(P32="A";N32-20;WENN(P32="B";N32-15;WENN(P32="C";N32-10;WENN(P32="D";N32-5;N32)))))+WENN(P31="A*";N31-50;WENN(P31="A";N31-40;WENN(P31="B";N31-30;WENN(P31="C";N31-20;WENN(P31="D";N31-10;WENN(P31="E";N31-5;N31))))))+WENN(P30="A*";N30-50;WENN(P30="A";N30-40;WENN(P30="B";N30-30;WENN(P30="C";N30-20;WENN(P30="D";N30-10;WENN(P30="E";N30-5;N30))))))+WENN(P28="A*";N28-50;WENN(P28="A";N28-40;WENN(P28="B";N28-30;WENN(P28="C";N28-20;WENN(P28="D";N28-10;WENN(P28="E";N28-5;N28))))))+WENN(P27="A*";N27-50;WENN(P27="A";N27-40;WENN(P27="B";N27-30;WENN(P27="C";N27-20;WENN(P27="D";N27-10;WENN(P27="E";N27-5;N27))))))
        '+ CInt(Get_Pkte_NutzerKoerper())
        Dim iBonus As Integer
        '+ CInt(Get_Pkte_Nutzer())
        iBonus = CInt(Get_Pkte_EW())
        iBonus = iBonus + CInt(Get_Pkte_NHZ())
        iBonus = iBonus + CInt(Get_Pkte_anordRaeume())
        iBonus = iBonus + CInt(Get_Pkte_Nachbarn())
        iBonus = iBonus + Get_Bonus_NutzerKoerper()
        iBonus = iBonus + Get_Bonus_Wasser()
        iBonus = iBonus + Get_Bonus_Tueren()
        iBonus = iBonus + Get_Bonus_LS_D()
        iBonus = iBonus + Get_Bonus_LS_W()
        iBonus = iBonus + Get_Bonus_TS_D()
        iBonus = iBonus + Get_Bonus_TS_TPLH()
        iBonus = iBonus + Get_Bonus_TS_BLT()

        Return "(incl. " & CStr(iBonus) & " Bonuspunkte)"
    End Function
    Public Function Get_Str_Geschoss() As String
        Dim gTyp As String = Get_Str_GeschossTyp(Projekt.Wohnung.Geschoss.Typ)
        Dim gNr As String = Projekt.Wohnung.Geschoss.Nr.ToString
        Dim gBez As String = Projekt.Wohnung.Geschoss.Bezeichnung

        Dim tmpG As String = ""

        If gTyp = "EG" Then
            tmpG = tmpG & "EG"
        ElseIf gTyp = "DG" Then
            tmpG = tmpG & "DG"
        ElseIf gTyp <> "" Then
            tmpG = tmpG & gNr & ". " & gTyp
        End If

        If gBez <> "" Then
            If tmpG <> "" Then tmpG = tmpG & ", "
            tmpG = tmpG & gBez
        End If
        Return tmpG
    End Function
    Public Function Get_Str_GeschossTyp(ByVal btGeschoss As Byte) As String
        Select Case btGeschoss
            Case UG
                Return "UG"
            Case EG
                Return "EG"
            Case OG
                Return "OG"
            Case DG
                Return "DG"
            Case Else
                Return ""
        End Select
    End Function
    Public Function Get_Klasse_I_Beschreibung() As String
        Dim Kl_I As String = Get_Klasse_I()

        If Kl_I = "" Then
            Return ""
        ElseIf Kl_I = "A*" Then
            Return "Sehr leises Wohngebiet"
        ElseIf Kl_I = "A" Then
            Return "Ruhiges Wohngebiet"
        ElseIf Kl_I = "B" Then
            Return "Wohngebiet ohne besondere Anforderungen an den Schallschutz der Außenbauteile"
        ElseIf Kl_I = "C" Then
            Return "Misch- bzw. Kerngebiet mit mäßiger Außenlärmbelastung und Anforderungen an den Schallschutz der Außenbauteile"
        ElseIf Kl_I = "D" Then
            Return "Misch- bzw. Kerngebiet mit hohen Anforderungen an den Schallschutz der Außenbauteile"
        ElseIf Kl_I = "E" Then
            Return "Gewerbegebiet oder hohe Außenlärmbelastung und sehr hohe Anforderungen an den Schallschutz der Außenbauteile"
        ElseIf Kl_I = "F" Then
            Return "Industriegebiet oder sehr hohe Außenlärmbelastung und sehr hohen Anforderungen an den Schallschutz der Außenbauteile"
        Else
            Return ""
        End If
    End Function
    Public Function Get_Klasse_II_Beschreibung() As String
        
        Dim Kl_II As String = Get_Klasse_II()

        If Kl_II = "" Then
            Return ""
        ElseIf Kl_II = "A*" Then
            Return "- Hoher Schallschutz in Doppel- und Reihenhäusern -" & Chr(13) & Chr(10) & "Wohneinheit mit sehr gutem Schallschutz, die ein ungestörtes Wohnen nahezu ohne Rücksichtnahme gegenüber den Nachbarn ermöglicht."
        ElseIf Kl_II = "A" Then
            Return "- Erhöhter Schallschutz in Doppel- und Reihenhäusern -" & Chr(13) & Chr(10) & "Wohneinheit mit sehr gutem Schallschutz, die ein ungestörtes Wohnen ohne große Rücksichtnahme gegenüber den Nachbarn ermöglicht."
        ElseIf Kl_II = "B" Then
            Return "- Hoher Schallschutz in Mehrfamilienhäusern. Normaler Schallschutz in 'Doppel- und Reihenhäusern -" & Chr(13) & Chr(10) & "Wohneinheit mit gutem Schallschutz, die bei gegenseitiger Rücksichtnahme zwischen den Nachbarn ein ruhiges Wohnen bei weitgehendem Schutz der Privatsphäre ermöglicht."
        ElseIf Kl_II = "C" Then
            Return "- Erhöhter Schallschutz in Mehrfamilienhäusern -" & Chr(13) & Chr(10) & "Wohneinheit mit gutem Schallschutz, in der die Bewohner bei üblichem rücksichtsvollen Wohnverhalten im allgemeinen Ruhe finden und die Vertraulichkeit gewahrt bleibt."
        ElseIf Kl_II = "D" Then
            Return "- Normaler Schallschutz in Mehrfamilienhäusern -" & Chr(13) & Chr(10) & "Wohneinheit mit einem Schallschutz, der die Anforderungen der DIN 4109-1:2016-07 für Geschosshäuser mit Wohnungen und Arbeitsräumen im Wesentlichen erfüllt (Ausnahmen: siehe II.3) und damit die Bewohner in Aufenthaltsräumen im Sinne des Gesundheitsschutzes vor unzumutbaren Belästigungen durch Schallübertragung aus fremden Wohneinheiten und von außen schützt. Es kann nicht erwartet werden, dass Geräusche aus fremden Wohneinheiten oder von außen nicht mehr wahrgenommen werden. Dies erfordert gegenseitige Rücksichtnahme durch Vermeidung unnötigen Lärms. Die Anforderungen setzen voraus, dass in benachbarten Räumen keine ungewöhnlich starken Geräusche verursacht werden."
        ElseIf Kl_II = "E" Then
            Return "Wohneinheit mit einem Schallschutz, der die Anforderungen der DIN 4109-1:2016-07 nicht erfüllt. Belästigungen durch Schallübertragung aus fremden Wohneinheiten und von außen sind möglich; besondere Rücksichtnahme ist unbedingt erforderlich. Die Vertraulichkeit ist nicht mehr gegeben."
        ElseIf Kl_II = "F" Then
            Return "Wohneinheit mit einem schlechten Schallschutz, der deutlich unter den Anforderungen der DIN 4109-1:2016-07 liegt. Mit Belästigungen durch Schallübertragung aus fremden Wohneinheiten und von außen muss auch bei bewusster Rücksichtnahme gerechnet weren; Vertraulichkeit kann nicht erwartet werden."
        Else
            Return ""
        End If
    End Function
    Public Function Get_Mes_ja() As String
        '=WENN(X32="";"";WENN(ODER('Schallschutzausweis ausführlich'!G26="X";'Schallschutzausweis ausführlich'!G27="X";
        'Schallschutzausweis ausführlich'!G29="X";'Schallschutzausweis ausführlich'!G30="X";'Schallschutzausweis ausführlich'!G31="X";
        'Schallschutzausweis ausführlich'!G33="X";'Schallschutzausweis ausführlich'!G34="X";'Schallschutzausweis ausführlich'!G39="X";
        'Schallschutzausweis ausführlich'!H41="X";'Schallschutzausweis ausführlich'!H42="X");"X";""))
        'Or .Koerper.Untersuchung = Messung 
        With Projekt
            If Get_Klasse_II() = "" Then
                Return ""
            ElseIf .LS_Wand.Untersuchung = Messung Or .LS_Decke.Untersuchung = Messung Or .TS_Decke.Untersuchung = Messung Or _
                    .TS_TPLH.Untersuchung = Messung Or .TS_BLT.Untersuchung = Messung Or (.Tueren.Ort <> nv And .Tueren.Untersuchung = Messung) Or _
                    .Wasser.Untersuchung = Messung Or .NutzerKoerper.MessungPrognose = Messung Then
                'Dim tMes As Boolean = False
                'If .LS_Wand.Untersuchung = Messung Then
                '    tMes = False
                'ElseIf .LS_Decke.Untersuchung = Messung Then
                '    tMes = False
                'ElseIf .TS_Decke.Untersuchung = Messung Then
                '    tMes = False
                'ElseIf .TS_TPLH.Untersuchung = Messung Then
                '    tMes = False
                'ElseIf .TS_BLT.Untersuchung = Messung Then
                '    tMes = False
                'ElseIf .Tueren.Untersuchung = Messung Then
                '    tMes = False
                'ElseIf .Wasser.Untersuchung = Messung Then
                '    tMes = False
                'ElseIf .NutzerKoerper.MessungPrognose = Messung Then
                '    tMes = False
                'End If
                Return "X"
            Else
                Return ""
            End If
        End With
    End Function
    Public Function Get_Mes_nein() As String
        '=WENN(X32="";"";WENN(ODER('Schallschutzausweis ausführlich'!G27="X";'Schallschutzausweis ausführlich'!G28="X";'Schallschutzausweis ausführlich'!G30="X";'Schallschutzausweis ausführlich'!G31="X";'Schallschutzausweis ausführlich'!G32="X";'Schallschutzausweis ausführlich'!G34="X";'Schallschutzausweis ausführlich'!G35="X";'Schallschutzausweis ausführlich'!G40="X";'Schallschutzausweis ausführlich'!H42="X";'Schallschutzausweis ausführlich'!H43="X");"";"X"))
        'Or .Koerper.Untersuchung = Messung 
        With Projekt
            If Get_Klasse_II() = "" Then
                Return ""
            ElseIf .LS_Wand.Untersuchung = Messung Or .LS_Decke.Untersuchung = Messung Or .TS_Decke.Untersuchung = Messung Or _
                    .TS_TPLH.Untersuchung = Messung Or .TS_BLT.Untersuchung = Messung Or (.Tueren.Ort <> nv And .Tueren.Untersuchung = Messung) Or _
                    .Wasser.Untersuchung = Messung Or .NutzerKoerper.MessungPrognose = Messung Then
                Return ""
            Else
                Return "X"
            End If
        End With
    End Function
    Public Function Get_gesKl_ja() As String
        '=WENN(R24="F";"X";
        'WENN(R24="E";WENN(UND(P27<>"F";P28<>"F";P30<>"F";P31<>"F";P32<>"F";P34<>"F";P35<>"F";P38<>"F";P40<>"F";P46<>"F";P47<>"F";P49<>"F";P51<>"F");"X";"");
        'WENN(R24="D";WENN(UND(P27<>"E";P28<>"E";P30<>"E";P31<>"E";P32<>"E";P34<>"E";P35<>"E";P38<>"E";P40<>"E";P46<>"E";P47<>"E";P49<>"E";P51<>"E");"X";"");
        'WENN(R24="C";WENN(UND(P27<>"D";P28<>"D";P30<>"D";P31<>"D";P32<>"D";P34<>"D";P35<>"D";P38<>"D";P40<>"D";P46<>"D";P47<>"D";P49<>"D";P51<>"D");"X";"");
        'WENN(R24="B";WENN(UND(P27<>"C";P28<>"C";P30<>"C";P31<>"C";P32<>"C";P34<>"C";P35<>"C";P38<>"C";P40<>"C";P46<>"C";P47<>"C";P49<>"C";P51<>"C");"X";"");
        'WENN(R24="A";WENN(UND(P27<>"B";P28<>"B";P30<>"B";P31<>"B";P32<>"B";P34<>"B";P35<>"B";P38<>"B";P40<>"B";P46<>"B";P47<>"B";P49<>"B";P51<>"B");"X";"");
        'WENN(R24="";"";"X")))))))

        Dim Kl_II As String = Get_Klasse_II()

        Dim Kl_LS_W As String = Get_Klasse_LS_Waende()
        Dim KL_LS_D As String = Get_Klasse_LS_Decken()
        Dim KL_TS_D As String = Get_Klasse_TS_Decken()
        Dim KL_TS_TPH As String = Get_Klasse_TS_TPLH()
        Dim KL_TS_BLLt As String = Get_Klasse_TS_BLT()
        Dim KL_Tueren As String = Get_Klasse_Tueren()
        Dim KL_Aussenbauteile As String = Get_Klasse_Aussenbauteile()
        Dim KL_Wasser As String = Get_Klasse_Wasser()
        Dim KL_lauteRaeume As String = Get_Klasse_lauteRaeume()

        With Projekt
            If Kl_II = "F" Then
                Return "X"
            ElseIf Kl_II = "E" Then
                If Kl_LS_W <> "F" And KL_LS_D <> "F" And KL_TS_D <> "F" And KL_TS_TPH <> "F" And KL_TS_BLLt <> "F" And KL_Tueren <> "F" _
                    And KL_Aussenbauteile <> "F" And KL_Wasser <> "F" And KL_lauteRaeume <> "F" Then
                    Return "X"
                Else
                    Return ""
                End If
            ElseIf Kl_II = "D" Then
                If Kl_LS_W <> "E" And KL_LS_D <> "E" And KL_TS_D <> "E" And KL_TS_TPH <> "E" And KL_TS_BLLt <> "E" And KL_Tueren <> "E" _
                        And KL_Aussenbauteile <> "E" And KL_Wasser <> "E" And KL_lauteRaeume <> "E" Then
                    Return "X"
                Else
                    Return ""
                End If
            ElseIf Kl_II = "C" Then
                If Kl_LS_W <> "D" And KL_LS_D <> "D" And KL_TS_D <> "D" And KL_TS_TPH <> "D" And KL_TS_BLLt <> "D" And KL_Tueren <> "D" _
                        And KL_Aussenbauteile <> "D" And KL_Wasser <> "D" And KL_lauteRaeume <> "D" Then
                    Return "X"
                Else
                    Return ""
                End If
            ElseIf Kl_II = "B" Then
                If Kl_LS_W <> "C" And KL_LS_D <> "C" And KL_TS_D <> "C" And KL_TS_TPH <> "C" And KL_TS_BLLt <> "C" And KL_Tueren <> "C" _
                        And KL_Aussenbauteile <> "C" And KL_Wasser <> "C" And KL_lauteRaeume <> "C" Then
                    Return "X"
                Else
                    Return ""
                End If
            ElseIf Kl_II = "A" Then
                If Kl_LS_W <> "B" And KL_LS_D <> "B" And KL_TS_D <> "B" And KL_TS_TPH <> "B" And KL_TS_BLLt <> "B" And KL_Tueren <> "B" _
                        And KL_Aussenbauteile <> "B" And KL_Wasser <> "B" And KL_lauteRaeume <> "B" Then
                    Return "X"
                Else
                    Return ""
                End If
            ElseIf Kl_II = "" Then
                Return ""
            Else
                Return "X"
            End If
        End With
    End Function
    Public Function Get_gesKl_nein() As String
        If Get_gesKl_ja() = "X" Then
            Return ""
        Else
            Return "X"
        End If
    End Function
End Module
