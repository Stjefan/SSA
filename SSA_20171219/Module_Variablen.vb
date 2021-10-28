Module Module_Variablen

    'Konstanten allgemeine Projektdaten
    '=================================

    Public Const LizenznehmerPR As String = "L#0 - Kurz und Fischer GmbH" '"L## - TESTVERSION" '"L#0 - Kurz und Fischer GmbH"
    Public LizenznehmerSW As String

    Public Const versNormal As Byte = 1
    Public Const versDEMOspezial As Byte = 2
    Public Const versDEMO As Byte = 3
    Public Const btDemo As Byte = versNormal

    Public DB_Antragsteller As ArrayList
    <VBFixedString(55)> Public LEER55 As String = ""
    <VBFixedString(100)> Public LEER100 As String = ""
    <VBFixedString(512)> Public LEER512 As String = ""
    <VBFixedString(2)> Public LEER2 As String = ""
    <VBFixedString(5)> Public LEER5 As String = ""
    <VBFixedString(10)> Public LEER10 As String = ""

    'Colors
    Public colBlue As Color = Color.FromArgb(255, 192, 255, 255)
    Public colRed As Color = Color.FromArgb(255, 255, 192, 128)
    'Klassen
    Public Const AA As Byte = 7
    Public Const A As Byte = 6
    Public Const B As Byte = 5
    Public Const C As Byte = 4
    Public Const D As Byte = 3
    Public Const E As Byte = 2
    Public Const F As Byte = 1

    'Gebietscharakter
    Public Const GC_WR As Byte = 1
    Public Const GC_WA As Byte = 2
    Public Const GC_MIWB As Byte = 3
    Public Const GC_GE As Byte = 4
    Public Const GC_GI As Byte = 5
    'Außenlärmpegel
    Public Const AP_bis55 As Byte = 1
    Public Const AP_56_60 As Byte = 2
    Public Const AP_61_65 As Byte = 3
    Public Const AP_66_70 As Byte = 4
    Public Const AP_71_75 As Byte = 5
    Public Const AP_76 As Byte = 6
    'Freibereich abgewand ja - nein
    Public Const aF_ja As Byte = 1
    Public Const aF_nein As Byte = 2
    'Geschosstyp
    Public Const UG As Byte = 1
    Public Const EG As Byte = 2
    Public Const OG As Byte = 3
    Public Const DG As Byte = 4
    Public Const Sonstiges As Byte = 5
    'Untersuchungen
    Public Const Prognose As Byte = 1
    Public Const Messung As Byte = 2
    Public Const nv As Byte = 3
    'fEstrich
    Public Const fE_kleiner50 As Byte = 1
    Public Const fE_groesser50 As Byte = 2
    'Bodenbelag
    Public Const Be_mit As Byte = 1
    Public Const Be_ohne As Byte = 2
    ''Messverfahren
    'Public Const KMV As Byte = 1
    'Public Const NMV As Byte = 2
    'Bauteile
    Public Const Treppe As Byte = 1
    Public Const Podest As Byte = 2
    Public Const Hausflur As Byte = 3
    Public Const Balkon As Byte = 4
    Public Const Laube As Byte = 5
    Public Const Loggie As Byte = 6
    Public Const Terrasse As Byte = 7
    'Türen
    Public Const Diele As Byte = 1
    Public Const Aufenthalt As Byte = 2
    Public Const Pruefzeugnis As Byte = 1
    'Aussenbauteile
    Public Const ohneNachweis As Byte = 1
    Public Const FensterMitDichtung As Byte = 2
    Public Const DINerfuellt As Byte = 3
    Public Const DINPlusErfuellt As Byte = 4
    'Wasserinstallation und Haustechn. Anlagen
    Public Const HT_kleiner20 As Byte = 1
    Public Const HT_20_24 As Byte = 2
    Public Const HT_24_27 As Byte = 3
    Public Const HT_27_30 As Byte = 4
    Public Const HT_30_35 As Byte = 5
    Public Const HT_groesser35 As Byte = 6

    Public Const erfuellt As Byte = 1
    Public Const keineAngabe As Byte = 2

    'Nutzergeräusche und Körperschallentkopplung
    Public Const Nutzer As Byte = 1
    Public Const Koerper As Byte = 2

    Public Const kleiner20 As Byte = 1
    Public Const L_20_25 As Byte = 2
    Public Const L_25_30 As Byte = 3
    Public Const L_30_35 As Byte = 4
    Public Const L_35_40 As Byte = 5
    Public Const L_40_45 As Byte = 6
    Public Const groesser45 As Byte = 7

    Public Const kleiner38 As Byte = 1
    Public Const L_38_43 As Byte = 2
    Public Const L_43_48 As Byte = 3
    Public Const L_48_53 As Byte = 4
    Public Const L_53_58 As Byte = 5
    Public Const L_58_63 As Byte = 6
    Public Const groesser63 As Byte = 7
    'Anordnung lauter Räume 
    Public Const guenstig As Byte = 1
    Public Const unguenstig As Byte = 2
    'laute Räume gem. DIN angrenzend
    Public Const keineAngrenzend As Byte = 1
    Public Const L_25_35 As Byte = 2
    Public Const L_30_40 As Byte = 3
    Public Const L_35_45 As Byte = 4
    'Treppenhaus Nachhallzeit
    Public Const NHZ_020 As Byte = 1
    Public Const NHZ_010 As Byte = 2
    Public Const NHZ_keine As Byte = 3
    'eigener Wohnbereich
    Public Const keineEmpfehlung As Byte = 1
    Public Const EW1 As Byte = 2
    Public Const EW2 As Byte = 3
    Public Const EW3 As Byte = 4

    'Variablen allgemeine Projektdaten
    '=================================

    Public stProjekt_Name As String                 'Projektname
    Public stProjekt_Pfad As String                 'Verzeichnispfad für das Projekt
    Public stProjekt_Grundriss As String            'Dateiname Grundriss ("Grundriss.*")
    Public stLageplan_Name As String                'Name und Verzeichnis Lageplan

    'Public Structure AusstellerAltData
    '    Public Adresse As AdresseAltData
    'End Structure
    Public Structure AusstellerData
        Public Adresse As AdresseData
    End Structure
    ' Public Aussteller As AusstellerData

    'Public DB_Antragsteller(30) As AdresseData
#Region "ProjektDATA"
    Public Structure ProjektData

        <VBFixedString(100)> Public Name As String          'Projektname
        <VBFixedString(512)> Public Pfad As String          'Verzeichnispfad für das Projekt
        <VBFixedString(512)> Public Grundriss As String      'Dateiname Lageplan ("Lageplan.*")

        'Status
        Public bStatus_Projekt As Boolean                   'Status Projekt (ist gerade ein Projekt geöffnet, bzw. in Bearbeitung)
        Public bStatus_Grundriss As Boolean                 'Status Grundriss (ist schon ein Lageplan eingelesen oder nicht)

        'Rubrik I
        Public Antragsteller As AdresseData
        Public Gebaeude As GebaeudeData
        Public Wohnung As WohnungData
        <VBFixedString(10)> Public Datum As String

        'Rubrik II
        Public Standort As StandortData
        'Rubrik III
        Public LS_Wand As Luftschall_Data
        Public LS_Decke As Luftschall_Data
        Public TS_Decke As Trittschall_Decke_Data
        Public TS_TPLH As Trittschall_TPLH_Data
        Public TS_BLT As Trittschall_BLT_Data

        Public Tueren As Tueren_Data
        Public Aussenbauteile As Byte   'ohne Nachweis, Fenster mit Dichtung, DIN erfüllt, DIN + 5 dB erfüllt
        <VBFixedString(55)> Public Bem_Aussenbauteile As String
        Public Wasser As Wasser_Data
        Public NutzerKoerper As Nutzer_Koerper_Data
        'Public Koerper As Nutzer_Koerper_Data
        Public Nachbarn As Byte
        <VBFixedString(55)> Public Bemerkung_Nachbarn As String
        Public anordnungRaeume As Byte
        <VBFixedString(55)> Public Bem_anordnungRaeume As String
        Public lauteRaeume As Byte
        <VBFixedString(55)> Public Bem_lauteRaeume As String
        Public NHZ As Byte
        <VBFixedString(55)> Public Bemerkung_AV As String
        Public eigenerWohnbereich As Byte
        <VBFixedString(55)> Public Bemerkung_eigenerWohnbereich As String
    End Structure

    'Public Structure ProjektAltData
    '    <VBFixedString(100)> Public Name As String          'Projektname
    '    <VBFixedString(512)> Public Pfad As String          'Verzeichnispfad für das Projekt
    '    <VBFixedString(512)> Public Grundriss As String      'Dateiname Lageplan ("Lageplan.*")

    '    'Status
    '    Public bStatus_Projekt As Boolean                   'Status Projekt (ist gerade ein Projekt geöffnet, bzw. in Bearbeitung)
    '    Public bStatus_Grundriss As Boolean                 'Status Grundriss (ist schon ein Lageplan eingelesen oder nicht)

    '    'Rubrik I
    '    Public Antragsteller As AdresseAltData
    '    Public Gebaeude As GebaeudeAltData
    '    Public Wohnung As WohnungData
    '    <VBFixedString(10)> Public Datum As String

    '    'Rubrik II
    '    Public Standort As StandortData_bisV20140818
    '    'Rubrik III
    '    Public LS_Wand As Luftschall_Data_bisV20140818
    '    Public LS_Decke As Luftschall_Data_bisV20140818
    '    Public TS_Decke As Trittschall_Decke_Data_bisV20140818
    '    Public TS_TPH As Trittschall_TPH_Data_bisV20140818
    '    Public TS_BLLT As Trittschall_BLLT_Data_bisV20140818

    '    Public Tueren As Tueren_Data_bisV20140818
    '    Public Aussenbauteile As Byte   'ohne Nachweis, Fenster mit Dichtung, DIN erfüllt, DIN + 5 dB erfüllt
    '    Public Wasser As Wasser_Data_bisV20140818
    '    Public Nutzer As Nutzer_Koerper_Data_bisV20140818
    '    Public Koerper As Nutzer_Koerper_Data_bisV20140818
    '    Public Nachbarn As Byte
    '    Public anordnungRaeume As Byte
    '    Public lauteRaeume As Byte
    '    Public eigenerWohnbereich As Byte
    'End Structure
    'Public Structure ProjektData_bis20171110

    '    <VBFixedString(100)> Public Name As String          'Projektname
    '    <VBFixedString(512)> Public Pfad As String          'Verzeichnispfad für das Projekt
    '    <VBFixedString(512)> Public Grundriss As String      'Dateiname Lageplan ("Lageplan.*")

    '    'Status
    '    Public bStatus_Projekt As Boolean                   'Status Projekt (ist gerade ein Projekt geöffnet, bzw. in Bearbeitung)
    '    Public bStatus_Grundriss As Boolean                 'Status Grundriss (ist schon ein Lageplan eingelesen oder nicht)

    '    'Rubrik I
    '    Public Antragsteller As AdresseData
    '    Public Gebaeude As GebaeudeData
    '    Public Wohnung As WohnungData
    '    <VBFixedString(10)> Public Datum As String

    '    'Rubrik II
    '    Public Standort As StandortData
    '    'Rubrik III
    '    Public LS_Wand As Luftschall_Data
    '    Public LS_Decke As Luftschall_Data
    '    Public TS_Decke As Trittschall_Decke_Data
    '    Public TS_TPH As Trittschall_TPH_Data_bis20171110
    '    Public TS_BLLT As Trittschall_BLLT_Data_bis20171110

    '    Public Tueren As Tueren_Data
    '    Public Aussenbauteile As Byte   'ohne Nachweis, Fenster mit Dichtung, DIN erfüllt, DIN + 5 dB erfüllt
    '    <VBFixedString(55)> Public Bem_Aussenbauteile As String
    '    Public Wasser As Wasser_Data
    '    Public Nutzer As Nutzer_Koerper_Data_bis20171110
    '    Public Koerper As Nutzer_Koerper_Data_bis20171110
    '    Public Nachbarn As Byte
    '    <VBFixedString(55)> Public Bemerkung_Nachbarn As String
    '    Public anordnungRaeume As Byte
    '    <VBFixedString(55)> Public Bem_anordnungRaeume As String
    '    Public lauteRaeume As Byte
    '    <VBFixedString(55)> Public Bem_lauteRaeume As String
    '    Public NHZ As Byte
    '    Public eigenerWohnbereich As Byte
    '    <VBFixedString(55)> Public Bemerkung_eigenerWohnbereich As String
    'End Structure
    'Public Structure ProjektData_bis20170803


    '    <VBFixedString(100)> Public Name As String          'Projektname
    '    <VBFixedString(512)> Public Pfad As String          'Verzeichnispfad für das Projekt
    '    <VBFixedString(512)> Public Grundriss As String      'Dateiname Lageplan ("Lageplan.*")

    '    'Status
    '    Public bStatus_Projekt As Boolean                   'Status Projekt (ist gerade ein Projekt geöffnet, bzw. in Bearbeitung)
    '    Public bStatus_Grundriss As Boolean                 'Status Grundriss (ist schon ein Lageplan eingelesen oder nicht)

    '    'Rubrik I
    '    Public Antragsteller As AdresseData
    '    Public Gebaeude As GebaeudeData
    '    Public Wohnung As WohnungData
    '    <VBFixedString(10)> Public Datum As String

    '    'Rubrik II
    '    Public Standort As StandortData
    '    'Rubrik III
    '    Public LS_Wand As Luftschall_Data
    '    Public LS_Decke As Luftschall_Data
    '    Public TS_Decke As Trittschall_Decke_Data
    '    Public TS_TPH As Trittschall_TPH_Data_bis20170803
    '    Public TS_BLLT As Trittschall_BLLT_Data_bis20170803

    '    Public Tueren As Tueren_Data
    '    Public Aussenbauteile As Byte   'ohne Nachweis, Fenster mit Dichtung, DIN erfüllt, DIN + 5 dB erfüllt
    '    <VBFixedString(55)> Public Bem_Aussenbauteile As String
    '    Public Wasser As Wasser_Data
    '    Public Nutzer As Nutzer_Koerper_Data
    '    Public Koerper As Nutzer_Koerper_Data
    '    Public Nachbarn As Byte
    '    <VBFixedString(55)> Public Bemerkung_Nachbarn As String
    '    Public anordnungRaeume As Byte
    '    <VBFixedString(55)> Public Bem_anordnungRaeume As String
    '    Public lauteRaeume As Byte
    '    <VBFixedString(55)> Public Bem_lauteRaeume As String
    '    Public eigenerWohnbereich As Byte
    '    <VBFixedString(55)> Public Bemerkung_eigenerWohnbereich As String
    'End Structure
    'Public Structure ProjektData_bisV20140818

    '    <VBFixedString(100)> Public Name As String          'Projektname
    '    <VBFixedString(512)> Public Pfad As String          'Verzeichnispfad für das Projekt
    '    <VBFixedString(512)> Public Grundriss As String      'Dateiname Lageplan ("Lageplan.*")

    '    'Status
    '    Public bStatus_Projekt As Boolean                   'Status Projekt (ist gerade ein Projekt geöffnet, bzw. in Bearbeitung)
    '    Public bStatus_Grundriss As Boolean                 'Status Grundriss (ist schon ein Lageplan eingelesen oder nicht)

    '    'Rubrik I
    '    Public Antragsteller As AdresseData
    '    Public Gebaeude As GebaeudeData
    '    Public Wohnung As WohnungData
    '    <VBFixedString(10)> Public Datum As String

    '    'Rubrik II
    '    Public Standort As StandortData_bisV20140818
    '    'Rubrik III
    '    Public LS_Wand As Luftschall_Data_bisV20140818
    '    Public LS_Decke As Luftschall_Data_bisV20140818
    '    Public TS_Decke As Trittschall_Decke_Data_bisV20140818
    '    Public TS_TPH As Trittschall_TPH_Data_bisV20140818
    '    Public TS_BLLT As Trittschall_BLLT_Data_bisV20140818

    '    Public Tueren As Tueren_Data_bisV20140818
    '    Public Aussenbauteile As Byte   'ohne Nachweis, Fenster mit Dichtung, DIN erfüllt, DIN + 5 dB erfüllt
    '    Public Wasser As Wasser_Data_bisV20140818
    '    Public Nutzer As Nutzer_Koerper_Data_bisV20140818
    '    Public Koerper As Nutzer_Koerper_Data_bisV20140818
    '    Public Nachbarn As Byte
    '    Public anordnungRaeume As Byte
    '    Public lauteRaeume As Byte
    '    Public eigenerWohnbereich As Byte
    'End Structure

    Public Projekt As ProjektData
#End Region

#Region "Nutzer_KoerperDATA"
    'Public Structure Nutzer_Koerper_Data_bisV20140818
    '    Public Untersuchung As Byte
    '    Public L_Intervall As Byte
    'End Structure
    'Public Structure Nutzer_Koerper_Data_bis20171110
    '    Public Untersuchung As Byte
    '    Public L_Intervall As Byte
    '    <VBFixedString(55)> Public Bemerkung As String
    'End Structure
    Public Structure Nutzer_Koerper_Data
        Public Untersuchung As Byte
        Public MessungPrognose As Byte
        Public L_Intervall As Byte
        <VBFixedString(55)> Public BemerkungNutzer As String
        <VBFixedString(55)> Public BemerkungKoerper As String
    End Structure

    'Public Structure Wasser_Data_bisV20140818
    '    Public Untersuchung As Byte
    '    Public L_Intervall As Byte
    '    Public LcLa_erfuellt As Byte
    'End Structure
    Public Structure Wasser_Data
        Public Untersuchung As Byte
        Public L_Intervall As Byte
        Public LcLa_erfuellt As Byte
        <VBFixedString(55)> Public Bem_Wasser As String
    End Structure
#End Region

#Region "TürenDATA"
    Public Structure Tueren_Data
        Public Ort As Byte
        Public Untersuchung As Byte
        Public L As Single
        <VBFixedString(55)> Public Bemerkung_Tueren As String
    End Structure
    'Public Structure Tueren_Data_bisV20140818
    '    Public Ort As Byte
    '    Public Untersuchung As Byte
    '    Public L As Single
    'End Structure
#End Region

#Region "TS_Decke_DATA"
    'Public Structure Trittschall_Decke_Data_bisV20140818
    '    Public Untersuchung As Byte
    '    Public Prognose As PegelData
    '    Public Messung As LS_TSD_MessungData_bis20170803
    'End Structure
    'Public Structure Trittschall_Decke_Data_bis20170803
    '    Public Untersuchung As Byte
    '    Public Prognose As PegelData
    '    Public Messung As LS_TSD_MessungData_bis20170803
    '    <VBFixedString(55)> Public Bemerkung_TS_Decke As String
    'End Structure
    'Public Structure Trittschall_Decke_Data_bis20171110
    '    Public Untersuchung As Byte
    '    Public fEstrich As Byte
    '    Public Prognose As TS_PrognoseDATA
    '    Public Messung As TS_Messung_TPLH_BLT_Data
    '    <VBFixedString(55)> Public Bemerkung_TS_Decke As String
    'End Structure
    Public Structure Trittschall_Decke_Data
        Public Untersuchung As Byte
        Public fEstrich As Byte
        Public Prognose As TS_PrognoseDATA
        Public Messung As TS_Messung_TPLH_BLT_Data
        <VBFixedString(55)> Public Bemerkung_TS_Decke As String
    End Structure
#End Region

#Region "TPH BLLT DATA"
#Region "Prognose"
    Public Structure TS_PrognoseDATA
        Public Bodenbelag As Byte
        Public Pegel As PegelData
    End Structure
    'Public Structure Prognose_TPH_Data_bis20171110
    '    Public Treppe As TS_PrognoseDATA
    '    Public Podest As TS_PrognoseDATA
    '    Public Hausflur As TS_PrognoseDATA
    'End Structure
    Public Structure Prognose_TPLH_Data
        Public Treppe As TS_PrognoseDATA
        Public Podest As TS_PrognoseDATA
        Public Laube As TS_PrognoseDATA
        Public Hausflur As TS_PrognoseDATA
    End Structure

    'Public Structure Prognose_BLLT_Data_bis20171110
    '    Public Balkon As TS_PrognoseDATA
    '    Public Laube As TS_PrognoseDATA
    '    Public Loggia As TS_PrognoseDATA
    '    Public Terrasse As TS_PrognoseDATA
    'End Structure
    Public Structure Prognose_BLT_Data
        Public Balkon As TS_PrognoseDATA
        Public Loggia As TS_PrognoseDATA
        Public Terrasse As TS_PrognoseDATA
    End Structure

#End Region
#Region "Messung"
#Region "TS_TPH_DATA"
#Region "Ältere Versionen"
    'Public Structure Trittschall_TPH_Data_bisV20140818
    '    Public Untersuchung As Byte
    '    Public Prognose As Prognose_TPH_Data_bis20171110
    '    Public Messung As TS_Messung_TPLH_BLT_Data '_bis20171110
    'End Structure
    'Public Structure Trittschall_TPH_Data_bis20170803
    '    Public Untersuchung As Byte
    '    Public Prognose As Prognose_TPH_Data_bis20171110
    '    Public Messung As TS_Messung_TPH_BLLT_Data_bis20170803
    '    <VBFixedString(55)> Public Bemerkung_TS_TPH As String
    'End Structure
    'Public Structure Trittschall_TPH_Data_bis20171110
    '    Public Untersuchung As Byte
    '    'Public Bodenbelag As Byte
    '    Public Prognose As Prognose_TPH_Data_bis20171110
    '    Public Messung As TS_Messung_TPLH_BLT_Data
    '    <VBFixedString(55)> Public Bemerkung_TS_TPH As String
    'End Structure
#End Region
    Public Structure Trittschall_TPLH_Data
        Public Untersuchung As Byte
        'Public Bodenbelag As Byte
        Public Prognose As Prognose_TPLH_Data
        Public Messung As TS_Messung_TPLH_BLT_Data
        <VBFixedString(55)> Public Bemerkung_TS_TPH As String
    End Structure

#End Region

#Region "TS_BLLT_DATA"
#Region "Ältere Versionen"
    'Public Structure Trittschall_BLLT_Data_bisV20140818
    '    Public Untersuchung As Byte
    '    Public Prognose As Prognose_BLLT_Data_bis20171110
    '    Public Messung As TS_Messung_TPLH_BLT_Data
    'End Structure
    'Public Structure Trittschall_BLLT_Data_bis20170803
    '    Public Untersuchung As Byte
    '    Public Prognose As Prognose_BLLT_Data_bis20171110
    '    Public Messung As TS_Messung_TPH_BLLT_Data_bis20170803
    '    <VBFixedString(55)> Public Bemerkung_TS_BLLT As String
    'End Structure
    'Public Structure Trittschall_BLLT_Data_bis20171110
    '    Public Untersuchung As Byte
    '    'Public Bodenbelag As Byte
    '    Public Prognose As Prognose_BLLT_Data_bis20171110
    '    Public Messung As TS_Messung_TPLH_BLT_Data
    '    <VBFixedString(55)> Public Bemerkung_TS_BLLT As String
    'End Structure
#End Region
   
    Public Structure Trittschall_BLT_Data
        Public Untersuchung As Byte
        'Public Bodenbelag As Byte
        Public Prognose As Prognose_BLT_Data
        Public Messung As TS_Messung_TPLH_BLT_Data
        <VBFixedString(55)> Public Bemerkung_TS_BLLT As String
    End Structure
#End Region

    Public Structure TS_Messung_TPLH_BLT_Data
        Public Anzahl As Byte
        Public Messung_1 As Messung_TPLH_BLT_Data
        Public Messung_2 As Messung_TPLH_BLT_Data
        Public Messung_3 As Messung_TPLH_BLT_Data
        Public Messung_4 As Messung_TPLH_BLT_Data
        Public Messung_5 As Messung_TPLH_BLT_Data
        Public Messung_6 As Messung_TPLH_BLT_Data
    End Structure
    Public Structure Messung_TPLH_BLT_Data
        Public Bodenbelag As Byte
        Public Bauteil As Byte
        Public Pegel As PegelData
    End Structure
#Region "Alte DATA"
    'Public Structure TS_Messung_TPH_BLLT_Data_bis20170803
    '    Public Anzahl As Byte
    '    Public Messung_1 As Messung_TPH_BLLT_Data_bis20170803
    '    Public Messung_2 As Messung_TPH_BLLT_Data_bis20170803
    '    Public Messung_3 As Messung_TPH_BLLT_Data_bis20170803
    '    Public Messung_4 As Messung_TPH_BLLT_Data_bis20170803
    '    Public Messung_5 As Messung_TPH_BLLT_Data_bis20170803
    '    Public Messung_6 As Messung_TPH_BLLT_Data_bis20170803
    'End Structure
    'Public Structure Messung_TPH_BLLT_Data_bis20170803
    '    Public Bauteil As Byte
    '    Public Messverfahren As Byte
    '    Public Pegel As PegelData
    'End Structure
#End Region
#End Region
#End Region

    'Public Structure Luftschall_Data_bisV20140818
    '    Public Untersuchung As Byte
    '    Public Prognose As PegelData
    '    Public Messung As LS_TSD_MessungData_bis20170803
    'End Structure
    Public Structure Luftschall_Data
        Public Untersuchung As Byte
        Public Prognose As PegelData
        Public Messung As LS_TSD_MessungData
        <VBFixedString(55)> Public Bemerkung_LS As String
    End Structure

    Public Structure PrognoseData
        Public Bauteil As Byte
        Public Pegel As PegelData
    End Structure

    Public Structure PegelData
        Public Pegel As Single
        <VBFixedString(2)> Public C As String
    End Structure

    Public Structure LS_TSD_MessungData
        Public Anzahl As Byte
        Public Messung_1 As PegelData
        Public Messung_2 As PegelData
        Public Messung_3 As PegelData
        Public Messung_4 As PegelData
        Public Messung_5 As PegelData
        Public Messung_6 As PegelData
    End Structure

    'Public Structure LS_TSD_MessungData_bis20170803
    '    Public Messanteil As Byte
    '    Public Anzahl As Byte
    '    Public Messung_1 As MessungData
    '    Public Messung_2 As MessungData
    '    Public Messung_3 As MessungData
    '    Public Messung_4 As MessungData
    '    Public Messung_5 As MessungData
    '    Public Messung_6 As MessungData
    'End Structure

    Public Structure MessungData
        Public Messverfahren As Byte
        Public Pegel As PegelData
    End Structure

    'Public Structure GebaeudeAltData
    '    Public Adresse As AdresseAltData
    '    <VBFixedString(100)> Public Gebaeudetyp As String
    '    Public Baujahr As Integer
    '    Public Wohneinheiten As String
    '    Public Kosten As Single
    'End Structure
    Public Structure GebaeudeData
        Public Adresse As AdresseData
        <VBFixedString(100)> Public Gebaeudetyp As String
        Public Baujahr As Integer
        Public Wohneinheiten As String
        Public Kosten As Single
    End Structure

    Public Structure WohnungData
        <VBFixedString(100)> Public Wohnungsbezeichnung As String
        Public Geschoss As GeschossData
        Public Raeume As String
        Public Wohnflaeche As Single
    End Structure

    Public Structure GeschossData
        Public Typ As Byte
        Public Nr As Integer
        <VBFixedString(100)> Public Bezeichnung As String
    End Structure

    'Public Structure AdresseAltData
    '    <VBFixedString(100)> Public Name As String
    '    <VBFixedString(100)> Public Zusatz As String
    '    <VBFixedString(100)> Public Strasse As String
    '    <VBFixedString(10)> Public Nr As String
    '    Public PLZ As Integer
    '    <VBFixedString(100)> Public Ort As String
    'End Structure
    Public Structure AdresseData
        <VBFixedString(100)> Public Name As String
        <VBFixedString(100)> Public Zusatz As String
        <VBFixedString(100)> Public Strasse As String
        <VBFixedString(10)> Public Nr As String
        <VBFixedString(5)> Public PLZ As String
        <VBFixedString(100)> Public Ort As String
    End Structure

    Public Structure StandortData
        Public Gebietscharakter As Byte
        <VBFixedString(55)> Public Bem_Gebietscharakter As String
        Public Aussenlaermpegel As Byte
        <VBFixedString(55)> Public Bem_Aussenlaermpegel As String
        Public abgewandFreibereich As Byte
    End Structure

    'Public Structure StandortData_bisV20140818
    '    Public Gebietscharakter As Byte
    '    Public Aussenlaermpegel As Byte
    '    Public abgewandFreibereich As Byte
    'End Structure
End Module
