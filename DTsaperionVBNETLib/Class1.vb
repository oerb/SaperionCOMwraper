Public Class Archivieren

    Public Function saveDokument(filelocation As String, Mandant As Integer,
                                 Unternehmen As Integer, WE As Integer,
                                 HausNr As Integer, Wohnung As Integer,
                                 FolgeNr As Integer,
                                 AdresseNr As Integer, DokuArt As String,
                                 VorgangKZ As String, Vorname As String,
                                 Name As String, Sachbearbeiter As String,
                                 Subject As String, MieterNr As Integer,
                                 Memo2 As String)
        Dim SapApp As Object
        Dim oDocument As Object
        Dim iRet As Object

        SapApp = CreateObject("Saperion.Application")

        'wenn saperion läuft, brauch nicht zu erfolgen
        'SapApp.login "win", "berlin"

        oDocument = CreateObject("Saperion.Document")
        oDocument.InsertFile(filelocation)  'hier die Datei, die archiviert werden soll
        'oDocument.InsertDocument(filelocation) 'Failed
        oDocument.DBName = "wowi"


        ' Mandant, Unternehmen, We, HausNr, Wohnung, WohnungZus, FolgeNr, AdressNr,
        ' OrdnerBezeichnung, DokuArt, VorgangKZ, BelegNr, BelegDatum, MieterNr, KontoNr,
        ' AuftragsNr, Gewerk, SeitenNr, Betrag, Vorname, Name, Autor, ArchivDatum,
        ' Sachbearbeiter, Memo1, Memo2, Memo3, UrkundenNr, Aktenzeichen, Bankverbindung,
        ' RecordMngtID , Grundbuch, GrundbuchBereich, GrundbuchBlatt


        '  bei Seriendruck Feldinhalt auslesen
        ' msgbox ActiveDocument:MailMerge:DataSource:DataFields:Item("MandantNr"):Value
        ' meineVar = ActiveDocument:MailMerge:DataSource:DataFields:Item("MandantNr"):Value
        ' oDocument.SetProperty "Mandant", ActiveDocument:MailMerge:DataSource:DataFields:Item("MandantNr"):Value

        'oDocument.SetProperty "Mandant", meineVariable
        oDocument.SetProperty("Mandant", 1)
        oDocument.SetProperty("Sachbearbeiter", Sachbearbeiter)
        Dim todaysdate As String = String.Format("{0:dd/MM/yyyy}", DateTime.Now)
        oDocument.SetProperty("ArchivDatum", todaysdate)
        oDocument.SetProperty("BelegDatum", todaysdate)
        If Subject IsNot "" Then
            oDocument.SetProperty("Memo1", Subject)
        End If
        'If Memo2 IsNot "" Then
        '    oDocument.SetProperty("Memo2", Memo2)
        'End If
        If Unternehmen > 0 Then
            oDocument.SetProperty("Unternehmen", Unternehmen)
        End If
        If WE > 0 Then
            oDocument.SetProperty("We", WE)
        End If
        If HausNr > 0 Then
            oDocument.SetProperty("HausNr", HausNr)
        End If
        If Wohnung > 0 Then
            oDocument.SetProperty("Wohnung", Wohnung)
        End If
        If AdresseNr > 0 Then
            oDocument.SetProperty("AdressNr", AdresseNr)
        End If
        If MieterNr > 0 Then
            oDocument.SetProperty("MieterNr", MieterNr)
        End If
        If FolgeNr > 0 Then
            oDocument.SetProperty("FolgeNr", FolgeNr)
        End If
        If DokuArt IsNot "" Then
            oDocument.SetProperty("DokuArt", DokuArt)
        End If
        If VorgangKZ IsNot "" Then
            oDocument.SetProperty("VorgangKZ", VorgangKZ)
        End If
        If Vorname IsNot "" Then
            oDocument.SetProperty("Vorname", Vorname)
        End If
        If Name IsNot "" Then
            oDocument.SetProperty("Name", Name)
        End If

        ' immer MIETER bei Mieterschreiben, wegen Plattenansteuerung in der Jukebox
        'oDocument.SetProperty("RecordMngtID", "MIETER")

        iRet = oDocument.Store()
        If iRet = True Then
            MsgBox("E-Mail erfolgreich archiviert", MsgBoxStyle.Information, "Saperion Archiv")
        Else
            MsgBox("E-Mail konnte nicht archiviert werden", MsgBoxStyle.Critical, "Warning")
        End If

        Return (iRet)

        SapApp = Nothing
        oDocument = Nothing

    End Function

    Public Function test()
        MsgBox("TEST Box in DTsaperionVBNETLib works")
        Return True
    End Function
End Class
