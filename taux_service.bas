Sub CalculTauxService()

    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim client As String, entrepot As String
    Dim qteCommandee As Double, qteLivree As Double
    Dim taux As Double

    ' Création ou nettoyage de la feuille Résultat
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("Résultat").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Résultat"
    
    Set ws = ThisWorkbook.Sheets("Données") ' nom de la feuille source
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim resultWs As Worksheet
    Set resultWs = Sheets("Résultat")
    
    ' En-têtes
    resultWs.Range("A1:D1").Value = Array("Client", "Entrepôt", "Taux de service (%)", "Nb de lignes")
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Traitement ligne par ligne
    For i = 2 To lastRow
        client = ws.Cells(i, 1).Value
        entrepot = ws.Cells(i, 2).Value
        qteCommandee = ws.Cells(i, 5).Value
        qteLivree = ws.Cells(i, 6).Value
        
        If qteCommandee > 0 Then
            key = client & "|" & entrepot
            If Not dict.exists(key) Then
                dict.Add key, Array(0, 0, 0) ' totalLivrée, totalCommandée, nbLignes
            End If
            dict(key)(0) = dict(key)(0) + qteLivree
            dict(key)(1) = dict(key)(1) + qteCommandee
            dict(key)(2) = dict(key)(2) + 1
        End If
    Next i

    ' Remplissage des résultats
    Dim ligne As Long: ligne = 2
    Dim k As Variant
    For Each k In dict.Keys
        Dim parts() As String
        parts = Split(k, "|")
        taux = dict(k)(0) / dict(k)(1) * 100
        
        resultWs.Cells(ligne, 1).Value = parts(0)
        resultWs.Cells(ligne, 2).Value = parts(1)
        resultWs.Cells(ligne, 3).Value = Round(taux, 2)
        resultWs.Cells(ligne, 4).Value = dict(k)(2)
        ligne = ligne + 1
    Next k

    MsgBox "Taux de service calculé avec succès dans la feuille 'Résultat'.", vbInformation

End Sub
