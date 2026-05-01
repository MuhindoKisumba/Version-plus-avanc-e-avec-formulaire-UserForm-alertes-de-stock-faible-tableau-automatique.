Option Explicit

' =========================
' INITIALISATION DU TABLEAU
' =========================
Public Sub InitialiserTable()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Stock")
    
    ws.Cells.Clear
    
    ws.Cells(1, 1).Value = "ID"
    ws.Cells(1, 2).Value = "Produit"
    ws.Cells(1, 3).Value = "Quantité"
    ws.Cells(1, 4).Value = "Prix"
    ws.Cells(1, 5).Value = "Seuil Alerte"
    ws.Cells(1, 6).Value = "Valeur Stock"
    
    MsgBox "Table initialisée", vbInformation
End Sub

' =========================
' CALCUL VALEUR STOCK
' =========================
Public Sub CalculValeurStock()
    Dim ws As Worksheet
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Stock")
    
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ws.Cells(i, 6).Value = ws.Cells(i, 3).Value * ws.Cells(i, 4).Value
    Next i
End Sub

' =========================
' ALERTE STOCK FAIBLE
' =========================
Public Sub AlerteStockFaible()
    Dim ws As Worksheet
    Dim i As Long
    Dim msg As String
    
    Set ws = ThisWorkbook.Sheets("Stock")
    
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 3).Value <= ws.Cells(i, 5).Value Then
            msg = msg & "- " & ws.Cells(i, 2).Value & vbCrLf
            ws.Rows(i).Interior.Color = RGB(255, 200, 200)
        Else
            ws.Rows(i).Interior.ColorIndex = xlNone
        End If
    Next i
    
    If msg <> "" Then
        MsgBox " Stock faible :" & vbCrLf & msg, vbExclamation
    Else
        MsgBox "Stock OK", vbInformation
    End If
End Sub

' =========================
' OUVRIR FORMULAIRE
' =========================
Sub OuvrirFormulaire()
    frmStock.Show
End Sub

' =========================
' ===== USERFORM CODE =====
' =========================

Private Sub btnAjouter_Click()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Sheets("Stock")
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ws.Cells(lastRow, 1).Value = txtID.Value
    ws.Cells(lastRow, 2).Value = txtNom.Value
    ws.Cells(lastRow, 3).Value = txtQte.Value
    ws.Cells(lastRow, 4).Value = txtPrix.Value
    ws.Cells(lastRow, 5).Value = txtSeuil.Value
    
    Call CalculValeurStock
    Call AlerteStockFaible
    
    MsgBox "Produit ajouté", vbInformation
End Sub

Private Sub btnModifier_Click()
    Dim ws As Worksheet
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Stock")
    
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 1).Value = txtID.Value Then
            
            ws.Cells(i, 2).Value = txtNom.Value
            ws.Cells(i, 3).Value = txtQte.Value
            ws.Cells(i, 4).Value = txtPrix.Value
            ws.Cells(i, 5).Value = txtSeuil.Value
            
            Call CalculValeurStock
            Call AlerteStockFaible
            
            MsgBox "Produit modifié", vbInformation
            Exit Sub
        End If
    Next i
    
    MsgBox "Produit non trouvé", vbExclamation
End Sub

Private Sub btnSupprimer_Click()
    Dim ws As Worksheet
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Stock")
    
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 1).Value = txtID.Value Then
            
            ws.Rows(i).Delete
            
            Call AlerteStockFaible
            
            MsgBox "Produit supprimé", vbInformation
            Exit Sub
        End If
    Next i
    
    MsgBox "Produit non trouvé", vbExclamation
End Sub
