Attribute VB_Name = "Distance_Chargement"
Sub Distance_C1()

    Dim C1 As Variant
    
Nombre_de_la_postion_de_la_butee_de_chargement:
    C1 = Range("G6") 'À Changer en Fonction de la Localistion de la Feuille
    If Not IsNumeric(P) Then GoTo Ici:
    If C1 < 300 Or C1 > (Range("G3") - (Range("G8") + 420 + 100 + Range("G4"))) Or C1 > (Range("G3") - (Range("G8") + 200 + Range("G4") * 2)) Or C1 = "" Then
Ici:
    ActiveSheet.Unprotect Password:="Test"
    Avertissement = MsgBox("Valeur incorrect." & Chr(10) & "Merci de la revoir.", vbInformation + vbOKOnly, "Avertissement")
    With Range("G6")
    .Select
    Application.Undo
    CreateObject("WScript.Shell").SendKeys "{F2}", True
    CreateObject("WScript.Shell").SendKeys ("^a"), True
    ActiveSheet.Protect Password:="Test"
    Exit Sub
    End With
    End If
    Call Retenue_Chargement
 End Sub
