Attribute VB_Name = "Distance_D�chargement"
Sub Distance_D1()

    Dim D1 As Variant
    
Nombre_de_la_postion_de_la_butee_de_chargement:
    D1 = Range("G8") 'À Changer en Fonction de la Localistion de la Feuille
    If Not IsNumeric(P) Then GoTo Ici:
    If D1 < 300 Or D1 > (Range("G3") - (Range("G6") + 420 + 100 + Range("G4"))) Or D1 > (Range("G3") - (Range("G6") + 200 + Range("G4") * 2)) Or D1 = "" Then
Ici:
    ActiveSheet.Unprotect Password:="Test" 'À Changer en Fonction du Mdp de la Feuille
    Avertissement = MsgBox("Valeur incorrect." & Chr(10) & "Merci de la revoir.", vbInformation + vbOKOnly, "Avertissement")
    With Range("G8")
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
