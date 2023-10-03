Attribute VB_Name = "Longueur_Pas_Accumulation"
Sub longueur_PA()

    Dim P As Variant
    
Nombre_de_la_postion_de_la_butee_de_chargement:
    P = Range("G4") 'À Changer en Fonction de la Localistion de la Feuille
    If Not IsNumeric(P) Then GoTo Ici:
    If P < 270 Or P > (Range("G3") - (Range("G6") + Range("G8") + 100 + 420)) Or P > (Range("G3") - (Range("G8") + 200 + Range("G6"))) / 2 Or P = "" Then
Ici:
    ActiveSheet.Unprotect Password:="Test" 'À Changer en Fonction du Mdp de la Feuille
    Avertissement = MsgBox("Valeur incorrect." & Chr(10) & "Merci de la revoir.", vbInformation + vbOKOnly, "Avertissement")
    With Range("G4")
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
