Attribute VB_Name = "Distance_Entreaxe_E"
Sub Distance_E()

    Dim E As Variant
    
Nombre_de_la_postion_de_la_butee_de_chargement:
    E = Range("G3") 'À Changer en Fonction de la Localistion de la Feuille
    If Not IsNumeric(E) Then GoTo Ici:
    If E < (Range("G6") + Range("G8") + 420 + 100 + Range("G4")) Or E < ((Range("G4") * 2) + Range("G6") + Range("G8") + 200) Or E = "" Then
Ici:
    ActiveSheet.Unprotect Password:="Test" 'À Changer en Fonction du Mdp de la Feuille
    Avertissement = MsgBox("Valeur incorrect." & Chr(10) & "Merci de la revoir.", vbInformation + vbOKOnly, "Avertissement")
    With Range("G3")
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

