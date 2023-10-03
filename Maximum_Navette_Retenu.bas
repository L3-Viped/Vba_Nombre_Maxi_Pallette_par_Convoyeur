Attribute VB_Name = "Maximum_Navette_Retenu"
Sub Max_Navette_Retenu()


Nombre_de_la_postion_de_la_butee_de_chargement:
    If Not IsNumeric(Range("G10")) Then GoTo Ici:
    If Range("G10") > (((Range("G3") - 600) / Range("G4")) + (1 / 2)) Or Range("G10") < 0 Or Range("G10") = "" Then
Ici:
    ActiveSheet.Unprotect Password:="Test" 'À Changer en Fonction du Mdp de la Feuille
    Avertissement = MsgBox("Valeur incorrect." & Chr(10) & "Merci de la revoir.", vbInformation + vbOKOnly, "Avertissement")
    With Range("G10")
    .Select
    Application.Undo
    CreateObject("WScript.Shell").SendKeys "{F2}", True
    CreateObject("WScript.Shell").SendKeys ("^a"), True
    ActiveSheet.Protect Password:="Test"
    Exit Sub
    End With
    End If
    If Not Range("G10").Value = Round(Range("G10"), 0) Then
    Range("G10").Value = WorksheetFunction.RoundDown(Range("G10"), 0)
    End If

 End Sub
