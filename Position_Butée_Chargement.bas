Attribute VB_Name = "Position_But�e_Chargement"
Sub Position_Butee_Chargement()


Varation_de_la_postion_de_la_butee_de_chargement:
Standard:
    ActiveSheet.Unprotect Password:="Test" 'À Changer en Fonction du Mdp de la Feuille
    If Range("G5") = "Oui" Then
    With Range("G6")
    .Locked = True
    .Interior.Color = RGB(220, 220, 220)
    .Font.Color = RGB(80, 80, 80)
    .Value = 300
    End With
Non_Standard:
    ElseIf Range("G5") = "Non" Then
    With Range("G6")
    .Locked = False
    .Interior.Color = RGB(255, 255, 255)
    .Font.Color = RGB(0, 0, 0)
    .Value = 300
    End With
    End If
    ActiveSheet.Protect Password:="Test"

 End Sub

