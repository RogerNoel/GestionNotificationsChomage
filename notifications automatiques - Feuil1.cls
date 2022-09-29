VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Ce code détecte un changement dans une cellule et demande une confirmation de ce changement
' si le changement est confirmé, rien de spécial
' si le changement est annulé, l'ancienne valeur sera restaurée
' IMPORTANT --> ce code doit être mis dans la FEUILLE concernée avec les options "worksheet"
' dans le menu de gauche et "change" dans le menu de droite

Private Sub CommandButton1_Click()
    Call notifications_chomage
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
Application.EnableEvents = False
    If Not Intersect(Target, Range("i:i")) Is Nothing Then
        ligne = Target.Row
        newvalue = Target.Value
        Application.Undo
        oldvalue = Target.Value
        Target.Value = newvalue
        Range("n" & ligne) = ""
        response = MsgBox("VERIFIEZ LA DATE", 1, "Confirmer ?")
        If response = vbCancel Then
            Target.Value = oldvalue
            Range("n" & ligne).Value = "OK"
        End If
    End If
    Application.EnableEvents = True
End Sub
