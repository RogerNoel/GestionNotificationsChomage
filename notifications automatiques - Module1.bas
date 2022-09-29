Attribute VB_Name = "Module1"
Dim numEmp As Integer
Function creation_alerte_agenda(dateAlerte As Date)
    Set outlookApp = New Outlook.Application ' -----> cr�er un nouveau rendez-vous
               Set task = outlookApp.CreateItem(olAppointmentItem)
               Dim nom As String
               With task
                   .Subject = "Notification CE � venir"
                   .Body = "Une copie de ce mail a �t� envoy�e sur l'adresse kchaouki@lecap.be"
                   .Start = dateAlerte + TimeValue("10:00:00")
                   .Save
                End With
    ' pour l'appeler �crire simplement: creation_alerte(item)
End Function
' ---------------------------------------------------------------------------------------------------------

' AJOUT ULTERIEUR DE LIGNES DANS LE CLASSEUR
' Afficher les colonnes de L � R pour v�rifier les calculs automatiques
' ATTENTION le n� employeur en colonne R doit �tre entr� � la main !!!

Sub notifications_chomage()
    Dim nbreLignes As Integer
    Dim mailGestionnaire As String
    Dim dateEnvoi As Date
    
    nbreLignes = Cells.Find(what:="*", searchdirection:=xlPrevious).Row
    For i = 2 To nbreLignes
        If Range("i" & i) <> "" Then 's'il y a une date de fin
            If Range("n" & i) = "" Then ' si le mail n'est pas encore programm�
                If Range("q" & i) <= 1 Then 'la colonne Q v�rifie s'il y a d�j� un envoi programm� � ce gestionnaire et � cette date pour ne pas spammer
                    mailGestionnaire = Range("o" & i)
                    dateEnvoiMail = Range("p" & i).Value
                    dateFin = Range("i" & i)
                    If IsDate(dateFin) Then
                        If Range("j" & i) = 13 Then ' si c'est une p�riode de 13 semaines
                            numEmp = Range("r" & i)
                            'creation_alerte_agenda (dateFin - 3)
                            Call send_mail(mailGestionnaire, dateEnvoiMail)
                            Range("n" & i).Value = "OK" ' ok mail programm�
                        ElseIf Range("c" & i) = "N" Then ' C est la colonne qui d�finit si CP 124 ou non
                            numEmp = Range("r" & i)
                            'creation_alerte_agenda (dateFin-3)
                            Call send_mail(mailGestionnaire, dateEnvoiMail)
                            Range("n" & i).Value = "OK" ' ok mail programm�
                        Else
                            'creation_alerte_agenda (dateFin)
                            numEmp = Range("r" & i)
                            Call send_mail(mailGestionnaire, dateEnvoiMail)
                            Range("n" & i).Value = "OK" ' ok mail programm�
                        End If
                    End If
                End If
            End If
        End If
    Next i
    MsgBox "Traitement termin�"
End Sub
' ---------------------------------------------------------------------------------------------------------

Sub send_mail(ByVal mailGestionnaire As String, ByVal dateEnvoi As Date) '------------> call
    Dim outlookApp As Outlook.Application
    Dim outlookMail As Outlook.MailItem
    
    Set outlookApp = New Outlook.Application
    Set outlookMail = outlookApp.CreateItem(olMailItem)
    corpsMail = "Notifications de CE � remplir pour l'employeur " & numEmp & "."
    With outlookMail
        .BodyFormat = olFormatHTML
        .HTMLBody = corpsMail
        .To = "kchaouki@lecap.be;aknops@lecap.be"
        .CC = mailGestionnaire
        .Subject = "Notifications CE"
        .DeferredDeliveryTime = dateEnvoi + TimeSerial(12, 0, 0)
        .Send
    End With
End Sub

Private Sub UserForm_Initialize()
    Application.Calculation = xlAutomatic
End Sub

