Attribute VB_Name = "Fin2Var"
Dim st As String
Dim mid As String
Dim fin As String
Dim ms As String 'male start
Dim fs As String 'female start
Dim ns As String 'nothing to start
Dim s As String 'start for many
Dim EmailAddr As String

Sub FIOMAIN()
    EditMail2var (0) 'запуск процедуры с параметром 0, и следовательно порядок ФИО
End Sub

Sub IOFMAIN()
    EditMail2var (1) 'запуск процедуры с параметром 1, и следовательно порядок ИОФ
End Sub



Private Sub EditMail2var(f)

Dim MW As Outlook.MailItem


Dim Msg As String 'message which already exist

'Итоговая фраза состоит из st+mid+fin+Msg
        'Исходники
        s = "Уважаемые" 'start for many
        ms = "Уважаемый" 'male start
        fs = "Уважаемая" 'female start
        ns = "" 'nothing to start
        st = ns
        fin = "добрый день!"
        
        
            Dim myinspector As Outlook.Inspector
            Set myinspector = Application.ActiveInspector
            Set MW = myinspector.CurrentItem 'присваиваем MW значение активного открытого окна outlook
            Msg = MW.HTMLBody 'запоминаем что уже есть в данном окне набранного (подпись, предыдущее сообщение)
            MW.HTMLBody = Empty 'очищаем тело сообщения
          
            EmailAddr = MW.To 'определяем кто адресат
            
            If InStr(EmailAddr, ";") <> 0 Then 'если есть ; значит адресотов несколько и обращаемся во множественном числе
                st = s
                mid = "коллеги"
            Else
            If InStr(EmailAddr, "(") <> 0 Then EmailAddr = Left(EmailAddr, InStr(EmailAddr, "(") - 2) 'если адресат один проверяем нет ли после адреса скобки и электронной почты, если есть то почту отрезаем
               
               
               
               If f = 0 Then IfFIO Else IfIOF
               
                
            End If
            
            MW.HTMLBody = "<HTML><BODY><div><ActiveDocument.Styles('Обычный')>" & _
            st & " " & mid & ", " & fin _
            & "</div>" & Msg & "</ActiveDocument.Styles('Обычный')></div></body></html>" 'собираем итоговое сообщение, с приветствием и тем что было изначально
            
            SendKeys "{End}~{NUMLOCK}" 'ставим курсор в конец строки и переходи на следующую для готовности сразу набирать текст, и заодно включаем нумлок
            
            MW.Display 'показываем активное окно
             
             
            
            
            
            
End Sub

Private Sub IfFIO()
                                            
                
                    
                mid = Right(EmailAddr, Len(EmailAddr) - InStr(EmailAddr, " ")) 'берём из адреса всё кроме первого слова
                
                    If InStr(Len(EmailAddr) - 3, EmailAddr, "вна") <> 0 Then 'если второе слово (отчество) заканчивается на "вна"
                    st = fs 'то обращаемся к женщине
                    Else
                        If InStr(Len(EmailAddr) - 3, EmailAddr, "вич") <> 0 Then st = ms 'если второе слово (отчество) заканчивается на "вич"
                    End If 'в противном случае определить пол не представляется возможным, поэтому обращение будет без "Уважаемый(ая)"
           

End Sub


Private Sub IfIOF()
                                              
                
                   
                For i = 1 To 2
                    x = InStr(x + 1, EmailAddr, " ") 'считаем сколько слов
                Next i
                
                If x = 0 And InStr(EmailAddr, " ") <> 0 Then 'если слов два
                mid = Left(EmailAddr, InStr(EmailAddr, " ") - 1) 'берём из адреса только первое слово
                Else
                    If x <> 0 Then mid = Left(EmailAddr, x - 1) Else mid = EmailAddr 'берём из адреса два первых слова
                End If
                    
                    If InStr(Len(EmailAddr) - 3, EmailAddr, "вна") <> 0 Then 'если второе слово (в идеале отчество) заканчивается на "вна"
                    st = fs 'то обращаемся к женщине
                    Else
                        If InStr(Len(EmailAddr) - 3, EmailAddr, "вич") <> 0 Then st = ms 'если второе слово (в идеале отчество) заканчивается на "вич"
                    End If 'в противном случае определить пол не представляется возможным, поэтому обращение будет без "Уважаемый(ая)"
            
End Sub








