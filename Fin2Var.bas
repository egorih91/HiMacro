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
    EditMail2var (0) '������ ��������� � ���������� 0, � ������������� ������� ���
End Sub

Sub IOFMAIN()
    EditMail2var (1) '������ ��������� � ���������� 1, � ������������� ������� ���
End Sub



Private Sub EditMail2var(f)

Dim MW As Outlook.MailItem


Dim Msg As String 'message which already exist

'�������� ����� ������� �� st+mid+fin+Msg
        '���������
        s = "���������" 'start for many
        ms = "���������" 'male start
        fs = "���������" 'female start
        ns = "" 'nothing to start
        st = ns
        fin = "������ ����!"
        
        
            Dim myinspector As Outlook.Inspector
            Set myinspector = Application.ActiveInspector
            Set MW = myinspector.CurrentItem '����������� MW �������� ��������� ��������� ���� outlook
            Msg = MW.HTMLBody '���������� ��� ��� ���� � ������ ���� ���������� (�������, ���������� ���������)
            MW.HTMLBody = Empty '������� ���� ���������
          
            EmailAddr = MW.To '���������� ��� �������
            
            If InStr(EmailAddr, ";") <> 0 Then '���� ���� ; ������ ��������� ��������� � ���������� �� ������������� �����
                st = s
                mid = "�������"
            Else
            If InStr(EmailAddr, "(") <> 0 Then EmailAddr = Left(EmailAddr, InStr(EmailAddr, "(") - 2) '���� ������� ���� ��������� ��� �� ����� ������ ������ � ����������� �����, ���� ���� �� ����� ��������
               
               
               
               If f = 0 Then IfFIO Else IfIOF
               
                
            End If
            
            MW.HTMLBody = "<HTML><BODY><div><ActiveDocument.Styles('�������')>" & _
            st & " " & mid & ", " & fin _
            & "</div>" & Msg & "</ActiveDocument.Styles('�������')></div></body></html>" '�������� �������� ���������, � ������������ � ��� ��� ���� ����������
            
            SendKeys "{End}~{NUMLOCK}" '������ ������ � ����� ������ � �������� �� ��������� ��� ���������� ����� �������� �����, � ������ �������� ������
            
            MW.Display '���������� �������� ����
             
             
            
            
            
            
End Sub

Private Sub IfFIO()
                                            
                
                    
                mid = Right(EmailAddr, Len(EmailAddr) - InStr(EmailAddr, " ")) '���� �� ������ �� ����� ������� �����
                
                    If InStr(Len(EmailAddr) - 3, EmailAddr, "���") <> 0 Then '���� ������ ����� (��������) ������������� �� "���"
                    st = fs '�� ���������� � �������
                    Else
                        If InStr(Len(EmailAddr) - 3, EmailAddr, "���") <> 0 Then st = ms '���� ������ ����� (��������) ������������� �� "���"
                    End If '� ��������� ������ ���������� ��� �� �������������� ���������, ������� ��������� ����� ��� "���������(��)"
           

End Sub


Private Sub IfIOF()
                                              
                
                   
                For i = 1 To 2
                    x = InStr(x + 1, EmailAddr, " ") '������� ������� ����
                Next i
                
                If x = 0 And InStr(EmailAddr, " ") <> 0 Then '���� ���� ���
                mid = Left(EmailAddr, InStr(EmailAddr, " ") - 1) '���� �� ������ ������ ������ �����
                Else
                    If x <> 0 Then mid = Left(EmailAddr, x - 1) Else mid = EmailAddr '���� �� ������ ��� ������ �����
                End If
                    
                    If InStr(Len(EmailAddr) - 3, EmailAddr, "���") <> 0 Then '���� ������ ����� (� ������ ��������) ������������� �� "���"
                    st = fs '�� ���������� � �������
                    Else
                        If InStr(Len(EmailAddr) - 3, EmailAddr, "���") <> 0 Then st = ms '���� ������ ����� (� ������ ��������) ������������� �� "���"
                    End If '� ��������� ������ ���������� ��� �� �������������� ���������, ������� ��������� ����� ��� "���������(��)"
            
End Sub








