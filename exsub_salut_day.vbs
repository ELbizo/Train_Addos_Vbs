
'Um primerio exercicio com Subrotinas. 
'Fazer um script p identificar corretamente o User.

'HEAD============================

Option Explicit 

Dim user

Do
MsgBox "Anyway... Qual seu nome, por favor: "
user=InputBox("Type it here.","Usuario_")
if Not IsNumeric(user) Then
Call Usuario(user)
Call Hodiau()
exit do
elseif IsNumeric(user) Then
MsgBox "Somente letras, por favor."
End if
Loop

'--> Making the subrotine1-----

Sub Usuario(string)
MsgBox "Entao, Sr(a). " & string& "..",VbInformation
End Sub

'--> Making the subrotine2-----

Sub Hodiau()
Dim verifica
verifica=Weekday(date)

if verifica=VbSunday Then
MsgBox "Que vc tenha um Otimo Domingo."
Elseif verifica=VbMonday Then
MsgBox "Que vc tenha uma boa Segunda-Feira."
Elseif verifica=VbTuesday Then
MsgBox "Que vc tenha uma boa Terca-Feira."
Elseif verifica=VbWednesday Then
MsgBox "Que vc tenha uma boa Quarta-Feira."
Elseif verifica=VbThursday Then
MsgBox "Que vc tenha uma boa Quinta-Feira."
Elseif verifica=VbFriday Then
MsgBox "Que vc tenha uma boa Sexta-Feira."
Elseif verifica=VbSaturday Then
MsgBox "Que vc tenha um otimo Sabado."
End if
End Sub

'-----

