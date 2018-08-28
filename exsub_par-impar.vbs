
'Exercicio Simples

'Gerar um peq. codigo p determinacao de se Nro eh Par ou Impar. 
'Introducao ao conceito de Sub_Rotinas.

'Head===========================================

Option Explicit

Dim nro
Dim op

'2)Call of the subrotine --> Par_Impar

Do
op=MsgBox("Verificar um numero.: ",VbYesNo+VbQuestion)
if op=VbYes Then
nro=InputBox("Type Here","Par ou Impar")
Call Par_Impar(nro)
elseif op=VbNo Then
MsgBox "Grato pelo Uso.",VbInformation
exit do
end if
Loop

'1)Make of the subrotine-----------------

Sub Par_Impar(x)
if x mod 2=0 Then
MsgBox "Esse nro eh par."
elseif x mod 2<>0 Then
MsgBox "Esse nro eh impar."
end if
End Sub

