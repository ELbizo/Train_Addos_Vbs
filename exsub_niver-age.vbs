
'If with InputBox and Sub(birthday -> age)


'Head========================================

Option Explicit

Dim resp
Dim name
Dim birth

'Body=========================================

resp=MsgBox ("Hi there! Vc pode nos fornecer o seu nome?", VbYesNo+VbQuestion)

if resp= VbYes Then
name=InputBox ("Type it here: ")
MsgBox "Hello " + name + ". Very good!",VbInformation

birth=InputBox("Agora, digite seu ano_de_nascimento, please.", " ", "Numeric Value")
Call NiverAge(birth)

elseif resp=VbNo Then
MsgBox "Cannot continue the program, Sorry.",VbOkOnly+vbCritical
end if

'SUB=========================================

Sub NiverAge(niver)
Dim year, age
year=2018
age=year-niver
MsgBox "Parabens! Vemos q hj vc esta com : " & age & " anos_de_idade."
MsgBox "Thanks 4your Attention. =]"
End Sub



