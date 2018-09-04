
'If with InputBox and Sub(birthday -> age)


'Head========================================

Option Explicit

Dim resp
Dim name
Dim birth

Dim objShell, nunaPath
Set objShell = CreateObject("WScript.Shell")
nunaPath=objShell.CurrentDirectory &"\"

'Body=========================================


REM resp=MsgBox ("Vc pode nos fornecer o seu nome?", VbYesNo+VbQuestion)

REM if resp= VbYes Then
REM name=InputBox ("Type it here: ")
REM MsgBox "Hello " + name + ". Very good!",VbInformation

resp=MsgBox("Vamos verificar a sua idade?", VbYesNo)
If resp=VbYes then
birth=InputBox("Digite seu ano_de_nascimento: ", " ", "Numeric Value")
Call NiverAge(birth)

Elseif resp=VbNo Then
MsgBox "Sad. We cannot continue the program.",VbOkOnly+vbCritical
end if

'SUB=========================================

Sub NiverAge(niver)
Dim year, age
year=2018
age=year-niver
MsgBox "Parabens! Vemos q hj vc esta com : " & age & " anos_de_idade."
MsgBox "Thanks 4your Attention. =]"
End Sub

'End of Procedure ----

Wscript.sleep 1500
objShell.run nunaPath &"exsub_salut_day.vbs"


