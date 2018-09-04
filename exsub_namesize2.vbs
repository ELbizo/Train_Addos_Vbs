
'Sub String-Name_Size [vs2] with "Object_Wscript.Shell"

'This only takes a name, put it to UpperCase and show 
'the length of the string.

'HEAD ================================

Option Explicit

Dim Objshell
Set Objshell=CreateObject("wscript.shell")

Dim yname
Dim choice

choice=MsgBox("Vamos verificar o seu nome?", VbYesNo+VbQuestion)
if choice=VbYes Then
yname=Trim(InputBox("Digite-o p. favor: "))
Call Maisc_Size(yname)
Elseif choice=VbNo Then
MsgBox "Ok. Grato pela utilizacao."
End if
wscript.quit

'SUB_Procedure [vs2]=====================

Sub Maisc_Size(string)
Dim recebe
Dim grando
Dim bytes
recebe=UCase(string)
grando=Len(string)
bytes=grando*8

Dim obj
set obj=CreateObject("wscript.shell")

wscript.sleep 2000
obj.run "notepad.exe"
wscript.sleep 2000

obj.sendkeys "Conferindo o seu nome: "
obj.sendkeys yname
wscript.sleep 2400
obj.sendkeys "{enter}"
obj.sendkeys "Isso esta correto, SIM?"
wscript.sleep 2600
obj.sendkeys "{enter 2}"
obj.sendkeys "Nome OK, portanto."
obj.sendkeys "{enter}"
wscript.sleep 2400

obj.sendkeys "{enter}"
wscript.sleep 1400
obj.sendkeys "%asn"
wscript.sleep 2500

MsgBox "Muito bem, " & recebe & "! " & "O tamanho do seu nome eh de: " & grando & " caracteres; " & VbCrlf & "o que corresponde a: " & bytes & " bytes de 'tamanho' =]"
wscript.sleep 1400
MsgBox "Grato pelo uso e consideracao."

REM wscript.sleep 1400
REM MsgBox "Vamos verificar tb o seu Niver(...)"
REM wscript.sleep 1400
REM Objshell.run "U:\Programming2\Vbscript_GitSent\exsub_niver-age.exe"

End Sub


