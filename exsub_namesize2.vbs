
'Sub String-Name_Size [vs2] with "Object_Wscript.Shell"

'This only takes a name, put it to UpperCase and show 
'the length of the string.

'HEAD ================================

Option Explicit

Dim yname
Dim choice

choice=MsgBox("Vamos verificar o seu nome?", VbYesNo+VbQuestion)
if choice=VbYes Then
yname=Trim(InputBox("Digite-o p. favor: "))
Call Maisc_Size(yname)

Elseif choice=VbNo Then
MsgBox "Ok. Grato pela utilizacao.", VbInformation
End if

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

MsgBox "procedure of: 'string.len' @honorama_code", VbInformation
End Sub

