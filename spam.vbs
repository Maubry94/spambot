set shell = createobject ("wscript.shell")

strtext = inputbox ("Ecris le message que tu veux spam")
strtimes = inputbox ("Combien de fois veux-tu spam ?")
strspeed = inputbox ("A quelle vitesse veux-tu spam (1000 = 1s) ?")
strtimeneed = inputbox ("Combien de temps as-tu besoin avant que le script ne se lance ?")

If not isnumeric (strtimes & strspeed & strtimeneed) then
msgbox "Tu as entre un caractere invalide. Recommence."
wscript.quit
End If
strtimeneed2 = strtimeneed * 1000
do
msgbox "Tu as " & strtimeneed & " secondes pour acceder a la zone de saisie ou tu veux envoyer le spam."
wscript.sleep strtimeneed2
for i=0 to strtimes
shell.sendkeys (strtext & "{enter}")
wscript.sleep strspeed
Next
wscript.sleep strspeed * strtimes / 10
returnvalue=MsgBox ("Encore ?",36)
If returnvalue=6 Then
Msgbox "Ok rebelote."
End If
If returnvalue=7 Then
msgbox "Ok, stop."
wscript.quit
End IF
loop