Attribute VB_Name = "modColor"
'___|..:::---=={Release Notes}==--:::..___|
'|     Code is in English and German!     |
'|      Code in Englisch und Deutsch      |
'|Eng:                                    |
'|                                        |
'|So now a example how to put this Code in|
'|              a Modul                   |
'|                                        |
'|Ger:                                    |
'|                                        |
'|   So jetzt gibts noch ein Beispiel,    |
'| wie man den Code in ein Modul stecken  |
'|                könnte                  |
'|________________________________________|
'                    © 2004 Benjamin Asbach
Option Explicit
Public Sub Code2Color(Source As String, Target As RichTextBox)
    '//Our ForCounter\\
    '//Unser ForZähler\\
    Dim x As Integer
    '//Our Buffer, perhaps is our Target our Source!?\\
    '//Unser Buffer, vielleicht ist unser Ziel unsere Quelle!?\\
    Dim Buffer As String
    '//Fill the Buffer\\
    '//Fülle den Buffer\\
    Buffer = Source
    '//Clear the Destination\\
    '//Bereinige das Ziel\
    Target.Text = ""
    Target.SelColor = vbBlack
    '//The Main loop\\
    '//Die Hauptschleife\\
    For x = 1 To Len(Buffer)
        '//Check if current Char is our ColorChar\\
        '///Überprüfung ob aktuelles Zeichen unser Farbzeichen ist\\
        If Mid(Buffer, x, 1) = "^" And IsNumeric(Mid(Buffer, x + 1, 1)) Then
            '//Check which Color was chose, so check x + 1\\
            If Mid(Buffer, x + 1, 1) = 0 Then
            '//O.K. Code = 0 is Black\\
            '//O.K. Code = 0 ist Schwarz\\
                Target.SelColor = vbBlack
            ElseIf Mid(Buffer, x + 1, 1) = 1 Then
            '//O.K. Code = 1 is Red\\
            '//O.K. Code = 1 ist Rot\\
                Target.SelColor = vbRed
            ElseIf Mid(Buffer, x + 1, 1) = 2 Then
            '//O.K. Code = 2 is Yellow\\
            '//O.K. Code = 2 ist Gelb\\
                Target.SelColor = vbGreen
            ElseIf Mid(Buffer, x + 1, 1) = 3 Then
            '//O.K. Code = 3 is Green\\
            '//O.K. Code = 3 ist Grün\\
                Target.SelColor = vbYellow
            ElseIf Mid(Buffer, x + 1, 1) = 4 Then
            '//O.K. Code = 4 is Blue\\
            '//O.K. Code = 4 ist Blau\\
                Target.SelColor = vbBlue
            ElseIf Mid(Buffer, x + 1, 1) = 5 Then
            '//O.K. Code = 4 is Blue\\
            '//O.K. Code = 4 ist Blau\\
                Target.SelColor = vbCyan
            ElseIf Mid(Buffer, x + 1, 1) = 6 Then
            '//O.K. Code = 4 is Blue\\
            '//O.K. Code = 4 ist Blau\\
                Target.SelColor = &HFF00FF
            ElseIf Mid(Buffer, x + 1, 1) = 7 Then
            '//O.K. Code = 4 is Blue\\
            '//O.K. Code = 4 ist Blau\\
                Target.SelColor = vbWhite
            ElseIf Mid(Buffer, x + 1, 1) = 8 Then
            '//O.K. Code = 4 is Blue\\
            '//O.K. Code = 4 ist Blau\\
                Target.SelColor = &HC0C0C0
            ElseIf Mid(Buffer, x + 1, 1) = 9 Then
            '//O.K. Code = 4 is Blue\\
            '//O.K. Code = 4 ist Blau\\
                Target.SelColor = &HE0E0E0
            End If
            x = x + 1
        Else
            '//If ColorChar isn't Found Copy the Char\\
            '//Wenn kein Farbzeichen gefunden wurde, Kopiere das Zeichen\\
            Target.SelText = Mid(Buffer, x, 1)
        End If
    Next x
End Sub
'//Function that returns the noraml name without Colorsymbols\\
'//Diese Funktion gibt den normalen Namen ohne die Farbsymbole zurück\\
Public Function Code2Normal(Source As RichTextBox) As String
    '//This is similar to the other Algorythm, we only cut out some code\\
    '//Dieser Abschnitt ist dem anderen Code sehr änlich, nur etwas wenige Code\\
    
    '//Our ForCounter\\
    '//Unser ForZähler\\
    Dim x As Integer
    '//The Main loop\\
    '//Die Hauptschleife\\
    For x = 1 To Len(Source.Text)
        '//Check if current Char is our ColorChar\\
        '///Überprüfung ob aktuelles Zeichen unser Farbzeichen ist\\
        If Mid(Source.Text, x, 1) = "°" Then
            '//We have to skip a char because to set a new Color you need 2 Chars ("°0" for Black)\\
            '//I know it's ugly to chnage the For-Counter but it's the easiest way :)\\
            '//Jetzt müssen wir ein Zeichen überspringen, da wir 2 Zeichen bruachen ume eine neue Farbe zu definieren\\
            '//Ich weiß ist ein unsauberer Weg den For-Zähler zu verändern, aber das is am einfachsten\\
            x = x + 1
        Else
            '//If ColorChar isn't Found Copy the Char\\
            '//Wenn kein Farbzeichen gefunden wurde, Kopiere das Zeichen\\
            Code2Normal = Code2Normal & Mid(Source.Text, x, 1)
        End If
    Next x
End Function
