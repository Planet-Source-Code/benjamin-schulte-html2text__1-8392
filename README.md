<div align="center">

## HTML2Text


</div>

### Description

This code is for converting a string with HTML tags and encodings into a text-only string. It rids multiple spaces and supports ALL encoded characters like &quot;, &nbsp;, &Auml; and so on.
 
### More Info
 
The original HTML String OrigHTML$

Text-only string


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Benjamin Schulte](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/benjamin-schulte.md)
**Level**          |Intermediate
**User Rating**    |3.9 (31 globes from 8 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/benjamin-schulte-html2text__1-8392/archive/master.zip)





### Source Code

```
'HTML2Text Copyright © 2000 Benjamin Schulte
'  (benni@bennisoft.de)
Public Function HTML2Text(ByVal OrigHTML$) As String
On Error Resume Next
If InStr(LCase$(OrigHTML$), "<body") > 0 Then
 OrigHTML$ = Mid$(OrigHTML$, InStr(LCase$(OrigHTML$), "<body"))
 OrigHTML$ = Mid$(OrigHTML$, InStr(OrigHTML$, ">") + 1)
 If InStr(LCase$(OrigHTML$), "</body>") > 0 Then _
 OrigHTML$ = Left$(OrigHTML$, InStr(LCase$(OrigHTML$), "</body>") - 1)
End If
Do While Len(OrigHTML$)
 CurrChar$ = Left$(OrigHTML$, 1)
 OrigHTML$ = Mid$(OrigHTML$, 2)
 Select Case CurrChar$
 Case " "
 OrigHTML$ = LTrim$(OrigHTML$)
 Case vbCr, vbLf
 CurrChar$ = ""
 If Left$(OrigHTML$, 1) = vbLf Then OrigHTML$ = Mid$(OrigHTML$, 2)
 OrigHTML$ = LTrim$(OrigHTML$)
 Case "<"
 CurrChar$ = ""
 If InStr(OrigHTML$, ">") > 0 Then
  CurrChar$ = Left$(OrigHTML$, InStr(OrigHTML$, ">") - 1)
  OrigHTML$ = Mid$(OrigHTML$, InStr(OrigHTML$, ">") + 1)
  Select Case LCase$(CurrChar$)
  Case "p", "/div"
  CurrChar$ = vbCrLf + vbCrLf
  Case "br"
  CurrChar$ = vbCrLf
  Case Else
  CurrChar$ = ""
  End Select
 End If
 Case "&"
 If InStr(OrigHTML$, ";") > 0 And InStr(OrigHTML$, ";") < InStr(OrigHTML$, " ") Then
  CurrChar$ = Left$(OrigHTML$, InStr(OrigHTML$, ";") - 1)
  OrigHTML$ = Mid$(OrigHTML$, InStr(OrigHTML$, ";") + 1)
  Select Case CurrChar$
  Case "amp"
  CurrChar$ = "&"
  Case "quot"
  CurrChar$ = """"
  Case "lt"
  CurrChar$ = "<"
  Case "gt"
  CurrChar$ = ">"
  Case "nbsp"
  CurrChar$ = " "
  Case "Auml"
  CurrChar$ = "Ä"
  Case "auml"
  CurrChar$ = "ä"
  Case "iexcl"
  CurrChar$ = "¡"
  Case "cent"
  CurrChar$ = "¢"
  Case "pound"
  CurrChar$ = "£"
  Case "curren"
  CurrChar$ = "¤"
  Case "yen"
  CurrChar$ = "¥"
  Case "brvbar"
  CurrChar$ = "|"
  Case "sect"
  CurrChar$ = "§"
  Case "uml"
  CurrChar$ = "¨"
  Case "copy"
  CurrChar$ = "©"
  Case "ordf"
  CurrChar$ = "ª"
  Case "laquo"
  CurrChar$ = "«"
  Case "not"
  CurrChar$ = "¬"
  Case "reg"
  CurrChar$ = "®"
  Case "macr"
  CurrChar$ = "¯"
  Case "deg"
  CurrChar$ = "°"
  Case "plusm"
  CurrChar$ = "±"
  Case "sup2"
  CurrChar$ = "²"
  Case "sup3"
  CurrChar$ = "³"
  Case "acute"
  CurrChar$ = "´"
  Case "micro"
  CurrChar$ = "µ"
  Case "para"
  CurrChar$ = "¶"
  Case "middot"
  CurrChar$ = "·"
  Case "cedil"
  CurrChar$ = "¸"
  Case "sup1"
  CurrChar$ = "¹"
  Case "ordm"
  CurrChar$ = "º"
  Case "raquo"
  CurrChar$ = "»"
  Case "frac14"
  CurrChar$ = "¼"
  Case "frac12"
  CurrChar$ = "½"
  Case "frac34"
  CurrChar$ = "¾"
  Case "iquest"
  CurrChar$ = "¿"
  Case "Agrave"
  CurrChar$ = "À"
  Case "Aacute"
  CurrChar$ = "Á"
  Case "Acirc"
  CurrChar$ = "Â"
  Case "Atilde"
  CurrChar$ = "Ã"
  Case "Aring"
  CurrChar$ = "Å"
  Case "AElig"
  CurrChar$ = "Æ"
  Case "Ccedil"
  CurrChar$ = "Ç"
  Case "Egrave"
  CurrChar$ = "È"
  Case "Eacute"
  CurrChar$ = "É"
  Case "Ecirc"
  CurrChar$ = "Ê"
  Case "Euml"
  CurrChar$ = "Ë"
  Case "Igrave"
  CurrChar$ = "Ì"
  Case "Iacute"
  CurrChar$ = "Í"
  Case "Icirc"
  CurrChar$ = "Î"
  Case "Iuml"
  CurrChar$ = "Ï"
  Case "ETH"
  CurrChar$ = "Ð"
  Case "Ntilde"
  CurrChar$ = "Ñ"
  Case "Ograve"
  CurrChar$ = "Ò"
  Case "Oacute"
  CurrChar$ = "Ó"
  Case "Ocirc"
  CurrChar$ = "Ô"
  Case "Otilde"
  CurrChar$ = "Õ"
  Case "Ouml"
  CurrChar$ = "Ö"
  Case "times"
  CurrChar$ = "×"
  Case "Oslash"
  CurrChar$ = "Ø"
  Case "Ugrave"
  CurrChar$ = "Ù"
  Case "Uacute"
  CurrChar$ = "Ú"
  Case "Ucirc"
  CurrChar$ = "Û"
  Case "Uuml"
  CurrChar$ = "Ü"
  Case "Yacute"
  CurrChar$ = "Ý"
  Case "THORN"
  CurrChar$ = "Þ"
  Case "szlig"
  CurrChar$ = "ß"
  Case "agrave"
  CurrChar$ = "à"
  Case "aacute"
  CurrChar$ = "á"
  Case "acirc"
  CurrChar$ = "â"
  Case "atilde"
  CurrChar$ = "ã"
  Case "aring"
  CurrChar$ = "å"
  Case "aelig"
  CurrChar$ = "æ"
  Case "ccedil"
  CurrChar$ = "ç"
  Case "egrave"
  CurrChar$ = "è"
  Case "eacute"
  CurrChar$ = "é"
  Case "ecirc"
  CurrChar$ = "ê"
  Case "euml"
  CurrChar$ = "ë"
  Case "igrave"
  CurrChar$ = "ì"
  Case "iacute"
  CurrChar$ = "í"
  Case "icirc"
  CurrChar$ = "î"
  Case "iuml"
  CurrChar$ = "ï"
  Case "eth"
  CurrChar$ = "ð"
  Case "ntilde"
  CurrChar$ = "ñ"
  Case "ograve"
  CurrChar$ = "ò"
  Case "oacute"
  CurrChar$ = "ó"
  Case "ocirc"
  CurrChar$ = "ô"
  Case "otilde"
  CurrChar$ = "õ"
  Case "ouml"
  CurrChar$ = "ö"
  Case "divide"
  CurrChar$ = "÷"
  Case "oslash"
  CurrChar$ = "ø"
  Case "ugrave"
  CurrChar$ = "ù"
  Case "uacute"
  CurrChar$ = "ú"
  Case "ucirc"
  CurrChar$ = "û"
  Case "uuml"
  CurrChar$ = "ü"
  Case "yacute"
  CurrChar$ = "ý"
  Case "thorn"
  CurrChar$ = "þ"
  Case "yuml"
  CurrChar$ = "ÿ"
  Case Else
  CurrChar$ = "&" + CurrChar$ + ";"
  End Select
 End If
 End Select
 NoHTML$ = NoHTML$ + CurrChar$
Loop
HTML2Text = NoHTML$
End Function
```

