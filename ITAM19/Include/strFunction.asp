<%
' File: strFunction.asp
' Version: 2.10 (2006-11-28, Kees Hiemstra)
' - Added the function ClearText
'
'***********************************************************************************
' Verify data entered using Required/Invalid characters and/or Minimum/Maximum length
'***********************************************************************************
Public Function ValidateText(strText, RequiredChars, InvalidChars, MinLength, MaxLength)
  Dim ivtCounter
  For ivtCounter = 1 To Len(RequiredChars)
    if instr(StrText, mid(RequiredChars, ivtCounter, 1)) = 0 then
    ValidateText = "Missing required characters (""" & RequiredChars & """)."
    exit for
   end if
  next
  if ValidateText <> "" then exit Function
  for ivtCounter = 1 to len(InvalidChars)
   if instr(StrText,mid(InvalidChars, ivtCounter, 1)) <> 0 then
    ValidateText = "Found invalid character (""" & mid(InvalidChars, ivtCounter ,1) & """)."
    Exit for
   end if
  next
  if ValidateText <> "" then exit Function
  if len(StrText) < MinLength then
   ValidateText =  "Text to short, minimum of " & MinLength & " characters."
  end if
  if len(StrText) > MaxLength and MaxLength > MinLength then
   ValidateText =  "Text to long, maximum of " & MaxLength & " characters."
  end if
end function

Public Function ClearText(strText)
  If Trim(strText) = "" Then
    ClearText = "&nbsp;"
  Else
    ClearText = Trim(strText)
  End If
End Function

%>
