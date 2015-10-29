On Error Resume Next
Set objSysInfo = CreateObject("ADSystemInfo")
set objShell = WScript.CreateObject("WScript.Shell")

strUser = objSysInfo.UserName

Set objUser = GetObject("LDAP://" & strUser)

If Err.Number <> 0 Then
'Wscript.Echo " ", Err.Description
Err.Clear
else
strProgramFiles = objShell.ExpandEnvironmentStrings("%PROGRAMFILES%")

Set objFSO = CreateObject("Scripting.FileSystemObject")
If (objFSO.FileExists(strProgramFiles & "\Microsoft Office\Office12\OUTLOOK.EXE")) Then
strFont = "Calibri"
ElseIf (objFSO.FileExists(strProgramFiles & "\Microsoft Office\Office11\OUTLOOK.EXE")) then
strFont = "Arial"
End If

'
strZpov = "С уважением,"
strDev = "–"
strPostIndex = objuser.postalCode
strName = objuser.FullName
strTitle = objuser.Title
strPhysicalDeliveryOfficeName = objuser.PhysicalDeliveryOfficeName
strDepartment = objuser.Department
strCompany = objuser.Company
strPhone = objUser.telephoneNumber
strHPhone = objUser.homePhone
strMobilePhone = objUser.mobile
strPager = objUser.pager

strFax = objUser.facsimileTelephoneNumber

strGorod = objuser.l

strStreet = objuser.streetAddress
strEmail = objuser.mail
strWeb = "www.domain.com"

Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

objSelection.ParagraphFormat.Space1

objSelection.TypeText strZpov & CHR(11) & CHR(11)
objSelection.font.bold = true
objSelection.font.bold = wdToggle
objSelection.TypeText strName & CHR(11)

objSelection.TypeText strTitle & CHR(11)

if (strDepartment="") then
objSelection.TypeText ""
else
objSelection.TypeText strDepartment & CHR(11)
end if

objSelection.TypeText strPhysicalDeliveryOfficeName & CHR(11) & CHR(11)

'objSelection.TypeText strDev
'objSelection.TypeText CHR(11)
objSelection.TypeText "Компания " & CHR(34) & "CompanyName" & CHR(34) & CHR(11)
objSelection.TypeText strstreet & CHR(11)
objSelection.TypeText strPostIndex & " " & strGorod & CHR(11)

if (strPhone<>"") then
objSelection.TypeText "внутренний: " & strPhone & CHR(11)
end if

if (strPhone<>"" AND strHphone<>"") then
strShortPhone = Right(strPhone,3)
objSelection.TypeText "Тел.: " & strHphone & " доб.: " & strShortPhone & CHR(11)
end if

if (strPager <> "") then
objSelection.TypeText "Прямой: " & strPager & CHR(11)
end if

if (strMobilePhone<>"") then
objSelection.TypeText "Моб.: " & strMobilePhone & CHR(11)
end if

if (strFax <> "") then
objSelection.TypeText "Факс: " & strFax & CHR(11)
end if

objSelection.TypeText "e-mail: "

objSelection.Hyperlinks.Add objselection.range, "mailto:" & strEMail, , , strEMail
objSelection.TypeText CHR(11)
objSelection.Hyperlinks.Add objSelection.Range, strWeb, "", "", strWeb

Set objSelection = objDoc.Range()

objSelection.Font.Name = strFont
objSelection.Font.Size = "10"

objSignatureEntries.Add "AD Signature", objSelection
objSignatureObject.NewMessageSignature = "AD Signature"
objSignatureObject.ReplyMessageSignature = "AD Signature"

objDoc.Saved = True
objWord.Quit
end if