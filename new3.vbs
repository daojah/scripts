On Error Resume Next

Set objSysInfo = CreateObject("ADSystemInfo")

strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)

strZpov = "S uvageniem, "
strPostIndex = ObjUser.postalCode
strName = objUser.FullName
strTitle = objUser.Title
strDepartment = objUser.Department
strCompany = objUser.Company
strPhone = objUser.telephoneNumber
strweb = objuser.wWWHomePage
strgorod = objuser.l
strstreet = objuser.streetAddress
strfax = objuser.facsimileTelephoneNumber
strIntPhone = objuser.ipPhone
strEmail = objuser.mail

Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

objSelection.ParagraphFormat.Space1
objselection.font.color = RGB(0, 0, 0)
objSelection.TypeText strZpov
objSelection.TypeText CHR(11)
objSelection.TypeText strName
objSelection.TypeText CHR(11)
objSelection.TypeText strTitle
objSelection.TypeText CHR(11)
objSelection.TypeText strCompany
objSelection.TypeText CHR(11)
objSelection.TypeText "Tel.    " & strPhone & " db. " & strintPhone
objSelection.TypeText CHR(11)
objselection.font.color = RGB(0, 0, 255)
objSelection.Hyperlinks.Add objSelection.range, "mailto:" & strEmail, , , strEmail 
objSelection.TypeText CHR(11)
objSelection.Hyperlinks.Add objSelection.Range, strWeb, "", "", strWeb
objSelection.TypeText CHR(11)
objselection.font.color = RGB(0, 0, 0)
objSelection.TypeText strPostIndex & strgorod & strstreet

Set objSelection = objDoc.Range()

objSignatureEntries.Add "AD Signature", objSelection
objSignatureObject.NewMessageSignature = "AD Signature"
objSignatureObject.ReplyMessageSignature = "AD Signature"

objDoc.Saved = True