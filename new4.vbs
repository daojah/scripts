Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

objSelection.TypeText "Ken Myer"
objSelection.TypeParagraph()
objSelection.TypeText "Fabrikam Corporation"

Set objSelection = objDoc.Range()

Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

objSignatureEntries.Add "Scripted Entry", objSelection

Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = _ 
    objEmailOptions.EmailSignature

objSignatureObject.NewMessageSignature = _ 
    "Scripted Entry"
objSignatureObject.ReplyMessageSignature = _ 
    "Scripted Entry"