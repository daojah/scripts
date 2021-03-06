' Скрипт создания корпоративной подписи и наведения порядка в MS Outlook
' Работает в Outlook 2000 - 2010.
' Делает очень полезные вещи:
' - Выставляет имя отправителя почты из поля DisplayName в домене
' - Отключает HTML-просмотр писем и отправку писем в HTML.
' - Создаёт простую текстовую подпись и выставляет её для всех учеток.
' Если есть вопросы или жгучее желание дать мне денег\набить морду -
' моя электропочта artem@brodetskiy.net .
 
' Конфигурационные параметры:
 
' Название фирмы:
prmOOO_Name =  "Syntegra Company"
' Сайт фирмы:
prmSite = "http://www.syntegra.com.ua/"
 
' Переменная BasePath определяет путь к файлам с данными пользователей.
' В случае, если её значение равняется ("") - данные пользователя берутся из домена
' Для доменной учетной записи используются поля DisplayName (Выводимое имя), mail (Эл. почта),
' telephoneNumber (Телефонный номер), title (Должность), mobile (Мобильный телефон);
' Для файлов используется построчное перечисление:
' 1 строка - ФИО,
' 2 строка - Должность
' 3 cтрока - e-mail
' 4 строка - служебный телефон
' 5 строка - мобильный телефон (необязательно).
' Полное имя файла должно выглядеть как BasePath\имя_компьютера\имя_пользователя.ini
' Например, \\server\UserData$\comp01\user04.ini (BasePath = "\\server\UserData$\")
BasePath = ""
 
 
'BasePath = "\\habraserver\HabraSignatures$\"    
'BasePath = "h:\habrascript\habratest\"
 
 
' ==========================================================================================================================
' Секция подпрограмм:
' ==========================================================================================================================
 
' Функция удаляет все файлы из папки.
Sub ClearFolder(parmPath)
Dim oSubDir, oSubFolder, oFile, n
 
   On Error Resume Next          
   Set oSubFolder = fso.getfolder(parmPath)
   For Each oFile In oSubFolder.Files    
      If Err.Number <> 0 Then    
         Err.Clear
      Else
             fso.DeleteFile oFile.Path, True
      End If
   Next
   For Each oSubDir In oSubFolder.Subfolders
      ClearFolder oSubDir.Path      
   Next
   On Error Goto 0              
End Sub
 
' Функция проверяет наличие значения в реестре
Function KeyExists(key)
            Dim key2
            On Error Resume Next
            key2 = WshShell.RegRead(key)
            If Err.Number <> 0 Then
                        KeyExists = False
            Else
                        KeyExists = True
            End If
            On Error GoTo 0
End Function
 
' Функция проверяет наличие ключа реестра
Function RegistryKeyExists (RegistryKey)
  If (Right(RegistryKey, 1) <> "\") Then
    RegistryKeyExists = false
  Else
    On Error Resume Next
    WshShell.RegRead RegistryKey
    Select Case Err
      Case 0:
        RegistryKeyExists = true
      Case &h80070002:
 
        ErrDescription = Replace(Err.description, RegistryKey, "")
        Err.clear
        WshShell.RegRead "HKEY_ERROR\"
        If (ErrDescription <> Replace(Err.description, "HKEY_ERROR\", "")) Then
          RegistryKeyExists = true
        Else
          RegistryKeyExists = false
        End If      
      Case Else:
        RegistryKeyExists = false
    End Select    
    On Error Goto 0
  End If
End Function
 
' Функция получает данные пользователя из LDAP
Sub GetDomainCreds()
        set LocalRoot = getObject("LDAP://RootDSE")
        DefNC = LocalRoot.get("DefaultNamingContext")
        strPathCopy = "<LDAP://" & DefNC & ">;"
        strCriteria = "(&(objectCategory=person)(objectClass=user)(sAMaccountname="&strUser&"));"
        strProperties = "DisplayName, userPrincipalName, telephoneNumber, title, mobile;"
        strScope = "Subtree"
        set objConnect = CreateObject("ADODB.Connection")
        objConnect.Provider = "ADsDSOObject"
        objConnect.Open = "Active Directory Provider"
        set objCommand = CreateObject("ADODB.Command")
        set objCommand.ActiveConnection = objConnect
        objCommand.CommandText = strPathCopy & strCriteria & strProperties & strScope
        objCommand.Properties("Page Size") = 1000
        objCommand.Properties("Size Limit") = 1
        objCommand.Properties("Timeout") = 30
        Set objRecordSet = objCommand.Execute
       
        strDisplayName = objRecordSet.Fields("DisplayName").Value
       
        strmail = objRecordSet.Fields("userPrincipalName").Value & vbcrlf
       
        strtelephoneNumber = objRecordSet.Fields("telephoneNumber").Value
		
       
        if (strtelephoneNumber <> "") then strtelephoneNumber = strtelephoneNumber & vbcrlf
        strtitle = objRecordSet.Fields("title").Value
       
        if (strtitle <> "") then
                strtitle = strtitle & vbcrlf
        else  
                strtitle = prmOOO_Name & vbcrlf
        end if
       
        strmobile = objRecordSet.Fields("mobile").Value
        if (strmobile <> "") then strmobile = strmobile & " (моб.)"& vbcrlf
End Sub
 
' Функция получает данные пользователя из файла
Sub GetFileCreds()
        strFile = BasePath & strComputerName & "\" & strUser & ".ini"
       
        'Если нет файла конфигурации, а пользователь сидит с оутлуком - он будет отправлять без подписи, непорядок!
        'Надо предупредить. К счастью, в домене такой проблемы не бывает.
        if not fso.FileExists (strFile) then
                Wscript.Echo "У вас не установлена подпись в Outlook. Обратитесь к сисадмину, он поможет. "
                Wscript.Quit   
        End If
 
        Set ts = fso.OpenTextFile(strFile, 1)
        strDisplayName = ts.ReadLine()
        strtitle = ts.ReadLine()
        if (strtitle <> "") then
                strtitle = strtitle & " " & prmOOO_Name & vbcrlf
        else  
        strtitle = prmOOO_Name & vbcrlf
        end if
       
        strmail = ts.ReadLine() & vbcrlf
        strtelephoneNumber = ts.ReadLine()
        if (strtelephoneNumber <> "") then strtelephoneNumber = strtelephoneNumber & vbcrlf
        if not ts.AtEndOfStream then
                strmobile = ts.ReadLine()
                        if (strmobile <> "") then strmobile = strmobile & " (моб.)"& vbcrlf
        end if
        ts.close
End Sub
 
 
' ==========================================================================================================================
' Основная секция:
' ==========================================================================================================================
 
' Определяем  переменные, в которых будем хранить данные пользователя
Dim strDisplayName
Dim strtitle
Dim strtelephoneNumber
Dim strmobile
Dim strmail
 
' Создаем нужные нам объекты
Set WshNetwork = WScript.CreateObject("WScript.Network")
set WshShell =  WScript.CreateObject("WScript.Shell")
Set fso = WScript.CreateObject("Scripting.FileSystemObject")
 
' Юзернейм, копьютернейм, имя папки Application Data (на висте\вин7 у неё другое название)
strUser = WshNetwork.UserName
strComputerName = WshNetwork.ComputerName
Folder = WshShell.SpecialFolders("AppData")
 
 
 
' Если у пользователя не стоит офис - он идёт лесом.
if not RegistryKeyExists("HKEY_CURRENT_USER\Software\Microsoft\Office\")  then 
        Wscript.Quit
        End If
       
 
' Проверяем BasePath и решаем, откуда нам брать учетные данные
If BasePath = "" then
                GetDomainCreds()
                else
                GetFileCreds()
                End If
 
 
 
' Делаем подпись
Signature = "<html xmlns:v=""urn:schemas-microsoft-com:vml""xmlns:o=""urn:schemas-microsoft-com:office:office""xmlns:w=""urn:schemas-microsoft-com:office:word""xmlns:m=""http://schemas.microsoft.com/office/2004/12/omml""xmlns=""http://www.w3.org/TR/REC-html40""><head><meta http-equiv=Content-Type content=""text/html; charset=windows-1251""><meta name=ProgId content=Word.Document><meta name=Generator content=""Microsoft Word 15""><meta name=Originator content=""Microsoft Word 15""><link rel=File-List href=""New.files/filelist.xml""><link rel=Edit-Time-Data href=""New.files/editdata.mso""><!--[if !mso]><style>v\:* {behavior:url(#default#VML);}o\:* {behavior:url(#default#VML);}w\:* {behavior:url(#default#VML);}.shape {behavior:url(#default#VML);}</style><![endif]--><!--[if gte mso 9]><xml> <o:OfficeDocumentSettings>  <o:AllowPNG/> </o:OfficeDocumentSettings></xml><![endif]--><link rel=themeData href=""New.files/themedata.thmx""><link rel=colorSchemeMapping href=""New.files/colorschememapping.xml""><!--[if gte mso 9]><xml> <w:WordDocument>  <w:View>Normal</w:View>  <w:Zoom>0</w:Zoom>  <w:TrackMoves/>  <w:TrackFormatting/>  <w:PunctuationKerning/>  <w:ValidateAgainstSchemas/>  <w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid>  <w:IgnoreMixedContent>false</w:IgnoreMixedContent>  <w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText>  <w:DoNotPromoteQF/>  <w:LidThemeOther>RU</w:LidThemeOther>  <w:LidThemeAsian>X-NONE</w:LidThemeAsian>  <w:LidThemeComplexScript>X-NONE</w:LidThemeComplexScript>  <w:DoNotShadeFormData/>  <w:Compatibility>   <w:BreakWrappedTables/>   <w:SnapToGridInCell/>   <w:WrapTextWithPunct/>   <w:UseAsianBreakRules/>   <w:DontGrowAutofit/>   <w:SplitPgBreakAndParaMark/>   <w:EnableOpenTypeKerning/>   <w:DontFlipMirrorIndents/>   <w:OverrideTableStyleHps/>   <w:UseFELayout/>  </w:Compatibility>  <m:mathPr>   <m:mathFont m:val=""Cambria Math""/>   <m:brkBin m:val=""before""/>   <m:brkBinSub m:val=""&#45;-""/>   <m:smallFrac m:val=""off""/>   <m:dispDef/>   <m:lMargin m:val=""0""/>   <m:rMargin m:val=""0""/>   <m:defJc m:val=""centerGroup""/>   <m:wrapIndent m:val=""1440""/>   <m:intLim m:val=""subSup""/>   <m:naryLim m:val=""undOvr""/>  </m:mathPr></w:WordDocument></xml><![endif]--><!--[if gte mso 9]><xml> <w:LatentStyles DefLockedState=""false"" DefUnhideWhenUsed=""false""  DefSemiHidden=""false"" DefQFormat=""false"" DefPriority=""99""  LatentStyleCount=""371"">  <w:LsdException Locked=""false"" Priority=""0"" QFormat=""true"" Name=""Normal""/>  <w:LsdException Locked=""false"" Priority=""9"" QFormat=""true"" Name=""heading 1""/>  <w:LsdException Locked=""false"" Priority=""9"" SemiHidden=""true""   UnhideWhenUsed=""true"" QFormat=""true"" Name=""heading 2""/>  <w:LsdException Locked=""false"" Priority=""9"" SemiHidden=""true""   UnhideWhenUsed=""true"" QFormat=""true"" Name=""heading 3""/>  <w:LsdException Locked=""false"" Priority=""9"" SemiHidden=""true""   UnhideWhenUsed=""true"" QFormat=""true"" Name=""heading 4""/>  <w:LsdException Locked=""false"" Priority=""9"" SemiHidden=""true""   UnhideWhenUsed=""true"" QFormat=""true"" Name=""heading 5""/>  <w:LsdException Locked=""false"" Priority=""9"" SemiHidden=""true""   UnhideWhenUsed=""true"" QFormat=""true"" Name=""heading 6""/>  <w:LsdException Locked=""false"" Priority=""9"" SemiHidden=""true""   UnhideWhenUsed=""true"" QFormat=""true"" Name=""heading 7""/>  <w:LsdException Locked=""false"" Priority=""9"" SemiHidden=""true""   UnhideWhenUsed=""true"" QFormat=""true"" Name=""heading 8""/>  <w:LsdException Locked=""false"" Priority=""9"" SemiHidden=""true""   UnhideWhenUsed=""true"" QFormat=""true"" Name=""heading 9""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""index 1""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""index 2""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""index 3""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""index 4""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""index 5""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""index 6""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""index 7""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""index 8""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""index 9""/>  <w:LsdException Locked=""false"" Priority=""39"" SemiHidden=""true""   UnhideWhenUsed=""true"" Name=""toc 1""/>  <w:LsdException Locked=""false"" Priority=""39"" SemiHidden=""true""   UnhideWhenUsed=""true"" Name=""toc 2""/>  <w:LsdException Locked=""false"" Priority=""39"" SemiHidden=""true""   UnhideWhenUsed=""true"" Name=""toc 3""/>  <w:LsdException Locked=""false"" Priority=""39"" SemiHidden=""true""   UnhideWhenUsed=""true"" Name=""toc 4""/>  <w:LsdException Locked=""false"" Priority=""39"" SemiHidden=""true""   UnhideWhenUsed=""true"" Name=""toc 5""/>  <w:LsdException Locked=""false"" Priority=""39"" SemiHidden=""true""   UnhideWhenUsed=""true"" Name=""toc 6""/>  <w:LsdException Locked=""false"" Priority=""39"" SemiHidden=""true""   UnhideWhenUsed=""true"" Name=""toc 7""/>  <w:LsdException Locked=""false"" Priority=""39"" SemiHidden=""true""   UnhideWhenUsed=""true"" Name=""toc 8""/>  <w:LsdException Locked=""false"" Priority=""39"" SemiHidden=""true""   UnhideWhenUsed=""true"" Name=""toc 9""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Normal Indent""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""footnote text""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""annotation text""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""header""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""footer""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""index heading""/>  <w:LsdException Locked=""false"" Priority=""35"" SemiHidden=""true""   UnhideWhenUsed=""true"" QFormat=""true"" Name=""caption""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""table of figures""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""envelope address""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""envelope return""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""footnote reference""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""annotation reference""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""line number""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""page number""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""endnote reference""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""endnote text""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""table of authorities""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""macro""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""toa heading""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""List""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""List Bullet""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""List Number""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""List 2""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""List 3""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""List 4""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""List 5""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""List Bullet 2""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""List Bullet 3""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""List Bullet 4""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""List Bullet 5""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""List Number 2""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""List Number 3""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""List Number 4""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""List Number 5""/>  <w:LsdException Locked=""false"" Priority=""10"" QFormat=""true"" Name=""Title""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Closing""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Signature""/>  <w:LsdException Locked=""false"" Priority=""1"" SemiHidden=""true""   UnhideWhenUsed=""true"" Name=""Default Paragraph Font""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Body Text""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Body Text Indent""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""List Continue""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""List Continue 2""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""List Continue 3""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""List Continue 4""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""List Continue 5""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Message Header""/>  <w:LsdException Locked=""false"" Priority=""11"" QFormat=""true"" Name=""Subtitle""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Salutation""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Date""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Body Text First Indent""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Body Text First Indent 2""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Note Heading""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Body Text 2""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Body Text 3""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Body Text Indent 2""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Body Text Indent 3""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Block Text""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Hyperlink""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""FollowedHyperlink""/>  <w:LsdException Locked=""false"" Priority=""22"" QFormat=""true"" Name=""Strong""/>  <w:LsdException Locked=""false"" Priority=""20"" QFormat=""true"" Name=""Emphasis""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Document Map""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Plain Text""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""E-mail Signature""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""HTML Top of Form""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""HTML Bottom of Form""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Normal (Web)""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""HTML Acronym""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""HTML Address""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""HTML Cite""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""HTML Code""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""HTML Definition""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""HTML Keyboard""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""HTML Preformatted""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""HTML Sample""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""HTML Typewriter""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""HTML Variable""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Normal Table""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""annotation subject""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""No List""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Outline List 1""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Outline List 2""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Outline List 3""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Simple 1""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Simple 2""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Simple 3""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Classic 1""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Classic 2""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Classic 3""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Classic 4""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Colorful 1""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Colorful 2""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Colorful 3""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Columns 1""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Columns 2""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Columns 3""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Columns 4""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Columns 5""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Grid 1""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Grid 2""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Grid 3""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Grid 4""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Grid 5""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Grid 6""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Grid 7""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Grid 8""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table List 1""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table List 2""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table List 3""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table List 4""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table List 5""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table List 6""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table List 7""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table List 8""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table 3D effects 1""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table 3D effects 2""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table 3D effects 3""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Contemporary""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Elegant""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Professional""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Subtle 1""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Subtle 2""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Web 1""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Web 2""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Web 3""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Balloon Text""/>  <w:LsdException Locked=""false"" Priority=""39"" Name=""Table Grid""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" UnhideWhenUsed=""true""   Name=""Table Theme""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" Name=""Placeholder Text""/>  <w:LsdException Locked=""false"" Priority=""1"" QFormat=""true"" Name=""No Spacing""/>  <w:LsdException Locked=""false"" Priority=""60"" Name=""Light Shading""/>  <w:LsdException Locked=""false"" Priority=""61"" Name=""Light List""/>  <w:LsdException Locked=""false"" Priority=""62"" Name=""Light Grid""/>  <w:LsdException Locked=""false"" Priority=""63"" Name=""Medium Shading 1""/>  <w:LsdException Locked=""false"" Priority=""64"" Name=""Medium Shading 2""/>  <w:LsdException Locked=""false"" Priority=""65"" Name=""Medium List 1""/>  <w:LsdException Locked=""false"" Priority=""66"" Name=""Medium List 2""/>  <w:LsdException Locked=""false"" Priority=""67"" Name=""Medium Grid 1""/>  <w:LsdException Locked=""false"" Priority=""68"" Name=""Medium Grid 2""/>  <w:LsdException Locked=""false"" Priority=""69"" Name=""Medium Grid 3""/>  <w:LsdException Locked=""false"" Priority=""70"" Name=""Dark List""/>  <w:LsdException Locked=""false"" Priority=""71"" Name=""Colorful Shading""/>  <w:LsdException Locked=""false"" Priority=""72"" Name=""Colorful List""/>  <w:LsdException Locked=""false"" Priority=""73"" Name=""Colorful Grid""/>  <w:LsdException Locked=""false"" Priority=""60"" Name=""Light Shading Accent 1""/>  <w:LsdException Locked=""false"" Priority=""61"" Name=""Light List Accent 1""/>  <w:LsdException Locked=""false"" Priority=""62"" Name=""Light Grid Accent 1""/>  <w:LsdException Locked=""false"" Priority=""63"" Name=""Medium Shading 1 Accent 1""/>  <w:LsdException Locked=""false"" Priority=""64"" Name=""Medium Shading 2 Accent 1""/>  <w:LsdException Locked=""false"" Priority=""65"" Name=""Medium List 1 Accent 1""/>  <w:LsdException Locked=""false"" SemiHidden=""true"" Name=""Revision""/>  <w:LsdException Locked=""false"" Priority=""34"" QFormat=""true""   Name=""List Paragraph""/>  <w:LsdException Locked=""false"" Priority=""29"" QFormat=""true"" Name=""Quote""/>  <w:LsdException Locked=""false"" Priority=""30"" QFormat=""true""   Name=""Intense Quote""/>  <w:LsdException Locked=""false"" Priority=""66"" Name=""Medium List 2 Accent 1""/>  <w:LsdException Locked=""false"" Priority=""67"" Name=""Medium Grid 1 Accent 1""/>  <w:LsdException Locked=""false"" Priority=""68"" Name=""Medium Grid 2 Accent 1""/>  <w:LsdException Locked=""false"" Priority=""69"" Name=""Medium Grid 3 Accent 1""/>  <w:LsdException Locked=""false"" Priority=""70"" Name=""Dark List Accent 1""/>  <w:LsdException Locked=""false"" Priority=""71"" Name=""Colorful Shading Accent 1""/>  <w:LsdException Locked=""false"" Priority=""72"" Name=""Colorful List Accent 1""/>  <w:LsdException Locked=""false"" Priority=""73"" Name=""Colorful Grid Accent 1""/>  <w:LsdException Locked=""false"" Priority=""60"" Name=""Light Shading Accent 2""/>  <w:LsdException Locked=""false"" Priority=""61"" Name=""Light List Accent 2""/>  <w:LsdException Locked=""false"" Priority=""62"" Name=""Light Grid Accent 2""/>  <w:LsdException Locked=""false"" Priority=""63"" Name=""Medium Shading 1 Accent 2""/>  <w:LsdException Locked=""false"" Priority=""64"" Name=""Medium Shading 2 Accent 2""/>  <w:LsdException Locked=""false"" Priority=""65"" Name=""Medium List 1 Accent 2""/>  <w:LsdException Locked=""false"" Priority=""66"" Name=""Medium List 2 Accent 2""/>  <w:LsdException Locked=""false"" Priority=""67"" Name=""Medium Grid 1 Accent 2""/>  <w:LsdException Locked=""false"" Priority=""68"" Name=""Medium Grid 2 Accent 2""/>  <w:LsdException Locked=""false"" Priority=""69"" Name=""Medium Grid 3 Accent 2""/>  <w:LsdException Locked=""false"" Priority=""70"" Name=""Dark List Accent 2""/>  <w:LsdException Locked=""false"" Priority=""71"" Name=""Colorful Shading Accent 2""/>  <w:LsdException Locked=""false"" Priority=""72"" Name=""Colorful List Accent 2""/>  <w:LsdException Locked=""false"" Priority=""73"" Name=""Colorful Grid Accent 2""/>  <w:LsdException Locked=""false"" Priority=""60"" Name=""Light Shading Accent 3""/>  <w:LsdException Locked=""false"" Priority=""61"" Name=""Light List Accent 3""/>  <w:LsdException Locked=""false"" Priority=""62"" Name=""Light Grid Accent 3""/>  <w:LsdException Locked=""false"" Priority=""63"" Name=""Medium Shading 1 Accent 3""/>  <w:LsdException Locked=""false"" Priority=""64"" Name=""Medium Shading 2 Accent 3""/>  <w:LsdException Locked=""false"" Priority=""65"" Name=""Medium List 1 Accent 3""/>  <w:LsdException Locked=""false"" Priority=""66"" Name=""Medium List 2 Accent 3""/>  <w:LsdException Locked=""false"" Priority=""67"" Name=""Medium Grid 1 Accent 3""/>  <w:LsdException Locked=""false"" Priority=""68"" Name=""Medium Grid 2 Accent 3""/>  <w:LsdException Locked=""false"" Priority=""69"" Name=""Medium Grid 3 Accent 3""/>  <w:LsdException Locked=""false"" Priority=""70"" Name=""Dark List Accent 3""/>  <w:LsdException Locked=""false"" Priority=""71"" Name=""Colorful Shading Accent 3""/>  <w:LsdException Locked=""false"" Priority=""72"" Name=""Colorful List Accent 3""/>  <w:LsdException Locked=""false"" Priority=""73"" Name=""Colorful Grid Accent 3""/>  <w:LsdException Locked=""false"" Priority=""60"" Name=""Light Shading Accent 4""/>  <w:LsdException Locked=""false"" Priority=""61"" Name=""Light List Accent 4""/>  <w:LsdException Locked=""false"" Priority=""62"" Name=""Light Grid Accent 4""/>  <w:LsdException Locked=""false"" Priority=""63"" Name=""Medium Shading 1 Accent 4""/>  <w:LsdException Locked=""false"" Priority=""64"" Name=""Medium Shading 2 Accent 4""/>  <w:LsdException Locked=""false"" Priority=""65"" Name=""Medium List 1 Accent 4""/>  <w:LsdException Locked=""false"" Priority=""66"" Name=""Medium List 2 Accent 4""/>  <w:LsdException Locked=""false"" Priority=""67"" Name=""Medium Grid 1 Accent 4""/>  <w:LsdException Locked=""false"" Priority=""68"" Name=""Medium Grid 2 Accent 4""/>  <w:LsdException Locked=""false"" Priority=""69"" Name=""Medium Grid 3 Accent 4""/>  <w:LsdException Locked=""false"" Priority=""70"" Name=""Dark List Accent 4""/>  <w:LsdException Locked=""false"" Priority=""71"" Name=""Colorful Shading Accent 4""/>  <w:LsdException Locked=""false"" Priority=""72"" Name=""Colorful List Accent 4""/>  <w:LsdException Locked=""false"" Priority=""73"" Name=""Colorful Grid Accent 4""/>  <w:LsdException Locked=""false"" Priority=""60"" Name=""Light Shading Accent 5""/>  <w:LsdException Locked=""false"" Priority=""61"" Name=""Light List Accent 5""/>  <w:LsdException Locked=""false"" Priority=""62"" Name=""Light Grid Accent 5""/>  <w:LsdException Locked=""false"" Priority=""63"" Name=""Medium Shading 1 Accent 5""/>  <w:LsdException Locked=""false"" Priority=""64"" Name=""Medium Shading 2 Accent 5""/>  <w:LsdException Locked=""false"" Priority=""65"" Name=""Medium List 1 Accent 5""/>  <w:LsdException Locked=""false"" Priority=""66"" Name=""Medium List 2 Accent 5""/>  <w:LsdException Locked=""false"" Priority=""67"" Name=""Medium Grid 1 Accent 5""/>  <w:LsdException Locked=""false"" Priority=""68"" Name=""Medium Grid 2 Accent 5""/>  <w:LsdException Locked=""false"" Priority=""69"" Name=""Medium Grid 3 Accent 5""/>  <w:LsdException Locked=""false"" Priority=""70"" Name=""Dark List Accent 5""/>  <w:LsdException Locked=""false"" Priority=""71"" Name=""Colorful Shading Accent 5""/>  <w:LsdException Locked=""false"" Priority=""72"" Name=""Colorful List Accent 5""/>  <w:LsdException Locked=""false"" Priority=""73"" Name=""Colorful Grid Accent 5""/>  <w:LsdException Locked=""false"" Priority=""60"" Name=""Light Shading Accent 6""/>  <w:LsdException Locked=""false"" Priority=""61"" Name=""Light List Accent 6""/>  <w:LsdException Locked=""false"" Priority=""62"" Name=""Light Grid Accent 6""/>  <w:LsdException Locked=""false"" Priority=""63"" Name=""Medium Shading 1 Accent 6""/>  <w:LsdException Locked=""false"" Priority=""64"" Name=""Medium Shading 2 Accent 6""/>  <w:LsdException Locked=""false"" Priority=""65"" Name=""Medium List 1 Accent 6""/>  <w:LsdException Locked=""false"" Priority=""66"" Name=""Medium List 2 Accent 6""/>  <w:LsdException Locked=""false"" Priority=""67"" Name=""Medium Grid 1 Accent 6""/>  <w:LsdException Locked=""false"" Priority=""68"" Name=""Medium Grid 2 Accent 6""/>  <w:LsdException Locked=""false"" Priority=""69"" Name=""Medium Grid 3 Accent 6""/>  <w:LsdException Locked=""false"" Priority=""70"" Name=""Dark List Accent 6""/>  <w:LsdException Locked=""false"" Priority=""71"" Name=""Colorful Shading Accent 6""/>  <w:LsdException Locked=""false"" Priority=""72"" Name=""Colorful List Accent 6""/>  <w:LsdException Locked=""false"" Priority=""73"" Name=""Colorful Grid Accent 6""/>  <w:LsdException Locked=""false"" Priority=""19"" QFormat=""true""   Name=""Subtle Emphasis""/>  <w:LsdException Locked=""false"" Priority=""21"" QFormat=""true""   Name=""Intense Emphasis""/>  <w:LsdException Locked=""false"" Priority=""31"" QFormat=""true""   Name=""Subtle Reference""/>  <w:LsdException Locked=""false"" Priority=""32"" QFormat=""true""   Name=""Intense Reference""/>  <w:LsdException Locked=""false"" Priority=""33"" QFormat=""true"" Name=""Book Title""/>  <w:LsdException Locked=""false"" Priority=""37"" SemiHidden=""true""   UnhideWhenUsed=""true"" Name=""Bibliography""/>  <w:LsdException Locked=""false"" Priority=""39"" SemiHidden=""true""   UnhideWhenUsed=""true"" QFormat=""true"" Name=""TOC Heading""/>  <w:LsdException Locked=""false"" Priority=""41"" Name=""Plain Table 1""/>  <w:LsdException Locked=""false"" Priority=""42"" Name=""Plain Table 2""/>  <w:LsdException Locked=""false"" Priority=""43"" Name=""Plain Table 3""/>  <w:LsdException Locked=""false"" Priority=""44"" Name=""Plain Table 4""/>  <w:LsdException Locked=""false"" Priority=""45"" Name=""Plain Table 5""/>  <w:LsdException Locked=""false"" Priority=""40"" Name=""Grid Table Light""/>  <w:LsdException Locked=""false"" Priority=""46"" Name=""Grid Table 1 Light""/>  <w:LsdException Locked=""false"" Priority=""47"" Name=""Grid Table 2""/>  <w:LsdException Locked=""false"" Priority=""48"" Name=""Grid Table 3""/>  <w:LsdException Locked=""false"" Priority=""49"" Name=""Grid Table 4""/>  <w:LsdException Locked=""false"" Priority=""50"" Name=""Grid Table 5 Dark""/>  <w:LsdException Locked=""false"" Priority=""51"" Name=""Grid Table 6 Colorful""/>  <w:LsdException Locked=""false"" Priority=""52"" Name=""Grid Table 7 Colorful""/>  <w:LsdException Locked=""false"" Priority=""46""   Name=""Grid Table 1 Light Accent 1""/>  <w:LsdException Locked=""false"" Priority=""47"" Name=""Grid Table 2 Accent 1""/>  <w:LsdException Locked=""false"" Priority=""48"" Name=""Grid Table 3 Accent 1""/>  <w:LsdException Locked=""false"" Priority=""49"" Name=""Grid Table 4 Accent 1""/>  <w:LsdException Locked=""false"" Priority=""50"" Name=""Grid Table 5 Dark Accent 1""/>  <w:LsdException Locked=""false"" Priority=""51""   Name=""Grid Table 6 Colorful Accent 1""/>  <w:LsdException Locked=""false"" Priority=""52""   Name=""Grid Table 7 Colorful Accent 1""/>  <w:LsdException Locked=""false"" Priority=""46""   Name=""Grid Table 1 Light Accent 2""/>  <w:LsdException Locked=""false"" Priority=""47"" Name=""Grid Table 2 Accent 2""/>  <w:LsdException Locked=""false"" Priority=""48"" Name=""Grid Table 3 Accent 2""/>  <w:LsdException Locked=""false"" Priority=""49"" Name=""Grid Table 4 Accent 2""/>  <w:LsdException Locked=""false"" Priority=""50"" Name=""Grid Table 5 Dark Accent 2""/>  <w:LsdException Locked=""false"" Priority=""51""   Name=""Grid Table 6 Colorful Accent 2""/>  <w:LsdException Locked=""false"" Priority=""52""   Name=""Grid Table 7 Colorful Accent 2""/>  <w:LsdException Locked=""false"" Priority=""46""   Name=""Grid Table 1 Light Accent 3""/>  <w:LsdException Locked=""false"" Priority=""47"" Name=""Grid Table 2 Accent 3""/>  <w:LsdException Locked=""false"" Priority=""48"" Name=""Grid Table 3 Accent 3""/>  <w:LsdException Locked=""false"" Priority=""49"" Name=""Grid Table 4 Accent 3""/>  <w:LsdException Locked=""false"" Priority=""50"" Name=""Grid Table 5 Dark Accent 3""/>  <w:LsdException Locked=""false"" Priority=""51""   Name=""Grid Table 6 Colorful Accent 3""/>  <w:LsdException Locked=""false"" Priority=""52""   Name=""Grid Table 7 Colorful Accent 3""/>  <w:LsdException Locked=""false"" Priority=""46""   Name=""Grid Table 1 Light Accent 4""/>  <w:LsdException Locked=""false"" Priority=""47"" Name=""Grid Table 2 Accent 4""/>  <w:LsdException Locked=""false"" Priority=""48"" Name=""Grid Table 3 Accent 4""/>  <w:LsdException Locked=""false"" Priority=""49"" Name=""Grid Table 4 Accent 4""/>  <w:LsdException Locked=""false"" Priority=""50"" Name=""Grid Table 5 Dark Accent 4""/>  <w:LsdException Locked=""false"" Priority=""51""   Name=""Grid Table 6 Colorful Accent 4""/>  <w:LsdException Locked=""false"" Priority=""52""   Name=""Grid Table 7 Colorful Accent 4""/>  <w:LsdException Locked=""false"" Priority=""46""   Name=""Grid Table 1 Light Accent 5""/>  <w:LsdException Locked=""false"" Priority=""47"" Name=""Grid Table 2 Accent 5""/>  <w:LsdException Locked=""false"" Priority=""48"" Name=""Grid Table 3 Accent 5""/>  <w:LsdException Locked=""false"" Priority=""49"" Name=""Grid Table 4 Accent 5""/>  <w:LsdException Locked=""false"" Priority=""50"" Name=""Grid Table 5 Dark Accent 5""/>  <w:LsdException Locked=""false"" Priority=""51""   Name=""Grid Table 6 Colorful Accent 5""/>  <w:LsdException Locked=""false"" Priority=""52""   Name=""Grid Table 7 Colorful Accent 5""/>  <w:LsdException Locked=""false"" Priority=""46""   Name=""Grid Table 1 Light Accent 6""/>  <w:LsdException Locked=""false"" Priority=""47"" Name=""Grid Table 2 Accent 6""/>  <w:LsdException Locked=""false"" Priority=""48"" Name=""Grid Table 3 Accent 6""/>  <w:LsdException Locked=""false"" Priority=""49"" Name=""Grid Table 4 Accent 6""/>  <w:LsdException Locked=""false"" Priority=""50"" Name=""Grid Table 5 Dark Accent 6""/>  <w:LsdException Locked=""false"" Priority=""51""   Name=""Grid Table 6 Colorful Accent 6""/>  <w:LsdException Locked=""false"" Priority=""52""   Name=""Grid Table 7 Colorful Accent 6""/>  <w:LsdException Locked=""false"" Priority=""46"" Name=""List Table 1 Light""/>  <w:LsdException Locked=""false"" Priority=""47"" Name=""List Table 2""/>  <w:LsdException Locked=""false"" Priority=""48"" Name=""List Table 3""/>  <w:LsdException Locked=""false"" Priority=""49"" Name=""List Table 4""/>  <w:LsdException Locked=""false"" Priority=""50"" Name=""List Table 5 Dark""/>  <w:LsdException Locked=""false"" Priority=""51"" Name=""List Table 6 Colorful""/>  <w:LsdException Locked=""false"" Priority=""52"" Name=""List Table 7 Colorful""/>  <w:LsdException Locked=""false"" Priority=""46""   Name=""List Table 1 Light Accent 1""/>  <w:LsdException Locked=""false"" Priority=""47"" Name=""List Table 2 Accent 1""/>  <w:LsdException Locked=""false"" Priority=""48"" Name=""List Table 3 Accent 1""/>  <w:LsdException Locked=""false"" Priority=""49"" Name=""List Table 4 Accent 1""/>  <w:LsdException Locked=""false"" Priority=""50"" Name=""List Table 5 Dark Accent 1""/>  <w:LsdException Locked=""false"" Priority=""51""   Name=""List Table 6 Colorful Accent 1""/>  <w:LsdException Locked=""false"" Priority=""52""   Name=""List Table 7 Colorful Accent 1""/>  <w:LsdException Locked=""false"" Priority=""46""   Name=""List Table 1 Light Accent 2""/>  <w:LsdException Locked=""false"" Priority=""47"" Name=""List Table 2 Accent 2""/>  <w:LsdException Locked=""false"" Priority=""48"" Name=""List Table 3 Accent 2""/>  <w:LsdException Locked=""false"" Priority=""49"" Name=""List Table 4 Accent 2""/>  <w:LsdException Locked=""false"" Priority=""50"" Name=""List Table 5 Dark Accent 2""/>  <w:LsdException Locked=""false"" Priority=""51""   Name=""List Table 6 Colorful Accent 2""/>  <w:LsdException Locked=""false"" Priority=""52""   Name=""List Table 7 Colorful Accent 2""/>  <w:LsdException Locked=""false"" Priority=""46""   Name=""List Table 1 Light Accent 3""/>  <w:LsdException Locked=""false"" Priority=""47"" Name=""List Table 2 Accent 3""/>  <w:LsdException Locked=""false"" Priority=""48"" Name=""List Table 3 Accent 3""/>  <w:LsdException Locked=""false"" Priority=""49"" Name=""List Table 4 Accent 3""/>  <w:LsdException Locked=""false"" Priority=""50"" Name=""List Table 5 Dark Accent 3""/>  <w:LsdException Locked=""false"" Priority=""51""   Name=""List Table 6 Colorful Accent 3""/>  <w:LsdException Locked=""false"" Priority=""52""   Name=""List Table 7 Colorful Accent 3""/>  <w:LsdException Locked=""false"" Priority=""46""   Name=""List Table 1 Light Accent 4""/>  <w:LsdException Locked=""false"" Priority=""47"" Name=""List Table 2 Accent 4""/>  <w:LsdException Locked=""false"" Priority=""48"" Name=""List Table 3 Accent 4""/>  <w:LsdException Locked=""false"" Priority=""49"" Name=""List Table 4 Accent 4""/>  <w:LsdException Locked=""false"" Priority=""50"" Name=""List Table 5 Dark Accent 4""/>  <w:LsdException Locked=""false"" Priority=""51""   Name=""List Table 6 Colorful Accent 4""/>  <w:LsdException Locked=""false"" Priority=""52""   Name=""List Table 7 Colorful Accent 4""/>  <w:LsdException Locked=""false"" Priority=""46""   Name=""List Table 1 Light Accent 5""/>  <w:LsdException Locked=""false"" Priority=""47"" Name=""List Table 2 Accent 5""/>  <w:LsdException Locked=""false"" Priority=""48"" Name=""List Table 3 Accent 5""/>  <w:LsdException Locked=""false"" Priority=""49"" Name=""List Table 4 Accent 5""/>  <w:LsdException Locked=""false"" Priority=""50"" Name=""List Table 5 Dark Accent 5""/>  <w:LsdException Locked=""false"" Priority=""51""   Name=""List Table 6 Colorful Accent 5""/>  <w:LsdException Locked=""false"" Priority=""52""   Name=""List Table 7 Colorful Accent 5""/>  <w:LsdException Locked=""false"" Priority=""46""   Name=""List Table 1 Light Accent 6""/>  <w:LsdException Locked=""false"" Priority=""47"" Name=""List Table 2 Accent 6""/>  <w:LsdException Locked=""false"" Priority=""48"" Name=""List Table 3 Accent 6""/>  <w:LsdException Locked=""false"" Priority=""49"" Name=""List Table 4 Accent 6""/>  <w:LsdException Locked=""false"" Priority=""50"" Name=""List Table 5 Dark Accent 6""/>  <w:LsdException Locked=""false"" Priority=""51""   Name=""List Table 6 Colorful Accent 6""/>  <w:LsdException Locked=""false"" Priority=""52""   Name=""List Table 7 Colorful Accent 6""/> </w:LatentStyles></xml><![endif]--><style><!-- /* Font Definitions */ @font-face	{font-family:Helvetica;	panose-1:2 11 6 4 2 2 2 2 2 4;	mso-font-charset:204;	mso-generic-font-family:swiss;	mso-font-pitch:variable;	mso-font-signature:-536859905 -1073711037 9 0 511 0;}@font-face	{font-family:""Cambria Math"";	panose-1:2 4 5 3 5 4 6 3 2 4;	mso-font-charset:204;	mso-generic-font-family:roman;	mso-font-pitch:variable;	mso-font-signature:-536870145 1107305727 0 0 415 0;}@font-face	{font-family:Calibri;	panose-1:2 15 5 2 2 2 4 3 2 4;	mso-font-charset:204;	mso-generic-font-family:swiss;	mso-font-pitch:variable;	mso-font-signature:-536870145 1073786111 1 0 415 0;} /* Style Definitions */ p.MsoNormal, li.MsoNormal, div.MsoNormal	{mso-style-unhide:no;	mso-style-qformat:yes;	mso-style-parent:"""";	margin:0cm;	margin-bottom:.0001pt;	mso-pagination:widow-orphan;	font-size:11.0pt;	font-family:""Calibri"",sans-serif;	mso-ascii-font-family:Calibri;	mso-ascii-theme-font:minor-latin;	mso-fareast-font-family:""Times New Roman"";	mso-fareast-theme-font:minor-fareast;	mso-hansi-font-family:Calibri;	mso-hansi-theme-font:minor-latin;	mso-bidi-font-family:""Times New Roman"";	mso-bidi-theme-font:minor-bidi;}a:link, span.MsoHyperlink	{mso-style-noshow:yes;	mso-style-priority:99;	mso-style-parent:"""";	color:blue;	text-decoration:underline;	text-underline:single;}a:visited, span.MsoHyperlinkFollowed	{mso-style-noshow:yes;	mso-style-priority:99;	color:#954F72;	mso-themecolor:followedhyperlink;	text-decoration:underline;	text-underline:single;}p.MsoAutoSig, li.MsoAutoSig, div.MsoAutoSig	{mso-style-noshow:yes;	mso-style-priority:99;	mso-style-link:""Электронная подпись Знак"";	margin:0cm;	margin-bottom:.0001pt;	mso-pagination:widow-orphan;	font-size:11.0pt;	font-family:""Calibri"",sans-serif;	mso-ascii-font-family:Calibri;	mso-ascii-theme-font:minor-latin;	mso-fareast-font-family:""Times New Roman"";	mso-fareast-theme-font:minor-fareast;	mso-hansi-font-family:Calibri;	mso-hansi-theme-font:minor-latin;	mso-bidi-font-family:""Times New Roman"";	mso-bidi-theme-font:minor-bidi;}span.a	{mso-style-name:""Электронная подпись Знак"";	mso-style-noshow:yes;	mso-style-priority:99;	mso-style-unhide:no;	mso-style-locked:yes;	mso-style-link:""Электронная подпись"";}.MsoChpDefault	{mso-style-type:export-only;	mso-default-props:yes;	font-size:11.0pt;	mso-ansi-font-size:11.0pt;	mso-bidi-font-size:11.0pt;	mso-ascii-font-family:Calibri;	mso-ascii-theme-font:minor-latin;	mso-fareast-font-family:""Times New Roman"";	mso-fareast-theme-font:minor-fareast;	mso-hansi-font-family:Calibri;	mso-hansi-theme-font:minor-latin;	mso-bidi-font-family:""Times New Roman"";	mso-bidi-theme-font:minor-bidi;}@page WordSection1	{size:612.0pt 792.0pt;	margin:2.0cm 42.5pt 2.0cm 3.0cm;	mso-header-margin:36.0pt;	mso-footer-margin:36.0pt;	mso-paper-source:0;}div.WordSection1	{page:WordSection1;}--></style><!--[if gte mso 10]><style> /* Style Definitions */ table.MsoNormalTable	{mso-style-name:""Обычная таблица"";	mso-tstyle-rowband-size:0;	mso-tstyle-colband-size:0;	mso-style-noshow:yes;	mso-style-priority:99;	mso-style-parent:"""";	mso-padding-alt:0cm 5.4pt 0cm 5.4pt;	mso-para-margin:0cm;	mso-para-margin-bottom:.0001pt;	mso-pagination:widow-orphan;	font-size:11.0pt;	font-family:""Calibri"",sans-serif;	mso-ascii-font-family:Calibri;	mso-ascii-theme-font:minor-latin;	mso-hansi-font-family:Calibri;	mso-hansi-theme-font:minor-latin;}</style><![endif]--></head><body lang=RU link=blue vlink=""#954F72"" style='tab-interval:35.4pt'><div class=WordSection1><table class=MsoNormalTable border=0 cellpadding=0 width=350 style='width:262.5pt; mso-cellspacing:1.5pt;mso-yfti-tbllook:1184;mso-padding-alt:0cm 0cm 0cm 0cm'> <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>   <td width=""90px""><p style=""text-align: center; vertical-align: central;""><a href=""http://www.syntegra.com.ua""><img src=""https://syntegra-public.sharepoint.com/Lists/Photos/"&strmail&".jpg"" alt=""SYNTEGRA"" height=""80"" width=""80""></a></p></td><td style='padding:0cm 0cm 0cm 0cm'>  <p class=MsoNormal style='line-height:9.0pt'><b><span lang=EN-US  style='font-size:7.5pt;font-family:""Helvetica"",sans-serif;mso-fareast-font-family:  Calibri;color:#212121;mso-ansi-language:EN-US;mso-fareast-language:EN-US'>"&strdisplayname&" </span></b><span lang=EN-US style='font-size:7.5pt;font-family:  ""Helvetica"",sans-serif;mso-fareast-font-family:Calibri;color:#212121;  mso-ansi-language:EN-US;mso-fareast-language:EN-US'>| "&strtitle&"<o:p></o:p></span></p>  <p class=MsoNormal style='line-height:9.0pt'><span lang=EN-US  style='font-size:7.5pt;font-family:""Helvetica"",sans-serif;mso-fareast-font-family:  Calibri;color:#212121;mso-ansi-language:EN-US;mso-fareast-language:EN-US'><a  href=""mailto:"&strmail&""">"&strmail&"</a> | "&strmobile&"<o:p></o:p></span></p>  <p class=MsoNormal style='mso-margin-top-alt:auto;line-height:9.0pt'><b><span  lang=EN-US style='font-size:7.5pt;font-family:""Helvetica"",sans-serif;  mso-fareast-font-family:Calibri;mso-ansi-language:EN-US;mso-fareast-language:  EN-US'>"&prmOOO_Name&"</span></b><span lang=EN-US style='font-size:7.5pt;font-family:  ""Helvetica"",sans-serif;mso-fareast-font-family:Calibri;mso-ansi-language:  EN-US;mso-fareast-language:EN-US'> <span style='display:none;mso-hide:all'>Office:  </span>+380 44 3647779 <span style='display:none;mso-hide:all'><span  style='mso-spacerun:yes'> </span></span>3 building 8b, Surikova str., Kiev,  Ukraine <o:p></o:p></span></p>  <p class=MsoNormal style='mso-margin-bottom-alt:auto;line-height:200%'><span  lang=EN-US style='font-size:7.5pt;line-height:200%;font-family:""Helvetica"",sans-serif;  mso-fareast-font-family:Calibri;mso-ansi-language:EN-US;mso-fareast-language:  EN-US'><a href=""http://www.syntegra.com.ua/"">http://www.syntegra.com.ua</a> <o:p></o:p></span></p>  <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;  mso-line-height-alt:9.0pt'><a href=""https://www.facebook.com/uasyntegra""><span  style='font-size:7.5pt;font-family:""Helvetica"",sans-serif;mso-fareast-font-family:  Calibri;mso-no-proof:yes;text-decoration:none;text-underline:none'><!--[if gte vml 1]><v:shape   id=""Рисунок_x0020_3"" o:spid=""_x0000_i1026"" type=""#_x0000_t75"" alt=""Facebook""   href=""https://www.facebook.com/uasyntegra"" style='width:12pt;height:12pt;   visibility:visible;mso-wrap-style:square' o:button=""t"">   <v:fill o:detectmouseclick=""t""/>   <v:imagedata src=""https://syntegra-public.sharepoint.com/Lists/Photos/facebook.png"" o:title=""Facebook""/>  </v:shape><![endif]--><![if !vml]><span style='mso-ignore:vglayout'><img  border=0 width=16 height=16 src=""https://syntegra-public.sharepoint.com/Lists/Photos/facebook.png"" alt=Facebook v:shapes=""Рисунок_x0020_3""></span><![endif]></span></a><a  href=""https://www.linkedin.com/company/syntegra-llc""><span style='font-size:  7.5pt;font-family:""Helvetica"",sans-serif;mso-fareast-font-family:Calibri;  mso-no-proof:yes;text-decoration:none;text-underline:none'><!--[if gte vml 1]><v:shape   id=""Рисунок_x0020_2"" o:spid=""_x0000_i1027"" type=""#_x0000_t75"" alt=""Linkedin""   href=""https://www.linkedin.com/company/syntegra-llc"" style='width:12pt;   height:12pt;visibility:visible;mso-wrap-style:square' o:button=""t"">   <v:fill o:detectmouseclick=""t""/>   <v:imagedata src=""https://syntegra-public.sharepoint.com/Lists/Photos/linkedin.png"" o:title=""Linkedin""/>  </v:shape><![endif]--><![if !vml]><span style='mso-ignore:vglayout'><img  border=0 width=16 height=16 src=""https://syntegra-public.sharepoint.com/Lists/Photos/linkedin.png"" alt=Linkedin v:shapes=""Рисунок_x0020_2""></span><![endif]></span></a><span  style='font-size:7.5pt;font-family:""Helvetica"",sans-serif;mso-fareast-font-family:  Calibri;mso-fareast-language:EN-US'><o:p></o:p></span></p>  </td> </tr> <tr style='mso-yfti-irow:1'>  <td colspan=2 style='padding:0cm 0cm 0cm 0cm'></td> </tr> <tr style='mso-yfti-irow:2;mso-yfti-lastrow:yes'>  <td colspan=2 style='padding:0cm 0cm 0cm 0cm'>  <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;  line-height:9.0pt'><span lang=EN-US style='font-size:7.0pt;font-family:""Helvetica"",sans-serif;  mso-fareast-font-family:Calibri;mso-ansi-language:EN-US;mso-fareast-language:  EN-US'>This e-mail message may contain confidential or legally privileged  information and is intended only for the use of the intended recipient(s). Any  unauthorized disclosure, dissemination, distribution, copying or the taking  of any action in reliance on the information herein is prohibited. E-mails  are not secure and cannot be guaranteed to be error free as they can be  intercepted, amended, or contain viruses. Anyone who communicates with us by  e-mail is deemed to have accepted these risks. Company Name is not  responsible for errors or omissions in this message and denies any  responsibility for any damage arising from the use of e-mail. Any opinion and  other statement contained in this message and any attachment are solely those  of the author and do not necessarily represent those of the company.<o:p></o:p></span></p>  </td> </tr></table><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p></div></body></html>"
            
 
 
' Подписи лежат в %APPDATA%\Microsoft\Signatures. Но если до этого никаких подписей не создавалось -
' этой папки может и не быть. Поэтому нужно создать.
If Not fso.FolderExists(Folder & "\Microsoft") Then
fso.CreateFolder(Folder & "\Microsoft")
End If
Folder = Folder & "\Microsoft"
 
If Not fso.FolderExists(Folder & "\Signatures") Then
fso.CreateFolder(Folder & "\Signatures")
End If
Folder = Folder & "\Signatures\"
 
' Удаляем все подписи из этой папки, в том числе и юзерские.
ClearFolder(Folder)
 
' Пишем подпись в текстовый файл.
Set ts = fso.OpenTextFile(Folder + "sev.txt", 2, True)
ts.WriteLine(Signature)
ts.Close
 
' Ставим аттрибут "только чтение", чтобы юзер сам её не отредактировал.
Set ts = fso.GetFile(Folder + "sev.txt")
ts.Attributes = 1
 
' Копируем ещё с тремя именами. Вообще оутлук перечисляет только файлы .txt, но на всякий случай.
fso.CopyFile Folder + "sev.txt", Folder + "sev.htm", OverwriteExistring
fso.CopyFile Folder + "sev.txt", Folder + "sev.rtf", OverwriteExistring
fso.CopyFile Folder + "sev.txt", Folder + "sev.html", OverwriteExistring
' Кстати, поскольку я использую только текстовые подписи, html у меня кривой получается.
' Туда неплохо бы добавить хотя бы теги <br>. Но мне это не надо.
 
' Теперь нам нужно понять, с какой версией офиса мы работаем. Кое-где стоят одновременно несколько
' версий, поэтому перебрать нужно все. К счастью, названия ключей реестра не менялись, поэтому
' достаточно просто перебрать номера версий.
Key1 = "HKEY_CURRENT_USER\Software\Microsoft\Office\"
Key2 = ".0\Outlook\Options\"
for i = 5 to 15  
        if RegistryKeyExists (Key1 & i & Key2 ) <> 0 then
                'Текстовый формат сообщения по умолчанию
                WshShell.RegWrite Key1 & i & Key2 & "Mail\EditorPreference", "65536", "REG_DWORD"                      
                'Читать все письма как текст
                WshShell.RegWrite Key1 & i & Key2 & "Mail\ReadAsPlain", "0", "REG_DWORD"                       
                'В том числе, и подписанные цифровой подписью.
                WshShell.RegWrite Key1 & i & Key2 & "Mail\ReadSignedAsPlain", "1", "REG_DWORD"                 
         End If
next
 
' Перечисляем все учетки и исправляем в них имена и дефолтные подписи
' Профили оутлука лежат здесь:
strKeyPath = "Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
 
' Нужно перечислить субключи реестра, здесь нужно немножко уличной магии
const HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
oReg.EnumKey HKEY_CURRENT_USER, strKeyPath, ProfileList
 
' Если профилей нет - обидно, идём лесом.
If IsNull(ProfileList) then
        Wscript.Quit
End If
 
' А вот если они есть - то нужно перебрать их все, вытащить из них учетные
' записи почты и навести в них "жыстачайшый парадак" (с)
For Each Profile in ProfileList
        ' И вновь уличная магия. Перечисляем субключи в профиле
        Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
        ' 9375CFF0413111d3B88A00104B2A6676 - это имя субключа, в который пишет и читает Оутлук.
        oReg.EnumKey HKEY_CURRENT_USER, strKeyPath & "\" & Profile &  "\9375CFF0413111d3B88A00104B2A6676", arrSubKeys
        ' Если в этом ключе что-то есть, тогда это всё нужно перебрать
        if not IsNull(arrSubKeys) then
                For Each subkey In arrSubKeys
                        keytext = "HKEY_CURRENT_USER\" & strKeyPath & "\" & Profile &  "\9375CFF0413111d3B88A00104B2A6676\" &  subkey & "\"
                        ' Если в этом ключе есть значение "Email" - это почтовый аккаунт! Начинаем исправлять
                                if KeyExists (keytext & "Email") then
                                ' Вообще там значения в юникоде написаны как REG_BINARY. Но и reg_sz прокатывает со свистом, если только англ. символы.
                                ' Имя пользователя
                                        WshShell.RegWrite keytext & "Display Name", strDisplayName , "REG_SZ"          
                                ' Используем нашу подпись для новых писем
                                        WshShell.RegWrite keytext & "New Signature", "sev", "REG_SZ"           
                                ' Используем нашу подпись для ответов на письма и форварда.
                                        WshShell.RegWrite keytext & "Reply-Forward Signature", "sev", "REG_SZ"         
                                end if
                Next
        End If
Next
Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Add()
Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature
objSignatureObject.NewMessageSignature = "sev"
objSignatureObject.ReplyMessageSignature = "sev"
objDoc.Saved = True
objDoc.Close
objWord.Quit
' Всё 
' Ну типа всё.