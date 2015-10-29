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
        strProperties = "DisplayName, mail, telephoneNumber, title, mobile;"
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
       
        strmail = objRecordSet.Fields("mail").Value & vbcrlf
       
        strtelephoneNumber = objRecordSet.Fields("telephoneNumber").Value
       
        if (strtelephoneNumber <> "") then strtelephoneNumber = strtelephoneNumber & vbcrlf
        strtitle = objRecordSet.Fields("title").Value
       
        if (strtitle <> "") then
                strtitle = strtitle & " " & prmOOO_Name & vbcrlf
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
Signature = "_____________________________________" & "<br>"  & "С уважением, " & "<br>"  &  strDisplayName  & "<br>"  & strtitle & "<br>" & strtelephoneNumber & strmobile & "<br>"&  strmail &"<br>" &  prmSite
            
 
 ' Подписи лежат в %APPDATA%\Microsoft\Signatures. Но если до этого никаких подписей не создавалось —
' этой папки может и не быть. Поэтому нужно создать.
If Not fso.FolderExists(Folder & "\Microsoft") Then
fso.CreateFolder(Folder & "\Microsoft")
End If
Folder = Folder & "\Microsoft"
If Not fso.FolderExists(Folder & "\Signatures") Then
fso.CreateFolder(Folder & "\Signatures")
End If
Folder = Folder & "\Signatures\"
' Удаляем все подписи из этой папки, в том числе и пользовательские.
ClearFolder(Folder)
' Пишем подпись в текстовый файл.
Set ts = fso.OpenTextFile(Folder + "AutoSign.txt", 2, True)
ts.WriteLine(Signature)
ts.Close
Set ts = fso.OpenTextFile(Folder + "AutoSign.htm", 2, True)
ts.WriteLine(Signature_html)
ts.Close
Set ts = fso.OpenTextFile(Folder + "AutoSign.rtf", 2, True)
ts.WriteLine(Signature_rtf)
ts.Close
' Ставим аттрибут "только чтение", чтобы юзер сам её не отредактировал.
Set ts = fso.GetFile(Folder + "AutoSign.txt")
Set ts = fso.GetFile(Folder + "AutoSign.htm")
Set ts = fso.GetFile(Folder + "AutoSign.rtf")
ts.Attributes = 1
' Копируем ещё с тремя именами. Вообще оутлук перечисляет только файлы .txt, но на всякий случай.
fso.CopyFile Folder + "AutoSign.htm", Folder + "AutoSign.html", OverwriteExistring
' Теперь нам нужно понять, с какой версией офиса мы работаем. Кое-где стоят одновременно несколько
' версий, поэтому перебрать нужно все. К счастью, названия ключей реестра не менялись, поэтому
' достаточно просто перебрать номера версий.
Key1 = "HKEY_CURRENT_USER\Software\Microsoft\Office\"
Key2 = ".0\Outlook\Options\"
for i = 5 to 17
if RegistryKeyExists (Key1 & i & Key2 ) <> 0 then
'Текстовый формат сообщения по умолчанию
WshShell.RegWrite Key1 & i & Key2 & "Mail\EditorPreference", "65536", "REG_DWORD"
'Читать все письма в формате отправителя (Если пользователь сам ставил галочку, чтобы письма присылались как тескт, то она снимится)
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
' Если профилей нет — обидно, идём лесом.
If IsNull(ProfileList) then
Wscript.Quit
End If
' А вот если они есть — то нужно перебрать их все, вытащить из них учетные
' записи почты и навести в них порядок
For Each Profile in ProfileList
' И вновь уличная магия. Перечисляем субключи в профиле
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
' 9375CFF0413111d3B88A00104B2A6676 — это имя субключа, в который пишет и читает Оутлук.
oReg.EnumKey HKEY_CURRENT_USER, strKeyPath & "\" & Profile & "\9375CFF0413111d3B88A00104B2A6676", arrSubKeys
' Если в этом ключе что-то есть, тогда это всё нужно перебрать
if not IsNull(arrSubKeys) then
For Each subkey In arrSubKeys
keytext = "HKEY_CURRENT_USER\" & strKeyPath & "\" & Profile & "\9375CFF0413111d3B88A00104B2A6676\" & subkey & "\"
' Если в этом ключе есть значение "Email" — это почтовый аккаунт! Начинаем исправлять
if KeyExists (keytext & "Email") then
' Вообще там значения в юникоде написаны как REG_BINARY. Но и reg_sz прокатывает со свистом, если только англ. символы.
' Имя пользователя
WshShell.RegWrite keytext & "Display Name", strDisplayName , "REG_SZ"
' Используем нашу подпись для новых писем
WshShell.RegWrite keytext & "New Signature", "AutoSign", "REG_SZ"
' Используем нашу подпись для ответов на письма и форварда.
WshShell.RegWrite keytext & "Reply-Forward Signature", "AutoSign", "REG_SZ"
end if
Next
End If
Next
' Ставим подпись по дефолту
Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Add()
Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature
objSignatureObject.NewMessageSignature = "AutoSign"
objSignatureObject.ReplyMessageSignature = "AutoSign"
objDoc.Saved = True
objDoc.Close
objWord.Quit
' Всё