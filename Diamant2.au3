#include <GUIConstantsEx.au3>
#include <IE.au3>
#include <Excel.au3>
#include <WindowsConstants.au3>


$x=InputBox("С какого пользователя начинать?","С какого пользователя начинать?")

$qual=InputBox("Введите количество пользователей","Введите количество пользователей")


$newpass = "P@ssw0rd1";InputBox("Пароль","Введите новый пароль по умолчанию")





While $x < $qual
$login = FileReadLine (@ScriptDir&"\users.txt", $x)
$first = FileReadLine (@ScriptDir&"\first.txt", $x)
$second = FileReadLine (@ScriptDir&"\second.txt", $x)
ConsoleWrite ($x&" "&$login&@CRLF)
;$login = "User"&$x&"@"&$domain
;run ("RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2")
RunWait ("RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2")

$oIE = _IECreate("https://portal.microsoftonline.com")
_IEAttach("Вход в Office")


$oForm = _IEFormGetCollection ($oIE, 0)
$oInput =_IEGetObjByName($oIE,"login")
$pInput =_IEGetObjByName($oIE,"passwd")
_IEFormElementSetValue($oInput, $login)
_IEFormElementSetValue($pInput, $newpass)

Sleep(2000)
_IEFormSubmit($oForm, 0)

_IELoadWait ($oIE)

_IENavigate ( $oIE, "https://www.yammer.com/office365" )

$oForm = _IEFormGetCollection ($oIE, 1)

$Inputfirstname =_IEGetObjByName($oIE,"user[first_name]")
$Inputsecondname =_IEGetObjByName($oIE,"user[last_name]")

_IEFormElementSetValue($Inputfirstname, $first)
_IEFormElementSetValue($Inputsecondname, $second)

Sleep(1000)

Local $hWnd = _IEPropertyGet($oIE, "hwnd")
ControlSend($hWnd, "", "[CLASS:Internet Explorer_Server; INSTANCE:1]", "{Enter}")

_IELoadWait ($oIE)

_IENavigate ($oIE, "https://www.yammer.com/office365" )

_IELoadWait ($oIE)

Sleep(1000)

ControlSend($hWnd, "", "[CLASS:Internet Explorer_Server; INSTANCE:1]", "{ESCAPE} ")

If WinActive($hWnd) Then
Else
   WinActivate ($hWnd)
EndIf

Sleep(1000)

$oPost = _IEGetObjById($oIE, "make_a_post")
; $href = $oLink.href
_IEAction($oPost, "click")

Sleep (15000)

Send ("Hi")

SLEEP (2000)
Send("{LSHIFT down}")
Sleep (30)
Send("{ENTER down}")
Sleep (30)
Send("{LSHIFT up}")
Sleep (30)
Send("{ENTER up}")

Sleep(2000)
_IEQuit ($oIE)
Sleep(2000)

$x=$x+1
WEnd

Exit