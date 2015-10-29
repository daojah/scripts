Start-Process PowerShell.exe -Verb runAs -windowstyle hidden { Get-ChildItem "C:\Users\*\AppData\Local\1C\1Cv8\*","C:\Users\*\AppData\Roaming\1C\1Cv8\*" | Where {$_.Name -as [guid]} |Remove-Item -Force -Recurse }
#ForEach-Object
