# 偽SMTPとして、メールを待ち受け、ファイルに出力します。
$port = 8587

# 何通受け取ったら終了するかを設定します。
$count = 1
# 設定数受け取らなくても、smtp.stop というファイルが存在したら、終了します。
# それまでは無限ループします。

function change_mime($encodedString) {
    # エンコーディングとエンコードされた部分を取得
    try {
        $match = [Regex]::Match($encodedString, "=\?([^?]+)\?B\?([^?]+)\?=")
        $encodingName = $match.Groups[1].Value
        $encodedPart = $match.Groups[2].Value
        
        # Base64デコード
        $decodedBytes = [System.Convert]::FromBase64String($encodedPart)
        
        # バイト配列を文字列に変換
        $encoding = [System.Text.Encoding]::GetEncoding($encodingName)
        $decodedString = $encoding.GetString($decodedBytes)
        
        # 結果
        $decodedString
    }
    catch {
        $encodedString    
    }    
}

function change_base64($base64String) {
    # Base64デコード
    $decodedBytes = [System.Convert]::FromBase64String($base64String)

    # バイト配列をShift-JISエンコーディングの文字列に変換
    $shiftJis = [System.Text.Encoding]::GetEncoding("utf-8")
    $decodedString = $shiftJis.GetString($decodedBytes)

    # 結果
    $decodedString
}

try {

    #    [System.Net.IPAddress]$IPAddress = "192.168.11.35"
    $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Any, $port)

    # リスナーを開始します。
    $listener.Start()

    while ($count -ne 0) {
        if (Test-Path -Path "smtp.stop") {
            break
        }
        # 接続要求があるか確認
        if (-not $listener.Pending()) {
            # 接続要求がない場合、少し待機
            Start-Sleep -Milliseconds 500
            continue
        }
        # 接続要求がある場合、

        # クライアントからの接続を待ちます。
        $client = $listener.AcceptTcpClient()
        $stream = $client.GetStream()
        $reader = New-Object System.IO.StreamReader($stream)
        $writer = New-Object System.IO.StreamWriter($stream)

        # SMTPセッションを開始します。
        $writer.WriteLine("220 PowerShell SMTP Server Ready")
        $writer.Flush()

        $mail = ""
        while ($null -ne ($line = $reader.ReadLine())) {
            # SMTPコマンドを解析します。
            if ($line.StartsWith("HELO") -or $line.StartsWith("EHLO")) {
                $writer.WriteLine("250 Hello")
            }
            elseif ($line.StartsWith("MAIL FROM:")) {
                $writer.WriteLine("250 OK")
            }
            elseif ($line.StartsWith("RCPT TO:")) {
                # ToとCCの両方を出力します。
                $mail += "Recipient: " + $line.Substring(8) + "`r`n"
                $writer.WriteLine("250 OK")
            }
            elseif ($line.StartsWith("DATA")) {
                $writer.WriteLine("354 Start mail input")
                $writer.Flush()

                # メールの内容をファイルに保存します。
                $subtext = ""
                $base64mode = $false
                $base64text = ""
                while (($line = $reader.ReadLine()) -ne ".") {
                    if ($base64mode) {
                        $base64text += $line
                    }
                    else {
                        if ($line.StartsWith("Subject:")) {
                            # ToとCCの両方を出力します。
                            $subtext = change_mime($line.Substring(9))
                            $mail += "Subject:" + $subtext + "`r`n" 
                        }
                        elseif ($line.StartsWith("Content-Transfer-Encoding: base64")) {
                            $base64mode = $true
                        }
                        else {
                            $mail += $line + "`r`n"
                        }
                    }
                }
                if ($base64mode) {
                    $mail += change_base64($base64text) + "`r`n"
                }
                $fname = "mail_" + $subtext + "_" + (Get-Date).ToString('yyyyMMddHHmmssfff') + ".txt"
                # $fname = "mail_"+(Get-Date).ToString('yyyyMMddHHmmssfff')+".txt"
                $mail | Out-File -FilePath $fname

                $writer.WriteLine("250 OK")
            }
            elseif ($line.StartsWith("QUIT")) {
                $writer.WriteLine("221 Bye")
                $writer.Flush()
                $count--
                break
            }
            else {
                $writer.WriteLine("500 Error: command not recognized")
            }
            $writer.Flush()
        }

        # クライアントとの接続を閉じます。
        $client.Close()
    }

}
finally {

    $listener.stop()

}
