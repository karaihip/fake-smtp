# �USMTP�Ƃ��āA���[����҂��󂯁A�t�@�C���ɏo�͂��܂��B
$port = 8587

# ���ʎ󂯎������I�����邩��ݒ肵�܂��B
$count = 1
# �ݒ萔�󂯎��Ȃ��Ă��Asmtp.stop �Ƃ����t�@�C�������݂�����A�I�����܂��B
# ����܂ł͖������[�v���܂��B

function change_mime($encodedString) {
    # �G���R�[�f�B���O�ƃG���R�[�h���ꂽ�������擾
    try {
        $match = [Regex]::Match($encodedString, "=\?([^?]+)\?B\?([^?]+)\?=")
        $encodingName = $match.Groups[1].Value
        $encodedPart = $match.Groups[2].Value
        
        # Base64�f�R�[�h
        $decodedBytes = [System.Convert]::FromBase64String($encodedPart)
        
        # �o�C�g�z��𕶎���ɕϊ�
        $encoding = [System.Text.Encoding]::GetEncoding($encodingName)
        $decodedString = $encoding.GetString($decodedBytes)
        
        # ����
        $decodedString
    }
    catch {
        $encodedString    
    }    
}

function change_base64($base64String) {
    # Base64�f�R�[�h
    $decodedBytes = [System.Convert]::FromBase64String($base64String)

    # �o�C�g�z���Shift-JIS�G���R�[�f�B���O�̕�����ɕϊ�
    $shiftJis = [System.Text.Encoding]::GetEncoding("utf-8")
    $decodedString = $shiftJis.GetString($decodedBytes)

    # ����
    $decodedString
}

try {

    #    [System.Net.IPAddress]$IPAddress = "192.168.11.35"
    $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Any, $port)

    # ���X�i�[���J�n���܂��B
    $listener.Start()

    while ($count -ne 0) {
        if (Test-Path -Path "smtp.stop") {
            break
        }
        # �ڑ��v�������邩�m�F
        if (-not $listener.Pending()) {
            # �ڑ��v�����Ȃ��ꍇ�A�����ҋ@
            Start-Sleep -Milliseconds 500
            continue
        }
        # �ڑ��v��������ꍇ�A

        # �N���C�A���g����̐ڑ���҂��܂��B
        $client = $listener.AcceptTcpClient()
        $stream = $client.GetStream()
        $reader = New-Object System.IO.StreamReader($stream)
        $writer = New-Object System.IO.StreamWriter($stream)

        # SMTP�Z�b�V�������J�n���܂��B
        $writer.WriteLine("220 PowerShell SMTP Server Ready")
        $writer.Flush()

        $mail = ""
        while ($null -ne ($line = $reader.ReadLine())) {
            # SMTP�R�}���h����͂��܂��B
            if ($line.StartsWith("HELO") -or $line.StartsWith("EHLO")) {
                $writer.WriteLine("250 Hello")
            }
            elseif ($line.StartsWith("MAIL FROM:")) {
                $writer.WriteLine("250 OK")
            }
            elseif ($line.StartsWith("RCPT TO:")) {
                # To��CC�̗������o�͂��܂��B
                $mail += "Recipient: " + $line.Substring(8) + "`r`n"
                $writer.WriteLine("250 OK")
            }
            elseif ($line.StartsWith("DATA")) {
                $writer.WriteLine("354 Start mail input")
                $writer.Flush()

                # ���[���̓��e���t�@�C���ɕۑ����܂��B
                $subtext = ""
                $base64mode = $false
                $base64text = ""
                while (($line = $reader.ReadLine()) -ne ".") {
                    if ($base64mode) {
                        $base64text += $line
                    }
                    else {
                        if ($line.StartsWith("Subject:")) {
                            # To��CC�̗������o�͂��܂��B
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

        # �N���C�A���g�Ƃ̐ڑ�����܂��B
        $client.Close()
    }

}
finally {

    $listener.stop()

}
