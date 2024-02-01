# MIMEエンコードされた文字列
# $encodedString = "=?utf-8?B?44K/44Kk44OI44Or?="
$encodedString = "test mail"

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

# 結果を表示
Write-Output $decodedString
} catch {
    Write-Output $encodedString    
}
