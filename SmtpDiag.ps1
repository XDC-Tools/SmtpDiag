Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Allows TLS 1.2 / 1.3 (as supported by OS/PowerShell 7+)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls13 -bor [Net.SecurityProtocolType]::Tls12

$form = New-Object Windows.Forms.Form
$form.Text = "SMTP Diagnostic Tool (Full Debug)"
$form.Size = [Drawing.Size]::new(600, 920)
$form.StartPosition = "CenterScreen"

# --- Fields ---
$fields = @{
    "SMTP Server:"     = "smtp.office365.com"
    "Port:"            = "587"
    "Sender Email:"    = "noreply@xdc.dk"
    "Recipient Email:" = "msr@xdc.dk"
    "Username:"        = "noreply@xdc.dk"
    "Password:"        = ""
    "Subject:"         = "Test Email from SMTP Diagnostic Tool"
    "Body:"            = "This is a test email sent from PowerShell."
}

$controls = @{}
$y = 10
foreach ($label in $fields.Keys) {
    $lbl = [Windows.Forms.Label]::new()
    $lbl.Text = $label
    $lbl.Location = [Drawing.Point]::new(10, $y)
    $lbl.Size = [Drawing.Size]::new(120, 20)

    $txt = [Windows.Forms.TextBox]::new()
    $txt.Location = [Drawing.Point]::new(140, $y)
    $txt.Size = [Drawing.Size]::new(420, 20)
    $txt.Text = $fields[$label]

    if ($label -eq "Password:") {
        $txt.UseSystemPasswordChar = $true
    }

    $form.Controls.Add($lbl)
    $form.Controls.Add($txt)
    $controls[$label] = $txt
    $y += 30
}

# --- Connection Type ---
$lblConnType = [Windows.Forms.Label]::new()
$lblConnType.Text = "Connection Type:"
$lblConnType.Location = [Drawing.Point]::new(10, $y)
$lblConnType.Size = [Drawing.Size]::new(120, 20)
$form.Controls.Add($lblConnType)

$cbConnType = [Windows.Forms.ComboBox]::new()
$cbConnType.Location = [Drawing.Point]::new(140, $y)
$cbConnType.Size = [Drawing.Size]::new(200, 20)
$cbConnType.Items.AddRange(@("STARTTLS (587)", "Implicit SSL (465)"))
$cbConnType.SelectedIndex = 0
$form.Controls.Add($cbConnType)
$y += 30

# --- Debug Options (Horizontal) ---
$gbDebug = [Windows.Forms.GroupBox]::new()
$gbDebug.Text = "Debug Options"
$gbDebug.Location = [Drawing.Point]::new(10, $y)
$gbDebug.Size = [Drawing.Size]::new(560, 60)

$cbTraceConnect = [Windows.Forms.CheckBox]::new()
$cbTraceConnect.Text = "Trace SMTP connection info"
$cbTraceConnect.Location = [Drawing.Point]::new(10, 25)
$cbTraceConnect.Size = [Drawing.Size]::new(180, 20)

$cbTraceProtocol = [Windows.Forms.CheckBox]::new()
$cbTraceProtocol.Text = "Trace protocol commands"
$cbTraceProtocol.Location = [Drawing.Point]::new(200, 25)
$cbTraceProtocol.Size = [Drawing.Size]::new(150, 20)

$cbTraceTls = [Windows.Forms.CheckBox]::new()
$cbTraceTls.Text = "Trace TLS handshake details"
$cbTraceTls.Location = [Drawing.Point]::new(360, 25)
$cbTraceTls.Size = [Drawing.Size]::new(190, 20)

$gbDebug.Controls.AddRange(@($cbTraceConnect, $cbTraceProtocol, $cbTraceTls))
$form.Controls.Add($gbDebug)

$y += 70   # move below the group box

# --- Debug Log ---
$lblLog = [Windows.Forms.Label]::new()
$lblLog.Text = "Debug Log:"
$lblLog.Location = [Drawing.Point]::new(10, $y)
$lblLog.Size = [Drawing.Size]::new(100, 20)
$form.Controls.Add($lblLog)

$y += 25
$txtLog = [Windows.Forms.TextBox]::new()
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.Location = [Drawing.Point]::new(10, $y)
$txtLog.Size = [Drawing.Size]::new(570, 350)
$txtLog.ReadOnly = $true
$form.Controls.Add($txtLog)

# --- Buttons ---
$btnConnect = [Windows.Forms.Button]::new()
$btnConnect.Text = "Test Connection"
$btnConnect.Size = [Drawing.Size]::new(150, 30)
$btnConnect.Location = [Drawing.Point]::new(10, 800)

$btnSend = [Windows.Forms.Button]::new()
$btnSend.Text = "Send Test Email"
$btnSend.Size = [Drawing.Size]::new(150, 30)
$btnSend.Location = [Drawing.Point]::new(170, 800)

$btnSendManual = [Windows.Forms.Button]::new()
$btnSendManual.Text = "Send Email (Full Debug)"
$btnSendManual.Size = [Drawing.Size]::new(170, 30)
$btnSendManual.Location = [Drawing.Point]::new(330, 800)

$btnClear = [Windows.Forms.Button]::new()
$btnClear.Text = "Clear Log"
$btnClear.Size = [Drawing.Size]::new(100, 30)
$btnClear.Location = [Drawing.Point]::new(510, 800)

$form.Controls.Add($btnConnect)
$form.Controls.Add($btnSend)
$form.Controls.Add($btnSendManual)
$form.Controls.Add($btnClear)

# --- Helper: logging w/ timestamp ---
function Write-Log {
    param([string]$message)
    $timestamp = (Get-Date -Format "HH:mm:ss.fff")
    $txtLog.AppendText("[$timestamp] $message`r`n")
}

$btnClear.Add_Click({ $txtLog.Clear() })

# --- Helper: Base64 ---
function To-Base64([string]$text) {
    [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($text))
}

# In your existing code, you used "S ->" and "S <-". 
# Now let's rename them to "CLIENT>>" for what the client sends, 
# and "SERVER>>" for what the server sends back.

# (Everything else stays the same logic, just renaming the lines in the transcripts.)

# --- Test Connection ---
$btnConnect.Add_Click({
    $server         = $controls["SMTP Server:"].Text
    $port           = [int]$controls["Port:"].Text
    $connType       = $cbConnType.SelectedItem
    $traceConnect   = $cbTraceConnect.Checked
    $traceProtocol  = $cbTraceProtocol.Checked
    $traceTls       = $cbTraceTls.Checked

    Write-Log ""
    Write-Log "Connecting to $server on port $port ($connType)..."
    try {
        $tcp = [Net.Sockets.TcpClient]::new()
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $tcp.Connect($server, $port)
        $sw.Stop()

        if ($traceConnect) {
            Write-Log "Local Endpoint:  $($tcp.Client.LocalEndPoint)"
            Write-Log "Remote Endpoint: $($tcp.Client.RemoteEndPoint)"
            Write-Log "Connection time: $($sw.ElapsedMilliseconds) ms"
        }

        $stream = $tcp.GetStream()

        # Read initial greet
        Start-Sleep -Milliseconds 300
        $buffer = New-Object byte[] 1024
        $count = $stream.Read($buffer, 0, $buffer.Length)
        $banner = [Text.Encoding]::ASCII.GetString($buffer, 0, $count)
        foreach ($line in $banner.Split("`n")) {
            $trim = $line.Trim()
            if ($trim) {
                if ($traceProtocol) { Write-Log "SERVER>> $trim" }
                Write-Log $trim
            }
        }

        if ($connType -like "*STARTTLS*") {
            # EHLO
            $cmdEhlo = "EHLO testclient.local`r`n"
            if ($traceProtocol) { Write-Log "CLIENT>> $cmdEhlo" }
            $ehloBytes = [Text.Encoding]::ASCII.GetBytes($cmdEhlo)
            $stream.Write($ehloBytes, 0, $ehloBytes.Length)
            Start-Sleep -Milliseconds 300

            # Read EHLO lines
            $ehloBuf = New-Object byte[] 2048
            $ehloCount = $stream.Read($ehloBuf, 0, $ehloBuf.Length)
            $ehloText = [Text.Encoding]::ASCII.GetString($ehloBuf, 0, $ehloCount)
            $ehloText.Split("`n") | ForEach-Object {
                $el = $_.Trim()
                if ($el) {
                    if ($traceProtocol) { Write-Log "SERVER>> $el" }
                    Write-Log $el
                }
            }

            if ($ehloText -notmatch "STARTTLS") {
                throw "STARTTLS not in EHLO response"
            }

            # STARTTLS
            $cmdTls = "STARTTLS`r`n"
            if ($traceProtocol) { Write-Log "CLIENT>> $cmdTls" }
            $tlsBytes = [Text.Encoding]::ASCII.GetBytes($cmdTls)
            $stream.Write($tlsBytes, 0, $tlsBytes.Length)
            Start-Sleep -Milliseconds 300

            $tlsBuf = New-Object byte[] 1024
            $tlsCount = $stream.Read($tlsBuf, 0, $tlsBuf.Length)
            $tlsResp = [Text.Encoding]::ASCII.GetString($tlsBuf, 0, $tlsCount).Trim()
            if ($traceProtocol) { Write-Log "SERVER>> $tlsResp" }

            if ($tlsResp -notlike "220*") {
                throw "STARTTLS failed: $tlsResp"
            }

            # Wrap in SSL
            $ssl = [Net.Security.SslStream]::new($stream, $false, { $true })
            $ssl.AuthenticateAsClient($server)
            Write-Log "✅ TLS handshake successful."

            if ($traceTls) {
                Write-Log "Protocol: $($ssl.SslProtocol)"
                Write-Log "Cipher: $($ssl.CipherAlgorithm) Strength: $($ssl.CipherStrength)"
                Write-Log "Hash: $($ssl.HashAlgorithm)"
                Write-Log "KeyEx: $($ssl.KeyExchangeAlgorithm)"

                if ($ssl.RemoteCertificate) {
                    Write-Log "Remote Cert Subject: $($ssl.RemoteCertificate.Subject)"
                    Write-Log "Remote Cert Issuer : $($ssl.RemoteCertificate.Issuer)"
                    Write-Log "Expires            : $($ssl.RemoteCertificate.GetExpirationDateString())"
                    Write-Log "Thumbprint         : $($ssl.RemoteCertificate.GetCertHashString())"

                    $chain = [System.Security.Cryptography.X509Certificates.X509Chain]::new()
                    $chain.Build($ssl.RemoteCertificate) | Out-Null
                    Write-Log "Certificate Chain:"
                    foreach ($element in $chain.ChainElements) {
                        Write-Log "  Subject : $($element.Certificate.Subject)"
                        Write-Log "  Issuer  : $($element.Certificate.Issuer)"
                        Write-Log "  Expires : $($element.Certificate.GetExpirationDateString())"
                        Write-Log "  Thumb   : $($element.Certificate.Thumbprint)"
                        Write-Log "  ------"
                    }
                }
            }
        }
        else {
            # Implicit SSL
            $ssl = [Net.Security.SslStream]::new($stream, $false, { $true })
            $ssl.AuthenticateAsClient($server)
            Write-Log "✅ TLS handshake successful."

            if ($traceTls) {
                Write-Log "Protocol: $($ssl.SslProtocol)"
                Write-Log "Cipher: $($ssl.CipherAlgorithm) Strength: $($ssl.CipherStrength)"
                Write-Log "Hash: $($ssl.HashAlgorithm)"
                Write-Log "KeyEx: $($ssl.KeyExchangeAlgorithm)"

                if ($ssl.RemoteCertificate) {
                    Write-Log "Remote Cert Subject: $($ssl.RemoteCertificate.Subject)"
                    Write-Log "Remote Cert Issuer : $($ssl.RemoteCertificate.Issuer)"
                    Write-Log "Expires            : $($ssl.RemoteCertificate.GetExpirationDateString())"
                    Write-Log "Thumbprint         : $($ssl.RemoteCertificate.GetCertHashString())"

                    $chain = [System.Security.Cryptography.X509Certificates.X509Chain]::new()
                    $chain.Build($ssl.RemoteCertificate) | Out-Null
                    Write-Log "Certificate Chain:"
                    foreach ($element in $chain.ChainElements) {
                        Write-Log "  Subject : $($element.Certificate.Subject)"
                        Write-Log "  Issuer  : $($element.Certificate.Issuer)"
                        Write-Log "  Expires : $($element.Certificate.GetExpirationDateString())"
                        Write-Log "  Thumb   : $($element.Certificate.Thumbprint)"
                        Write-Log "  ------"
                    }
                }
            }
        }

        $tcp.Close()
        Write-Log "✅ Connection test completed."
    }
    catch {
        Write-Log "❌ Error: $_"
    }
})

# --- Send Email (SmtpClient) ---
$btnSend.Add_Click({
    $smtpServer  = $controls["SMTP Server:"].Text
    $port        = [int]$controls["Port:"].Text
    $from        = $controls["Sender Email:"].Text
    $to          = $controls["Recipient Email:"].Text
    $username    = $controls["Username:"].Text
    $password    = $controls["Password:"].Text
    $subject     = $controls["Subject:"].Text
    $body        = $controls["Body:"].Text

    $traceProtocol = $cbTraceProtocol.Checked
    Write-Log ""
    Write-Log "Sending email to $to (SmtpClient) ..."
    if ($traceProtocol) {
        Write-Log "CLIENT>> Email From: $from"
        Write-Log "CLIENT>> Email To:   $to"
        Write-Log "CLIENT>> Subject:    $subject"
        Write-Log "CLIENT>> Body:       $body"
    }

    try {
        $smtp = [Net.Mail.SmtpClient]::new($smtpServer, $port)
        $smtp.EnableSsl = $true
        $smtp.Credentials = [Net.NetworkCredential]::new($username, $password)

        $mail = [Net.Mail.MailMessage]::new($from, $to, $subject, $body)
        $smtp.Send($mail)

        Write-Log "✅ Email sent successfully."
        Write-Log "Note: Full raw transcript not exposed by SmtpClient."
    }
    catch {
        $errorText = $_.Exception.Message
        if ($errorText -match "Authentication unsuccessful|535") {
            Write-Log "❌ Authentication failed: Please check username and password."
        } else {
            Write-Log "❌ Failed to send email: $errorText"
        }
    }
})

# --- Send Email (Full Debug) ---
# ... identical to prior version, but "S ->" / "S <-" replaced with "CLIENT>>" / "SERVER>>"
$btnSendManual.Add_Click({
    $server    = $controls["SMTP Server:"].Text
    $port      = [int]$controls["Port:"].Text
    $from      = $controls["Sender Email:"].Text
    $to        = $controls["Recipient Email:"].Text
    $username  = $controls["Username:"].Text
    $password  = $controls["Password:"].Text
    $subject   = $controls["Subject:"].Text
    $body      = $controls["Body:"].Text

    $traceConnect  = $cbTraceConnect.Checked
    $traceProtocol = $cbTraceProtocol.Checked
    $traceTls      = $cbTraceTls.Checked

    Write-Log ""
    Write-Log "Sending email to $to (Manual SMTP Full Debug) ..."
    if ($traceProtocol) {
        Write-Log "CLIENT>> Email From: $from"
        Write-Log "CLIENT>> Email To:   $to"
        Write-Log "CLIENT>> Subject:    $subject"
        Write-Log "CLIENT>> Body:       $body"
    }

    try {
        # 1) TCP Connect
        $tcp = [Net.Sockets.TcpClient]::new()
        $sw  = [System.Diagnostics.Stopwatch]::StartNew()
        $tcp.Connect($server, $port)
        $sw.Stop()

        if ($traceConnect) {
            Write-Log "Local:  $($tcp.Client.LocalEndPoint)"
            Write-Log "Remote: $($tcp.Client.RemoteEndPoint)"
            Write-Log "Connect time: $($sw.ElapsedMilliseconds) ms"
        }

        $netStream = $tcp.GetStream()
        Start-Sleep -Milliseconds 300

        # Read greeting
        $buffer = New-Object byte[] 1024
        $count = $netStream.Read($buffer, 0, $buffer.Length)
        $banner = [Text.Encoding]::ASCII.GetString($buffer, 0, $count)
        foreach ($line in $banner.Split("`n")) {
            $l = $line.Trim()
            if ($l) {
                if ($traceProtocol) { Write-Log "SERVER>> $l" }
                Write-Log $l
            }
        }

        # 2) Possibly do STARTTLS
        $doStartTls = $cbConnType.SelectedItem -like "*STARTTLS*"
        if ($doStartTls) {
            # EHLO
            $ehloCmd = "EHLO testclient.local`r`n"
            if ($traceProtocol) { Write-Log "CLIENT>> $ehloCmd" }
            $ehloBytes = [Text.Encoding]::ASCII.GetBytes($ehloCmd)
            $netStream.Write($ehloBytes, 0, $ehloBytes.Length)

            Start-Sleep -Milliseconds 300
            $ehloBuf = New-Object byte[] 2048
            $ehloCnt = $netStream.Read($ehloBuf, 0, $ehloBuf.Length)
            $ehloResp = [Text.Encoding]::ASCII.GetString($ehloBuf, 0, $ehloCnt)
            foreach ($row in $ehloResp.Split("`n")) {
                $r = $row.Trim()
                if ($r) {
                    if ($traceProtocol) { Write-Log "SERVER>> $r" }
                    Write-Log $r
                }
            }
            if ($ehloResp -notmatch "STARTTLS") {
                throw "STARTTLS not supported by server"
            }

            # STARTTLS command
            $startTlsCmd = "STARTTLS`r`n"
            if ($traceProtocol) { Write-Log "CLIENT>> $startTlsCmd" }
            $startTlsBytes = [Text.Encoding]::ASCII.GetBytes($startTlsCmd)
            $netStream.Write($startTlsBytes, 0, $startTlsBytes.Length)

            Start-Sleep -Milliseconds 300
            $startTlsBuf = New-Object byte[] 1024
            $startTlsCnt = $netStream.Read($startTlsBuf, 0, $startTlsBuf.Length)
            $startTlsResp = [Text.Encoding]::ASCII.GetString($startTlsBuf, 0, $startTlsCnt).Trim()
            if ($traceProtocol) { Write-Log "SERVER>> $startTlsResp" }
            if ($startTlsResp -notlike "220*") {
                throw "STARTTLS failed: $startTlsResp"
            }

            # TLS handshake
            $ssl = [Net.Security.SslStream]::new($netStream, $false, { $true })
            $ssl.AuthenticateAsClient($server)
            Write-Log "✅ TLS handshake done."

            if ($traceTls) {
                Write-Log "Protocol: $($ssl.SslProtocol)"
                Write-Log "Cipher: $($ssl.CipherAlgorithm) Strength: $($ssl.CipherStrength)"
                Write-Log "Hash: $($ssl.HashAlgorithm)"
                Write-Log "KeyEx: $($ssl.KeyExchangeAlgorithm)"
                if ($ssl.RemoteCertificate) {
                    Write-Log "Remote Cert: $($ssl.RemoteCertificate.Subject)"
                }
            }

            # Re-EHLO
            $ehlo2Cmd = "EHLO testclient.local`r`n"
            if ($traceProtocol) { Write-Log "CLIENT>> $ehlo2Cmd" }
            $ehlo2Bytes = [Text.Encoding]::ASCII.GetBytes($ehlo2Cmd)
            $ssl.Write($ehlo2Bytes, 0, $ehlo2Bytes.Length)
            Start-Sleep -Milliseconds 200

            $ehlo2Resp = ''
            if ($ssl.CanRead) {
                $respBuf = New-Object byte[] 2048
                $respCnt = $ssl.Read($respBuf, 0, $respBuf.Length)
                if ($respCnt -gt 0) {
                    $ehlo2Resp = [Text.Encoding]::ASCII.GetString($respBuf, 0, $respCnt)
                    foreach ($row in $ehlo2Resp.Split("`n")) {
                        $r = $row.Trim()
                        if ($r) {
                            if ($traceProtocol) { Write-Log "SERVER>> $r" }
                            Write-Log $r
                        }
                    }
                }
            }
            $smtpStream = $ssl
        }
        else {
            # Implicit SSL
            $ssl = [Net.Security.SslStream]::new($netStream, $false, { $true })
            $ssl.AuthenticateAsClient($server)
            Write-Log "✅ TLS handshake done."
            if ($traceTls) {
                Write-Log "Protocol: $($ssl.SslProtocol)"
                Write-Log "Cipher: $($ssl.CipherAlgorithm) Strength: $($ssl.CipherStrength)"
                Write-Log "Hash: $($ssl.HashAlgorithm)"
                Write-Log "KeyEx: $($ssl.KeyExchangeAlgorithm)"
                if ($ssl.RemoteCertificate) {
                    Write-Log "Remote Cert: $($ssl.RemoteCertificate.Subject)"
                }
            }
            # EHLO
            $ehloCmd = "EHLO testclient.local`r`n"
            if ($traceProtocol) { Write-Log "CLIENT>> $ehloCmd" }
            $ehloBytes = [Text.Encoding]::ASCII.GetBytes($ehloCmd)
            $ssl.Write($ehloBytes, 0, $ehloBytes.Length)
            Start-Sleep -Milliseconds 300

            $ehloBuf = New-Object byte[] 2048
            $ehloCnt = $ssl.Read($ehloBuf, 0, $ehloBuf.Length)
            if ($ehloCnt -gt 0) {
                $ehloResp = [Text.Encoding]::ASCII.GetString($ehloBuf, 0, $ehloCnt)
                foreach ($row in $ehloResp.Split("`n")) {
                    $r = $row.Trim()
                    if ($r) {
                        if ($traceProtocol) { Write-Log "SERVER>> $r" }
                        Write-Log $r
                    }
                }
            }
            $smtpStream = $ssl
        }

        # 3) AUTH LOGIN
        $authCmd = "AUTH LOGIN`r`n"
        if ($traceProtocol) { Write-Log "CLIENT>> $authCmd" }
        $smtpStream.Write([Text.Encoding]::ASCII.GetBytes($authCmd), 0, $authCmd.Length)
        Start-Sleep -Milliseconds 150

        $respBuf = New-Object byte[] 512
        $respCnt = $smtpStream.Read($respBuf, 0, $respBuf.Length)
        $respTxt = [Text.Encoding]::ASCII.GetString($respBuf, 0, $respCnt).Trim()
        if ($traceProtocol) { Write-Log "SERVER>> $respTxt" }
        if ($respTxt -notmatch "334") {
            throw "AUTH did not proceed. Server said: $respTxt"
        }

        # username
        $u64 = To-Base64($username)
        if ($traceProtocol) { Write-Log "CLIENT>> [base64-username]" }
        $smtpStream.Write([Text.Encoding]::ASCII.GetBytes("$u64`r`n"), 0, $u64.Length + 2)
        Start-Sleep -Milliseconds 150

        $respCnt = $smtpStream.Read($respBuf, 0, $respBuf.Length)
        $respTxt = [Text.Encoding]::ASCII.GetString($respBuf, 0, $respCnt).Trim()
        if ($traceProtocol) { Write-Log "SERVER>> $respTxt" }
        if ($respTxt -notmatch "334") {
            throw "AUTH did not proceed (password prompt). Server said: $respTxt"
        }

        # password
        $p64 = To-Base64($password)
        if ($traceProtocol) { Write-Log "CLIENT>> [base64-password]" }
        $smtpStream.Write([Text.Encoding]::ASCII.GetBytes("$p64`r`n"), 0, $p64.Length + 2)
        Start-Sleep -Milliseconds 150

        $respCnt = $smtpStream.Read($respBuf, 0, $respBuf.Length)
        $respTxt = [Text.Encoding]::ASCII.GetString($respBuf, 0, $respCnt).Trim()
        if ($traceProtocol) { Write-Log "SERVER>> $respTxt" }
        if ($respTxt -notmatch "^235") {
            throw "AUTH LOGIN failed: $respTxt"
        }

        # 4) MAIL FROM
        $mailFromCmd = "MAIL FROM:<$from>`r`n"
        if ($traceProtocol) { Write-Log "CLIENT>> $mailFromCmd" }
        $smtpStream.Write([Text.Encoding]::ASCII.GetBytes($mailFromCmd), 0, $mailFromCmd.Length)
        Start-Sleep -Milliseconds 150

        $respCnt = $smtpStream.Read($respBuf, 0, $respBuf.Length)
        $respTxt = [Text.Encoding]::ASCII.GetString($respBuf, 0, $respCnt).Trim()
        if ($traceProtocol) { Write-Log "SERVER>> $respTxt" }
        if ($respTxt -notmatch "^250") {
            throw "MAIL FROM failed: $respTxt"
        }

        # 5) RCPT TO
        $rcptCmd = "RCPT TO:<$to>`r`n"
        if ($traceProtocol) { Write-Log "CLIENT>> $rcptCmd" }
        $smtpStream.Write([Text.Encoding]::ASCII.GetBytes($rcptCmd), 0, $rcptCmd.Length)
        Start-Sleep -Milliseconds 150

        $respCnt = $smtpStream.Read($respBuf, 0, $respBuf.Length)
        $respTxt = [Text.Encoding]::ASCII.GetString($respBuf, 0, $respCnt).Trim()
        if ($traceProtocol) { Write-Log "SERVER>> $respTxt" }
        if ($respTxt -notmatch "^250") {
            throw "RCPT TO failed: $respTxt"
        }

        # 6) DATA
        $dataCmd = "DATA`r`n"
        if ($traceProtocol) { Write-Log "CLIENT>> $dataCmd" }
        $smtpStream.Write([Text.Encoding]::ASCII.GetBytes($dataCmd), 0, $dataCmd.Length)
        Start-Sleep -Milliseconds 150

        $respCnt = $smtpStream.Read($respBuf, 0, $respBuf.Length)
        $respTxt = [Text.Encoding]::ASCII.GetString($respBuf, 0, $respCnt).Trim()
        if ($traceProtocol) { Write-Log "SERVER>> $respTxt" }
        if ($respTxt -notmatch "^354") {
            throw "DATA not accepted: $respTxt"
        }

        # Headers + body
        $smtpHeaders = "From: <$from>`r`nTo: <$to>`r`nSubject: $subject`r`nContent-Type: text/plain; charset=UTF-8`r`n`r`n"
        $smtpMsg = $smtpHeaders + $body + "`r`n.`r`n"
        if ($traceProtocol) { Write-Log "CLIENT>> [Message Body + <CRLF>.<CRLF>]" }
        $msgBytes = [Text.Encoding]::ASCII.GetBytes($smtpMsg)
        $smtpStream.Write($msgBytes, 0, $msgBytes.Length)
        Start-Sleep -Milliseconds 150

        $respCnt = $smtpStream.Read($respBuf, 0, $respBuf.Length)
        $respTxt = [Text.Encoding]::ASCII.GetString($respBuf, 0, $respCnt).Trim()
        if ($traceProtocol) { Write-Log "SERVER>> $respTxt" }
        if ($respTxt -notmatch "^250") {
            throw "Message not accepted: $respTxt"
        }

        # 7) QUIT
        $quitCmd = "QUIT`r`n"
        if ($traceProtocol) { Write-Log "CLIENT>> $quitCmd" }
        $smtpStream.Write([Text.Encoding]::ASCII.GetBytes($quitCmd), 0, $quitCmd.Length)
        Start-Sleep -Milliseconds 150

        $respCnt = $smtpStream.Read($respBuf, 0, $respBuf.Length)
        if ($respCnt -gt 0) {
            $respTxt = [Text.Encoding]::ASCII.GetString($respBuf, 0, $respCnt).Trim()
            if ($traceProtocol) { Write-Log "SERVER>> $respTxt" }
        }

        $tcp.Close()
        Write-Log "✅ Email sent successfully (Manual SMTP)."
    }
    catch {
        Write-Log "❌ Error (Manual SMTP): $_"
    }
})

[void]$form.ShowDialog()
