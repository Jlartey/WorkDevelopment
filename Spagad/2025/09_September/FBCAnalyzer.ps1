# Configuration
$analyzerIP = '192.168.1.250'    # Medical analyzer IP
$analyzerPort = 5100             # Medical analyzer port
$destinationIP = '192.168.1.241' # Destination server IP
$destinationPort = 1111          # Destination port
$logFile = "$env:USERPROFILE\Documents\HL7_Forwarder.log"
$bufferSize = 8192               # Buffer size for reading data
$connectionTimeout = 10000       # 10 second timeout
$destinationSendTimeout = 3000   # 3 second timeout for sending
$minDataSize = 10                # Minimum data size to process
$maxIdleTime = 30000             # 30 seconds max idle time (ms)

# HL7 Message Delimiters
[byte]$hl7StartChar = 0x0B       # VT character
[byte]$hl7EndChar = 0x1C         # FS character
[byte]$hl7CRChar = 0x0D          # Carriage Return
[byte]$hl7EOTChar = 0x04         # End of Transmission character

# Create log file directory if it doesn't exist
$logDirectory = Split-Path -Path $logFile -Parent
if (-not (Test-Path -Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
}

# Enhanced logging function with error highlighting
function Write-Log {
    param(
        [string]$message,
        [string]$level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - [$level] - $message"
    
    $logEntry | Out-File -FilePath $logFile -Append
    
    switch ($level) {
        "ERROR" { Write-Host $logEntry -ForegroundColor Red }
        "WARN"  { Write-Host $logEntry -ForegroundColor Yellow }
        "DEBUG" { Write-Host $logEntry -ForegroundColor Gray }
        default { Write-Host $logEntry }
    }
}

# Function to establish TCP connection with timeout and keepalive
function Connect-TCP {
    param(
        [string]$ip,
        [int]$port,
        [int]$timeoutMs,
        [bool]$enableKeepAlive = $false
    )
    
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $client.ReceiveTimeout = $timeoutMs
        $client.SendTimeout = $timeoutMs
        
        # Enable keepalive if requested
        if ($enableKeepAlive) {
            $client.Client.SetSocketOption([System.Net.Sockets.SocketOptionLevel]::Socket, 
                                         [System.Net.Sockets.SocketOptionName]::KeepAlive, 
                                         $true)
        }
        
        $connectResult = $client.BeginConnect($ip, $port, $null, $null)
        
        if ($connectResult.AsyncWaitHandle.WaitOne($timeoutMs, $true)) {
            $client.EndConnect($connectResult)
            $stream = $client.GetStream()
            $stream.ReadTimeout = $timeoutMs
            Write-Log "Connected to $($ip):$($port)" "DEBUG"
            return @{
                Client = $client
                Stream = $stream
            }
        }
        else {
            $client.Close()
            throw "Connection timeout after $($timeoutMs)ms"
        }
    }
    catch {
        Write-Log "Connection error to $($ip):$($port): $_" "ERROR"
        throw
    }
}

# Function to send data with timeout and EOT character
function Send-HL7Message {
    param(
        [System.Net.Sockets.NetworkStream]$stream,
        [byte[]]$hl7Message,
        [int]$timeoutMs
    )
    
    try {
        # Add EOT character to the message
        $messageWithEOT = $hl7Message + $hl7EOTChar
        
        # Set timeout for this operation
        $stream.WriteTimeout = $timeoutMs
        
        # Send the message
        $stream.Write($messageWithEOT, 0, $messageWithEOT.Length)
        $stream.Flush()
        
        #Write-Log "Sent $($messageWithEOT.Length) bytes (including EOT)" "DEBUG"
        return $true
    }
    catch {
        Write-Log "Failed to send message: $_" "ERROR"
        return $false
    }
    finally {
        # Reset timeout to default
        $stream.WriteTimeout = $connectionTimeout
    }
}

# Function to properly close connection
function Close-Connection {
    param(
        [System.Net.Sockets.TcpClient]$client,
        [System.Net.Sockets.NetworkStream]$stream,
        [string]$connectionType
    )
    
    try {
        if ($null -ne $stream) {
            $stream.Dispose()
            Write-Log "Closed $connectionType stream" "DEBUG"
        }
    }
    catch {
        Write-Log "Error closing $connectionType stream: $_" "WARN"
    }
    
    try {
        if ($null -ne $client) {
            $client.Dispose()
            Write-Log "Closed $connectionType client" "DEBUG"
        }
    }
    catch {
        Write-Log "Error closing $connectionType client: $_" "WARN"
    }
}

# Main processing function
function Start-Forwarding {
    $analyzerConnection = $null
    $lastActivityTime = [System.Diagnostics.Stopwatch]::StartNew()
    
    try {
        # Establish analyzer connection with keepalive
        Write-Log "Connecting to analyzer at ${analyzerIP}:${analyzerPort}"
        $analyzerConnection = Connect-TCP -ip $analyzerIP -port $analyzerPort -timeoutMs $connectionTimeout -enableKeepAlive $true
        Write-Log "Analyzer connection established"
        
        $buffer = New-Object byte[] $bufferSize
        $messageBuffer = New-Object System.Collections.Generic.List[byte]
        $inMessage = $false
        
        while ($true) {
            # Check for analyzer disconnection
            if ($null -eq $analyzerConnection -or !$analyzerConnection.Client.Connected) {
                throw "Analyzer connection lost"
            }
            
            # Check for data available
            if ($analyzerConnection.Stream.DataAvailable) {
                $bytesRead = $analyzerConnection.Stream.Read($buffer, 0, $buffer.Length)
                $lastActivityTime.Restart()
                
                if ($bytesRead -gt 0) {
                    for ($i = 0; $i -lt $bytesRead; $i++) {
                        $currentByte = $buffer[$i]
                        
                        # Detect HL7 message start
                        if ($currentByte -eq $hl7StartChar) {
                            $inMessage = $true
                            $messageBuffer.Clear()
                            $messageBuffer.Add($currentByte)
                            continue
                        }
                        
                        # Buffer the byte if we're in a message
                        if ($inMessage) {
                            $messageBuffer.Add($currentByte)
                        }
                        
                        # Detect HL7 message end
                        if ($currentByte -eq $hl7EndChar -and $inMessage) {
                            $inMessage = $false
                            
                            # Process complete HL7 message
                            if ($messageBuffer.Count -ge $minDataSize) {
                                $messageBytes = $messageBuffer.ToArray()
                                Write-Log "Detected complete HL7 message (${$messageBytes.Length} bytes)"
                                
                                # Forward to destination
                                $destinationConnection = $null
                                try {
                                    Write-Log "Connecting to destination server"
                                    $destinationConnection = Connect-TCP -ip $destinationIP -port $destinationPort -timeoutMs $connectionTimeout
                                    
                                    if (Send-HL7Message -stream $destinationConnection.Stream -hl7Message $messageBytes -timeoutMs $destinationSendTimeout) {
                                        Write-Log "Successfully forwarded HL7 message with EOT"
                                    }
                                    else {
                                        Write-Log "Failed to send HL7 message" "WARN"
                                    }
                                }
                                catch {
                                    Write-Log "Destination error: $_" "ERROR"
                                }
                                finally {
                                    Close-Connection -client $destinationConnection.Client -stream $destinationConnection.Stream -connectionType "destination"
                                }
                            }
                        }
                    }
                }
            }
            # Check for idle timeout
            elseif ($lastActivityTime.ElapsedMilliseconds -gt $maxIdleTime) {
                throw "Analyzer connection timeout (no data for $($maxIdleTime/1000) seconds)"
            }
            else {
                Start-Sleep -Milliseconds 100
            }
        }
    }
    catch {
        Write-Log "Processing error: $_" "ERROR"
        throw
    }
    finally {
        Close-Connection -client $analyzerConnection.Client -stream $analyzerConnection.Stream -connectionType "analyzer"
    }
}

# Main execution loop
Write-Log "=== HL7 Forwarder Started ==="
Write-Log "Press Ctrl+C to stop"

$retryCount = 0
$maxRetries = 10
$retryDelay = 5

while ($retryCount -lt $maxRetries) {
    try {
        Start-Forwarding
    }
    catch {
        $retryCount++
        Write-Log "Error occurred (attempt $retryCount of $maxRetries): $_" "ERROR"
        
        if ($retryCount -lt $maxRetries) {
            Write-Log "Retrying in $retryDelay seconds..." "WARN"
            Start-Sleep -Seconds $retryDelay
        }
        else {
            Write-Log "Max retries reached. Exiting." "ERROR"
            exit 1
        }
    }
}