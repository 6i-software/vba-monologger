# ------------------------------------- #
#                                       #
#    VBA Monologger                     #
#    Copyright © 2024, 6i software      #
#                                       #
# ------------------------------------- #
#
# Simple web HTTP server for displaying log message sent by request's clients.
#
# It is used by the handler `VBAMonlologger.handler.handlerConsole`, in order to stream logs messages from 
# VBA procedures. A VBA client sends request (i.e. with a body wich contains a log message) and the server 
# simply show this message into this output console. The server will do nothing more than display the log 
# message sent by the client in the output console.
#
# Notes that, the client can stop the server with a special request. If the request contains the special 
# message "exit" or "stop-server". In this case, the server interprets it as a command to halt its execution 
# and release the listening port.
# 
# **Usage:**
# ```
# powershell .\VBAMonologgerServerLogsViewer.ps1 -port 20100 -hostname "localhost"
# ```
#
# This PowerShell code is directly embedded within the VBAMonlogger xlam library. You can locate it in the VBA 
# function `getPowershellCodeServerLogsViewer` into the handler `VBAMonlogger.handler.handlerConsole`.
##

# To enable CLI option -v/-Verbose (and more...) 
[CmdletBinding()]
param (
    [Parameter(Position = 0)][string] $hostname = "localhost",
    [Parameter(Position = 1)][int] $port = 20100
)

Add-Type -AssemblyName System.Net.Http


# region - Setting of global variables
$Global:prefixConsoleOutput = "[VBAMonologger server]"
$Global:defaultForegroundColor = "White"
$Global:defaultBackgroundColor = "Black"
$Global:highlightForegroundColor = "DarkGreen"
$Global:highlightBackgroundColor = "Black"
$Global:hostname = $hostname
$Global:port = $port
# endregion


# region - Helpers console
function _consoleWriteStyles
{
    param (
        [Parameter(Mandatory = $true)] [string]$message
    )

    # Search tag {h}{/h} and replace it by its default styles
    $message = $message -replace '<h>(.*?)</h>', "<style=`"foregroundColor:$highlightForegroundColor; backgroundColor:$highlightBackgroundColor;`">`${1}`</s>"

    # Regex used to capture styles
    $regex = '<style="([^"]+)">(.+?)</s>'
    $regexStyleColor = '(?:foregroundColor:(?<fgColor>[^;]+);?)?(?:\s*backgroundColor:(?<bgColor>[^;]+);?)?'

    $lastIndex = 0
    $matchesStyles = [regex]::Matches($message, $regex)
    if ($matchesStyles.Count -eq 0)
    {
        # If no matches, write text with default style
        [console]::WriteLine($message)
    }
    else
    {
        # Saved current console colors
        $currentForegroundColor = [console]::ForegroundColor
        $currentBackgroundColor = [console]::BackgroundColor

        foreach ($match in $matchesStyles)
        {
            $captureStyles = $match.Groups[1].Value
            $captureTextStyled = $match.Groups[2].Value
            $startIndex = $match.Index

            # Write text with default style before style bloc
            if ($startIndex - $lastIndex -gt 0)
            {
                $textBefore = $message.Substring($lastIndex, $startIndex - $lastIndex)
                [console]::Write($textBefore)
            }

            # Capture styles properties colors
            $styleMatch = [regex]::Match($captureStyles, $regexStyleColor)
            $foregroundColor = if ($styleMatch.Groups["fgColor"].Value)
            {
                $styleMatch.Groups["fgColor"].Value
            }
            else
            {
                $currentForegroundColor
            }
            $backgroundColor = if ($styleMatch.Groups["bgColor"].Value)
            {
                $styleMatch.Groups["bgColor"].Value
            }
            else
            {
                $currentBackgroundColor
            }

            # Write capture text with colors
            [console]::ForegroundColor = $foregroundColor
            [console]::BackgroundColor = $backgroundColor
            [console]::Write($captureTextStyled)
            [console]::ForegroundColor = $currentForegroundColor
            [console]::BackgroundColor = $currentBackgroundColor

            # Update position after regexp match
            $lastIndex = $startIndex + $match.Length
        }

        # Write remaining text after the last matches with default style
        if ($lastIndex -lt $message.Length)
        {
            $remainingText = $message.Substring($lastIndex)
            [console]::Write($remainingText)
        }
        [console]::WriteLine("")
    }
}

function _consoleLog
{
    param (
        [Parameter(Mandatory = $true)] [string] $message,
        [Parameter(Mandatory = $true)] [string] $type
    )

    # Save current colors
    $currentForegroundColor = [console]::ForegroundColor
    $currentBackgroundColor = [console]::BackgroundColor
    $currentHighlightForegroundColor = $Global:highlightForegroundColor
    $currentHighlightBackgroundColor = $Global:highlightBackgroundColor

    # Prepare output message
    $type = $type.ToLower()
    switch ($type)
    {
        "debug" {
            if ($VerbosePreference -eq "Continue")
            {
                [console]::ForegroundColor = "DarkGray"
                $message = $Global:prefixConsoleOutput + " [{0}] {1}" -f (Get-Date).ToString(), $message
            }
        }
        "error" {
            [console]::ForegroundColor = "Red"
            $Global:highlightForegroundColor = "Yellow"
            $Global:highlightBackgroundColor = "DarkRed"
            $message = $Global:prefixConsoleOutput + " [{0}] [ERROR] {1}" -f (Get-Date).ToString(), $message
        }
        "warning" {
            [console]::ForegroundColor = "DarkYellow"
            $Global:highlightBackgroundColor = "DarkMagenta"
            $message = $Global:prefixConsoleOutput + " [{0}] [WARNING] {1}" -f (Get-Date).ToString(), $message
        }
        default {
            [console]::ForegroundColor = "White"
            $message = $Global:prefixConsoleOutput + " [{0}] {1}" -f (Get-Date).ToString(), $message
        }
    }
    _consoleWriteStyles($message)

    # Restore previous colors
    [console]::ForegroundColor = $currentForegroundColor
    [console]::BackgroundColor = $currentBackgroundColor
    $Global:highlightForegroundColor = $currentHighlightForegroundColor
    $Global:highlightBackgroundColor = $currentHighlightBackgroundColor
}

function _consoleDebug()
{
    param ([Parameter(Mandatory = $true)] [string]$message)
    if ($VerbosePreference -eq "Continue")
    {
        _consoleLog -message $message -type "debug"
    }
}

function _consoleError()
{
    param ([Parameter(Mandatory = $true)] [string]$message)
    _consoleLog -message $message -type "error"
}

function _consoleWarning()
{
    param ([Parameter(Mandatory = $true)] [string]$message)
    _consoleLog -message $message -type "warning"
}

function _splashscreen()
{
    # Clear output console
    [System.Console]::Clear()
    $Host.UI.RawUI.WindowTitle = "VBAMonologger server logs viewer"

    # Write splashscreen message
    $currentForegroundColor = [console]::ForegroundColor
    [console]::ForegroundColor = "Blue"
    [console]::WriteLine("`n=== Start VBAMonologger server logs viewer ===`n")

    # Restore previous color
    [console]::ForegroundColor = $currentForegroundColor
}
# endregion

# region - HTTP server
function _createHTTPServer
{
    try
    {
        $server = [System.Net.HttpListener]::new()
        $server.Prefixes.Add("http://" + $Global:hostname + ":" + $Global:port + "/")
        $server.Start()
        _consoleDebug("Server is listening on : <h>""http://" + $Global:hostname + ":" + $Global:port + """</h>")
    }
    catch
    {
        _consoleError("Creation of server encountered an <h>critical error</h>. It is possible that the HTTP server's port: <h>" + $Global:port + "</h>, is already in use by another application or process.`n$_")
        Exit 1
    }

    return $server
}

function _stopHTTPServer
{
    param ([Parameter(Mandatory = $true)] [System.Net.HttpListener] $server)
    $server.Stop()
    [console]::WriteLine("`nServer shutdown, bye bye !")
}

function _startHTTPServer
{
    param ([Parameter(Mandatory = $true)] [System.Net.HttpListener] $server)

    try
    {
        $continue = $true
        while ($continue)
        {
            _consoleDebug("Waiting for new client connection...")
            $context = $server.GetContext()
            $request = $context.Request
            $response = $context.Response
            _consoleDebug("Connection established by a new client.")

            # Read client request (with support encoding UTF-8)
            $reader = [System.IO.StreamReader]::new($request.InputStream, [System.Text.Encoding]::UTF8)
            $message = $reader.ReadToEnd()
            _consoleDebug("Request received from the new client.")

            # Process client's request
            $command = $message.ToLower()

            # Check if `-wait` option  is given in request's command
            $waitTime = 0
            if ($command -like '*-wait*')
            {
                $parts = $command.Split()
                $mainCommand = $parts[0]
                $waitForIndex = [Array]::IndexOf($parts, '-wait')
                if ($waitForIndex -ne -1 -and $waitForIndex + 1 -lt $parts.Length)
                {
                    $waitTime = [int]$parts[$waitForIndex + 1]
                }
            }
            else
            {
                $mainCommand = $command
            }

            if ($mainCommand -eq 'exit' -or $mainCommand -eq 'stop' -or $mainCommand -eq 'stop-server')
            {
                if ($waitTime -gt 0)
                {
                    _consoleDebug("Stop command received, with a wait time for its execution of : <h>""" + $waitTime + """</h>.")
                    Start-Sleep -Milliseconds $waitTime
                } else {
                    _consoleDebug("Stop command received.")
                }
                
                $responseString = "Stop command received and executed by the server."
                $continue = $false
            }
            else
            {
                $responseString = "Request received! " + $message
                # Simply output message's request into console, in order to show log record
                [console]::WriteLine($message)
            }


            # Add custom headers 
            $response.Headers.Add("Server", "VBAMonologger HTTP Server")
            $response.Headers.Add("X-Powered-By", "PowerShell 5")
            $response.Headers.Add("X-Request-Received",(Get-Date).ToString("yyyy-MM-ddTHH:mm:ss"))

            # Send server response (with support encoding UTF-8)
            $response.StatusCode = 200
            $response.ContentType = "text/plain; charset=utf-8"
            $buffer = [System.Text.Encoding]::UTF8.GetBytes($responseString)
            $response.OutputStream.Write($buffer, 0, $buffer.Length)
            $response.Close()
        }
    }
    catch
    {
        _consoleError("The server encountered an <h>critical error</h>, preventing it from continuing to operate.`n$_")
    }
    finally
    {
        _stopHTTPServer -Server $server
    }
}
# endregion

function _main
{
    _splashscreen

    $server = _createHTTPServer
    _startHTTPServer -Server $server
}

_main