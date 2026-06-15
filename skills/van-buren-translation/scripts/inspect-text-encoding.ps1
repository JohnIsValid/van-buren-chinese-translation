param(
    [Parameter(Mandatory = $true)]
    [string]$Path
)

$ErrorActionPreference = "Stop"
$resolved = (Resolve-Path -LiteralPath $Path).Path
$bytes = [System.IO.File]::ReadAllBytes($resolved)
$containerHeader = ""
$contentOffset = 0

if ($bytes.Length -ge 3 -and $bytes[0] -eq 0x7E -and $bytes[1] -eq 0xE3 -and $bytes[2] -eq 0x03) {
    $containerHeader = "vbtext-7e-e3-03"
    $contentOffset = 3
}

$contentBytes = if ($contentOffset -gt 0) {
    [byte[]]$bytes[$contentOffset..($bytes.Length - 1)]
}
else {
    $bytes
}

function Test-StrictUtf8 {
    param([byte[]]$Data)
    try {
        $utf8 = New-Object System.Text.UTF8Encoding($false, $true)
        [void]$utf8.GetString($Data)
        return $true
    }
    catch {
        return $false
    }
}

$encoding = "unknown"
$bomLength = 0
if ($contentBytes.Length -ge 3 -and $contentBytes[0] -eq 0xEF -and $contentBytes[1] -eq 0xBB -and $contentBytes[2] -eq 0xBF) {
    $encoding = "utf-8-bom"
    $bomLength = 3
}
elseif ($contentBytes.Length -ge 2 -and $contentBytes[0] -eq 0xFF -and $contentBytes[1] -eq 0xFE) {
    $encoding = "utf-16-le"
    $bomLength = 2
}
elseif ($contentBytes.Length -ge 2 -and $contentBytes[0] -eq 0xFE -and $contentBytes[1] -eq 0xFF) {
    $encoding = "utf-16-be"
    $bomLength = 2
}
elseif (($contentBytes | Where-Object { $_ -ge 0x80 }).Count -eq 0) {
    $encoding = "ascii"
}
elseif (Test-StrictUtf8 -Data $contentBytes) {
    $encoding = "utf-8"
}
elseif ($containerHeader -eq "vbtext-7e-e3-03") {
    try {
        $strictGbk = [System.Text.Encoding]::GetEncoding(
            936,
            (New-Object System.Text.EncoderExceptionFallback),
            (New-Object System.Text.DecoderExceptionFallback)
        )
        [void]$strictGbk.GetString($contentBytes)
        $encoding = "gbk"
    }
    catch {
        try {
            $strict1252 = [System.Text.Encoding]::GetEncoding(
                1252,
                (New-Object System.Text.EncoderExceptionFallback),
                (New-Object System.Text.DecoderExceptionFallback)
            )
            [void]$strict1252.GetString($contentBytes)
            $encoding = "windows-1252"
        }
        catch {
            $encoding = "unknown"
        }
    }
}

$decoder = switch ($encoding) {
    "utf-8-bom" { New-Object System.Text.UTF8Encoding($true, $true) }
    "utf-8" { New-Object System.Text.UTF8Encoding($false, $true) }
    "ascii" { [System.Text.Encoding]::ASCII }
    "utf-16-le" { [System.Text.Encoding]::Unicode }
    "utf-16-be" { [System.Text.Encoding]::BigEndianUnicode }
    "gbk" { [System.Text.Encoding]::GetEncoding(936) }
    "windows-1252" { [System.Text.Encoding]::GetEncoding(1252) }
    default { $null }
}

$raw = if ($decoder -ne $null) {
    $decoder.GetString($contentBytes, $bomLength, $contentBytes.Length - $bomLength)
}
else {
    ""
}

$crlf = ([regex]::Matches($raw, "`r`n")).Count
$lf = ([regex]::Matches($raw, "(?<!`r)`n")).Count
$lineEnding = if ($crlf -gt 0 -and $lf -eq 0) { "crlf" } elseif ($lf -gt 0 -and $crlf -eq 0) { "lf" } elseif ($crlf -eq 0 -and $lf -eq 0) { "none" } else { "mixed" }
$replacementCount = if ($raw.Length -gt 0) { ([regex]::Matches($raw, [string][char]0xFFFD)).Count } else { 0 }
$nullCount = if ($raw.Length -gt 0) { ([regex]::Matches($raw, [string][char]0x0000)).Count } else { 0 }

[PSCustomObject]@{
    path = $resolved
    bytes = $bytes.Length
    container_header = $containerHeader
    content_offset = $contentOffset
    detected_encoding = $encoding
    bom_bytes = $bomLength
    line_endings = $lineEnding
    replacement_character_count = $replacementCount
    unexpected_null_count = $nullCount
    requires_confirmation = ($encoding -eq "unknown" -or $lineEnding -eq "mixed" -or $replacementCount -gt 0 -or $nullCount -gt 0)
} | ConvertTo-Json -Depth 3
