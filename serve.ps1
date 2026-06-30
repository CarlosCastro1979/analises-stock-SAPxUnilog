# Servidor local para Performance Logística (SAP x Unilog)
$port = if ($env:PORT) { [int]$env:PORT } else { 5501 }
$root = $PSScriptRoot

$mime = @{
    '.html' = 'text/html; charset=utf-8'
    '.htm'  = 'text/html; charset=utf-8'
    '.js'   = 'application/javascript; charset=utf-8'
    '.css'  = 'text/css; charset=utf-8'
    '.json' = 'application/json; charset=utf-8'
    '.png'  = 'image/png'
    '.jpg'  = 'image/jpeg'
    '.svg'  = 'image/svg+xml'
    '.ico'  = 'image/x-icon'
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
$listener.Prefixes.Add("http://127.0.0.1:$port/")
$listener.Start()

Write-Host ""
Write-Host "  Performance Logistica - servidor local" -ForegroundColor Green
Write-Host "  http://localhost:$port/" -ForegroundColor Cyan
Write-Host "  http://127.0.0.1:$port/" -ForegroundColor Cyan
Write-Host "  Pasta: $root" -ForegroundColor DarkGray
Write-Host "  Ctrl+C para parar" -ForegroundColor DarkGray
Write-Host ""

try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response

        $rel = [System.Uri]::UnescapeDataString($request.Url.LocalPath).TrimStart('/')
        if (-not $rel) { $rel = 'index.html' }
        $file = Join-Path $root ($rel -replace '/', [IO.Path]::DirectorySeparatorChar)

        if ((Test-Path $file -PathType Leaf) -and $file.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) {
            $ext = [IO.Path]::GetExtension($file).ToLower()
            $response.ContentType = if ($mime.ContainsKey($ext)) { $mime[$ext] } else { 'application/octet-stream' }
            if ($mime.ContainsKey($ext)) {
                $response.Headers.Add('Cache-Control', 'no-cache, no-store, must-revalidate')
                $response.Headers.Add('Pragma', 'no-cache')
                $response.Headers.Add('Expires', '0')
                $lm = (Get-Item $file).LastWriteTimeUtc
                $response.Headers.Add('Last-Modified', $lm.ToString('R'))
            }
            $bytes = [IO.File]::ReadAllBytes($file)
            $response.ContentLength64 = $bytes.Length
            $response.StatusCode = 200
            $response.OutputStream.Write($bytes, 0, $bytes.Length)
            Write-Host "  200  $($request.HttpMethod)  /$rel" -ForegroundColor DarkGray
        } else {
            $response.StatusCode = 404
            $msg = [Text.Encoding]::UTF8.GetBytes('404 Not Found')
            $response.ContentLength64 = $msg.Length
            $response.OutputStream.Write($msg, 0, $msg.Length)
            Write-Host "  404  $($request.HttpMethod)  /$rel" -ForegroundColor Yellow
        }
        $response.Close()
    }
} finally {
    $listener.Stop()
    $listener.Close()
}
