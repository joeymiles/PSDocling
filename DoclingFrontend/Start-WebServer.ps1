param([int]$Port = 8081)

$http = New-Object System.Net.HttpListener
$prefix = "http://localhost:$Port/"
$http.Prefixes.Add($prefix)
$http.Start()

Write-Host "Web server running at $prefix" -ForegroundColor Green

# Simple mime map
$MimeMap = @{
  ".html" = "text/html; charset=utf-8"
  ".htm"  = "text/html; charset=utf-8"
  ".css"  = "text/css; charset=utf-8"
  ".js"   = "application/javascript; charset=utf-8"
  ".json" = "application/json; charset=utf-8"
  ".svg"  = "image/svg+xml"
  ".png"  = "image/png"
  ".jpg"  = "image/jpeg"
  ".jpeg" = "image/jpeg"
  ".gif"  = "image/gif"
  ".webp" = "image/webp"
  ".ico"  = "image/x-icon"
  ".txt"  = "text/plain; charset=utf-8"
  ".map"  = "application/json; charset=utf-8"
  ".xml"  = "application/xml; charset=utf-8"
}

try {
  while ($http.IsListening) {
    $context = $http.GetContext()
    $response = $context.Response
    $request  = $context.Request

    $path = $request.Url.LocalPath
    if ($path -eq "/" -or [string]::IsNullOrEmpty($path)) { $path = "/index.html" }

    $filePath = Join-Path $PSScriptRoot $path.TrimStart('/')

    if (Test-Path $filePath) {
      try {
        $bytes = [System.IO.File]::ReadAllBytes($filePath)
        $ext = [System.IO.Path]::GetExtension($filePath).ToLowerInvariant()
        $contentType = $MimeMap[$ext]
        if (-not $contentType) { $contentType = "application/octet-stream" }
        $response.ContentType = $contentType
        # Basic cache headers for static content
        $response.AddHeader("Cache-Control", "no-cache, no-store, must-revalidate")
        $response.AddHeader("Pragma", "no-cache")
        $response.AddHeader("Expires", "0")
        $response.ContentLength64 = $bytes.Length
        $response.OutputStream.Write($bytes, 0, $bytes.Length)
      } catch {
        $response.StatusCode = 500
        $msg = [System.Text.Encoding]::UTF8.GetBytes("Internal server error: $($_.Exception.Message)")
        $response.OutputStream.Write($msg, 0, $msg.Length)
      }
    } else {
      $response.StatusCode = 404
      $body = "<!doctype html><meta charset=`"utf-8`"><title>404</title><body style='font-family:Segoe UI,Roboto,Arial; padding:24px'><h1>404 Not Found</h1><p>$([System.Web.HttpUtility]::HtmlEncode($path))</p></body>"
      $msg = [System.Text.Encoding]::UTF8.GetBytes($body)
      $response.ContentType = "text/html; charset=utf-8"
      $response.ContentLength64 = $msg.Length
      $response.OutputStream.Write($msg, 0, $msg.Length)
    }

    $response.Close()
  }
} finally {
  $http.Stop()
}
