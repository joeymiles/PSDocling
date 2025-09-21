param([int]$Port = 8081)

$http = New-Object System.Net.HttpListener
$http.Prefixes.Add("http://localhost:$Port/")
$http.Start()

Write-Host "Web server running at http://localhost:$Port" -ForegroundColor Green

try {
    while ($http.IsListening) {
        $context = $http.GetContext()
        $response = $context.Response

        $path = $context.Request.Url.LocalPath
        if ($path -eq "/") { $path = "/index.html" }

        $filePath = Join-Path $PSScriptRoot $path.TrimStart('/')

        if (Test-Path $filePath) {
            $content = [System.IO.File]::ReadAllBytes($filePath)
            $response.ContentType = "text/html"
            $response.ContentLength64 = $content.Length
            $response.OutputStream.Write($content, 0, $content.Length)
        } else {
            $response.StatusCode = 404
        }

        $response.Close()
    }
} finally {
    $http.Stop()
}
