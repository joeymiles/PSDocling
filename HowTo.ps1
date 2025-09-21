
Import-Module .\DoclingSystem.psm1

# Quick start
Initialize-DoclingSystem -GenerateFrontend
Start-DoclingSystem -OpenBrowser

Add-DocumentToQueue -Path "C:\Users\Jaga\Downloads\4bde0d00-1a4f-4e13-8438-498ee1fe8f6b.pdf"

Get-NextQueueItem  

Get-DoclingSystemStatus



