try {
    $ppt = New-Object -ComObject PowerPoint.Application
    Write-Host "PowerPoint Application is available!"
    $ppt.Quit()
} catch {
    Write-Host "PowerPoint COM error: $_"
}
