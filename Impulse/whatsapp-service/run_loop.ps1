while ($true) {
    Write-Host "Starting WhatsApp Microservice..."
    node server.js
    $exitCode = $LASTEXITCODE
    Write-Host "Node exited with code $exitCode"
    Write-Host "Cleaning up auth directory to prevent EBUSY corruption..."
    Remove-Item -Path ".wwebjs_auth" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "Restarting in 3 seconds..."
    Start-Sleep -Seconds 3
}
