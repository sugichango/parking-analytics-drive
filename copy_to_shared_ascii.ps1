param(
    [string]$DestPath
)

Write-Host "Script started"

# Local source
$src = "C:\Users\sugitamasahiko\Documents\parking_system\downloads"

# Destination from argument
$dest = $DestPath

# Log file
$logpath = "C:\Users\sugitamasahiko\Documents\parking_system\copy_to_shared.log"

Write-Host "Source: $src"
Write-Host "Dest  : $dest"
Write-Host "Start robocopy..."

# robocopy (mirror)
robocopy $src $dest /E /MIR /FFT /R:3 /W:5 /NP /LOG+:"$logpath"

Write-Host "Robocopy finished"

exit $LASTEXITCODE

