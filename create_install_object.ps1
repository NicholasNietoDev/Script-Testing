<#
.SYNOPSIS
  Select installer(s), test common silent switches, and produce JSON manifest entries.

.PARAMETER LocalDir
  Where to copy or extract all files for local execution.

.PARAMETER OutputFile
  File to which individual JSON objects will be appended.

#>
[CmdletBinding()]
param(
    [string]$LocalDir   = (Join-Path -Path "{SET PATH}" -ChildPath 'InstallerDownloads'),
    [string]$OutputFile = (Join-Path -Path $LocalDir -ChildPath 'install_manifest_entries.txt')
)

# Ensure directory exists
if (-not (Test-Path $LocalDir)) { New-Item -Path $LocalDir -ItemType Directory | Out-Null }
# Prepare output file
if (-not (Test-Path $OutputFile)) { New-Item -Path $OutputFile -ItemType File | Out-Null }

Add-Type -AssemblyName System.Windows.Forms

function Find-ContainedInstallerInZip {
    param(
        [string]$ZipPath,
        [string]$OutDir
    )
    $extractDir = Join-Path $OutDir ([IO.Path]::GetFileNameWithoutExtension($ZipPath))
    Expand-Archive -Path $ZipPath -DestinationPath $extractDir -Force
    # look for the first .exe/.msi/.ps1/.bat inside
    $found = Get-ChildItem -Path $extractDir -Recurse |
             Where-Object { $_.Extension -in '.exe','.msi','.ps1','.bat' } |
             Select-Object -First 1
    return $found?.FullName
}

function Test-SilentSwitches {
    param(
        [string]$InstallerPath,
        [string[]]$Switches,
        [string]$UserArgs,
        [switch]$IsMsi
    )
    foreach ($s in $Switches) {
        Write-Host "`nTesting silent switch: $s" -ForegroundColor Cyan
        Start-Sleep -Seconds 2  # give you time to prepare
        if ($IsMsi) {
            $args = "/i `"$InstallerPath`" $s $UserArgs"
            Start-Process -FilePath msiexec.exe -ArgumentList $args -Wait
        } else {
            $args = "$s $UserArgs"
            Start-Process -FilePath $InstallerPath -ArgumentList $args -Wait -NoNewWindow
        }
        $resp = Read-Host "Did that run completely silently (no UI)? (Y/N)"
        if ($resp.Trim().ToUpper() -eq 'Y') {
            return $s
        }
    }
    return ''
}

# Launch file picker
$ofd = New-Object System.Windows.Forms.OpenFileDialog
$ofd.InitialDirectory = [Environment]::GetFolderPath('Desktop')
$ofd.Multiselect      = $true
$ofd.Filter           = 'All files (*.*)|*.*'
if ($ofd.ShowDialog() -ne 'OK') {
    Write-Error "No file selected. Exiting."
    exit
}

foreach ($file in $ofd.FileNames) {
    $ext   = [IO.Path]::GetExtension($file).ToLower()
    $name  = [IO.Path]::GetFileNameWithoutExtension($file)
    $drive = Split-Path -Path $ofd.FileName -Qualifier
    $fullPath = Split-Path -Path $ofd.FileName -NoQualifier
    $driveData = ((& net use $drive | Select-String -Pattern 'Remote name') -split '\s+', 3)[2]

    $fqdn  = (Join-Path -Path $driveData -ChildPath $fullPath)

    # Copy or extract into LocalDir
    switch ($ext) {
        '.zip' {
            Write-Host "Extracting ZIP to find installer..." -ForegroundColor Yellow
            $installer = Find-ContainedInstallerInZip -ZipPath $file -OutDir $LocalDir
            if (-not $installer) {
                Write-Warning "No installer found inside $file skipping."
                continue
            }
        }
        default {
            $installer = Join-Path $LocalDir ([IO.Path]::GetFileName($file))
            Copy-Item -Path $file -Destination $installer -Force
        }
    }

    # Prompt for extra args
    $userArgs = Read-Host "Enter any additional arguments (or press Enter for none)"

    # Decide silent-switch list
    if ($installer -match '\.msi$') {
        $switches = '/qn','/quiet'
        $isMsi    = $true
    } elseif ($installer -match '\.exe$') {
        $switches = '/S','/silent','/verysilent','/quiet','/qn'
        $isMsi    = $false
    } else {
        $switches = @()
        $isMsi    = $false
    }

    # Test those switches
    $chosenSwitch = ''
    if ($switches.Count -gt 0) {
        $chosenSwitch = Test-SilentSwitches -InstallerPath $installer `
            -Switches $switches -UserArgs $userArgs -IsMsi:$isMsi
    }

    $friendlyName = Read-Host "Enter Friendly Name (or press Enter for none)"

    # Build JSON object
    $obj = [PSCustomObject]@{
        Name         = $name
        FriendlyName = $friendlyName
        Type         = 'Installer'
        FQDN         = $fqdn
        Arguments    = ($chosenSwitch + ' ' + $userArgs).Trim()
    }

    # Append to output file (one JSON object per line, with trailing comma)
    $jsonLine = ($obj | ConvertTo-Json -Depth 2) + ','
    $jsonLine | Out-File -FilePath $OutputFile -Append -Encoding UTF8

    Write-Host "Appended manifest entry for '$name' to $OutputFile"
}

Write-Host "All done."
