function Parse-MOI {
    param (
        [string]$filename
    )

    $data = [System.IO.File]::ReadAllBytes($filename)
    $version = $data[0..1]
    $reversed_size = $data[5..2]
    $size = [BitConverter]::ToUInt32($reversed_size, 0)
    $reversed_year = $data[7..6]
    $year = [BitConverter]::ToUInt16($reversed_year,0)
    $month = $data[8]
    $day = $data[9]
    $hour = $data[0xa]
    $minutes = $data[0xb]
    $reversed_mili = $data[0xd..0xc]
    $milliseconds = [BitConverter]::ToUInt16($reversed_mili, 0)
    $seconds = $milliseconds / 1000
    Write-Host $data[0xc..0xd]
    
    return @{
        'version' = $version
        'size' = $size
        'year' = $year.ToString()
        'month' = $month.ToString().PadLeft(2, '0')
        'day' = $day.ToString().PadLeft(2, '0')
        'hour' = $hour.ToString().PadLeft(2, '0')
        'minutes' = $minutes.ToString().PadLeft(2, '0')
        'seconds' = $seconds.ToString().PadLeft(2, '0')
    }
}

if ($args.Length -eq 0 -or $args[0] -eq '-h' -or $args[0] -eq '--help') {
    Write-Host 'usage: convert.ps1 <directory>...'
    Write-Host '  Convert one or more directories full of MOI and MOD files into MP4 files with correct metadata'
    return
}

$ffmpegPath = "C:\Progra`m Files\ffmpeg\bin\ffmpeg.exe"

foreach ($dir in $args) {
    $moiFiles = Get-ChildItem -Path $dir -Filter "*.MOI"

    foreach ($moi in $moiFiles) {
        $data = Parse-MOI $moi.FullName
        Write-Host "$($moi.FullName): $($data)"

        $mod = $moi.FullName -replace "MOI$", "MOD"
        if (-not (Test-Path $mod)) {
            Write-Error "Video file doesn't exist: $mod"
            return
        }

        $mp4 = $moi.FullName -replace "MOI$", "mp4"
        if (Test-Path $mp4) {
            Write-Host "$mp4 already exists, skipping"
            continue
        }

        $ffmpegCommand = "& `'$ffmpegPath`' -i `"$mod`" -vcodec copy -acodec aac -metadata `"" + "creation_time=$($data['year'])-$($data['month'])-$($data['day'])T$($data['hour']):$($data['minutes']):$($data['seconds'])Z`" `"$mp4`""
        $newDate = Get-Date -Year $data['year'] -Month $data['month'] -Day $data['day'] -Hour $data['hour'] -Minute $data['minutes'] -Second $data['seconds']

        Invoke-Expression $ffmpegCommand
        # set file modification time
        (Get-Item $mp4).LastWriteTime = $newDate

        # set file creation time
        (Get-Item $mp4).CreationTime = $newDate

    }
}
