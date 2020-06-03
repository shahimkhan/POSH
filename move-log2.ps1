function move-log { 
<# 
    .NOTES  
        Name:   Shahim Khan
        Date:   10-Apr-2017  
        .DESCRIPTION 
         The script accomplished the following tasks:
			1. Move Audit trace files older than 3 days from the default SQL data directory to \\eqxim01\SQLTraces
			2. Use 7-zip to compress the recently copied files into a single highly compressed archive
			3. delete the moved tracefiles & clear the space.
  
#> 
    [CmdletBinding()] 
    Param 
    ( 
        [Parameter(Mandatory=$false)][ValidateScript({ Test-Path $_ -PathType Container })] 
        [string]$LogPath="E:\Program Files\Microsoft SQL Server\MSSQL10_50.SQLSAPB1\MSSQL\DATA", 
        [Parameter(Mandatory=$false)][string]$ArchivePath="\\eqxim01\SQLTraces\", 
        [Parameter(Mandatory=$false)][int]$daysBack="2" 
    ) 
 
    Begin 
    { 
       
		$tz = "C:\Program Files\7-Zip\7z.exe" 
        Set-Alias sz $tz 
        $staging = "\\eqxim01\SQLTraces\DATA\*.trc" 
        Write-Verbose "Checking to see if archive folder exists" 
        if (!(Test-Path $Archivepath)) { 
            Write-Verbose "Archive folder doesn't exist. It will be created now" 
            try { 
               $null=New-Item $Archivepath -ItemType directory -Force -ErrorAction Stop -ErrorVariable DirectoryError 
            } 
            Catch { 
                write-Error "An error occurred created the archive folder $Archivepath. Error: $DirectoryError" 
            } 
        } 
    } 
    Process 
    { 
        Write-Verbose "Starting the process block"        
        $refDate=(Get-Date).AddDays(-$daysBack) 
        $oldFiles=Get-ChildItem -filter *.trc -Path $LogPath | where {$_.LastWriteTime -lt "$refDate"} 
        foreach ($oldFile in $oldFiles) { 
                    Write-Output "Adding Permissions for $oldFile" 
                    cacls $oldFile /E /G mesoblastltd\labtech:F /T 
        $parentFolder = $oldFile.Directory.Name 
        $ArchiveFolder=join-path -path $ArchivePath -childpath $parentFolder 
             if (!(Test-Path $ArchiveFolder)) { 
                 try 
                { 
                 $null=New-Item $ArchiveFolder -ItemType directory -Force -ea stop -ErrorVariable ArchDirError 
                } 
                Catch 
                { 
                    Write-Error "Error moving file $oldfile to destination $ArchiveFolder. Error: $ArchDirError"     
                 } 
            } 
            try  
            { 
                Move-Item "$($oldFile.DirectoryName)\$oldFile" $ArchiveFolder -Force -ea stop -ErrorVariable MoveError 
            } 
            Catch  
            { 
                Write-Error "Error moving file $oldfile to destination $ArchiveFolder. Error: $MoveError" 
            } 
        } 
    } 
    End 
    { 
     $nodename = [system.environment]::MachineName
     $date = (Get-Date).tostring("dd-MM-yyyy")
     $topath = $ArchivePath + $nodename + "_" + $date + "_SQLTrace.zip"
    sz a -mmt=off $topath $staging
    Remove-Item $staging -Force 
    } 
} 
move-log 