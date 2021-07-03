$oldPrefix = "?:\???\Documents"
$newPrefix = "%Documents%"

$searchPath = "D:\Scott\Documents\Bible\Sermons\Links - New"

$dryRun = $TRUE

$shell = new-object -com wscript.shell

if ( $dryRun ) {
   write-host "Executing dry run" -foregroundcolor green -backgroundcolor black
}
else {
   write-host "Executing real run" -foregroundcolor red -backgroundcolor black
}

Get-ChildItem $searchPath -filter *.lnk -recurse | ForEach-Object {
   $lnk = $shell.createShortcut( $_.fullname )
   $oldPath = $lnk.targetPath
   $oldWork = $lnk.workingDirectory

   $lnkRegex = "^" + [regex]::escape( $oldPrefix )

   if ( ( $oldPath -match $lnkRegex ) -or ( $oldWork -match $lnkRegex ) ) {
      write-host "Found: " + $_.fullname -foregroundcolor yellow -backgroundcolor black
      if ( $oldPath -match $lnkRegex ) {
         $newPath = $oldPath -replace $lnkRegex, $newPrefix

         write-host " Target Path:" -foregroundcolor green -backgroundcolor black
         write-host "  Replace: " + $oldPath
         write-host "  With:    " + $newPath

         if ( !$dryRun ) {
            $lnk.targetPath = $newPath
         }
      }
      if ( $oldWork -match $lnkRegex ) {
         $newWork = $oldWork -replace $lnkRegex, $newPrefix

         write-host " Working Directory:" -foregroundcolor green -backgroundcolor black
         write-host "  Replace: " + $oldWork
         write-host "  With:    " + $newWork

         if ( !$dryRun ) {
            $lnk.workingDirectory = $newWork
         }
      }
      if ( !$dryRun ) {
         $lnk.Save()
      }
   }
}
