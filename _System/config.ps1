<#
 Shared Configuration for iCloud Photos Organizer scripts
#>

$sources = @(
    @{ Path = "C:\Users\User1\Pictures\iCloud Photos\Photos";  Person = "User1" },
    @{ Path = "C:\Users\User2\Pictures\iCloud Photos\Photos"; Person = "User2" },
    @{ Path = "C:\Users\User3\Pictures\iCloud Photos\Photos";  Person = "User3" }
)

$dest             = "D:\Photos\Phone Media"
$sysDir           = "D:\Photos\Phone Media\_System"
$tempDir          = "$sysDir\Temp"
$logsDir          = "$sysDir\Logs"
$dataDir          = "$sysDir\Data"
$logFile          = "$dataDir\processed.log"
$hashLog          = "$dataDir\hashes.csv"
$errorLog         = "$logsDir\errors.log"
$slowLog          = "$logsDir\slow-downloads.log"
$exiftool         = "C:\Tools\Exiftool\exiftool.exe"
$maxWaitSecPhoto  = 120  # 2 minutes for photos
$maxWaitSecVideo  = 600  # 10 minutes for videos
$processFileLimit = 2000

$photoExt = @(".jpg",".jpeg",".heic",".png",".gif")
$videoExt = @(".mov",".mp4",".m4v")
