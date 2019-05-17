using namespace System.DirectoryServices

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true,ParameterSetName="LDAP")][Switch]$LDAP,
    [Parameter(Mandatory=$true,ParameterSetName="EDMS")][Switch]$EDMS,
    [Parameter(Mandatory=$true)][String]$HostFQDN,
    [Parameter(Mandatory=$true)][String]$SearchRootDN,
    [Parameter(Mandatory=$false)][String]$LDAPFilter = "(objectClass=*)",
    [Parameter(Mandatory=$false)][object]$PropertiesToLoad = @("description","displayName","name","objectClass","objectGUID","objectSid","whenChanged","whenCreated","edsvaNamingContextDN","distinguishedName"),
    [Parameter(Mandatory=$false)][String]$LogFile = "C:\Temp\ADSITest.log",
    [Parameter(Mandatory=$false)][Switch]$ShowLog,
    [Parameter(Mandatory=$false)][pscredential]$Credential
    
)

function WriteLog($type, $data){
    $time = (Get-Date -Format "yyyy/MM/dd hh:mm:ss:fff tt")
    $logEntry = $time.ToString() + " | " + $type + " " + "$data"
    if ($ShowLog){
        Write-Host $logEntry
    }
    $logEntry | Out-File -FilePath $LogFile -Append
}

function WriteProgress($currentCount, $totalCount){
    [int]$percentComplete = (($currentCount / $totalCount) * 100)
    Write-Progress -Activity "Getting Directory Entries" -Status ("$percentComplete% Complete:") -PercentComplete $percentComplete
}

function Main(){
    WriteLog "[INFO]" "-------------------------BEGIN-------------------------"
    if ($EDMS){
        if ($Credential){
            $RootDirectory = New-Object DirectoryEntry("EDMS://$HostFQDN/$SearchRootDN", $Credential.UserName, $Credential.GetNetworkCredential().Password)
        } else {
            $RootDirectory = New-Object DirectoryEntry("EDMS://$HostFQDN/$SearchRootDN")
        }
    } elseif ($LDAP){
        if ($Credential){
            $RootDirectory = New-Object DirectoryEntry("LDAP://$HostFQDN/$SearchRootDN", $Credential.UserName, $Credential.GetNetworkCredential().Password)
        } else {
            $RootDirectory = New-Object DirectoryEntry("LDAP://$HostFQDN/$SearchRootDN")
        }
    } else {
        return
    }


    if (!$RootDirectory.Username){
        WriteLog "[ERROR]" "Access Denied"
        return
    }

    WriteLog "[INFO]" ("Logged in as: " + $RootDirectory.Username)
    WriteLog "[ADSI]" $RootDirectory.Path

    WriteLog "[FILTER]" $LDAPFilter
    $Search = [DirectorySearcher]($RootDirectory)

    if ($EDMS){
        $Search.PageSize = 0
    } elseif ($LDAP){
        $Search.PageSize = 100000000
    }

    $Search.Filter = $LDAPFilter

    $Search.PropertiesToLoad.Clear() | Out-Null
    
    foreach($property in $PropertiesToLoad){
        $Search.PropertiesToLoad.Add($property) | Out-Null
        WriteLog "[PROPERTY]" $property
    }

    WriteLog "[INFO]" "Querying directory..."
    try {
        $Collection = $Search.FindAll()
        
    } catch {
        WriteLog "[ERROR]" $_
        if ($EDMS){
            WriteLog "[ERROR]" "Make sure you have the Active Roles ADSI Provider installed and you can access to the Active Roles service."
        }
        $_
        return
    }
    WriteLog "[INFO]" "Querying directory COMPLETE."

    WriteLog "[INFO]" "Getting directory entries..."

    $totalCount = $Collection.Count
    $currentCount = 1

    foreach($Object in $Collection){
        WriteLog "[DIRECTORY OBJECT]" $Object.Path

        WriteProgress $currentCount $totalCount
        $currentCount++

        try {
            $TempObject = $Object.GetDirectoryEntry()
            WriteLog "       \_____[SUCCESS]" $TempObject.distinguishedName
            $TempObject.Dispose()
        } catch {
            WriteLog "       \_____[ERROR]" $_
            $_
        }
    }

    $RootDirectory.Dispose()
    $Search.Dispose()
    WriteLog "[INFO]" "Getting directory entries COMPLETE."
}

$timer = [system.diagnostics.stopwatch]::StartNew()

Main

$timer.Stop()
WriteLog "[ELAPSED]" $timer.Elapsed