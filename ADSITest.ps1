using namespace System.DirectoryServices

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true,ParameterSetName="LDAP")][Switch]$LDAP,
    [Parameter(Mandatory=$true,ParameterSetName="EDMS")][Switch]$EDMS,
    [Parameter(Mandatory=$true)][String]$HostFQDN,
    [Parameter(Mandatory=$true)][String]$SearchRootDN,
    [Parameter(Mandatory=$false)][String]$LDAPFilter = "(objectClass=*)",
    [Parameter(Mandatory=$false)][object]$PropertiesToLoad = @("description","displayName","name","objectClass","objectGUID","objectSid","whenChanged","whenCreated","distinguishedName"),
    [Parameter(Mandatory=$false)][switch]$LogPropertyValues,
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

    if (!$RootDirectory){
        WriteLog "[ERROR]" "Unable to connect."
        return
    }

    if ($Credential){
        WriteLog "[INFO]" ("Logged in as: " + $RootDirectory.Username)
    } else {
        WriteLog "[INFO]" ("Logged in as: " + ([Security.Principal.WindowsIdentity]::GetCurrent().Name))
    }
    
    WriteLog "[ADSI]" $RootDirectory.Path

    WriteLog "[FILTER]" $LDAPFilter
    $Search = New-Object DirectorySearcher($RootDirectory)

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
        $LifetimeService = $Collection.InitializeLifetimeService()
        $Timespan = New-Object Timespan(10,0,0)
        $LifetimeService.Renew($Timespan) | Out-Null
        WriteLog "[INFO]" ("Lease time: " + $LifetimeService.CurrentLeaseTime)
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

    try{
        foreach($Object in $Collection){
            WriteLog "[DIRECTORY OBJECT]" $Object.Path

            WriteProgress $currentCount $totalCount
            $currentCount++

            try {
                WriteLog "       \_____[SUCCESS]" ""
                if ($LogPropertyValues){
                    foreach($prop in $PropertiesToLoad){
                        WriteLog "           \_____[PROPERTY]" ($prop + " : " + $Object.Properties.($prop.toLower()))
                    }
                }
            } catch {
                WriteLog "       \_____[ERROR]" $_
                $_
            }
        }
    } catch {
        WriteLog "       \_____[ERROR]" $_
        $_
    }
    
    $Collection.Dispose()
    $RootDirectory.Dispose()
    $Search.Dispose()
    WriteLog "[INFO]" "Getting directory entries COMPLETE."
}

$timer = [system.diagnostics.stopwatch]::StartNew()

Main

$timer.Stop()
WriteLog "[ELAPSED]" $timer.Elapsed