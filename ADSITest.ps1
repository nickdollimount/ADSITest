using namespace System.DirectoryServices

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true,ParameterSetName="LDAP")][Switch]$LDAP,
    [Parameter(Mandatory=$true,ParameterSetName="EDMS")][Switch]$EDMS,
    [Parameter(Mandatory=$true)][String]$HostFQDN,
    [Parameter(Mandatory=$true)][String]$SearchRootDN,
    [Parameter(Mandatory=$false)][String]$LDAPFilter = "(objectClass=*)",
    [Parameter(Mandatory=$false)][String]$LogFile = "C:\Temp\ADSITest.log",
    [Parameter(Mandatory=$false)][Switch]$ShowLog
    
)

function WriteLog($type, $data){
    $time = (Get-Date -Format "yyyy/MM/dd hh:mm:ss:fff tt")
    $logEntry = $time.ToString() + " | " + $type + " " + "$data"
    if ($ShowLog){
        Write-Host $logEntry
    }
    $logEntry | Out-File -FilePath $LogFile -Append
}

function Execute(){
    if ($EDMS){
        $RootDirectory = [DirectoryEntry]("EDMS://$HostFQDN/$SearchRootDN")
        WriteLog "[ADSI]" "EDMS://$HostFQDN/$SearchRootDN"
    } elseif ($LDAP){
        $RootDirectory = [DirectoryEntry]("LDAP://$HostFQDN/$SearchRootDN")
        WriteLog "[ADSI]" "LDAP://$HostFQDN/$SearchRootDN"
    }
        
    WriteLog "[FILTER]" $LDAPFilter
    $Search = [DirectorySearcher]($RootDirectory)

    if ($EDMS){
        $Search.PageSize = 0
    } elseif ($LDAP){
        $Search.PageSize = 100000000
    }

    $Search.Filter = $LDAPFilter

    $Search.PropertiesToLoad.Clear() | Out-Null
    $Search.PropertiesToLoad.Add("description") | Out-Null
    $Search.PropertiesToLoad.Add("displayName") | Out-Null
    $Search.PropertiesToLoad.Add("name") | Out-Null
    $Search.PropertiesToLoad.Add("objectClass") | Out-Null
    $Search.PropertiesToLoad.Add("objectGUID") | Out-Null
    $Search.PropertiesToLoad.Add("objectSid") | Out-Null
    $Search.PropertiesToLoad.Add("whenChanged") | Out-Null
    $Search.PropertiesToLoad.Add("whenCreated") | Out-Null
    $Search.PropertiesToLoad.Add("edsvaNamingContextDN") | Out-Null
    $Search.PropertiesToLoad.Add("distinguishedName") | Out-Null

    WriteLog "[INFO]" "Querying directory..."
    try {
        $Collection = $Search.FindAll()
        
    } catch {
        WriteLog "[ERROR]" $_
        $_
        return
    }
    WriteLog "[INFO]" "Querying directory COMPLETE."

    WriteLog "[INFO]" "Getting directory entries..."
    foreach($Object in $Collection){
        WriteLog "[DIRECTORY OBJECT]" $Object.Path
        try {
            $TempObject = $Object.GetDirectoryEntry()
            WriteLog "       \_____[SUCCESS]" $TempObject.distinguishedName
            $TempObject.Dispose()
        } catch {
            WriteLog "       \_____[ERROR]" $_
            $_
        }
    }

    $Search.Dispose()
    
    WriteLog "[INFO]" "Getting directory entries COMPLETE."
}

Execute