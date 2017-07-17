

#Create Credential Global Variables
$global:UserName;
$global:SecurePassword;
$global:Credential;

#Create O365 global Session Variables
$global:PSOutlookSession;
$global:importPSOutlookSession;

#Get SID
$objUser = New-Object System.Security.Principal.NTAccount($env:USERNAME)
$strSID = $objUser.Translate([System.Security.Principal.SecurityIdentifier])
$global:userSID = $strSID.Value

#Sets Password and date in registry
function SplunkO365Init{
    #User Name with O365 Get-MessageTrace Permissions 
    $UserName = Read-Host -Prompt "Enter UserName"
    
    #Prompt for password, convert secure string to store in registry
    $SecurePassword = Read-Host -Prompt "Enter Password" -AsSecureString
    $SecStringAsPlainText = $SecurePassword | ConvertFrom-SecureString

    #Create Registry Items for Main to use. 
    New-Item -Path Registry::HKEY_USERS\$($global:userSID)\Software -Name O365 -Force 
    New-ItemProperty -Path Registry::HKEY_USERS\$($global:userSID)\Software\O365 -Name SECSTRING -Value $SecStringAsPlainText -PropertyType String
    New-ItemProperty -Path Registry::HKEY_USERS\$($global:userSID)\Software\O365 -Name USERNAME -Value $UserName -PropertyType String
    New-ItemProperty -Path Registry::HKEY_USERS\$($global:userSID)\Software\O365 -Name STARTDATE -Value $(get-date).AddMinutes(-15) -PropertyType String
}

# Open a session with exchange
function Enter-ExchangeSession{
<#
    .SYNOPSIS

    Enter an Exchange Session. 

    .DESCRIPTION
    
    This function creates a remote session with "ps.outlook.com/powershell" and the Security and Compliance Center. Imports it globally allowing the user to interact with Exchange Online in the current PS Window. Imports Compliance Search Specific Commands, Get-Mailbox, Search-Mailbox, and Get-MessageTrace.  

    .EXAMPLE

    Enter-ExchangeSession
       
    Create a session with Exchange


#>

    Param(
        [System.Object]$Credential = $(Get-Credential -Message "Enter your Office 365 Credential")
    )
        #Set Session Options. Proxy and 10 Minute TimeOut
        $soProxySettings = $($(New-PSSessionOption -IdleTimeout 600000),$global:proxysettings)
        $so = New-PSSessionOption -IdleTimeout 600000

        #PowerShell Session URL's
        $PSOutlookURL = "https://ps.outlook.com/powershell"
        
        #Check to see if System Proxy Flag is Set.    
        #Create new remote powershell session with exchange
        $global:PSOutlookSession = New-PSSession -Name "PSOutlook" -ConfigurationName Microsoft.Exchange -ConnectionUri $PSOutlookURL -Credential $Credential -WarningAction SilentlyContinue -Authentication Basic -AllowRedirection -SessionOption $so -ErrorAction Stop
        
        #Specify import command names
        $tmpPSOutlookCommandNames = @("Get-MessageTrace")
 
        #import PSSessions
        Import-Module ($global:importPSOutlookSession = Import-PSSession $global:PSOutlookSession -CommandName $tmpPSOutlookCommandNames -AllowClobber -DisableNameChecking -ErrorAction Stop) -Global -DisableNameChecking -Force 
 
}

# Close a session with exchange
function Remove-ExchangeSession{
<#
    .SYNOPSIS
    Remove current Exchange Session based on Session Name: PSSecComp and PSOutlook.

    .DESCRIPTION
    
    This function removes the remote session specified in $session variable.
 
    .EXAMPLE
    Exit-ExchangeSession
       
#>
    #Remove TMP_* Modules connected to Outlook 
    Get-Module tmp_* |% {if($_.Description -match "outlook.com" ) {Write-Output "Remove Module $($_.Name)" ; remove-module $_.name}}
    #Remove PSSession's that Match Outlook.com
    Get-PSSession | % { if( $_.ComputerName -match "outlook.com" ){Write-Output "Remove PSSession $($_.Name)"; Remove-PSSession -Id $_.id}  } 
}

function Reset-ExchangeSessionState{
    Param(
        [System.Object]$Credential = $(Get-Credential -Message "Enter your Office 365 Credential")
    )    

    #end session with Microsoft Exchange
    if(Get-PSSession){
        Get-PSSession | % { 
            if( $_.ComputerName -match "outlook.com" ){
                if($_.State -notcontains "Opened"){
                    #Remove-ExchangeSession if they exist. 
                    Remove-ExchangeSession 
                    Enter-ExchangeSession -Credential $Credential
                
                    break;
                }
                else{
                    write-output "`"$($_.Name)`" is Open." 
                }
            }
            else{
                Enter-ExchangeSession -Credential $Credential
                break;
            }       
        }
    }
    else{
        Enter-ExchangeSession -Credential $Credential
        break;
    }
}

#Remove Old Log Files
function Remove-OldTranscripts {

    Param(
        $TRANSCRIPTPATH = ""
    )

    #Remove Log Items Greater Than 30-days Old
    Get-ChildItem -Path $TRANSCRIPTPATH  | where LastWriteTime -LT $(get-date).AddDays(-30) | ForEach-Object {
        Remove-Item $_.FullName -Force 
    }
}

#Collect MessageTrace logs and output in JSON file
function SplunkMessageTrace{
    #Get Start Date from Registry, Set End Date
    $StartDate = get-date -date $(Get-ItemPropertyValue -Path Registry::HKEY_USERS\$($global:userSID)\Software\O365 -Name StartDate)
    $EndDate = $(Get-Date).AddMinutes(-10)

    #Initiate Page
    $Page = 0 

    #This is to keep track of how long each session takes
    $start = get-date

    do{
        #Increment MessageTrace Page
        $Page ++
        
        #Output current page number
        Write-Host -ForegroundColor Yellow Starting Page $page
        
        #ensure that session is open
        Reset-ExchangeSessionState -Credential $Credential

        #Set output path 
        $Path = "C:\Scripts\data\Office365\$($Page)_MessageTrace.json"

        #Test if output path exists. If exists, Delete/Recreate or create path.
        if(Test-Path -path $Path){
            Remove-Item $Path -Force 
            New-Item -Path $Path -ItemType File | Out-Null
            }
        else{
            New-Item -Path $Path -ItemType File | Out-Null
        }

        #Start StreamWriter
        $sw = new-object system.IO.StreamWriter($Path)
                
        #Collect Messages from server
        Get-MessageTrace -PageSize 5000 -Page $Page -StartDate $StartDate.ToUniversalTime() -EndDate $EndDate.ToUniversalTime() | ForEach-Object {
                #Convert output to JSON and store in file                 
                $sw.WriteLine($(@{
                        "source"="MessageTrace";
                        "host"=$_.organization;
                        "runspace_id"=$_.RunspaceId;
                        "sender"=$_.SenderAddress;
                        "subject"=$_.Subject;
                        "recipient"=$_.RecipientAddress;
                        "action"=$_.Status;
                        "dest_ip"=$_.ToIP;
                        "src_ip"=$_.FromIP;
                        "@time"=[math]::round($(get-date -date $_.Received -UFormat "%s"));
                        "size"=$_.Size;
                        "message_id"=$_.MessageId;
                        "message_trace_id"=$_.MessageTraceId
                        "@timezone"=$(get-date -date $_.Received -UFormat "%Z");
                    } | ConvertTo-Json -Compress)
                )

                    
            }


        #Close Stream Writer
        $sw.close()

        #Output total time for collection
        $end = get-date 
        $total = $end - $start
        $total

    } #Run until Path is empty. When path is empty there are no longer results returning from MessageTrace for that time period
    until ($(Get-Item $Path).Length -eq 0)

    #update registry date
    Set-ItemProperty -Path Registry::HKEY_USERS\$($global:userSID)\Software\O365 -Name STARTDATE -Value $EndDate
}

#Main function
function main{

    #Create Credential Object
    $global:UserName = Get-ItemPropertyValue -Path Registry::HKEY_USERS\$($global:userSID)\Software\O365 -Name USERNAME
    $global:SecurePassword = Get-ItemPropertyValue -Path Registry::HKEY_USERS\$($global:userSID)\Software\O365 -Name SECSTRING | ConvertTo-SecureString 
    $global:Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $SecurePassword

    #Create Log Path Variable
    $TRANSCRIPTPATH = 'C:\scripts\Office365\log\' 
    $NEWTRANSCRIPT = $TRANSCRIPTPATH + $(get-date -Format "yyyyMMdd") + '_messagetrace.log'
    Remove-OldTranscripts -TRANSCRIPTPATH $TRANSCRIPTPATH 
    
    $TaskCount = 1 

    #Loop 30 times
    while($TaskCount -le 30){
        #Start Logging 
        Start-Transcript -Path $NEWTRANSCRIPT -Append -Force 
        
        #Run Message Trace
        SplunkMessageTrace

        #Output Complete
        write-host -ForegroundColor Green Complete Task $TaskCount

        #Sleep for 90 seconds (This is to allow the Splunk forward time to collect the logs and close the file before the script restarts)
        Start-Sleep 90

        $TaskCount++

        #End Logging 
        Stop-Transcript
    }

    #Convert Password and save to Registry
    $secString = $Credential.Password | ConvertFrom-SecureString 
    Set-ItemProperty -Path Registry::HKEY_USERS\$($global:userSID)\Software\O365 -Name SECSTRING -Value $secString


}

#Initialization : Review Splunk O365 Init function to validate Correct settings 
#SplunkO365Init

#SplunkMessageTrace Main 
#main
