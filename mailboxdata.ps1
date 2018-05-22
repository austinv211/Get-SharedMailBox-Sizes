<#
Name: mailboxdata.ps1
Description: gets mailbox statistics for a list of mailboxes
prerequisites: must be Exchange Online Admin
Author: Austin Vargason
Date Modified: 5/22/18
#>

function Connect-ExchangeOnline () {

    #if connecttion already exists do not do anything
    if ( (Get-PSSession | Select -ExpandProperty ConfigurationName) -contains "Microsoft.Exchange") {
        Write-Host "Already Connected :)" -ForegroundColor Cyan -BackgroundColor Black
    }
    else {
        Write-Host "Connecting to Exchange Online" -ForegroundColor Cyan -BackgroundColor Black

        #connect to Exchange Online
        $UserCredential = Get-Credential

        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication  Basic -AllowRedirection

        #import the session
        Import-PSSession $Session | Out-Null
    }
}

function Get-SharedMailBoxData () {

    param (
        [Parameter(Mandatory=$true)]
        [String]$filePath
    )

    #connect to Exchange Online
    Connect-ExchangeOnline

    #get the content for the mailboxes
    $file = Get-Content -Path $filePath

    #array to store the objects
    $resultArray = @()

    #set a counter
    $i = 0

    #get the mailbox data for each mailbox in the file
    foreach ($mailbox in $file) {
        #create a custom object
        $obj = New-Object -TypeName PSObject

        #save the name
        $name = $mailbox

        $size = Get-Mailbox -Identity $name -IncludeInactiveMailbox | 
                Get-MailboxStatistics | 
                Select -ExpandProperty TotalItemSize

        #add the results to the object
        $obj | Add-Member -Name "MailboxName" -Value $name -MemberType NoteProperty
        $obj | Add-Member -Name "TotalItemSize" -Value $size -MemberType NoteProperty

        #add the object to the result array
        $resultArray += $obj

        #increase the counter 
        $i++

        #write the progress
        Write-Progress -Activity "Getting Mailbox Data" -Status "Recieved Data for Mailbox: $name" -PercentComplete (($i/$file.Count) * 100)
        
    }


    #return the array
    return $resultArray
}

Get-SharedMailBoxData -filePath .\SharedmailBoxes.txt | Export-Csv -NoTypeInformation -Path .\mailboxDataResults.csv




