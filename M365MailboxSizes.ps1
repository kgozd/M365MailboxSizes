function connect_ms_apps {
  
    $credentials = Get-Credential -Message "Please enter your credentials for Microsoft 365"
    $password  = $credentials.Password
    $username  = $credentials.UserName
    $UserCredential = New-Object System.Management.Automation.PSCredential ($username, $password)

    Connect-AzureAD -Credential $UserCredential 
    Connect-ExchangeOnline -Credential $UserCredential 
    Clear-Host
}

function close_ms_sessions {
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-AzureAD -Confirm:$false
}

function get_user_data {
    $group_name = Read-Host  "Type group name"
    
    $MS365_group = (Get-AzureADGroup -SearchString $group_name)
    Write-Host "Full group name:" $MS365_group.DisplayName
    Write-Host "Group Email:" $MS365_group.Mail `n`n
    $Users = Get-AzureADGroupMember -All:$true -ObjectId  $MS365_group.ObjectId | Where-Object { $_.AssignedLicenses.Count -gt 0 -and $_.AccountEnabled -eq $true }

    $UserInformation = @()

    foreach ($User in $Users) {
        
        $user_total_mailbox_size = (Get-MailboxStatistics -Identity $User.UserPrincipalName).TotalItemSize
    
        $UserInfo = @{
            UserDisplayName    = $User.DisplayName
            UserMailboxSize = $user_total_mailbox_size
            UserMail           = $User.UserPrincipalName
            
        }

        $UserInformation += New-Object PSObject -Property $UserInfo 
       
    }

    return $UserInformation
}

function main{
    connect_ms_apps
    $dane_userów = get_user_data
    #  $dane_userów | ForEach-Object {
    #     $UserInfo = $_
    #    Write-Host $UserInfo.UserDisplayName, $UserInfo.UserMailboxSize
       
    # }
    $dane_userów
    $dane_userów   | Export-Csv  -Path ".\free_space_quota_ms365.csv" -NoTypeInformation -Encoding UTF8   
    close_ms_sessions




}
main
