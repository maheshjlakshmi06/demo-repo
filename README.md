# Demo

some description

function accountExpiresToString($accountExpires) {
    if (($_.AccountExpires -eq 0) -or 
        ($_.AccountExpires -eq [int64]::MaxValue)) {
        "Never expires"
    }
    else {
        [datetime]::fromfiletime($accountExpires)
    }
}

SamAccountName,PasswordLastSet,employeeID,msExchUsageLocation,Description,DistinguishedName,DisplayName,EmailAddress,UserPrincipalName,Enabled,whenCreated,whenChanged,GivenName,sn,employeeType,extensionAttribute1,extensionAttribute2,extensionAttribute3,extensionAttribute4,@{l="AccountExpires";e={ accountExpiresToString($_.AccountExpires)}},@{n='LastLogon';e={[DateTime]::FromFileTime($_.LastLogon)}},LastLogonDate,objectSid,@{ Name = 'msExchArchiveGUID'; Expression = { $_.msExchArchiveGUID -join ''; }; },@{ Name = 'msExchArchiveName'; Expression = { $_.msExchArchiveName -join ''; }; },msExchArchiveStatus,msDS-ExternalDirectoryObjectId,Title,Office,Department,physicalDeliveryOfficeName,Company,@{name="MemberOf";expression={$_.memberof -join ";"}},Name,@{N='Manager';E={(Get-ADUser $_.Manager).userPrincipalName}},@{L='ProxyAddress_1';E={$_.proxyaddresses[0]}},@{L='PrxyAddress_2';E={$_.ProxyAddresses[1]}},@{L='ProxyAddress_3';E={$_.ProxyAddresses[2]}},@{L='ProxyAddress_4';E={$_.ProxyAddresses[3]}},@{L='ProxyAddress_5';E={$_.ProxyAddresses[4]}},@{L='ProxyAddress_6';E={$_.ProxyAddresses[5]}},@{L='ProxyAddress_7';E={$_.ProxyAddresses[6]}},@{L='ProxyAddress_8';E={$_.ProxyAddresses[7]}},@{L='ProxyAddress_9';E={$_.ProxyAddresses[8]}},@{L='ProxyAddress_10';E={$_.ProxyAddresses[9]}},@{L='ProxyAddress_11';E={$_.ProxyAddresses[10]}},@{L='ProxyAddress_12';E={$_.ProxyAddresses[11]}},@{L='ProxyAddress_13';E={$_.ProxyAddresses[12]}},@{L='ProxyAddress_14';E={$_.ProxyAddresses[13]}}


$Accountstatus = $res1 | Group-Object enabled | select Name,Count
$OfficeSite = $res1 | Group-Object Office | select Name,Count
$Accounts = $Accountstatus | ft -Wrap -AutoSize | Out-String
$office = $OfficeSite | ft -Wrap -AutoSize | Out-String

Send-MailMessage 
