$global:datetime = $((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString())
$global:logfile = ".\log_$datetime.txt"
$global:hashfile = ".\md5hash_$datetime.txt"

function checkNeededModule() {
    $Modules=Get-Module -Name MSOnline -ListAvailable
    if($Modules.count -eq 0)
    {
    Write-Host  Please install MSOnline module using below command: `nInstall-Module MSOnline  -ForegroundColor yellow
    $confirm = Read-Host Are you sure you want to install module? [Y] Yes [N] No  
        if ($confirm -match "[yY]") { 
            Write-host "Installing MSOnline"
            Install-Module MSOnline
            Write-host "MSOnline module is installed in the machine successfully."
        }
        elseif ($confirm -cnotmatch "[yY]" ) { 
            Write-host "Exiting. `nNote: MSOnline PowerShell module must be available in your system to run the script." 
            Exit 
        }
    }
    $Exchange = (get-module ExchangeOnlineManagement -ListAvailable).Name
    if ($Exchange -eq $null) {
        Write-host "Important: ExchangeOnline PowerShell module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
        $confirm = Read-Host Are you sure you want to install module? [Y] Yes [N] No  
        if ($confirm -match "[yY]") { 
            Write-host "Installing ExchangeOnlineManagement"
            Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
            Write-host "ExchangeOnline PowerShell module is installed in the machine successfully."
        }
        elseif ($confirm -cnotmatch "[yY]" ) { 
            Write-host "Exiting. `nNote: ExchangeOnline PowerShell module must be available in your system to run the script." 
            Exit 
        }
    }
}

function loginMSOnline(){
    Get-PSSession | Remove-PSSession
    Import-Module MSOnline
    #Storing credential in script for scheduling purpose/ Passing credential as parameter
    if(($UserName -ne "") -and ($Password -ne ""))
    {
        $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
        $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
        Connect-MsolService -Credential $credential
    }
    else
    {
        Connect-MsolService | Out-Null
    }
    write-host "Logged in M365 at $datetime"
    "Logged in m365 $datetime" | Out-File -FilePath $logfile -Append

}

function ExportO365Users(){
    $OutputFile = ".\Office365Users_$datetime.csv"

    write-host "Exporting o365 users at $datetime"
    "Exporting o365 users at $datetime" | Out-File -FilePath $logfile -Append
    
    #Export user details to CSV.
    Get-MsolUser -All | Select DisplayName,UserPrincipalName, IsLicensed, BlockCredential | Export-CSV $OutputFile -NoTypeInformation -Encoding UTF8
    Get-FileHash $OutputFile | Format-Table -Wrap | Out-File -FilePath $hashfile -Append

    write-host "Finish Exporting o365 users at $datetime"
    "Finish Exporting o365 users at $datetime" | Out-File -FilePath $logfile -Append
}

function getMFAStatus(){
    $Result=""
    $Results=@()
    $UserCount=0
    $PrintedEnabledUser=0
    $PrintedDisabledUser=0

    #Output file declaration
    $ExportCSV=".\MFADisabledUserReport_$datetime.csv"
    $ExportCSVReport=".\MFAEnabledUserReport_$datetime.csv"

    #Loop through each user
    Get-MsolUser -All | foreach{
        $UserCount++
        $DisplayName=$_.DisplayName
        $Upn=$_.UserPrincipalName
        $MFAStatus=$_.StrongAuthenticationRequirements.State
        $MethodTypes=$_.StrongAuthenticationMethods
        $RolesAssigned=""
        Write-Progress -Activity "`n     Processed user count: $UserCount "`n"  Currently Processing: $DisplayName"
        if($_.BlockCredential -eq "True"){
            $SignInStatus="False"
            $SignInStat="Denied"
        }
        else{
            $SignInStatus="True"
            $SignInStat="Allowed"
        }

        #Filter result based on SignIn status
        if(($SignInAllowed -ne $null) -and ([string]$SignInAllowed -ne [string]$SignInStatus)){
            return
        }

        #Filter result based on License status
        if(($LicensedUserOnly.IsPresent) -and ($_.IsLicensed -eq $False))
        {
            return
        }

        if($_.IsLicensed -eq $true)
        {
            $LicenseStat="Licensed"
        }
        else
        {
            $LicenseStat="Unlicensed"
        }

        #Check for user's Admin role
        $Roles=(Get-MsolUserRole -UserPrincipalName $upn).Name
        if($Roles.count -eq 0)
        {
            $RolesAssigned="No roles"
            $IsAdmin="False"
        }
        else
        {
            $IsAdmin="True"
            foreach($Role in $Roles){
                $RolesAssigned=$RolesAssigned+$Role
                if($Roles.indexof($role) -lt (($Roles.count)-1))
                {
                    $RolesAssigned=$RolesAssigned+","
                }
            }
        }

        #Filter result based on Admin users
        if(($AdminOnly.IsPresent) -and ([string]$IsAdmin -eq "False"))
        {
            return
        }

        #Check for MFA enabled user
        if(($MethodTypes -ne $Null) -and ($MFAStatus -ne $Null) -and (-Not ($DisabledOnly.IsPresent) ))
        {
            #Check for Conditional Access
            if($MFAStatus -eq $null)
            {
                $MFAStatus='Enabled via Conditional Access'
            }

            #Filter result based on EnforcedOnly filter
            if((([string]$MFAStatus -eq "Enabled") -or ([string]$MFAStatus -eq "Enabled via Conditional Access")) -and ($EnforcedOnly.IsPresent))
            {
                return
            }

            #Filter result based on EnabledOnly filter
            if(([string]$MFAStatus -eq "Enforced") -and ($EnabledOnly.IsPresent))
            {
                return
            }

            #Filter result based on MFA enabled via Other source
            if((($MFAStatus -eq "Enabled") -or ($MFAStatus -eq "Enforced")) -and ($ConditionalAccessOnly.IsPresent))
            {
                return
            }

            $Methods=""
            $MethodTypes=""
            $MethodTypes=$_.StrongAuthenticationMethods.MethodType
            $DefaultMFAMethod=($_.StrongAuthenticationMethods | where{$_.IsDefault -eq "True"}).MethodType
            $MFAPhone=$_.StrongAuthenticationUserDetails.PhoneNumber
            $MFAEmail=$_.StrongAuthenticationUserDetails.Email

            if($MFAPhone -eq $Null)
            { $MFAPhone="-"}
            if($MFAEmail -eq $Null)
            { $MFAEmail="-"}

            if($MethodTypes -ne $Null)
            {
                $ActivationStatus="Yes"
                foreach($MethodType in $MethodTypes)
                {
                    if($Methods -ne "")
                    {
                        $Methods=$Methods+","
                    }
                    $Methods=$Methods+$MethodType
                }
            }
            else
            {
                $ActivationStatus="No"
                $Methods="-"
                $DefaultMFAMethod="-"
                $MFAPhone="-"
                $MFAEmail="-"
            }

            #Print to output file
            $PrintedEnabledUser++
            $Result=@{'DisplayName'=$DisplayName;'UserPrincipalName'=$upn;'MFAStatus'=$MFAStatus;'ActivationStatus'=$ActivationStatus;'DefaultMFAMethod'=$DefaultMFAMethod;'AllMFAMethods'=$Methods;'MFAPhone'=$MFAPhone;'MFAEmail'=$MFAEmail;'LicenseStatus'=$LicenseStat;'IsAdmin'=$IsAdmin;'AdminRoles'=$RolesAssigned;'SignInStatus'=$SigninStat}
            $Results= New-Object PSObject -Property $Result
            $Results | Select-Object DisplayName,UserPrincipalName,MFAStatus,ActivationStatus,DefaultMFAMethod,AllMFAMethods,MFAPhone,MFAEmail,LicenseStatus,IsAdmin,AdminRoles,SignInStatus | Export-Csv -Path $ExportCSVReport -Notype -Append
        }

        #Check for MFA disabled user
        #elseif(($DisabledOnly.IsPresent) -and ($MFAStatus -eq $Null) -and ($_.StrongAuthenticationMethods.MethodType -eq $Null))
        #{
        $MFAStatus="Disabled"
        $Department=$_.Department
        if($Department -eq $Null)
        { 
            $Department="-"}
            $PrintedDisabledUser++
            $Result=@{'DisplayName'=$DisplayName;'UserPrincipalName'=$upn;'Department'=$Department;'MFAStatus'=$MFAStatus;'LicenseStatus'=$LicenseStat;'IsAdmin'=$IsAdmin;'AdminRoles'=$RolesAssigned; 'SignInStatus'=$SigninStat}
            $Results= New-Object PSObject -Property $Result
            $Results | Select-Object DisplayName,UserPrincipalName,Department,MFAStatus,LicenseStatus,IsAdmin,AdminRoles,SignInStatus | Export-Csv -Path $ExportCSV -Notype -Append
        }
        #}

    #Open output file after execution
    Write-Host "`nGot MFA status successfully at $datetime"
    "Got MFA status successfully $datetime" | Out-File -FilePath $logfile -Append

    if((Test-Path -Path $ExportCSV) -eq "True")
    {
        Write-Host "MFA Disabled user report available in: $ExportCSV"
        "MFA Disabled user report available in: $ExportCSV" | Out-File -FilePath $logfile -Append

        # $Prompt = New-Object -ComObject wscript.shell
        # $UserInput = $Prompt.popup("Do you want to open output file?",`
        # 0,"Open Output File",4)
        # If ($UserInput -eq 6)
        # {
        #     Invoke-Item "$ExportCSV"
        # }
        Get-FileHash $ExportCSV | Format-Table -Wrap | Out-File -FilePath $hashfile -Append
        Write-Host "Exported report has $PrintedEnabledUser users"
        "Exported report has $PrintedEnabledUser users" | Out-File -FilePath $logfile -Append
    }
    if((Test-Path -Path $ExportCSVReport) -eq "True")
    {
        Write-Host "MFA Enabled user report available in: $ExportCSVReport"
        "MFA Enabled user report available in: $ExportCSVReport"  | Out-File -FilePath $logfile -Append
        # $Prompt = New-Object -ComObject wscript.shell
        # $UserInput = $Prompt.popup("Do you want to open output file?",`
        # 0,"Open Output File",4)
        # If ($UserInput -eq 6)
        # {
        #     Invoke-Item "$ExportCSVReport"
        # }
        Get-FileHash $ExportCSVReport | Format-Table -Wrap | Out-File -FilePath $hashfile -Append
        Write-Host "Exported report has $PrintedDisabledUser users"
        "Exported report has $PrintedDisabledUser users" | Out-File -FilePath $logfile -Append
    }
    Else
    {
        Write-Host "No user found that matches your criteria."
        "No user found that matches your criteria." | Out-File -FilePath $logfile -Append
    }
}

function getO365UserLicenseReport(){
    #Using this script administrator can identify all licensed users with their assigned licenses, services, and its status.

    Param
    (
        [Parameter(Mandatory = $false)]
        [string]$UserNamesFile
    )


    Function Get_UsersLicenseInfo
    {
        $LicensePlanWithEnabledService=""
        $FriendlyNameOfLicensePlanWithService=""
        $upn=$_.userprincipalname
        $Country=$_.Country
        if([string]$Country -eq "")
        {
            $Country="-"
        }
        Write-Progress -Activity "`n     Exported user count:$LicensedUserCount "`n"Currently Processing:$upn"
        #Get all asssigned SKU for current user
        $Skus=$_.licenses.accountSKUId
        $LicenseCount=$skus.count
        $count=0
        #Loop through each SKUid
        foreach($Sku in $Skus)  #License loop
        {
            #Convert Skuid to friendly name
            $LicenseItem= $Sku -Split ":" | Select-Object -Last 1
            $EasyName=$FriendlyNameHash[$LicenseItem]
            if(!($EasyName))
                {$NamePrint=$LicenseItem}
            else
                {$NamePrint=$EasyName}
            #Get all services for current SKUId
            $Services=$_.licenses[$count].ServiceStatus
            if(($Count -gt 0) -and ($count -lt $LicenseCount))
            {
                $LicensePlanWithEnabledService=$LicensePlanWithEnabledService+","
                $FriendlyNameOfLicensePlanWithService=$FriendlyNameOfLicensePlanWithService+","
            }
            $DisabledServiceCount = 0
            $EnabledServiceCount=0
            $serviceExceptDisabled=""
            $FriendlyNameOfServiceExceptDisabled=""
            foreach($Service in $Services) #Service loop
            {
                $flag=0
                $ServiceName=$Service.ServicePlan.ServiceName
                if($service.ProvisioningStatus -eq "Disabled")
                {
                    $DisabledServiceCount++
                }
                else
                {
                    $EnabledServiceCount++
                    if($EnabledServiceCount -ne 1)
                    {
                        $serviceExceptDisabled =$serviceExceptDisabled+","
                    }
                    $serviceExceptDisabled =$serviceExceptDisabled+$ServiceName
                    $flag=1
                }
                #Convert ServiceName to friendly name
                for($i=0;$i -lt $ServiceArray.length;$i +=2)
                {
                    $ServiceFriendlyName = $ServiceName
                    $Condition = $ServiceName -Match $ServiceArray[$i]
                    if($Condition -eq "True")
                    {
                        $ServiceFriendlyName=$ServiceArray[$i+1]
                        break
                    }
                }
                if($flag -eq 1)
                {
                    if($EnabledServiceCount -ne 1)
                    {
                        $FriendlyNameOfServiceExceptDisabled =$FriendlyNameOfServiceExceptDisabled+","
                    }
                    $FriendlyNameOfServiceExceptDisabled =$FriendlyNameOfServiceExceptDisabled+$ServiceFriendlyName
                }
                #Store Service and its status in Hash table
                $Result = @{'DisplayName'=$_.Displayname;'UserPrinciPalName'=$upn;'LicensePlan'=$Licenseitem;'FriendlyNameofLicensePlan'=$nameprint;'ServiceName'=$service.ServicePlan.ServiceName;
                'FriendlyNameofServiceName'=$serviceFriendlyName;'ProvisioningStatus'=$service.ProvisioningStatus}
                $Results = New-Object PSObject -Property $Result
                $Results |select-object DisplayName,UserPrinciPalName,LicensePlan,FriendlyNameofLicensePlan,ServiceName,FriendlyNameofServiceName,
                ProvisioningStatus | Export-Csv -Path $ExportCSV -Notype -Append
            }
            if($Disabledservicecount -eq 0)
            {
                $serviceExceptDisabled ="All services"
                $FriendlyNameOfServiceExceptDisabled="All services"
            }
            $LicensePlanWithEnabledService=$LicensePlanWithEnabledService + $Licenseitem +"[" +$serviceExceptDisabled +"]"
            $FriendlyNameOfLicensePlanWithService=$FriendlyNameOfLicensePlanWithService+ $NamePrint + "[" + $FriendlyNameOfServiceExceptDisabled +"]"
            #Increment SKUid count
            $count++
        }
        $Output=@{'Displayname'=$_.Displayname;'UserPrincipalName'=$upn;Country=$Country;'LicensePlanWithEnabledService'=$LicensePlanWithEnabledService;
        'FriendlyNameOfLicensePlanAndEnabledService'=$FriendlyNameOfLicensePlanWithService}
        $Outputs= New-Object PSObject -Property $output
        $Outputs | Select-Object Displayname,userprincipalname,Country,LicensePlanWithEnabledService,FriendlyNameOfLicensePlanAndEnabledService | Export-Csv -path $ExportSimpleCSV -NoTypeInformation -Append
    }


    Function getAndExport()
    {
        #Clean up session
        #Get-PSSession | Remove-PSSession
        #Connect AzureAD from PowerShell
        #Connect-MsolService
        #Set output file
        $ExportCSV=".\DetailedO365UserLicenseReport_$datetime.csv"
        $ExportSimpleCSV=".\SimpleO365UserLicenseReport_$datetime.csv"
        #FriendlyName list for license plan and service
        $FriendlyNameHash=Get-Content -Raw -Path .\LicenseFriendlyName.txt -ErrorAction Stop | ConvertFrom-StringData
        $ServiceArray=Get-Content -Path .\ServiceFriendlyName.txt -ErrorAction Stop
        #Hash table declaration
        $Result=""
        $Results=@()
        $output=""
        $outputs=@()
        #Get licensed user
        $LicensedUserCount=0

        #Check for input file/Get users from input file
        if([string]$UserNamesFile -ne "")
        {
            #We have an input file, read it into memory
            $UserNames=@()
            $UserNames=Import-Csv -Header "DisplayName" $UserNamesFile
            $userNames
            foreach($item in $UserNames)
            {
                Get-MsolUser -UserPrincipalName $item.displayname | where{$_.islicensed -eq "true"} | Foreach{
                    Get_UsersLicenseInfo
                    $LicensedUserCount++}
            }
        }
        #Get all licensed users
        else
        {
            Get-MsolUser -All | where{$_.islicensed -eq "true"} | Foreach{
                Get_UsersLicenseInfo
                $LicensedUserCount++}
        }

        #Open output file after execution
        Write-Host Detailed report available in: $ExportCSV
        "Detailed report available in: $ExportCSV" | Out-File -FilePath $logfile -Append
        Write-host Simple report available in: $ExportSimpleCSV
        "Simple report available in: $ExportSimpleCSV" | Out-File -FilePath $logfile -Append
        $Prompt = New-Object -ComObject wscript.shell
        $UserInput = $Prompt.popup("Do you want to open output files?",`
        0,"Open Files",4)
        If ($UserInput -eq 6)
        {
            Invoke-Item "$ExportCSV"
            Invoke-Item "$ExportSimpleCSV"
        }
        Get-FileHash $ExportCSV | Format-Table -Wrap | Out-File -FilePath $hashfile -Append
        Get-FileHash $ExportSimpleCSV | Format-Table -Wrap | Out-File -FilePath $hashfile -Append
    }
    . getAndExport
}

function loginExchange(){
    Get-PSSession | Remove-PSSession
    #Storing credential in script for scheduling purpose/ Passing credential as parameter
    if(($UserName -ne "") -and ($Password -ne ""))
    {
        $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
        $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
        Connect-ExchangeOnline -Credential $credential
    }
    else
    {
        Connect-ExchangeOnline | Out-Null
    }
    write-host "Logged in Exchange at $datetime"
    "Logged in Exchange at $datetime" | Out-File -FilePath $logfile -Append
}

function getUnifiedGroup(){
    $ExportCSV = ".\UnifiedGroup_$datetime.csv"
    Get-UnifiedGroup | Select-Object DisplayName, GroupType, PrimarySmtpAddress | Export-csv -Path $ExportCSV

    if((Test-Path -Path $ExportCSV) -eq "True"){
        Write-Host "`nGot Unified Group successfully at $datetime"
        "Got Unified Group successfully $datetime" | Out-File -FilePath $logfile -Append
        Get-FileHash $ExportCSV | Format-Table -Wrap | Out-File -FilePath $hashfile -Append
    }
}

function getConfigAnalyzer(){
    $ExportCSVstandard = ".\configAnalyzerStandard_$datetime.csv"
    $ExportCSVstrict = ".\configAnalyzerStrict_$datetime.csv"

    $results = Get-ConfigAnalyzerPolicyRecommendation -RecommendedPolicyType Standard
    foreach($result in $results){
        $Output=@{'Policy'=$result.Policy;'PolicyGroup'=$result.PolicyGroup;'SettingNameDescription'=$result.SettingNameDescription;'Currentconfiguration'=$result.Currentconfiguration}
        $Outputs = New-Object PSObject -Property $Output
        $Outputs | Select-Object Policy,PolicyGroup,SettingNameDescription,Currentconfiguration | Export-csv -Path $ExportCSVstandard -NoTypeInformation -Append
    }

    $results = Get-ConfigAnalyzerPolicyRecommendation -RecommendedPolicyType Strict
    foreach($result in $results){
        $Output=@{'Policy'=$result.Policy;'PolicyGroup'=$result.PolicyGroup;'SettingNameDescription'=$result.SettingNameDescription;'Currentconfiguration'=$result.Currentconfiguration}
        $Outputs = New-Object PSObject -Property $Output
        $Outputs | Select-Object Policy,PolicyGroup,SettingNameDescription,Currentconfiguration | Export-csv -Path $ExportCSVstrict -NoTypeInformation -Append
    }


    #Get-ConfigAnalyzerPolicyRecommendation -RecommendedPolicyType Standard | Export-csv -Path $ExportCSVstandard -Append
    #Get-ConfigAnalyzerPolicyRecommendation -RecommendedPolicyType Strict | Format-Table | Select-Object Policy, PolicyGroup, SettingName, SettingNameDescription, Currentconfiguration | Export-csv -Path $ExportCSVstrict

    if((Test-Path -Path $ExportCSVstandard) -eq "True"){
        Write-Host "`nGot Config Analyzer standard policy successfully at $datetime"
        "Got Config Analyzer standard policy successfully $datetime" | Out-File -FilePath $logfile -Append
        Get-FileHash $ExportCSVstandard | Format-Table -Wrap | Out-File -FilePath $hashfile -Append
    }
    if((Test-Path -Path $ExportCSVstrict) -eq "True"){
        Write-Host "`nGot Config Analyzer strict policy successfully at $datetime"
        "Got Config Analyzer strict policy successfully $datetime" | Out-File -FilePath $logfile -Append
        Get-FileHash $ExportCSVstrict | Format-Table -Wrap | Out-File -FilePath $hashfile -Append
    }
}

function getAtpPolicy(){
    $ExportCSV = ".\ATPPolicy_$datetime.csv"
    $results = Get-AtpPolicyForO365
    foreach($result in $results){
        $Output=@{'Name'=$result.Name;'IsValid'=$result.IsValid;'AdminDisplayName'=$result.AdminDisplayName;
            'TrackClicks'=$result.TrackClicks;
            'AllowClickThrough'=$result.AllowClickThrough;
            'EnableSafeLinksForClients'=$result.EnableSafeLinksForClients;
            'EnableSafeLinksForWebAccessCompanion'=$result.EnableSafeLinksForWebAccessCompanion;
            'EnableSafeLinksForO365Clients'=$result.EnableSafeLinksForO365Clients;
            'BlockUrls'=$result.BlockUrls;
            'EnableATPForSPOTeamsODB'=$result.EnableATPForSPOTeamsODB;
            'EnableSafeDocs'=$result.EnableSafeDocs;
            'AllowSafeDocsOpen'=$result.AllowSafeDocsOpen;
            'WhenChangedUTC'=$result.WhenChangedUTC;
            'WhenCreatedUTC'=$result.WhenCreatedUTC;
            }
        $Outputs = New-Object PSObject -Property $Output
        $Outputs | Select-Object Name,AdminDisplayName,IsValid,WhenCreatedUTC,WhenChangedUTC,TrackClicks,
        AllowClickThrough,EnableSafeLinksForClients,EnableSafeLinksForWebAccessCompanion,EnableSafeLinksForO365Clients,
        BlockUrls, EnableATPForSPOTeamsODB,EnableSafeDocs,AllowSafeDocsOpen | Export-csv -Path $ExportCSV -NoTypeInformation -Append
    }

    if((Test-Path -Path $ExportCSV) -eq "True"){
        Write-Host "`nGot ATPPolicy successfully at $datetime"
        "Got ATPPolicy successfully $datetime" | Out-File -FilePath $logfile -Append
        Get-FileHash $ExportCSV | Format-Table -Wrap | Out-File -FilePath $hashfile -Append
    }

}

function getAntiPhishPolicy(){
    $ExportCSV = ".\AntiPhishPolicy_$datetime.csv"
    $results = Get-AntiPhishPolicy
    foreach($result in $results){
        $Output=@{'Name'=$result.Name;'AdminDisplayName'=$result.AdminDisplayName;
            'Enabled'=$result.Enabled;'IsDefault'=$result.IsDefault;
            'WhenChangedUTC'=$result.WhenChangedUTC;
            'WhenCreatedUTC'=$result.WhenCreatedUTC;
            'PhishThresholdLevel'=$result.PhishThresholdLevel;
            'ImpersonationProtectionState'=$result.ImpersonationProtectionState;
            'EnableMailboxIntelligenceProtection'=$result.EnableMailboxIntelligenceProtection;
            'EnableSpoofIntelligence'=$result.EnableSpoofIntelligence;
            'TargetedUsersToProtect'=$result.TargetedUsersToProtect;
            'TargetedDomainsToProtect'=$result.TargetedDomainsToProtect;
            'ExcludedDomains'=$result.EnableATPForSPOTeamsODB;
            'ExcludedSenders'=$result.EnableSafeDocs;
            'RecommendedPolicyType'=$result.AllowSafeDocsOpen;
            }
        $Outputs = New-Object PSObject -Property $Output
        $Outputs | Select-Object Name,AdminDisplayName,Enabled,WhenCreatedUTC,WhenChangedUTC,PhishThresholdLevel,
        ImpersonationProtectionState,EnableMailboxIntelligenceProtection,EnableSpoofIntelligence,
        TargetedUsersToProtect,TargetedDomainsToProtect, ExcludedDomains,ExcludedSenders,
        RecommendedPolicyType | Export-csv -Path $ExportCSV -NoTypeInformation -Append
    }

    if((Test-Path -Path $ExportCSV) -eq "True"){
        Write-Host "`nGot AntiPhishPolicy successfully at $datetime"
        "Got AntiPhishPolicy successfully $datetime" | Out-File -FilePath $logfile -Append
        Get-FileHash $ExportCSV | Format-Table -Wrap | Out-File -FilePath $hashfile -Append
    }
}

function getSafeAttachmentPolicy(){
    $ExportCSV = ".\SafeAttachmentPolicy_$datetime.csv"
    $results = Get-SafeAttachmentPolicy
    foreach($result in $results){
        $Output=@{'Name'=$result.Name;
            'Action'=$result.Action;
            'Enable'=$result.Enable;
            'IsDefault'=$result.IsDefault;
            }
        $Outputs = New-Object PSObject -Property $Output
        $Outputs | Select-Object Name,Action,Enable,
        IsDefault | Export-csv -Path $ExportCSV -NoTypeInformation -Append
    }

    if((Test-Path -Path $ExportCSV) -eq "True"){
        Write-Host "`nGot SafeAttachmentPolicy successfully at $datetime"
        "Got SafeAttachmentPolicy successfully $datetime" | Out-File -FilePath $logfile -Append
        Get-FileHash $ExportCSV | Format-Table -Wrap | Out-File -FilePath $hashfile -Append
    }
}

function getSafeLinksPolicy(){
    $ExportCSV = ".\SafeLinksPolicy_$datetime.csv"
    $results = Get-SafeLinksPolicy
    foreach($result in $results){
        $Output=@{'Name'=$result.Name;
            'AdminDisplayName'=$result.AdminDisplayName
            'IsEnabled'=$result.IsEnabled;
            'IsDefault'=$result.IsDefault;
            'WhenChangedUTC'=$result.WhenChangedUTC;
            'WhenCreatedUTC'=$result.WhenCreatedUTC;
            'EnableSafeLinksForEmail'=$result.EnableSafeLinksForEmail;
            'EnableSafeLinksForTeams'=$result.EnableSafeLinksForTeams;
            'EnableSafeLinksForOffice'=$result.EnableSafeLinksForOffice;
            'TrackClicks'=$result.TrackClicks;
            'AllowClickThrough'=$result.AllowClickThrough;
            'ScanUrls'=$result.ScanUrls;
            'ExcludedUrls'=$result.ExcludedUrls;
            }
        $Outputs = New-Object PSObject -Property $Output
        $Outputs | Select-Object Name,AdminDisplayName,IsEnabled,IsDefault,WhenChangedUTC,WhenCreatedUTC,
        EnableSafeLinksForEmail,EnableSafeLinksForTeams,EnableSafeLinksForOffice,TrackClicks,AllowClickThrough,
        ScanUrls,ExcludedUrls | Export-csv -Path $ExportCSV -NoTypeInformation -Append
    }

    if((Test-Path -Path $ExportCSV) -eq "True"){
        Write-Host "`nGot SafeLinksPolicy successfully at $datetime"
        "Got SafeLinksPolicy successfully $datetime" | Out-File -FilePath $logfile -Append
        Get-FileHash $ExportCSV | Format-Table -Wrap | Out-File -FilePath $hashfile -Append
    }
}

function getTransportRule(){
    $ExportCSV = ".\TransportRule_$datetime.csv"
    $results = Get-TransportRule
    foreach($result in $results){
        $Output=@{'Name'=$result.Name;
            'State'=$result.State
            'Mode'=$result.Mode;
            'Priority'=$result.Priority;
            'Comments'=$result.Comments;
            }
        $Outputs = New-Object PSObject -Property $Output
        $Outputs | Select-Object Name,State,Mode,Priority,Comments| Export-csv -Path $ExportCSV -NoTypeInformation -Append
    }

    if((Test-Path -Path $ExportCSV) -eq "True"){
        Write-Host "`nGot TransportRule successfully at $datetime"
        "Got TransportRule successfully $datetime" | Out-File -FilePath $logfile -Append
        Get-FileHash $ExportCSV | Format-Table -Wrap | Out-File -FilePath $hashfile -Append
    }
}

function getRetentionPolicy(){
    $ExportCSV = ".\RetentionPolicy_$datetime.csv"
    $results = Get-RetentionPolicy
    foreach($result in $results){
        $Output=@{'Name'=$result.Name;
            'RetentionPolicyTagLinks'=$result.RetentionPolicyTagLinks
            }
        $Outputs = New-Object PSObject -Property $Output
        $Outputs | Select-Object Name,RetentionPolicyTagLinks | Export-csv -Path $ExportCSV -NoTypeInformation -Append
    }

    if((Test-Path -Path $ExportCSV) -eq "True"){
        Write-Host "`nGot RetentionPolicy successfully at $datetime"
        "Got RetentionPolicy successfully $datetime" | Out-File -FilePath $logfile -Append
        Get-FileHash $ExportCSV | Format-Table -Wrap | Out-File -FilePath $hashfile -Append
    }
}

function getUserInboxRule(){
    $ExportCSV = ".\getUserInboxRule_$datetime.csv"
    Get-ExoMailbox -ResultSize Unlimited | 
    Select-Object -ExpandProperty UserPrincipalName | 
    Foreach-Object {Get-InboxRule -Mailbox $_ | 
    Select-Object -Property MailboxOwnerID,Name,Enabled,From,Description,RedirectTo,ForwardTo} |
    Export-csv -Path $ExportCSV -NoTypeInformation -Append

    if((Test-Path -Path $ExportCSV) -eq "True"){
        Write-Host "`nGot UserInboxRule successfully at $datetime"
        "Got UserInboxRule successfully $datetime" | Out-File -FilePath $logfile -Append
        Get-FileHash $ExportCSV | Format-Table -Wrap | Out-File -FilePath $hashfile -Append
    }

}

function main(){
    . checkNeededModule
    . loginMSOnline
    . ExportO365Users
    . getMFAStatus
    . getO365UserLicenseReport
    . loginExchange
    . getUnifiedGroup
    . getConfigAnalyzer
    . getAtpPolicy
    . getAntiPhishPolicy
    . getSafeAttachmentPolicy
    . getSafeLinksPolicy
    . getTransportRule
    . getRetentionPolicy
    . getUserInboxRule

}
. main