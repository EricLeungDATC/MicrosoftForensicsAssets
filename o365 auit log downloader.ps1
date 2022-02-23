### Microsoft 365 Audit Logs Downloader 

    # =============== CHANGE THESE =============== #     

    $days = 1 # days to check 

    $hoursPerBatch = 12  # Download 12 hours at a time, until the start date/time is reached 

    $currentDate = Get-Date 

    $OutputFolder = "UnifiedAuditLog_"+$tenant+"_"+$currentDate.ToString('yyyyMMdd')     

    # ============================================ # 

 

    # Microsoft 365 Login with MFA Enabled 

    $Session = Connect-ExchangeOnline -UserPrincipalName $upn -ShowProgress $false 

 

    # Set Current Date 12:00:00 am as $EndDate and set $StartDate - days to check 

    $ErrorActionPreference = "Stop" 

    $EndDate = $currentDate.Date.ToUniversalTime() 

    $InitialDate = $EndDate.Date.AddDays(-$days) 

 

    If(!(Test-Path $OutputFolder)) { 

          New-Item -ItemType Directory -Force -Path $OutputFolder 

    } 

 

    "Window is from: $InitialDate -> $EndDate" 

    do { 

        # Download 12 hours at a time, until the start date/time is reached 

        $StartDate = $EndDate.AddHours(-$hoursPerBatch) 

        "Running search for $StartDate -> $EndDate" 

        do { 

            $time = $EndDate 

            # Search the defined date(s), SessionId + SessionCommand in combination with the loop will return and append 5000 object per iteration until all objects are returned (maximum limit is 50k objects) 

            do { 

                try { 

                    $auditOutput = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -SessionId "$StartDate -> $EndDate" -SessionCommand ReturnLargeSet -ResultSize 5000 

                    $errorvar = $false 

                } 

                catch { 

                    $errorvar = $true 

                    "An error occured. Trying again." 

                } 

            } while ($errorvar) 

            if ($auditOutput) { 

                "Results received (" + ($(Get-Date) - $time) + "), writing to file..." 

                $filename = "$StartDate - $EndDate.csv".Replace(":",".").Replace("/","-") 

                #$auditOutput | Select-Object -Property AuditData | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | % {$_ -replace '""','"'} | % {$_ -replace '^"','' -replace '"$',''} | Add-Content -Path "$OutputFolder\$filename" -Encoding UTF8 

                $auditOutput | Select-Object -ExpandProperty AuditData | Select-Object -Skip 1 | Add-Content -Path "$OutputFolder\$filename" -Encoding UTF8 

                #Write-Output $auditOutput 

                Get-FileHash -Path "$OutputFolder\$filename" | Format-List >> "$OutputFolder\FileHash.txt" 

            } 

        } while ($AuditOutput) 

        $endDate = $EndDate.AddHours(-$hoursPerBatch) 

    } while ($InitialDate -lt $StartDate) 