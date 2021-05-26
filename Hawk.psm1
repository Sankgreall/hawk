# Make sure we get back all of the unified audit log results for the search we are doing
Function Get-AllUnifiedAuditLogEntry {
    param 
    (
        [Parameter(Mandatory = $true)]
        [string]$UnifiedSearch,
        [datetime]$StartDate = $Hawk.StartDate,
        [datetime]$EndDate = $Hawk.EndDate
    )
	
    # Validate the incoming search command
    if (($UnifiedSearch -match "-StartDate") -or ($UnifiedSearch -match "-EndDate") -or ($UnifiedSearch -match "-SessionCommand") -or ($UnifiedSearch -match "-ResultSize") -or ($UnifiedSearch -match "-SessionId")) {
        Out-LogFile "Do not include any of the following in the Search Command"
        Out-LogFile "-StartDate, -EndDate, -SessionCommand, -ResultSize, -SessionID"
        Write-Error -Message "Unable to process search command, switch in UnifiedSearch that is handled by this cmdlet specified" -ErrorAction Stop
    }
		
    # Make sure key variables are null
    [string]$cmd = $null
    
    $query_value = 50000
    $query_chunk = 1000
    # build our initial search command to execute
    $cmd = $UnifiedSearch + " -StartDate `'" + $StartDate + "`' -EndDate `'" + $EndDate + "`' -ResultSize 1"
    Out-Logfile ("Searching for data between " + $StartDate + " and " + $EndDate) -Notice "Action" 

    # Run the initial command
    $Output = $null
    $Output += (Invoke-Expression $cmd)
    # Sort our result set to make sure the higest number is in the last position
    $Output = $Output | Sort-Object -Property ResultIndex

    # If no events, we can just return
    if ($null -eq $Output) 
    {
        $Run = $false
        # Convert our list to an array and return it
        [array]$Output = $Output
        return $Output
    }

    # Else, we have data
    Out-LogFile ("There are " + $Output[-1].ResultCount + " records within your collection period")
    # Sort our result set to make sure the higest number is in the last position
    $Output = $Output | Sort-Object -Property ResultIndex
    $recordCount = $Output[-1].ResultCount
    $timeDifference = (New-TimeSpan -Start $StartDate -End $EndDate).Days
    if($timeDifference -lt 1)
    {
        $timeDifference = 1
    }


    # Calculate interval for log collection
    $interval = ([int]$recordCount / [int]$query_value)
    #As a fraction of the auditable period:
    $ratio = $timeDifference / $interval
    # Hours in a day
    $hours = 24 * $ratio

    if($hours -gt ($timeDifference * 24))
    {
        $hours = ($timeDifference * 24)
        Out-LogFile("Will attempt to collect records within a single session")
    }
    else
    {
        Out-LogFile("Will attempt to collect records in chunks of " + [int]$hours + " hours")   
    }
    $masterHours = $hours
       

    # Now we begin the loop!
    # Setup our run variable
    $Run = $true

    $Output = $null
    $errorCounter = 0

    while($Run)
    {
        Test-EXOConnection
        $sessionID = (Get-Date -UFormat %H%M%S)

        # If timespan goes beyond end date, replace value with end date
        if($StartDate.AddHours($hours) -gt $EndDate)
        {
            $QueryEnd = $EndDate
        }

        else
        {
            $QueryEnd = $StartDate.AddHours($hours)
        }

        # If the difference between start and end is tiny, something's wrong. Skip
        $secondsDifference = (New-TimeSpan -Start $StartDate -End $QueryEnd).Seconds
        if ($secondsDifference -lt 10) # if there's more than 50,000 logs in 10 seconds...
        {
            $StartDate = $StartDate.AddSeconds($secondsDifference)
            $QueryEnd = $QueryEnd.AddSeconds($secondsDifference)
        } 


        $cmd = $UnifiedSearch + " -StartDate `'" + $StartDate + "`' -EndDate `'" + $QueryEnd + "`'  -ResultSize " + $query_chunk + " -SessionCommand ReturnLargeSet -sessionid " + $sessionId       
        if ($StartDate.AddHours($hours) -gt $EndDate -Or $StartDate.AddHours($hours) -eq $EndDate)
        {
            # Time to end :/
            $Run = $false
        }

        Out-LogFile ("Searching between " + $StartDate + " and " + $QueryEnd) -Notice "action"
        
        try
        {
            $Output += (Invoke-Expression $cmd)
            # Sort our result set to make sure the higest number is in the last position
            $Output = $Output | Sort-Object -Property ResultIndex
        }
        catch
        {
            $errorCounter = $errorCounter + 1
            if($errorCounter -gt 3)
            {
                $Output = $null
                $StartDate = $StartDate.AddHours($hours)
                
            }
            continue
        }
        
        if ($Output.Length -eq 0 -or $Output[-1].ResultCount -eq 0)
        {
            # No logs, continue
            Out-LogFile ("Returned Result count was 0, continuing")
            $Output = $null
            $StartDate = $StartDate.AddHours($hours)
            continue
        }

        # Check for returned size
        $returnedSize = $Output[-1].ResultCount
        if ([int]$returnedSize -gt $query_chunk -And [int]$returnedSize -lt $query_value)
        {
            $count = 0
            $counter = 0
            $session = $true
            $seconderrorCounter = 0
            $weirdErrorCounter = 0
            # Output progress before loop
            Out-LogFile ("Retrieved:" + $Output[-1].ResultIndex.tostring().PadRight(5, " ") + " Total: " + $Output[-1].ResultCount)
            while($session)
            {
                Test-EXOConnection
                try
                {
                    $Output += (Invoke-Expression $cmd)
                    # Sort our result set to make sure the higest number is in the last position
                    $Output = $Output | Sort-Object -Property ResultIndex
                }
                catch
                {
                    $seconderrorCounter = $seconderrorCounter + 1
                    if($seconderrorCounter -gt 3)
                    {
                        Out-LogFile ("Something went wrong, we're moving on")
                        $session = $false
                    }
                    continue
                }

                # There is a weird error whereby the log count doesn't increase
                # Eventually this leads to a loop and the script can't advance
                # This checks for this bug, adds a counter, and  continues if counter
                # is > 3.
                if($count -eq $Output[-1].ResultIndex)
                {
                    # Haven't increased since previous, increment counter 
                    $counter = $counter + 1
                    if($counter -gt 3)
                    {
                        # We should actually try again with reduced hours
                        if ($weirdErrorCounter -lt 3)
                        {
                            $reduceHourFlag = $true # If not set, we just move on
                            Out-LogFile ("Something went wrong, we're going to try again with reduced timeframe")
                        }
                        else
                        {
                            # MS being MS
                            Out-LogFile ("[ERROR] Something went wrong and we can't fix it - we have to move on")
                            $weirdErrorCounter = 0
                        }                        
                        $weirdErrorCounter = $weirdErrorCounter + 1
                        $session = $false
                    }
                }

                # On the other hand, of the index has DECREASED, it means MS reset our session (why?)
                # It seems that the most effective treatment is to retry with lower volumes of logs
                if ($Output[-1].ResultIndex -lt $count) # New index is LESS than previous index
                {
                    $reduceHourFlag = $true # If not set, we just move on
                    Out-LogFile ("Something went wrong, we're going to try again with reduced timeframe")
                    $session = $false
                }

                # Update the new count for the next iteration
                $count = $Output[-1].ResultIndex

                if($Output[-1].ResultIndex -eq $Output[-1].ResultCount)
                {
                    Out-LogFile ("Retrieved:" + $Output[-1].ResultIndex.tostring().PadRight(5, " ") + " Total: " + $Output[-1].ResultCount)
                    Out-LogFile ("Retrieved all logs from session, writing to file...")
                    $logFlag = $true
                    $session = $false
                }
                else
                {
                    if ($logFlag -eq $true)
                    {
                        $logFlag = $false
                    }
                    else
                    {
                        Out-LogFile ("Retrieved:" + $Output[-1].ResultIndex.tostring().PadRight(5, " ") + " Total: " + $Output[-1].ResultCount)
                    }
                    
                }
            }
        }

        elseif ([int]$returnedSize -gt $query_value)
        {
            Out-LogFile ("WARNING - day interval too large, returned: " + $returnedSize)
            $reduceHourFlag = $true
        }

        # Output is less than $query_chunk
        else
        {
            # Do nothing
            $logFlag = $false
        }

        if($reduceHourFlag -eq $true)
        {
            $reduceHourFlag = $false
            $hours = ($hours / 2)
            $adjustTime = $true
            $Output = $null
            continue
        }

        # Output the current progress
        if ($logFlag -eq $false)
        {
            Out-LogFile ("Retrieved:" + $Output[-1].ResultIndex.tostring().PadRight(5, " ") + " Total: " + $Output[-1].ResultCount)
        }
        else
        {
            $logFlag = $false    
        }
        
        # Write data to file
        $filename = Join-Path $Hawk.FilePath ("UnifiedAuditLog.txt")            
        $stream = New-Object IO.StreamWriter $filename, $true
        Foreach ($event in $Output)
        {
            $stream.WriteLine($event.auditdata)
        }
        $stream.Dispose()                    
        
        $Output = $null
        $StartDate = $StartDate.AddHours($hours)

        # Reset hours if it's been adjusted
        if($adjustTime = $true)
        {
            $hours = $masterHours
            $adjustTime = $false
        }

    }	

    Out-LogFile ("Retrieved all results.")
    # Convert our list to an array and return it
    [array]$Output = $Output
    return $Output
}

# Writes output to a log file with a time date stamp
Function Out-LogFile {
    Param 
    ( 
        [string]$string,
        [string]$notice="INFO"
    )
	
    # Get our log file path
    $LogFile = Join-path $Hawk.FilePath "Hawk.log"
    $ScreenOutput = $true
    $LogOutput = $true
	
    # Get the current date
    [string]$date = Get-Date -Format G
		
    # Deal with each switch and what log string it should put out and if any special output

    # Action indicates that we are starting to do something
    if ($null -ne $notice)
    {
        $notice = $notice.ToUpper()
        [string]$logstring = ( ("[" + $date + "] - [" + $notice + "] ") + $string)

    }
    # Normal output
    else
    {
        [string]$logstring = ( "[" + $date + "] - " + $string)
    }

    # Write everything to our log file
    if ($LogOutput)
    {
        $logstring | Out-File -FilePath $LogFile -Append
    }
	
    # Output to the screen
    if ($ScreenOutput)
    {
        Write-Information -MessageData $logstring -InformationAction Continue
    }

}

# Test if we are connected to the compliance center online and connect if now
Function Test-CCOConnection 
{
    Write-Output "Not yet implemented"
}

# Test if we are connected to Exchange Online and connect if not
Function Test-EXOConnection 
{

    # Check our token cache and if it will expire in less than 15 min renew the session
    $Expires = (Get-TokenCache | Where-Object { $_.resource -like "*outlook.office365.com*" }).ExpiresOn

    # if Expires is null we want to just move on
    if ($null -eq $Expires) { }
    else 
    {
        # If it is not null then we need to see if it is expiring soon
        if (($Expires - ((get-date).AddMinutes(15)) -le 0)) 
        {
            Write-Host "Token Near Expiry - rebuilding EXO connection"
            Connect-ExchangeOnline
        }
    }

    # In all cases make sure we are "connected" to EXO
    try
    { 
        $null = Get-OrganizationConfig -erroraction stop
                    
    }
    catch [System.Management.Automation.CommandNotFoundException] 
    {
        # Connect to EXO if we couldn't find the command
        Write-Host "Not Connected to Exchange Online"
        Write-Host "Connecting to EXO using CloudConnect Module"
        Connect-ExchangeOnline
    }
}

# Test if we are connected to MSOL and connect if we are not
Function Test-MSOLConnection 
{
	
    try { $null = Get-MsolCompanyInformation -ErrorAction Stop }
    catch [Microsoft.Online.Administration.Automation.MicrosoftOnlineException]
    {
		
        # Write to the screen if we don't have a log file path yet
        if ([string]::IsNullOrEmpty($Hawk.Logfile)) 
        {
            Write-Host "Connecting to MSOLService using MSOnline Module"
        }
        # Otherwise output to the log file
        else
        {
            Write-Host "Connecting to MSOLService using MSOnline Module"
        }

        # Connect to the MSOl Service (This should have been installed with the CloudConnect Module)
        Connect-MsolService

    }
}

# Test if we have a connection with the AzureAD Cmdlets
Function Test-AzureADConnection
{
    
    $TestModule = Get-Module AzureAD -ListAvailable -ErrorAction SilentlyContinue
    $MinimumVersion = New-Object -TypeName Version -ArgumentList "2.0.0.131"

    if ($null -eq $TestModule) 
    {
        Write-Host "Please Install the AzureAD Module with the following command:"
        Write-Host "Install-Module AzureAD"
        break
    }
    # Since we are not null pull the highest version
    else
    {
        $TestModuleVersion = ($TestModule | Sort-Object -Property Version -Descending)[0].version
    }
	
    # Test the version we need at least 2.0.0.131
    if ($TestModuleVersion -lt $MinimumVersion)
    {
        Write-Host ("AzureAD Module Installed Version: " + $TestModuleVersion)
        Write-Host ("Miniumum Required Version: " + $MinimumVersion)
        Write-Host"Please update the module with: Update-Module AzureAD"
        break
    }
    # Do nothing
    else { }

    try
    { 
        $Null = Get-AzureADTenantDetail -ErrorAction Stop
    }
    catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException]
    {
        Write-Host "Please connect to AzureAD prior to running this cmdlet"
        Write-Host "Connect-AzureAD"
        break
    }
}


