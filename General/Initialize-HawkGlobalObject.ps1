# Create the hawk global object for use by other cmdlets in the hawk module
Function Initialize-HawkGlobalObject 
{
    param
    (
        [string]$StartDate=$null,
		[string]$EndDate="",
		[int]$Lookback=0,
		[string]$FilePath
    )

    Function Test-LoggingPath 
    {
        param([string]$PathToTest)

        # First test if the path we were given exists
        if (Test-Path $PathToTest)
        {

            # If the path exists verify that it is a folder
            if ((Get-Item $PathToTest).PSIsContainer -eq $true)
            {
                Return $true
            }
            # If it is not a folder return false and write an error
            else
            {
                Write-Host ("Path provided " + $PathToTest + " was not found to be a folder.") 
                Return $false
            }
        }
        # If it doesn't exist then return false and write an error
        else
        {
            Write-Host ("Directory " + $PathToTest + " Not Found") 
            Return $false
        }
    }

    Function New-LoggingFolder
    {
        param([string]$RootPath)

        # Create a folder ID based on date
        [string]$FolderID = "UnifiedAuditLogs-Collection"

        # Add that ID to the given path
        $FullOutputPath = Join-Path $RootPath $FolderID
        
        # Just in case we run this twice in a min lets not throw an error
        if (Test-Path $FullOutputPath)
        {
            Write-Host "`n!!! Appending to existing project !!!" 
        }
        # If it is not there make it
        else
        {
            Write-Host ("### Creating subfolder with name " + $FullOutputPath)  
            $null = New-Item $FullOutputPath -ItemType Directory
        }

        Return $FullOutputPath
    }

    Function Set-LoggingPath
    {
        param ([string]$Path)

        # If a value if provided go from there
        # If the provided path is valid then we can create the subfolder
        if (Test-LoggingPath -PathToTest $Path)
        {
            $Folder = New-LoggingFolder -RootPath $Path
        }
        # If the provided path fails validation then we just need to stop
        else
        {
            Write-Error ("Provided Path is not valid " + $Path) -ErrorAction Stop
        }

        Return $Folder
    }
 
    ### Main ###

    Write-Host "`n########### JustGiveMeAuditLogsPlease ###########"
  
    # Test if we have a connection to O365
    Test-MSOLConnection
    Test-EXOConnection	

    [string]$OutputPath = Set-LoggingPath -path $FilePath

    if ($Lookback -gt 0)
    {
        [DateTime]$StartDate = ((Get-Date).AddDays(-$Lookback)).Date
        Write-Host ("### Setting StartDate to " + $StartDate)
        [DateTime]$EndDate = ((Get-Date).AddDays(1)).Date
        Write-Host ("### Setting EndDate to " + $EndDate + "`n")
    }

    else
    {
        # StartDate must have been set
        [DateTime]$StartDate = (Get-Date $StartDate)

        if ($StartDate -ge ((Get-date).AddDays(-90)).Date -and ($StartDate -le (Get-Date).Date)) 
        {
            Write-Host ("### Setting StartDate to " + $StartDate)
        }

        else 
        {
            Write-Host ("Your StartDate is either greater than 90 days or in the future")
            Write-Host ($StartDate)
            return
        }

        # Set EndDate if it's specified
        if ($EndDate -ne "")
        {
            [DateTime]$EndDate = (Get-Date $EndDate)

            # Check that EndDate is greater than StartDate
            if ($StartDate -gt $EndDate)
            {
                Write-Host "your EndDate is set before the StartDate."
                Write-Host ($EndDate)
                return
            }
            elseif ($EndDate -gt (Get-Date))
            {
                Write-Host "your EndDate is set too far in the future."
                Write-Host ($EndDate)
                return              
            }
            Write-Host ("### Setting EndDate to " + $EndDate + "`n")
        }
        # Else, if EndDate is not specified, set to end of log collection
        else
        {
            [DateTime]$EndDate = (Get-Date).AddDays(1).Date # The one is just to make sure we compensate for local time zone issues
            Write-Host ("### Setting EndDate to " + $EndDate + "`n")
        }
    }
    
    # Determine if we have access to a P1 or P2 Azure Ad License
    # EMS SKU contains Azure P1 as part of the sku
    if ([bool](Get-MsolAccountSku | Where-Object { ($_.accountskuid -like "*aad_premium*") -or ($_.accountskuid -like "*EMS*") }))
    {
        [bool]$AdvancedAzureLicense = $true
    }
    else
    {
        [bool]$AdvancedAzureLicense = $false
    }

    # Null our object then create it
    $Output = $null
    $Output = New-Object -TypeName PSObject

    # Build the output object from what we have collected
    $Output | Add-Member -MemberType NoteProperty -Name FilePath -Value $OutputPath
    $Output | Add-Member -MemberType NoteProperty -Name Lookback -Value $Lookback
    $Output | Add-Member -MemberType NoteProperty -Name StartDate -Value $StartDate
    $Output | Add-Member -MemberType NoteProperty -Name EndDate -Value $EndDate
    $Output | Add-Member -MemberType NoteProperty -Name AdvancedAzureLicense -Value $AdvancedAzureLicense
    $Output | Add-Member -MemberType NoteProperty -Name WhenCreated -Value (Get-Date -Format g)

    # Create the global hawk variable
    New-Variable -Name Hawk -Scope Global -value $Output -Force

    #### End of IF
}

