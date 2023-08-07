#Requires -Modules MicrosoftTeams
Function Set-HolidaysInTeamsAdministration {
<#
    .SYNOPSIS
    This function generates public holidays based on French government APIs for Teams Admin Center.

    .DESCRIPTION
    Author :        DUBOIS Dylan
    Last Update :   07/08/2023

    .EXAMPLE
    PS> Set-HolidaysInTeamsAdministration

    .LINK
    GITHUB :        https://github.com/duboisdylan/Set-HolidaysInTeamsAdministration
    THANKS FOR :    https://github.com/12Knocksinna/Office365itpros/blob/master/PopulateTeamsHolidays.PS1
#>
    Begin {
        # Get this year.
        $ThisYear = Get-Date -Format "yyyy"

        # Use API FR Calendar of government
        $APIRequest = "https://calendrier.api.gouv.fr/jours-feries/metropole/$($ThisYear).json"
        $AllHolidaysInThisYear = Invoke-RestMethod -Method "Get" -Uri $APIRequest -ContentType "application/json"
        $AllHolidaysDaysInThisYear = $AllHolidaysInThisYear.psobject.properties.name

        # Check if Microsoft Teams is load.
        $Module = Get-Module | Select-Object -ExpandProperty "Name"
        If ("MicrosoftTeams" -notin $Module)
        {
            Connect-MicrosoftTeams
        }
       
        # Fetch current Teams holiday schedule
        Write-Host "Retrieving current Teams holiday schedule..."

        $TeamsSchedule = @{}
        [array]$CurrentSchedule = Get-CsOnlineSchedule | Where-Object {$_.Type -eq "Fixed"} | Select-Object Name, FixedSchedule
        # Build hash table of current events
        ForEach ($Event in $CurrentSchedule) {
            $EventDate = Get-Date($Event.FixedSchedule.DateTimeRanges.Start) -format d
            $TeamsSchedule.Add([string]$Event.Name,$EventDate)
        }
    }

    Process {
        Try {
            Foreach ($Day in $AllHolidaysDaysInThisYear)
            {
                # Get $Day in good format
                $GoodFormatDate = Get-Date -Date $Day -Format "yyyy-MM-dd"

                If ((Get-Date -Date $GoodFormatDate -Format "dd/MM/yyyy") -in $TeamsSchedule.Values)
                {
                    Write-Host ("{0} event already registered for {1}" -f $AllHolidaysInThisYear.$Day, $GoodFormatDate)
                }
                Else 
                {
                    Write-Host "Add $($AllHolidaysInThisYear.$Day) $(Get-Date -Format "yyyy")"
                    # Create DateTimeRange and Schedule Holiday in Teams
                    $HolidayDay = New-CsOnlineDateTimeRange -Start (Get-Date -Date $GoodFormatDate -Format "dd/MM/yyyy")
                    New-CsOnlineSchedule -Name "$($AllHolidaysInThisYear.$Day) $(Get-Date -Format "yyyy")" -FixedSchedule -DateTimeRanges @($HolidayDay)
                }
            }
        }
        Catch {
            write-warning "Une erreur est survenue" 
            write-host $_.Exception.Message
        }
    }

    End {
        # We check the date configured.
        Get-CsOnlineSchedule | Where-Object {$_.Type -eq "Fixed"} | Select-Object Name
    }
}