﻿<#
Annual Leave Fact Sheet: https://www.opm.gov/policy-data-oversight/pay-leave/leave-administration/fact-sheets/annual-leave/
Sick Leave Fact Sheet: https://www.opm.gov/policy-data-oversight/pay-leave/leave-administration/fact-sheets/sick-leave-general-information/
Leave Year Beginning/Ending Dates: https://www.opm.gov/policy-data-oversight/pay-leave/leave-administration/fact-sheets/leave-year-beginning-and-ending-dates/
Federal Holidays: https://www.opm.gov/policy-data-oversight/pay-leave/federal-holidays/
#>

<#
See about handling "Enter" key press events on all forms.

Make it say "Leave Balances as of Date" instead of just Leave Balances. Move this to the center and just make the one over the box say Leave Balances.

Add AutoSize to everything that should have it. Make things FlowLayoutPanels if it makes more sense (like employee type). Change font. Place all controls correctly.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#Enable nicer looking visual styles, especially on the DateTimePickers.
[System.Windows.Forms.Application]::EnableVisualStyles()

$Script:FormFont  = New-Object System.Drawing.Font("Microsoft Sans Serif", 8.25)
$Script:IconsFont = New-Object System.Drawing.Font("Segoe MDL2 Assets", 10, [System.Drawing.FontStyle]::Bold)

$Script:CurrentDate                = (Get-Date).Date
$Script:LeaveYearBeginningBaseline = (Get-Date -Year 2023 -Month 1 -Day 1).Date #In 2023 the first pay period began on 1 Jan. Use this as a baseline.
$Script:BeginningOfPayPeriod       = $Null
$Script:LastSelectableDate         = $Null
$Script:ThreeYearMark              = $Null
$Script:FifteenYearMark            = $Null
$Script:WorkHoursPerPayPeriod      = 0

$Script:LengthOfServiceString = $Null
$Script:DateOfMileStoneString = $Null
$Script:TimeUntilMilestone    = $Null

$Script:MaximumAnnual = 936 #With SES ceiling of 720 and SES accrual and rare year with 27 pay periods.
$Script:MaximumSick   = 9999 #Assuming you work full time, every year has 27 pay periods, and you never use a single hour, it would take almost 93 years to reach this amount.

$Script:LeaveBalances  = New-Object System.Collections.Generic.List[PSCustomObject]
$Script:ProjectedLeave = New-Object System.Collections.Generic.List[PSCustomObject]

$Script:ConfigFile               = "$env:APPDATA\PowerShell Scripts\Federal Civilian Leave Calculator\Federal Civilian Leave Calculator.ini"
$Script:HolidayWebsite           = "https://www.opm.gov/policy-data-oversight/pay-leave/federal-holidays/"
$Script:HolidaysHashTable        = @{}
$Script:InaugurationDayHashTable = @{}

$Script:UnsavedProjectedLeave = $False

#region Default Settings -- Changing values here only affects the first launch of the program. After that, it loads from a config file.

$Script:LastLaunchedDate = $Script:CurrentDate
$Script:SCDLeaveDate     = $Script:CurrentDate

$Script:DisplayLengthOrTimeUntilMilestone = "LengthOfService" #Options are "LengthOfService", "DateOfMilestone", and "TimeUntilMilestone"

$Script:EmployeeType        = "Full-Time" #Options are "Full-Time", "Part-Time", and "SES"
$Script:LeaveCeiling        = 240 #Options are 240, 360, and 720
$Script:InaugurationHoliday = $False
$Script:WorkSchedule = @{
    #Week 1
    PayPeriodDay1 = 0 #Sunday
    PayPeriodDay2 = 8 #Monday
    PayPeriodDay3 = 8 #Tuesday
    PayPeriodDay4 = 8 #Wednesday
    PayPeriodDay5 = 8 #Thursday
    PayPeriodDay6 = 8 #Friday
    PayPeriodDay7 = 0 #Saturday

    #Week 2
    PayPeriodDay8  = 0 #Sunday
    PayPeriodDay9  = 8 #Monday
    PayPeriodDay10 = 8 #Tuesday
    PayPeriodDay11 = 8 #Wednesday
    PayPeriodDay12 = 8 #Thursday
    PayPeriodDay13 = 8 #Friday
    PayPeriodDay14 = 0 #Saturday
}

$Script:ProjectOrGoal = "Project" #Options are "Project" and "Goal"
$Script:ProjectToDate = $Script:CurrentDate
$Script:AnnualGoal    = 0
$Script:SickGoal      = 0
$Script:AnnualDecimal = 0.0
$Script:SickDecimal   = 0.0

$Script:DisplayAfterEachLeave = $False
$Script:DisplayAfterEachPP    = $False
$Script:DisplayHighsAndLows   = $False

$AnnualLeaveCustomObject = [PSCustomObject] @{
    Name      = [String]"Annual"
    Balance   = [Int32]0
    Threshold = [Int32]0
    Static    = [Boolean]$True
}

$SickLeaveCustomObject = [PSCustomObject] @{
    Name      = [String]"Sick"
    Balance   = [Int32]0
    Threshold = [Int32]0
    Static    = [Boolean]$True
}

$Script:LeaveBalances.Add($AnnualLeaveCustomObject)
$Script:LeaveBalances.Add($SickLeaveCustomObject)

#endregion Default Settings

#region Functions

function Main
{
    $Script:BeginningOfPayPeriod = GetBeginningOfPayPeriodForDate -Date $Script:CurrentDate
    $Script:LastSelectableDate   = GetLeaveYearEndForDate -Date ((Get-Date -Year ($Script:BeginningOfPayPeriod.Year + 2) -Month $Script:BeginningOfPayPeriod.Month -Day $Script:BeginningOfPayPeriod.Day)).Date
    GetOpmHolidaysForYears
    LoadConfig
    SetWorkHoursPerPayPeriod
    GetAccrualRateDateChange
    UpdateExistingBalancesAndProjectedLeaveAtLaunch

    BuildMainForm
    
    [System.Windows.Forms.Application]::Run($Script:MainForm)
}

function GetAccrualRateDateChange
{
    $Script:ThreeYearMark   = $Script:SCDLeaveDate.AddYears(3)
    $Script:FifteenYearMark = $Script:SCDLeaveDate.AddYears(15)

    if($Script:ThreeYearMark -ne (GetBeginningOfPayPeriodForDate -Date $Script:ThreeYearMark))
    {
        $Script:ThreeYearMark = (GetEndingOfPayPeriodForDate -Date $Script:ThreeYearMark).AddDays(1)
    }

    if($Script:FifteenYearMark -ne (GetBeginningOfPayPeriodForDate -Date $Script:FifteenYearMark))
    {
        $Script:FifteenYearMark = (GetEndingOfPayPeriodForDate -Date $Script:FifteenYearMark).AddDays(1)
    }
}

function GetAnnualLeaveAccrualHours
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)][DateTime] $PayPeriod
    )
    
    $AccruedHours = 0

    if($Script:EmployeeType -eq "SES")
    {
        $AccruedHours = 8
    }

    else
    {
        if($Script:EmployeeType -eq "Full-Time")
        {
            if($PayPeriod -ge $Script:FifteenYearMark)
            {
                $AccruedHours = 8
            }

            elseif($PayPeriod -ge $Script:ThreeYearMark)
            {
                if((GetEndingOfPayPeriodForDate -Date $PayPeriod) -eq (GetLeaveYearEndForDate -Date $PayPeriod))
                {
                    $AccruedHours = 10
                }

                else
                {
                    $AccruedHours = 6
                }
            }

            else
            {
                $AccruedHours = 4
            }
        }

        else #Part-Time
        {
            if($PayPeriod -ge $Script:FifteenYearMark)
            {
                $AccruedHours = $Script:WorkHoursPerPayPeriod / 10
            }

            elseif($PayPeriod -ge $Script:ThreeYearMark)
            {
                $AccruedHours = $Script:WorkHoursPerPayPeriod / 13
            }

            else
            {
                $AccruedHours = $Script:WorkHoursPerPayPeriod / 20
            }
        }
    }

    return $AccruedHours
}

function GetBeginningOfPayPeriodForDate
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)] $Date
    )

    #Get the modulus of the number of days between the baseline and provided date, then subtract it from the provided date to get the beginning of the pay period.
    return $Date.AddDays(-($Date - $Script:LeaveYearBeginningBaseline).Days % 14)
}

function GetEndingOfPayPeriodForDate
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)] $Date
    )

    #Get the beginning of the pay period and then just add 13 days.
    return (GetBeginningOfPayPeriodForDate -Date $Date).AddDays(13)
}

function GetHoursForWorkDay
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)][DateTime] $Day
    )
    
    $Hours = 0
    
    if($Script:HolidaysHashTable.ContainsKey($Day.ToString("MM/dd/yyyy")) -eq $False)
    {
        if($Script:InaugurationHoliday -eq $False -or
           $Script:InaugurationDayHashTable.ContainsKey($Day.ToString("MM/dd/yyyy")) -eq $False)
        {
            $DayOfPayPeriod = (($Day - $Script:BeginningOfPayPeriod).Days % 14) + 1

            $Hours = $Script:WorkSchedule.("PayPeriodDay" + $DayOfPayPeriod)
        }
    }
    
    return $Hours
}

function GetLeaveYearEndForDate
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)] $Date
    )

    $Year = (GetBeginningOfPayPeriodForDate -Date $Date).Year
    
    return (GetEndingOfPayPeriodForDate -Date (Get-Date -Year $Year -Month 12 -Day 31).Date)
}

function GetLeaveYearScheduleDeadline
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)] $Date
    )

    return $Date.AddDays(-42) #42 Days is the day before the start of the third biweekly pay period prior to the end of the leave year.
}

function GetOpmHolidaysForYears
{
    $BeginningYear = ($Script:BeginningOfPayPeriod).Year
    $EndingYear    = $Script:LastSelectableDate.Year

    $Failed = $False
    
    try
    {
        $HolidayContent = (Invoke-WebRequest -DisableKeepAlive -Uri $Script:HolidayWebsite).Content
    }

    catch
    {
        $Failed = $True
    }

    for($TargetYear = $BeginningYear; $TargetYear -le $EndingYear; $TargetYear++)
    {
        $YearBeginningIndex = $Null
        $YearEndingIndex    = $Null
        $YearContent        = $Null

        try
        {
            if(($HolidayContent -match "<section class=`"tab-content`" title=`"$TargetYear`">") -eq $True) #If the year is in the tab list and not under Historical Data.
            {
                $YearBeginningIndex = $HolidayContent.IndexOf("<section class=`"tab-content`" title=`"$TargetYear`">")
                $YearEndingIndex = $HolidayContent.IndexOf("</section>", $YearBeginningIndex)
                $YearContent = $HolidayContent.Substring($YearBeginningIndex, ($YearEndingIndex - $YearBeginningIndex + 10)) #The +10 is so it includes the </section>
            }

            else #If the year is under the Historical Data tab.
            {
                $YearBeginningIndex = $HolidayContent.IndexOf("<table class=`"DataTable HolidayTable`"><caption>$TargetYear Holiday Schedule")
                $YearEndingIndex = $HolidayContent.IndexOf("<p class=`"top`"><a href=`"#content`">Back to top</a></p>", $YearBeginningIndex)

                if($YearEndingIndex -eq -1)
                        {
                $YearEndingIndex = $HolidayContent.LastIndexOf("</p>") + 4 #The + 4 is so we get the "</p>" so every holiday content we get should be identical for further processing.
            }

                $YearContent = $HolidayContent.Substring($YearBeginningIndex, ($YearEndingIndex - $YearBeginningIndex))
            }
        }

        catch
        {
            $Failed = $True
        }
        
        try
        {
            $HolidayList = $YearContent.Split("`n") | Select-String -Pattern "<td>" -SimpleMatch #For some reason even though it has new lines, Select-String returns the entire thing unless I split it. Only the actual dates have the <td> tag which is why we filter on that.
        }

        catch
        {
            $Failed = $True
        }

        if($Failed -eq $False)
        {
            foreach($Line in $HolidayList)
            {
                if($HolidayList.IndexOf($Line) % 2 -eq 0) #So we only get the actual dates, not which holiday it is (while useful for humans, not useful for the script)
                {
                    $ModifiedLine = $Line.ToString().Replace("<td>", "").Replace("</td>", "").Trim() #Get rid of the additional HTML tags

                    if($ModifiedLine.Contains("<") -eq $True)
                    {
                        $ModifiedLine = $ModifiedLine.Substring(0, ($ModifiedLine.IndexOf("<"))).Trim() #If there are any notes on the date marked with an *, clear them out
                    }

                    $ModifiedLine = $ModifiedLine.Substring($ModifiedLine.IndexOf(" ") + 1) #Gets rid of the day of the week in front.

                    if(($HolidayList.IndexOf($Line) -eq 0) -and ($Line.ToString().Contains("December"))) #Sometimes OPM includes New Years in the previous year because it falls on a Saturday, so the holiday is given on a Friday.
                    {
                        $PreviousYear = $TargetYear - 1
                    
                        $ModifiedLine += ", $PreviousYear" #Add the previous year at the end
                    }

                    else
                    {
                        $ModifiedLine += ", $TargetYear" #Add the year at the end
                    }

                    if($ModifiedLine -match "\w* \d{2}, \d{4}, \d{4}") #Some of the entries rarely have the year already listed in the date, so we need to strip that off.
                    {
                        $ModifiedLine = $ModifiedLine.Substring(0, $ModifiedLine.Length - 6) #Trim off the year, space, and comma.
                    }

                    $DateObject = Get-Date $ModifiedLine

                    $DateString = $DateObject.ToString("MM/dd/yyyy")

                    if($HolidaysHashTable.Contains($DateString) -eq $False)
                    {
                        $HolidayNameString = $HolidayList[$HolidayList.IndexOf($Line) + 1].ToString().Replace("<td>", "").Replace("</td>", "").Trim()

                        if($HolidayNameString.ToLower().Contains("inauguration") -eq $True)
                        {
                            $InaugurationDayHashTable[$DateString] = $HolidayNameString
                        }

                        else
                        {
                            $HolidaysHashTable[$DateString] = $HolidayNameString
                        }
                    }
                }
            }
        }
    }

    if($Failed -eq $True)
    {
        $Null = ShowMessageBox -Text "Unable to automatically get holidays. When creating/editing projected leave, manually check that hours on holidays are correct." -Caption "Error Retrieving Holidays" -Buttons "OK" -Icon "Error"
    }
}

function GetSickLeaveAccrualHours
{
    $AccruedHours = 4

    if($Script:EmployeeType -eq "Part-Time")
    {
        $AccruedHours = $Script:WorkHoursPerPayPeriod / 20
    }

    return $AccruedHours
}

function SetWorkHoursPerPayPeriod
{
    $Script:WorkHoursPerPayPeriod = 0

    foreach($Day in $Script:WorkSchedule.Keys)
    {
        $Script:WorkHoursPerPayPeriod += $Script:WorkSchedule[$Day]
    }
}

function LoadConfig
{
    if((Test-Path -Path $Script:ConfigFile) -eq $True)
    {
        $Errors = $False
            
        try
        {
            $LoadedConfig = Import-Clixml -Path $Script:ConfigFile

            $LeaveNameHashTable = @{}

            if($LoadedConfig.GetType().Name -eq "Object[]")
            {
                if($LoadedConfig[0] -is [DateTime]) #Load LastLaunchedDate DateTime
                {
                    $Script:LastLaunchedDate = $LoadedConfig[0]
                }

                else
                {
                    $Errors = $True
                }

                if($LoadedConfig[1] -is [DateTime] -and
                   $LoadedConfig[1] -le $Script:CurrentDate) #Load SCDLeaveDate DateTime
                {
                    $Script:SCDLeaveDate = $LoadedConfig[1]
                }

                else
                {
                    $Errors = $True
                }

                if($LoadedConfig[2] -is [String] -and
                  ($LoadedConfig[2] -eq "LengthOfService" -or
                   $LoadedConfig[2] -eq "DateOfMilestone" -or
                   $LoadedConfig[2] -eq "TimeUntilMilestone")) #Load DisplayLengthOrtimeUntilMilestone String
                {
                    $Script:DisplayLengthOrTimeUntilMilestone = $LoadedConfig[2]
                }

                else
                {
                    $Errors = $True
                }
                    
                if($LoadedConfig[3] -is [String] -and
                  ($LoadedConfig[3] -eq "Full-Time" -or
                   $LoadedConfig[3] -eq "Part-Time" -or
                   $LoadedConfig[3] -eq "SES")) #Load EmployeeType String
                {
                    $Script:EmployeeType = $LoadedConfig[3]
                }

                else
                {
                    $Errors = $True
                }

                if($LoadedConfig[4] -is [Int32] -and
                  ($LoadedConfig[4] -eq 240 -or
                   $LoadedConfig[4] -eq 360 -or
                   $LoadedConfig[4] -eq 720)) #Load LeaveCeiling Int32
                {
                    $Script:LeaveCeiling = $LoadedConfig[4]
                }

                else
                {
                    $Errors = $True
                }
                    
                if($LoadedConfig[5] -is [Boolean]) #Load InaugurationHoliday boolean
                {
                    $Script:InaugurationHoliday = $LoadedConfig[5]
                }

                else
                {
                    $Errors = $True
                }
                    
                if($LoadedConfig[6] -is [Hashtable] -and
                   $LoadedConfig[6].Count -eq 14) #Load the WorkSchedule Hashtable
                {
                    $LoadWorkSchedule = $True
                        
                    foreach($Key in $LoadedConfig[6].Keys)
                    {
                        #Regex Matches:
                        #Line start     ^
                        #Literal text    PayPeriodDay
                        #Non-capturing group         (?:            )
                        #One digit 1-9                  [1-9]
                        #Or                                  |
                        #1 followed by...                     1
                        #One digit 0-4                         [0-4]
                        #Line end                                    $
                        if($Key -match "^PayPeriodDay(?:[1-9]|1[0-4])$" -eq $False -and
                           $LoadedConfig[6][$Key] -isnot [Int32] -and
                           $LoadedConfig[6][$Key] -lt 0 -and
                           $LoadedConfig[6][$Key] -gt 24)
                        {
                            $LoadWorkSchedule = $False
                        }
                    }

                    if($LoadWorkSchedule -eq $True)
                    {
                        $Script:WorkSchedule = $LoadedConfig[6]
                    }

                    else
                    {
                        $Errors = $True
                    }
                }

                else
                {
                    $Errors = $True
                }

                if($LoadedConfig[7] -is [String] -and
                  ($LoadedConfig[7] -eq "Project" -or
                   $LoadedConfig[7] -eq "Goal")) #Load ProjectOrGoal String
                {
                    $Script:ProjectOrGoal = $LoadedConfig[7]
                }

                else
                {
                    $Errors = $True
                }

                if($LoadedConfig[8] -is [DateTime] -and
                   $LoadedConfig[8] -le $Script:LastSelectableDate) #Load ProjectToDate DateTime
                {
                    $Script:ProjectToDate = $LoadedConfig[8]
                }

                else
                {
                    $Errors = $True
                }

                if($LoadedConfig[9] -is [Int32] -and
                   $LoadedConfig[9] -ge 0 -and
                   $LoadedConfig[9] -le $Script:MaximumAnnual) #Load AnnualGoal Int32
                {
                    $Script:AnnualGoal = $LoadedConfig[9]
                }

                else
                {
                    $Errors = $True
                }

                if($LoadedConfig[10] -is [Int32] -and
                   $LoadedConfig[10] -ge 0 -and
                   $LoadedConfig[10] -le $Script:MaximumSick) #Load SickGoal Int32
                {
                    $Script:SickGoal = $LoadedConfig[10]
                }

                else
                {
                    $Errors = $True
                }

                if($LoadedConfig[11] -is [Double] -and
                   $LoadedConfig[11] -ge 0 -and
                   $LoadedConfig[11] -lt 1) #Load AnnualDecimal Int32
                {
                    $Script:AnnualDecimal = $LoadedConfig[11]
                }

                else
                {
                    $Errors = $True
                }

                if($LoadedConfig[12] -is [Double] -and
                   $LoadedConfig[12] -ge 0 -and
                   $LoadedConfig[12] -lt 1) #Load SickDecimal Int32
                {
                    $Script:SickDecimal = $LoadedConfig[12]
                }

                else
                {
                    $Errors = $True
                }

                if($LoadedConfig[13] -is [Boolean]) #Load DisplayAfterEachLeave boolean
                {
                    $Script:DisplayAfterEachLeave = $LoadedConfig[13]
                }

                else
                {
                    $Errors = $True
                }

                if($LoadedConfig[14] -is [Boolean]) #Load DisplayAfterEachPP boolean
                {
                    $Script:DisplayAfterEachPP = $LoadedConfig[14]
                }

                else
                {
                    $Errors = $True
                }

                if($LoadedConfig[15] -is [Boolean]) #Load DisplayHighsAndLows boolean
                {
                    $Script:DisplayHighsAndLows = $LoadedConfig[15]
                }

                else
                {
                    $Errors = $True
                }

                $Index = 16 #First item in the array that belongs to a List.
                    
                while($Index -lt $LoadedConfig.Count -and
                      $LoadedConfig[$Index] -is [PSCustomObject] -and
                      $LoadedConfig[$Index].PSobject.Properties.Name.Contains("Name") -eq $True -and
                      $LoadedConfig[$Index].Name -is [String]) #Loop to add all Leave Balances.
                {
                    if($LoadedConfig[$Index].Name -eq "Annual")
                    {
                        if($LoadedConfig[$Index].Balance -is [Decimal] -and
                           $LoadedConfig[$Index].Balance -ge 0 -and
                           $LoadedConfig[$Index].Balance -le $Script:MaximumAnnual)
                        {
                            $Script:LeaveBalances[0].Balance = [Int32]$LoadedConfig[$Index].Balance
                        }

                        else
                        {
                            $Errors = $True
                        }

                        if($LoadedConfig[$Index].Threshold -is [Decimal] -and
                           $LoadedConfig[$Index].Threshold -ge 0 -and
                           $LoadedConfig[$Index].Threshold -le $Script:MaximumAnnual)
                        {
                            $Script:LeaveBalances[0].Threshold = [Int32]$LoadedConfig[$Index].Threshold
                        }

                        else
                        {
                            $Errors = $True
                        }
                    }

                    elseif($LoadedConfig[$Index].Name -eq "Sick")
                    {
                        if($LoadedConfig[$Index].Balance -is [Decimal] -and
                           $LoadedConfig[$Index].Balance -ge 0 -and
                           $LoadedConfig[$Index].Balance -le $Script:MaximumSick)
                        {
                            $Script:LeaveBalances[1].Balance = [Int32]$LoadedConfig[$Index].Balance
                        }

                        else
                        {
                            $Errors = $True
                        }

                        if($LoadedConfig[$Index].Threshold -is [Decimal] -and
                           $LoadedConfig[$Index].Threshold -ge 0 -and
                           $LoadedConfig[$Index].Threshold -le $Script:MaximumSick)
                        {
                            $Script:LeaveBalances[1].Threshold = [Int32]$LoadedConfig[$Index].Threshold
                        }

                        else
                        {
                            $Errors = $True
                        }
                    }

                    else
                    {
                            #Regex matches a set                     [           ]
                            #Exludes this set                         ^
                            #Capital letters                           A-Z
                            #Lowercase letters                            a-z
                            #Digits                                          0-9
                            #Also a space                                       <space>
                        
                        if($LoadedConfig[$Index].Name.Trim().Length -gt 0 -and
                           $LoadedConfig[$Index].Name.Trim().Length -le 1000 -and #Todo Set maximum length here to match whatever is on the input field.
                           $LoadedConfig[$Index].Name.Trim() -match "[^A-Za-z0-9 ]" -eq $False -and
                           $LoadedConfig[$Index].Name.ToLower().Trim() -ne "annual" -and
                           $LoadedConfig[$Index].Name.ToLower().Trim() -ne "sick" -and
                           $LoadedConfig[$Index].Balance -is [Decimal] -and #For some reason after importing the file the Int32 is now a Decimal. We cast it back later.
                           $LoadedConfig[$Index].Balance -ge 0 -and
                           $LoadedConfig[$Index].Balance -le $Script:MaximumSick -and
                           $LoadedConfig[$Index].Expires -is [Boolean] -and
                           $LoadedConfig[$Index].ExpiresOn -is [DateTime])
                        {
                            $NewLeaveBalance = [PSCustomObject] @{
                                Name      = [String]$LoadedConfig[$Index].Name.Trim()
                                Balance   = [Int32]$LoadedConfig[$Index].Balance
                                Expires   = [Boolean]$LoadedConfig[$Index].Expires
                                ExpiresOn = [DateTime]$LoadedConfig[$Index].ExpiresOn.Date
                                Static    = [Boolean]$False
                            }
                            
                            $Script:LeaveBalances.Add($NewLeaveBalance)
                        }
                        
                        else
                        {
                            $Errors = $True
                        }
                    }
                    
                    $Index++
                }

                #Loop through the leave balances and ensure they are valid (leave can have the same name if one expires and the other doesn't, or both expire on different days).
                
                $LeaveIndex = 2 #Skip Annual/Sick

                while($LeaveIndex -lt $Script:LeaveBalances.Count)
                {
                    $LeaveName = $Script:LeaveBalances[$LeaveIndex].Name

                    $InnerIndex = $LeaveIndex + 1
                    
                    while($InnerIndex -lt $Script:LeaveBalances.Count)
                    {
                        if($LeaveName -eq $Script:LeaveBalances[$InnerIndex].Name -and
                         (($Script:LeaveBalances[$LeaveIndex].Expires -eq $False -and
                           $Script:LeaveBalances[$InnerIndex].Expires -eq $False) -or
                          ($Script:LeaveBalances[$LeaveIndex].Expires -eq $True -and
                           $Script:LeaveBalances[$InnerIndex].Expires -eq $True -and
                           $Script:LeaveBalances[$LeaveIndex].ExpiresOn -eq $Script:LeaveBalances[$InnerIndex].ExpiresOn)))
                        {
                            $Script:LeaveBalances.RemoveAt($InnerIndex)
                        }
                        
                        else
                        {
                            $InnerIndex++
                        }
                    }
                    
                    $LeaveIndex++
                }

                #Sort the balances to prevent any funny business.
                $Script:LeaveBalances = [System.Collections.Generic.List[PSCustomObject]] ($Script:LeaveBalances | Sort-Object -Property @{Expression = "Static"; Descending = $True},
                                                                                                                                         @{Expression = "Name"; Descending = $False},
                                                                                                                                         @{Expression = "Expires"; Descending = $True},
                                                                                                                                         @{Expression = "ExpiresOn"; Descending = $False})

                foreach($Balance in $Script:LeaveBalances) #Get a list of valid leave balance names for the LeaveBank and the expiration dates.
                {
                    $LatestDate = $Script:CurrentDate
                    
                    if($Balance.Name.ToLower() -eq "annual" -or
                       $Balance.Name.ToLower() -eq "sick")
                    {
                        $LatestDate = $Script:LastSelectableDate
                    }

                    else
                    {
                        if($Balance.Expires -eq $True -and
                           $Balance.ExpiresOn -gt $LatestDate)
                        {
                            $LatestDate = $Balance.ExpiresOn
                        }

                        elseif($Balance.Expires -eq $False)
                        {
                            $LatestDate = $Script:LastSelectableDate
                        }
                    }

                    $LeaveNameHashTable.($Balance.Name) = $LatestDate
                }
                
                while($Index -lt $LoadedConfig.Count -and
                      $LoadedConfig[$Index] -is [PSCustomObject] -and
                      $LoadedConfig[$Index].PSobject.Properties.Name.Contains("LeaveBank") -eq $True -and
                      $LoadedConfig[$Index].LeaveBank -is [String]) #Loop to add all Projected Leave.
                {
                    if($LeaveNameHashTable.ContainsKey($LoadedConfig[$Index].LeaveBank.Trim().ToLower()) -eq $True -and
                       $LoadedConfig[$Index].StartDate -is [DateTime] -and
                       $LoadedConfig[$Index].StartDate -le $Script:LastSelectableDate -and
                       $LoadedConfig[$Index].EndDate -is [DateTime] -and
                       $LoadedConfig[$Index].EndDate -le (GetEndingOfPayPeriodForDate -Date $LoadedConfig[$Index].StartDate) -and
                       $LoadedConfig[$Index].EndDate -ge $LoadedConfig[$Index].StartDate)
                    {
                        $NewProjectedLeave = [PSCustomObject] @{
                            LeaveBank      = [String]$LoadedConfig[$Index].LeaveBank.Trim()
                            StartDate      = [DateTime]$LoadedConfig[$Index].StartDate.Date
                            EndDate        = [DateTime]$LoadedConfig[$Index].EndDate.Date
                            HoursHashTable = @{}
                        }

                        #Populate the HoursHashTable.
                        for($Date = $NewProjectedLeave.StartDate; $Date -le $NewProjectedLeave.EndDate; $Date = $Date.AddDays(1))
                        {
                            if($LoadedConfig[$Index].HoursHashTable.ContainsKey($Date.ToString("MM/dd/yyyy")) -eq $True -and
                               $LoadedConfig[$Index].HoursHashTable[$Date.ToString("MM/dd/yyyy")] -ge 0 -and
                               $LoadedConfig[$Index].HoursHashTable[$Date.ToString("MM/dd/yyyy")] -le 24) #If the hashtable value exists and is valid.
                            {
                                $NewProjectedLeave.HoursHashTable[$Date.ToString("MM/dd/yyyy")] = $LoadedConfig[$Index].HoursHashTable[$Date.ToString("MM/dd/yyyy")]
                            }

                            else
                            {
                                $NewProjectedLeave.HoursHashTable[$Date.ToString("MM/dd/yyyy")] = GetHoursForWorkDay -Day $Date.ToString("MM/dd/yyyy")

                                $Errors = $True
                            }
                        }
                        
                        $Script:ProjectedLeave.Add($NewProjectedLeave)
                    }

                    else
                    {
                        $Errors = $True
                    }
                    
                    $Index++
                }

                #Sort the projected leave to prevent any funny business.
                $Script:ProjectedLeave = [System.Collections.Generic.List[PSCustomObject]] ($Script:ProjectedLeave | Sort-Object -Property "StartDate", "EndDate", "LeaveBank")
            }

            else
            {
                $Errors = $True
            }
        }

        catch
        {
            $Errors = $True
        }

        if($Errors -eq $True)
        {
            $Null = ShowMessageBox -Text "There was a problem opening the configuration file. This is either due to a corrupted file or invalid data. Please verify that your information is correct." -Caption "Error Opening Config File" -Buttons "OK" -Icon "Error"
        }
    }
}

function NumberGetsLetterS($String, $Number)
{
    if($Number -ne 1)
    {
        $String += "s"
    }

    return $String
}

function OutputFormAppendText
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)]  $Text,
        [parameter(Mandatory=$False)] $Alignment,
        [parameter(Mandatory=$False)] $Color,
        [parameter(Mandatory=$False)] $FontSize,
        [parameter(Mandatory=$False)] $FontStyle
    )

    $RichTextBox = $Script:OutputForm.Controls["OutputRichTextBox"]

    $RichTextBox.SelectionStart  = $RichTextBox.TextLength
    $RichTextBox.SelectionLength = 0

    $RichTextBox.SelectionAlignment = "Left"
    $RichTextBox.SelectionColor     = "Black"

    $NewFontSize  = $RichTextBox.Font.Size
    $NewFontStyle = $RichTextBox.Font.Style

    if($Alignment -ne $Null) #Can be "Center", "Left", or "Right"
    {
        $RichTextBox.SelectionAlignment = $Alignment
    }

    if($Color -ne $Null) #Can be any color here: https://learn.microsoft.com/en-us/dotnet/api/system.windows.media.colors?view=windowsdesktop-8.0
    {
        $RichTextBox.SelectionColor = $Color
    }

    if($FontSize -ne $Null) #Pick a point size for the font.
    {
        $NewFontSize = $FontSize
    }

    if($FontStyle -ne $Null) #Can be Bold, Italic, Regular, Strikeout, or Underline. Look at the Enum to do combinations. https://learn.microsoft.com/en-us/dotnet/api/system.drawing.fontstyle?view=windowsdesktop-8.0
    {
        $NewFontStyle = $FontStyle
    }

    if($FontSize -ne $Null -or
       $FontStyle -ne $Null)
    {
        $RichTextBox.SelectionFont = New-Object System.Drawing.Font($RichTextBox.Font.Name, $NewFontSize, [System.Drawing.FontStyle]::$NewFontStyle)
    }

    $RichTextBox.AppendText($Text)
}

function PopulateLeaveBalanceListBox
{
    $LeaveBalanceListBox = $Script:MainForm.Controls["LeavePanel"].Controls["BalanceListBox"]

    $LeaveBalanceListBox.BeginUpdate()

    $LeaveBalanceListBox.Items.Clear()
    
    foreach($LeaveItem in $Script:LeaveBalances) #Todo might adjust the `t tabbing in here for spacing.
    {
        $String = $LeaveItem.Name + "`t" + $LeaveItem.Balance

        if(($LeaveItem.Name -eq "Annual" -or
            $LeaveItem.Name -eq "Sick") -and
            $LeaveItem.Threshold -ne 0)
        {
            $String += "`tThreshold: " + $LeaveItem.Threshold
        }

        if($LeaveItem.Name -ne "Annual" -and
           $LeaveItem.Name -ne "Sick" -and
           $LeaveItem.Expires -eq $True)
        {
            $ExpireDate = $LeaveItem.ExpiresOn.ToString("MM/dd/yyyy")

            $String += "`tExpires: $ExpireDate"
        }

        $LeaveBalanceListBox.Items.Add($String) | Out-Null
    }

    $LeaveBalanceListBox.EndUpdate()
}

function PopulateOutputFormRichTextBox
{
    $LeaveBalancesCopy   = New-Object System.Collections.Generic.List[PSCustomObject]
    $LeaveExpiresOnList  = New-Object System.Collections.Generic.List[PSCustomObject]
    $ProjectedLeaveIndex = 0
    $EndOfPayPeriod      = GetEndingOfPayPeriodForDate -Date $Script:BeginningOfPayPeriod
    $LeaveYearEnd        = GetLeaveYearEndForDate -Date $EndOfPayPeriod

    $AnnualHigh = $Script:LeaveBalances[0].Balance
    $AnnualLow  = $Script:LeaveBalances[0].Balance
    $SickHigh   = $Script:LeaveBalances[1].Balance
    $SickLow    = $Script:LeaveBalances[1].Balance

    $ProjectionAnnualDecimal = $Script:AnnualDecimal
    $ProjectionSickDecimal   = $Script:SickDecimal

    $GoalsMet = $False

    foreach($LeaveBalance in $Script:LeaveBalances)
    {
        $LeaveBalancesCopy.Add($LeaveBalance.PSObject.Copy())

        if($LeaveBalance.Expires -eq $True)
        {
            $LeaveExpiresOnList.Add($LeaveBalance.ExpiresOn)
        }
    }
    
    $StartDate = $Script:BeginningOfPayPeriod

    if($Script:ProjectOrGoal -eq "Project")
    {
        $EndDate = GetEndingOfPayPeriodForDate -Date $Script:ProjectToDate
    }

    else
    {
        $EndDate = $Script:LastSelectableDate
    }

    $TitleString = ""

    if($Script:ProjectOrGoal -eq "Project")
    {
        $StartDateString = $StartDate.ToString("MM/dd/yyyy")
        $EndDateString   = $EndDate.ToString("MM/dd/yyyy")
        
        $TitleString = "Leave Projection From $StartDateString Through $EndDateString"
    }

    else
    {
        $TitleString = "Projection to Reach Goals"
    }

    $TitleString += "`n"

    #Add the title.
    OutputFormAppendText -Text $TitleString -Alignment "Center" -FontSize ($Script:OutputForm.Controls["OutputRichTextBox"].Font.Size + 2) -FontStyle "Bold"

    #Add the section header
    OutputFormAppendText -Text "`nStarting Balances:`n" -FontStyle "Bold"

    #Add Annual/Sick
    OutputFormAppendText -Text ($LeaveBalancesCopy[0].Name + ":`t" + $LeaveBalancesCopy[0].Balance)
    OutputFormAppendText -Text ("`n" + $LeaveBalancesCopy[1].Name + ":`t" + $LeaveBalancesCopy[1].Balance)

    #If projecting to date, append the rest of the balances and their expiration if they expire.
    if($Script:ProjectOrGoal -eq "Project")
    {
        for($Index = 2; $Index -lt $LeaveBalancesCopy.Count; $Index++)
        {
            $String = "`n" + $LeaveBalancesCopy[$Index].Name + ":`t" + $LeaveBalancesCopy[$Index].Balance

            if($LeaveBalancesCopy[$Index].Expires -eq $True)
            {
                $String += "`tExpires: " + $LeaveBalancesCopy[$Index].ExpiresOn.ToString("MM/dd/yyyy")
            }

            OutputFormAppendText -Text $String
        }
    }

    if($Script:ProjectOrGoal -eq "Goal" -and
        $LeaveBalancesCopy[0].Balance -ge $Script:AnnualGoal -and
        $LeaveBalancesCopy[1].Balance -ge $Script:SickGoal)
    {
        $GoalsMet = $True
    }

    #Loop through dates to determine when to accrue leave, subtract leave from hours, etc...
    while($EndOfPayPeriod -le $EndDate -and
          $GoalsMet -eq $False)
    {
        if($Script:EmployeeType -ne "SES")
        {
            if($EndOfPayPeriod.AddDays(-14) -lt $Script:FifteenYearMark -and
                   $EndOfPayPeriod -gt $Script:FifteenYearMark)
            {
                $String = "`n`nAnnual Leave Accrual Rate Changed to the Greater Than 15 Years Category on " + $Script:FifteenYearMark.ToString("MM/dd/yyyy") + "."
                
                OutputFormAppendText -Text $String -Color "Green"
            }

            elseif($EndOfPayPeriod.AddDays(-14) -lt $Script:ThreeYearMark -and
               $EndOfPayPeriod -gt $Script:ThreeYearMark)
            {
                $String = "`n`nAnnual Leave Accrual Rate Changed to the 3 to 15 Years Category on " + $Script:ThreeYearMark.ToString("MM/dd/yyyy") + "."
                
                OutputFormAppendText -Text $String -Color "Green"
            }
        }
        
        $LeaveExpiresThisPayPeriod = $False

        foreach($Expire in $LeaveExpiresOnList)
        {
            if($Expire -gt $EndOfPayPeriod.AddDays(-14) -and
               $Expire -le $EndOfPayPeriod)
            {
                $LeaveExpiresThisPayPeriod = $True
            }
        }

        if($Script:ProjectedLeave[$ProjectedLeaveIndex].StartDate -le $EndOfPayPeriod -or
           $LeaveExpiresThisPayPeriod -eq $True) #Need to simulate progressing the days one at a time for two weeks since something happens.
        {
            $Date = GetBeginningOfPayPeriodForDate -Date $EndOfPayPeriod
            
            $ProjectedLeaveEndIndex = $ProjectedLeaveIndex

            while($ProjectedLeaveEndIndex -lt $Script:ProjectedLeave.Count -and
                  $Script:ProjectedLeave[$ProjectedLeaveEndIndex].StartDate -le $EndOfPayPeriod)
            {
                $ProjectedLeaveEndIndex++
            }

            for($Day = 0; $Day -lt 14; $Day++)
            {
                for($ProjectedIndex = $ProjectedLeaveIndex; $ProjectedIndex -lt $ProjectedLeaveEndIndex; $ProjectedIndex++)
                {
                    if($Script:ProjectedLeave[$ProjectedIndex].HoursHashTable.ContainsKey($Date.ToString("MM/dd/yyyy")) -eq $True -and
                       $Script:ProjectedLeave[$ProjectedIndex].HoursHashTable[$Date.ToString("MM/dd/yyyy")] -gt 0)
                    {
                        $BalanceFound = $False
                        $BalanceIndex = 0

                        while($BalanceIndex -lt $LeaveBalancesCopy.Count -and
                              $LeaveBalancesCopy[$BalanceIndex].Name -ne $Script:ProjectedLeave[$ProjectedIndex].LeaveBank)
                        {
                            $BalanceIndex++
                        }

                        if($BalanceIndex -lt $LeaveBalancesCopy.Count -and
                           $LeaveBalancesCopy[$BalanceIndex].Name -eq $Script:ProjectedLeave[$ProjectedIndex].LeaveBank)
                        {
                            $BalanceFound = $True
                        }

                        if($BalanceFound -eq $True) #This is where the leave is subtracted off the balance.
                        {
                            $BalanceBeforeSubtraction = $LeaveBalancesCopy[$BalanceIndex].Balance
                            
                            $LeaveBalancesCopy[$BalanceIndex].Balance -= $Script:ProjectedLeave[$ProjectedIndex].HoursHashTable[$Date.ToString("MM/dd/yyyy")]

                            while($LeaveBalancesCopy[$BalanceIndex].Balance -lt 0 -and
                                 ($BalanceIndex + 1) -lt $LeaveBalancesCopy.Count -and
                                  $LeaveBalancesCopy[$BalanceIndex + 1].Name -eq $LeaveBalancesCopy[$BalanceIndex].Name)
                            {
                                $LeaveBalancesCopy[$BalanceIndex + 1].Balance += $LeaveBalancesCopy[$BalanceIndex].Balance  #Take that negative balance off of the next entry with the same name.
                                
                                $Count = 0

                                foreach($LeaveBalance in $LeaveBalancesCopy)
                                {
                                    if($LeaveBalance.Expires -eq $True -and
                                       $LeaveBalance.ExpiresOn -eq $LeaveBalancesCopy[$BalanceIndex].ExpiresOn)
                                    {
                                        $Count++
                                    }
                                }

                                if($Count -gt 1)
                                {
                                    $LeaveExpiresOnList.Remove($LeaveBalance.ExpiresOn.ToString("MM/dd/yyyy"))
                                    Write-Host "Flag"
                                }

                                $LeaveBalancesCopy.RemoveAt($BalanceIndex)
                            }

                            if($Script:DisplayHighsAndLows -eq $True)
                            {
                                if($LeaveBalancesCopy[0].Balance -lt $AnnualLow)
                                {
                                    $AnnualLow = $LeaveBalancesCopy[0].Balance
                                }

                                if($LeaveBalancesCopy[1].Balance -lt $SickLow)
                                {
                                    $SickLow = $LeaveBalancesCopy[1].Balance
                                }
                            }

                            if($Script:DisplayAfterEachLeave -eq $True -and
                              ($Script:ProjectOrGoal -eq "Project" -or
                               $LeaveBalancesCopy[$BalanceIndex].Name -eq "Annual" -or
                               $LeaveBalancesCopy[$BalanceIndex].Name -eq "Sick"))
                            {
                                $LeaveName = $LeaveBalancesCopy[$BalanceIndex].Name

                                if($LeaveName.ToLower().Contains("leave") -eq $False)
                                {
                                   $LeaveName += " Leave"
                                }
                                
                                $String  = "`n`n" + $Script:ProjectedLeave[$ProjectedIndex].HoursHashTable[$Date.ToString("MM/dd/yyyy")] + " Hours of $LeaveName Taken on " + $Date.ToString("MM/dd/yyyy") + "."
                                $String += "`n$LeaveName Balance is now: " + [Math]::Floor($LeaveBalancesCopy[$BalanceIndex].Balance)

                                OutputFormAppendText -Text $String
                            }

                            if($LeaveBalancesCopy[$BalanceIndex].Name -eq "Annual" -or
                               $LeaveBalancesCopy[$BalanceIndex].Name -eq "Sick" -and
                              ($LeaveBalancesCopy[$BalanceIndex].Threshold -gt 0 -and
                               $BalanceBeforeSubtraction -ge $LeaveBalancesCopy[$BalanceIndex].Threshold -and
                               $LeaveBalancesCopy[$BalanceIndex].Balance -lt $LeaveBalancesCopy[$BalanceIndex].Threshold))
                            {
                                $String = "`n`n" + $LeaveBalancesCopy[$BalanceIndex].Name + " Leave balance is " + $LeaveBalancesCopy[$BalanceIndex].Balance + " which is below the set threshold of " + $LeaveBalancesCopy[$BalanceIndex].Threshold + " after taking leave on " + $Date.ToString("MM/dd/yyyy") + "."
                                
                                OutputFormAppendText -Text $String -Color "Blue"
                            }

                            if($LeaveBalancesCopy[$BalanceIndex].Balance -lt 0)
                            {
                                $LeaveName = $LeaveBalancesCopy[$BalanceIndex].Name

                                if($LeaveName.ToLower().Contains("leave") -eq $False)
                                {
                                   $LeaveName += " Leave"
                                }
                                
                                $String = "`n`n$LeaveName balance is negative (" + $LeaveBalancesCopy[$BalanceIndex].Balance + ") after taking leave on " + $Date.ToString("MM/dd/yyyy") + "."

                                OutputFormAppendText -Text $String -Color "Red"
                            }
                        }
                    }
                }

                if($LeaveExpiresThisPayPeriod -eq $True)
                {
                    $Index = 0
                    
                    while($Index -lt $LeaveBalancesCopy.Count)
                    {
                        if($LeaveBalancesCopy[$Index].Expires -eq $True -and
                           $LeaveBalancesCopy[$Index].Balance -gt 0 -and
                           $LeaveBalancesCopy[$Index].ExpiresOn -eq $Date)
                        {
                            $LeaveName = $LeaveBalancesCopy[$Index].Name

                            if($LeaveName.ToLower().Contains("leave") -eq $False)
                            {
                                $LeaveName += " Leave"
                            }
                            
                            $String = "`n`n" + $LeaveBalancesCopy[$Index].Balance + " hours of $LeaveName will expire on " + $Date.ToString("MM/dd/yyyy") + "."

                            OutputFormAppendText -Text $String -Color "Red"
                            
                            $LeaveBalancesCopy.RemoveAt($Index)
                        }
                        
                        else
                        {
                            $Index++
                        }
                    }
                }
                
                $Date = $Date.AddDays(1)
            }

            $ProjectedLeaveIndex = $ProjectedLeaveEndIndex
        }

        $LeaveBalancesCopy[0].Balance += (GetAnnualLeaveAccrualHours -PayPeriod $EndOfPayPeriod)
        $LeaveBalancesCopy[1].Balance += GetSickLeaveAccrualHours

        if($EndOfPayPeriod -eq $LeaveYearEnd)
        {
            if($LeaveBalancesCopy[0].Balance -gt $Script:LeaveCeiling)
            {
                $String  = "`n`n" + ($LeaveBalancesCopy[0].Balance - $Script:LeaveCeiling) + " hours of Annual Leave will be forfeited on " + $LeaveYearEnd.ToString("MM/dd/yyyy") + " which is the Leave Year End. "
                $String += "The Leave Year Schedule Deadline for the " + (GetBeginningOfPayPeriodForDate -Date $LeaveYearEnd).Year + " Leave Year is " + (GetLeaveYearScheduleDeadline -Date $LeaveYearEnd).ToString("MM/dd/yyyy") + " in order for the leave to be eligible to be restored in certain circumstances."
                
                OutputFormAppendText -Text $String -Color "Red"

                $LeaveBalancesCopy[0].Balance = $Script:LeaveCeiling

                if($Script:EmployeeType -eq "Part-Time")
                {
                    $ProjectionAnnualDecimal = 0.0
                    $ProjectionSickDecimal   = 0.0
                }
            }
            
            $LeaveYearEnd = GetLeaveYearEndForDate -Date $LeaveYearEnd.AddDays(1) #Get the next leave year end and store it. 
        }

        if($Script:DisplayHighsAndLows -eq $True)
        {
            if($LeaveBalancesCopy[0].Balance -gt $AnnualHigh)
            {
                $AnnualHigh = $LeaveBalancesCopy[0].Balance
            }

            if($LeaveBalancesCopy[1].Balance -gt $SickHigh)
            {
                $SickHigh = $LeaveBalancesCopy[1].Balance
            }
        }

        if($Script:DisplayAfterEachPP -eq $True -and
           $EndOfPayPeriod -lt $EndDate)
        {
            #Add descriptive date
            OutputFormAppendText -Text ("`n`nBalances as of Pay Period Ending " + $EndOfPayPeriod.ToString("MM/dd/yyyy") + ":`n")
            
            #Add Annual with color
            $AnnualString = $LeaveBalancesCopy[0].Name + ":`t" + [Math]::Floor($LeaveBalancesCopy[0].Balance)

            if($LeaveBalancesCopy[0].Balance -lt 0)
            {
                OutputFormAppendText -Text $AnnualString -Color "Red"
            }

            elseif($LeaveBalancesCopy[0].Balance -gt $Script:LeaveCeiling)
            {
                $AnnualString += "`tBalance is greater than your Annual Leave ceiling."
        
                OutputFormAppendText -Text $AnnualString -Color "Blue"
            }

            else
            {
                OutputFormAppendText -Text $AnnualString
            }

            #Add Sick with color
            $SickString = "`n" + $LeaveBalancesCopy[1].Name + ":`t" + [Math]::Floor($LeaveBalancesCopy[1].Balance)

            if($LeaveBalancesCopy[1].Balance -lt 0)
            {
                OutputFormAppendText -Text $SickString -Color "Red"
            }

            else
            {
                OutputFormAppendText -Text $SickString
            }
    
            #If projecting to date, append the rest of the balances and their expiration if they expire.
            if($Script:ProjectOrGoal -eq "Project")
            {
                for($Index = 2; $Index -lt $LeaveBalancesCopy.Count; $Index++)
                {
                    if($LeaveBalancesCopy[$Index].Expires -eq $False -or
                      ($LeaveBalancesCopy[$Index].Expires -eq $True -and
                       $LeaveBalancesCopy[$Index].ExpiresOn -ge $EndOfPayPeriod))
                    {
                        $String = "`n" + $LeaveBalancesCopy[$Index].Name + ":`t" + $LeaveBalancesCopy[$Index].Balance

                        if($LeaveBalancesCopy[$Index].Expires -eq $True)
                        {
                            $String += "`tExpires: " + $LeaveBalancesCopy[$Index].ExpiresOn.ToString("MM/dd/yyyy")
                        }

                        if($LeaveBalancesCopy[$Index].Balance -lt 0)
                        {
                            OutputFormAppendText -Text $String -Color "Red"
                        }

                        else
                        {
                            OutputFormAppendText -Text $String
                        }
                    }
                }
            }
        }

        if($Script:ProjectOrGoal -eq "Goal" -and
           $LeaveBalancesCopy[0].Balance -ge $Script:AnnualGoal -and
           $LeaveBalancesCopy[1].Balance -ge $Script:SickGoal)
        {
            $GoalsMet = $True
        }

        $EndOfPayPeriod = $EndOfPayPeriod.AddDays(14)
    }

    #Once we exit the loop, set the EndOfPayPeriod back to what it should be.
    $EndOfPayPeriod = $EndOfPayPeriod.AddDays(-14)

    if($Script:EmployeeType -eq "Part-Time") #Handle decimal leave amounts if part-time.
    {
        $LeaveBalancesCopy[0].Balance += $ProjectionAnnualDecimal
        $ProjectionAnnualDecimal       = $LeaveBalancesCopy[0].Balance - [Math]::Floor($LeaveBalancesCopy[0].Balance)
        $LeaveBalancesCopy[0].Balance  = [Math]::Floor($LeaveBalancesCopy[0].Balance)

        $LeaveBalancesCopy[1].Balance += $ProjectionSickDecimal
        $ProjectionSickDecimal         = $LeaveBalancesCopy[1].Balance - [Math]::Floor($LeaveBalancesCopy[1].Balance)
        $LeaveBalancesCopy[1].Balance  = [Math]::Floor($LeaveBalancesCopy[1].Balance)
    }

    if($Script:DisplayHighsAndLows -eq $True)
    {
        $String = "`n`nAnnual Leave High:`t" + [Math]::Floor($AnnualHigh) + "`nSick Leave High:`t" + [Math]::Floor($SickHigh) + "`nAnnual Leave Low:`t" + [Math]::Floor($AnnualLow) + "`nSick Leave Low:`t" + [Math]::Floor($SickLow)
        
        OutputFormAppendText -Text $String
    }

    if($Script:ProjectOrGoal -eq "Goal")
    {
        if($GoalsMet -eq $True)
        {
            $String = "`n`nAnnual Leave and Sick Leave Goals Achieved After Pay Period Ending " + $EndOfPayPeriod.ToString("MM/dd/yyyy") + "."

            OutputFormAppendText -Text $String
        }

        else
        {
            $String = "`n`nGoals not met by " + $EndOfPayPeriod.ToString("MM/dd/yyyy") + "."

            OutputFormAppendText -Text $String
        }

        $String = "`n`nGoals:`nAnnual:`t" + $Script:AnnualGoal + "`nSick:`t" + $Script:SickGoal

        OutputFormAppendText -Text $String
    }

    OutputFormAppendText -Text "`n"

    #Add the section header
    OutputFormAppendText -Text ("`nEnding Balances After Pay Period Ending " + $EndOfPayPeriod.ToString("MM/dd/yyyy") + ":`n") -Alignment "Center" -FontSize ($Script:OutputForm.Controls["OutputRichTextBox"].Font.Size + 2) -FontStyle "Bold"

    #Add Annual with color
    $AnnualString = $LeaveBalancesCopy[0].Name + ":`t" + $LeaveBalancesCopy[0].Balance

    if($LeaveBalancesCopy[0].Balance -lt 0)
    {
        OutputFormAppendText -Text $AnnualString -Color "Red"
    }

    elseif($LeaveBalancesCopy[0].Balance -gt $Script:LeaveCeiling)
    {
        $AnnualString += "`tBalance is greater than your Annual Leave ceiling."
        
        OutputFormAppendText -Text $AnnualString -Color "Blue"
    }

    else
    {
        OutputFormAppendText -Text $AnnualString
    }

    #Add Sick with color
    $SickString = "`n" + $LeaveBalancesCopy[1].Name + ":`t" + $LeaveBalancesCopy[1].Balance

    if($LeaveBalancesCopy[1].Balance -lt 0)
    {
        OutputFormAppendText -Text $SickString -Color "Red"
    }

    else
    {
        OutputFormAppendText -Text $SickString
    }
    
    #If projecting to date, append the rest of the balances and their expiration if they expire.
    if($Script:ProjectOrGoal -eq "Project")
    {
        for($Index = 2; $Index -lt $LeaveBalancesCopy.Count; $Index++)
        {
            if($LeaveBalancesCopy[$Index].Expires -eq $False -or
              ($LeaveBalancesCopy[$Index].Expires -eq $True -and
               $LeaveBalancesCopy[$Index].ExpiresOn -ge $EndOfPayPeriod))
            {
                $String = "`n" + $LeaveBalancesCopy[$Index].Name + ":`t" + $LeaveBalancesCopy[$Index].Balance

                if($LeaveBalancesCopy[$Index].Expires -eq $True)
                {
                    $String += "`tExpires: " + $LeaveBalancesCopy[$Index].ExpiresOn.ToString("MM/dd/yyyy")
                }

                if($LeaveBalancesCopy[$Index].Balance -lt 0)
                {
                    OutputFormAppendText -Text $String -Color "Red"
                }

                else
                {
                    OutputFormAppendText -Text $String
                }
            }
        }
    }

    #Get the Leave Year End again so we're working with the correct year.
    $LeaveYearEnd = GetLeaveYearEndForDate -Date $Script:CurrentDate

    if($EndOfPayPeriod -lt $LeaveYearEnd)
    {
        foreach($Leave in $Script:ProjectedLeave)
        {
            if($Leave.LeaveBank -eq "Annual" -and
               $Leave.StartDate -gt $EndOfPayPeriod -and
               $Leave.StartDate -le $LeaveYearEnd)
            {
                $TotalHours = 0

                foreach($Day in $Leave.HoursHashTable.Keys)
                {
                    $TotalHours += $Leave.HoursHashTable[$Day]
                }

                $LeaveBalancesCopy[0].Balance -= $TotalHours
            }
        }
        
        while($EndOfPayPeriod -lt $LeaveYearEnd)
        {
            $LeaveBalancesCopy[0].Balance += (GetAnnualLeaveAccrualHours -PayPeriod $EndOfPayPeriod)
            
            $EndOfPayPeriod = $EndOfPayPeriod.AddDays(14)
        }

        if($LeaveBalancesCopy[0].Balance -gt $Script:LeaveCeiling)
        {
            $String  = "`n`nYour Annual Leave Balance is expected to be " + $LeaveBalancesCopy[0].Balance + " hours at the Leave Year End on " + $LeaveYearEnd.ToString("MM/dd/yyyy") + ". "
            $String += "This is " + ($LeaveBalancesCopy[0].Balance - $Script:LeaveCeiling) + " hours over your leave ceiling of " + $Script:LeaveCeiling + " hours. "
            $String += "If you schedule no additional Annual Leave before the Leave Year End, you will forfeit these hours."
            $String += "`n`nIf possible, you should submit the leave requests on or before " + (GetLeaveYearScheduleDeadline -Date $LeaveYearEnd).ToString("MM/dd/yyyy") + " so "
            $String += "these hours are eligible to be restored if you are unable to take your scheduled Annual Leave for a few specific reasons which causes you to forfeit these hours."

            OutputFormAppendText -Text $String -Color "Blue"
        }
    }
}

function PopulateProjectedLeaveDays
{
    $SelectedLeaveDetails = $Script:ProjectedLeave[$Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"].SelectedIndex]
    
    $Script:EditProjectedLeaveForm.Controls["LeaveDatesFlowLayoutPanel"].Visible = $False

    #For some reason it's super silly and will sometimes fail to dispose of these, so here, do it manually.
    foreach($Control in $Script:EditProjectedLeaveForm.Controls["LeaveDatesFlowLayoutPanel"].Controls)
    {
        foreach($SubControl in $Control.Controls)
        {
            $SubControl.Dispose()
        }

        $Control.Dispose()
    }
    
    $Script:EditProjectedLeaveForm.Controls["LeaveDatesFlowLayoutPanel"].Controls.Clear()
    
    $Date = $Script:EditProjectedLeaveForm.Controls["LeaveStartDateTimePicker"].Value
    $EndDate = $Script:EditProjectedLeaveForm.Controls["LeaveEndDateTimePicker"].Value
    
    while($Date -le $EndDate)
    {
        $DateString = $Date.ToString("MM/dd/yyyy")
        
        $NewPanel = New-Object System.Windows.Forms.Panel
        $NewPanel.Name = "DatePanel"
        $NewPanel.Height = 20
        $NewPanel.Width = 390

        $NewNumericUpDown = New-Object System.Windows.Forms.NumericUpDown
        $NewNumericUpDown.Name = "HoursNumericUpDown"
        $NewNumericUpDown.Height = 20
        $NewNumericUpDown.Maximum = 24
        $NewNumericUpDown.Tag = $DateString
        $NewNumericUpDown.Width = 40

        $DayOfWeekLabel = New-Object System.Windows.Forms.Label
        $DayOfWeekLabel.Left = 45
        $DayOfWeekLabel.Text = $Date.DayOfWeek.ToString()
        $DayOfWeekLabel.Width = 70

        $DateLabel = New-Object System.Windows.Forms.Label
        $DateLabel.AutoSize = $True
        $DateLabel.Left = 113
        $DateLabel.Text = $DateString
        
        if($SelectedLeaveDetails.HoursHashTable.ContainsKey($DateString) -eq $True)
        {
            $NewNumericUpDown.Value = $SelectedLeaveDetails.HoursHashTable[$DateString]
        }

        else
        {
            $NewNumericUpDown.Value = GetHoursForWorkDay -Day $DateString
        }

        if($Script:HolidaysHashTable.ContainsKey($DateString) -eq $True)
        {
            $DateLabel.Text = $DateString + "   -   " + $Script:HolidaysHashTable[$DateString]
        }

        elseif($InaugurationHoliday -eq $True -and
               $Script:InaugurationDayHashTable.ContainsKey($DateString) -eq $True)
        {
            $DateLabel.Text = $DateString + "   -   " + $Script:InaugurationDayHashTable[$DateString]
        }

        $NewPanel.Controls.AddRange(($NewNumericUpDown, $DayOfWeekLabel, $DateLabel))
        $Script:EditProjectedLeaveForm.Controls["LeaveDatesFlowLayoutPanel"].Controls.Add($NewPanel)

        $NewNumericUpDown.Add_TextChanged({HourNumericUpDownChanged})
        
        $Date = $Date.AddDays(1)
    }

    $Script:EditProjectedLeaveForm.Controls["LeaveDatesFlowLayoutPanel"].Visible = $True
}

function PopulateProjectedLeaveListBox
{
    $ProjectedLeaveListBox = $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"]

    $SelectedItem = $Script:ProjectedLeave[$ProjectedLeaveListBox.SelectedIndex]

    $ProjectedLeaveListBox.BeginUpdate()

    $ProjectedLeaveListBox.Items.Clear()
    
    foreach($LeaveItem in $Script:ProjectedLeave) #Todo might adjust the `t tabbing in here for spacing.
    {
        $NameString = $LeaveItem.LeaveBank
        $DateString = ""
        $TotalHours = 0

        if($NameString.ToLower().Contains("leave") -eq $False)
        {
            $NameString += " Leave"
        }

        if($LeaveItem.StartDate -eq $LeaveItem.EndDate)
        {
            $DateString = "on " + $LeaveItem.StartDate.ToString("MM/dd/yyyy")
        }

        else
        {
            $DateString = "from " + $LeaveItem.StartDate.ToString("MM/dd/yyyy") + " to " + $LeaveItem.EndDate.ToString("MM/dd/yyyy")
        }

        foreach($Day in $LeaveItem.HoursHashTable.Keys)
        {
            $TotalHours += $LeaveItem.HoursHashTable[$Day]
        }
        
        $String = "$TotalHours hours of $NameString " + $DateString

        $ProjectedLeaveListBox.Items.Add($String) | Out-Null
    }

    $ProjectedLeaveListBox.EndUpdate()

    $ProjectedLeaveListbox.SelectedIndex = $Script:ProjectedLeave.IndexOf($SelectedItem)
}

function UpdateExistingBalancesAndProjectedLeaveAtLaunch
{
    if((GetBeginningOfPayPeriodForDate -Date $Script:CurrentDate) -gt (GetBeginningOfPayPeriodForDate -Date $Script:LastLaunchedDate))
    {
        $EndOfPayPeriod = GetEndingOfPayPeriodForDate -Date $Script:LastLaunchedDate
        $LeaveYearEnd   = GetLeaveYearEndForDate -Date $EndOfPayPeriod
    
        while($EndOfPayPeriod -lt $Script:BeginningOfPayPeriod)
        {
            while($Script:ProjectedLeave.Count -gt 0 -and
                  $Script:ProjectedLeave[0].StartDate -le $EndOfPayPeriod)
            {
                $TotalHours = 0

                foreach($Day in $Script:ProjectedLeave[0].HoursHashTable.Keys)
                {
                    $TotalHours += $Script:ProjectedLeave[0].HoursHashTable[$Day]
                }

                $Found = $False
                $Index = 0

                while($Index -lt $Script:LeaveBalances.Count -and
                      $Script:LeaveBalances[$Index].Name -ne $Script:ProjectedLeave[0].LeaveBank)
                {
                    $Index++
                }

                if($Index -lt $Script:LeaveBalances.Count -and
                   $Script:LeaveBalances[$Index].Name -eq $Script:ProjectedLeave[0].LeaveBank)
                {
                    $Found = $True
                }

                if($Found -eq $True)
                {
                    $Script:LeaveBalances[$Index].Balance -= $TotalHours

                    while($Script:LeaveBalances[$Index].Balance -le 0 -and
                         ($Index + 1) -lt $Script:LeaveBalances.Count -and
                          $Script:LeaveBalances[$Index + 1].Name -eq $Script:LeaveBalances[$Index].Name)
                    {
                        $Script:LeaveBalances[$Index + 1].Balance += $Script:LeaveBalances[$Index].Balance #Take that negative balance off of the next entry with the same name.

                        $Script:LeaveBalances.RemoveAt($Index)
                    }
                }

                $Script:ProjectedLeave.RemoveAt(0)
            }
            
            $Script:LeaveBalances[0].Balance += (GetAnnualLeaveAccrualHours -PayPeriod $EndOfPayPeriod)
            $Script:LeaveBalances[1].Balance += GetSickLeaveAccrualHours

            if($EndOfPayPeriod -eq $LeaveYearEnd)
            {
                if($Script:LeaveBalances[0].Balance -gt $Script:LeaveCeiling)
                {
                    $Script:LeaveBalances[0].Balance = $Script:LeaveCeiling

                    if($Script:EmployeeType -eq "Part-Time")
                    {
                        $Script:AnnualDecimal = 0.0
                        $Script:SickDecimal   = 0.0
                    }
                }
                
                $LeaveYearEnd = GetLeaveYearEndForDate -Date $LeaveYearEnd.AddDays(1) #Get the next leave year end and store it.
            }

            $EndOfPayPeriod = $EndOfPayPeriod.AddDays(14)
        }

        if($Script:EmployeeType -eq "Part-Time") #Handle decimal leave amounts if part-time.
        {
            $Script:LeaveBalances[0].Balance += $Script:AnnualDecimal
            $Script:AnnualDecimal             = $Script:LeaveBalances[0].Balance - [Math]::Floor($Script:LeaveBalances[0].Balance)
            $Script:LeaveBalances[0].Balance  = [Math]::Floor($Script:LeaveBalances[0].Balance)

            $Script:LeaveBalances[1].Balance += $Script:SickDecimal
            $Script:SickDecimal               = $Script:LeaveBalances[1].Balance - [Math]::Floor($Script:LeaveBalances[1].Balance)
            $Script:LeaveBalances[1].Balance  = [Math]::Floor($Script:LeaveBalances[1].Balance)
        }

        if($Script:LeaveBalances[0].Balance -lt 0) #Ensure Annual Leave isn't negative.
        {
            $Script:LeaveBalances[0].Balance = 0

            if($Script:EmployeeType -eq "Part-Time")
            {
                $Script:AnnualDecimal = 0.0
            }
        }

        if($Script:LeaveBalances[1].Balance -lt 0) #Ensure Sick Leave isn't negative.
        {
            $Script:LeaveBalances[1].Balance = 0

            if($Script:EmployeeType -eq "Part-Time")
            {
                $Script:SickDecimal = 0.0
            }
        }

        $Index = 2 #Skip annual and sick

        while($Index -lt $Script:LeaveBalances.Count)
        {
            if(($Script:LeaveBalances[$Index].Expires -eq $True -and
                $Script:LeaveBalances[$Index].ExpiresOn -lt $Script:CurrentDate) -or
                $Script:LeaveBalances[$Index].Balance -le 0)
            {
                $Script:LeaveBalances.RemoveAt($Index)
            }

            else
            {
                $Index++
            }
        }
    }

    $Script:LastLaunchedDate = $Script:CurrentDate
}

function UpdateLengthOfServiceStrings
{
    $YearDifference  = $Script:CurrentDate.Year - $Script:SCDLeaveDate.Year
    $MonthDifference = $Script:CurrentDate.Month - $Script:SCDLeaveDate.Month
    $DayDifference   = $Script:CurrentDate.Day - $Script:SCDLeaveDate.Day

    if($DayDifference -lt 0)
    {
        $MonthDifference--

        $Month = $CurrentDate.Month - 1

        if($Month -eq 0)
        {
            $DayDifference = ($Script:CurrentDate - (Get-Date -Year ($Script:CurrentDate.Year - 1) -Month 12 -Day $Script:SCDLeaveDate.Day).Date).Days
        }

        else
        {
            $DayDifference = $Script:CurrentDate.DayOfYear - (Get-Date -Month ($Script:CurrentDate.Month - 1) -Day $Script:SCDLeaveDate.Day).DayOfYear
        }
    }

    if($MonthDifference -lt 0)
    {
        $YearDifference--

        $MonthDifference += 12
    }

    $Script:LengthOfServiceString = "Length of Service: " + $YearDifference + " " + (NumberGetsLetters -String "Year" -Number $YearDifference) + ", " + $MonthDifference + " " + (NumberGetsLetters -String "Month" -Number $MonthDifference) + ", " + $DayDifference + " " + (NumberGetsLetters -String "Day" -Number $DayDifference)

    if($Script:CurrentDate -lt $Script:FifteenYearMark)
    {
        $MilestoneDate = $Script:FifteenYearMark
        $AccrualString = "Greater Than 15 Years"

        if($Script:CurrentDate -lt $Script:ThreeYearMark)
        {
            $MilestoneDate = $Script:ThreeYearMark
            $AccrualString = "3 to 15 Years"
        }

        $Script:DateOfMileStoneString = "Accrue Rate for " + $AccrualString + " on: " + $MilestoneDate.ToString("MM/dd/yyyy")
        
        $YearDifference  = $MilestoneDate.Year - $Script:CurrentDate.Year
        $MonthDifference = $MilestoneDate.Month - $Script:CurrentDate.Month
        $DayDifference   = $MilestoneDate.Day - $Script:CurrentDate.Day

        if($DayDifference -lt 0)
        {
            $MonthDifference--

            $Month = $MilestoneDate.Month - 1

            if($Month -eq 0)
            {
                $DayDifference = ($MilestoneDate - (Get-Date -Year ($MilestoneDate.Year - 1) -Month 12).Date).Days
            }

            else
            {
                $DayDifference = $MilestoneDate.DayOfYear - (Get-Date -Year ($MilestoneDate.Year) -Month ($MilestoneDate.Month - 1)).DayOfYear
            }
        }

        if($MonthDifference -lt 0)
        {
            $YearDifference--

            $MonthDifference += 12
        }

        $Script:TimeUntilMilestone = "Accrue Rate for " + $AccrualString + " in: " + $YearDifference + " " + (NumberGetsLetters -String "Year" -Number $YearDifference) + ", " + $MonthDifference + " " + (NumberGetsLetters -String "Month" -Number $MonthDifference) + ", " + $DayDifference + " " + (NumberGetsLetters -String "Day" -Number $DayDifference)
    }

    else
    {
        $Script:TimeUntilMilestone = $Null
        $Script:DateOfMileStoneString = $Null
    }

    if($Script:DisplayLengthOrTimeUntilMilestone -eq "DateOfMilestone")
    {
        $Script:MainForm.Controls["SettingsPanel"].Controls["LengthOfServiceTextBox"].Text = $Script:DateOfMileStoneString
    }

    elseif($Script:DisplayLengthOrTimeUntilMilestone -eq "TimeUntilMilestone")
    {
        $Script:MainForm.Controls["SettingsPanel"].Controls["LengthOfServiceTextBox"].Text = $Script:TimeUntilMilestone
    }

    else
    {
        $Script:MainForm.Controls["SettingsPanel"].Controls["LengthOfServiceTextBox"].Text = $Script:LengthOfServiceString
    }
}

function UpdateProjectedLeaveDatesIfLeaveBalanceExpiresOrChangesOrDeleted
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)][System.String] $LeaveBankName
    )
    
    $ProjectedLeaveUpdated = $False
    $NoExpireLeave = $False
    $LatestDate    = $Null

    foreach($LeaveBalance in $Script:LeaveBalances)
    {
        if($LeaveBalance.Name -eq $LeaveBankName -and
            $NoExpireLeave -eq $False -and
            $LatestDate -lt $LeaveBalance.ExpiresOn)
        {
            if($LeaveBalance.Expires -eq $False)
            {
                $NoExpireLeave = $True
            }

            else
            {
                $LatestDate = $LeaveBalance.ExpiresOn
            }
        }
    }
    
    if($NoExpireLeave -eq $False)
    {
        foreach($Leave in $Script:ProjectedLeave)
        {
            if($LeaveBankName -eq $Leave.LeaveBank)
            {
                $Changed = $False
                
                if($Leave.StartDate -gt $LatestDate)
                {
                    $Leave.StartDate = $LatestDate
                    
                    $Changed = $True
                }

                if($Leave.EndDate -gt $LatestDate)
                {
                    $Leave.EndDate = $LatestDate
                    
                    $Changed = $True
                }

                if($Changed -eq $True)
                {
                    foreach($Date in $($Leave.HoursHashTable.Keys)) #Can't update a hashtable while you're iterating through it, so have to do some magic to make it actually work. That's why the extra paranthesis.
                    {
                        if(($Date | Get-Date) -gt $LatestDate)
                        {
                            $Leave.HoursHashTable.Remove($Date)
                        }
                    }

                    $MovingDate = $Leave.StartDate

                    while($MovingDate -le $Leave.EndDate)
                    {
                        $DateString = $MovingDate.ToString("MM/dd/yyyy")
                            
                        if($Leave.HoursHashTable.ContainsKey($DateString) -eq $False)
                        {
                            $Leave.HoursHashTable.$DateString = GetHoursForWorkDay -Day $DateString
                        }
                            
                        $MovingDate = $MovingDate.AddDays(1)
                    }

                    $ProjectedLeaveUpdated = $True
                }
            }
        }
    }

    return $ProjectedLeaveUpdated
}

function UpdateTypeOfEmployeeTextBoxString
{
    $Script:MainForm.Controls["SettingsPanel"].Controls["EmployeeTypeTextBox"].Text = $Script:EmployeeType + " (" + $Script:WorkHoursPerPayPeriod + " Hours per Pay Period / " + $Script:LeaveCeiling + " Hour Ceiling)"
}

#endregion Functions

#region Event Handlers

#region Main Form

function MainFormClosing
{
    if($Script:EmployeeType -ne "Part-Time")
    {
        $Script:AnnualDecimal = 0.0
        $Script:SickDecimal   = 0.0
    }
    
    $SettingsArray = @(
        $Script:LastLaunchedDate
        $Script:SCDLeaveDate
        $Script:DisplayLengthOrTimeUntilMilestone
        $Script:EmployeeType
        $Script:LeaveCeiling
        $Script:InaugurationHoliday
        $Script:WorkSchedule
        $Script:ProjectOrGoal
        $Script:ProjectToDate
        $Script:AnnualGoal
        $Script:SickGoal
        $Script:AnnualDecimal
        $Script:SickDecimal
        $Script:DisplayAfterEachLeave
        $Script:DisplayAfterEachPP
        $Script:DisplayHighsAndLows
        $Script:LeaveBalances
        $Script:ProjectedLeave
    )

    try
    {
        if((Test-Path -Path $Script:ConfigFile) -eq $False)
        {
            New-Item -Path (Split-Path -Path $Script:ConfigFile) -ItemType "Directory" -ErrorAction "SilentlyContinue" | Out-Null
            New-Item -Path $Script:ConfigFile -ItemType "File" -ErrorAction "SilentlyContinue" | Out-Null
        }

        if((Test-Path -Path $Script:ConfigFile) -eq $True)
        {
            $SettingsArray | Export-Clixml -Path $Script:ConfigFile
        }
    }

    catch
    {
        Write-Error "Exception happened while attempting to create/save the configuration file."
    }
}

function MainFormSCDLeaveDateTimePickerValueChanged
{
    $Script:SCDLeaveDate = $Script:MainForm.Controls["SettingsPanel"].Controls["SCDLeaveDateDateTimePicker"].Value.Date
    
    GetAccrualRateDateChange

    UpdateLengthOfServiceStrings
}

function MainFormUpdateInfoButtonClick
{
    BuildEmployeeInfoForm

    $Script:EmployeeInfoForm.ShowDialog()
}

function MainFormLengthOfServiceTextBoxClick
{
    if($Script:EmployeeType -eq "SES" -or
       $Script:DisplayLengthOrTimeUntilMilestone -eq "TimeUntilMilestone" -or
       $Script:CurrentDate -ge $Script:FifteenYearMark)
    {
        $Script:DisplayLengthOrTimeUntilMilestone = "LengthOfService"
        
        $Script:MainForm.Controls["SettingsPanel"].Controls["LengthOfServiceTextBox"].Text = $Script:LengthOfServiceString
    }

    elseif($Script:DisplayLengthOrTimeUntilMilestone -eq "DateOfMilestone")
    {
        $Script:DisplayLengthOrTimeUntilMilestone = "TimeUntilMilestone"
        
        $Script:MainForm.Controls["SettingsPanel"].Controls["LengthOfServiceTextBox"].Text = $Script:TimeUntilMilestone
    }

    else
    {
        $Script:DisplayLengthOrTimeUntilMilestone = "DateOfMilestone"

        $Script:MainForm.Controls["SettingsPanel"].Controls["LengthOfServiceTextBox"].Text = $Script:DateOfMileStoneString
    }
}

function MainFormProjectBalanceRadioButtonClick
{
    $Script:ProjectOrGoal = "Project"
    
    $Script:MainForm.Controls["ReportPanel"].Controls["ProjectToDateDateTimePicker"].Enabled = $True

    $Script:MainForm.Controls["ReportPanel"].Controls["AnnualGoalNumericUpDown"].Enabled = $False
    $Script:MainForm.Controls["ReportPanel"].Controls["SickGoalNumericUpDown"].Enabled   = $False
}

function MainFormReachGoalRadioButtonClick
{
    $Script:ProjectOrGoal = "Goal"
    
    $Script:MainForm.Controls["ReportPanel"].Controls["ProjectToDateDateTimePicker"].Enabled = $False

    $Script:MainForm.Controls["ReportPanel"].Controls["AnnualGoalNumericUpDown"].Enabled = $True
    $Script:MainForm.Controls["ReportPanel"].Controls["SickGoalNumericUpDown"].Enabled   = $True
}

function MainFormProjectButtonClick
{
    BuildOutputForm

    $Script:OutputForm.ShowDialog()
}

function MainFormLeaveBalanceListBoxIndexChanged
{
    $SelectedLeaveBalance = $Script:LeaveBalances[$Script:MainForm.Controls["LeavePanel"].Controls["BalanceListBox"].SelectedIndex]
    
    if($SelectedLeaveBalance.Name -eq "Annual" -or
       $SelectedLeaveBalance.Name -eq "Sick")
    {
        $Script:MainForm.Controls["LeavePanel"].Controls["BalanceDeleteButton"].Enabled = $False
    }

    else
    {
        $Script:MainForm.Controls["LeavePanel"].Controls["BalanceDeleteButton"].Enabled = $True
    }
}

function MainFormLeaveBalanceListBoxDoubleClick
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)][System.EventArgs] $EventArguments
    )
    
    $DoubleClickedIndex = $Script:MainForm.Controls["LeavePanel"].Controls["BalanceListBox"].IndexFromPoint($EventArguments.Location)

    if($DoubleClickedIndex -ne -1)
    {
        MainFormBalanceEditButtonClick
    }
}

function MainFormBalanceAddButtonClick
{
    $NewLeaveBalance = [PSCustomObject] @{
    Name      = [String]"New Leave"
    Balance   = [Int32]0
    Expires   = [Boolean]$False
    ExpiresOn = [DateTime]$Script:CurrentDate
    Static    = [Boolean]$False
    }

    $Script:LeaveBalances.Add($NewLeaveBalance)

    PopulateLeaveBalanceListBox

    $Script:LeaveBalances[$Script:LeaveBalances.Count - 1].Name = ""

    $Script:MainForm.Controls["LeavePanel"].Controls["BalanceListBox"].SelectedIndex = $Script:LeaveBalances.Count - 1

    MainFormBalanceEditButtonClick
}

function MainFormBalanceEditButtonClick
{
    BuildEditLeaveBalanceForm

    $SelectedLeaveBalance = $Script:LeaveBalances[$Script:MainForm.Controls["LeavePanel"].Controls["BalanceListBox"].SelectedIndex]

    if($SelectedLeaveBalance.Name -eq "")
    {
        $Script:EditLeaveBalanceForm.Controls["EditLeaveOkButton"].Enabled = $False
    }

    $Script:EditLeaveBalanceForm.ShowDialog()
}

function MainFormBalanceDeleteButtonClick
{
    $SelectedIndex = $Script:MainForm.Controls["LeavePanel"].Controls["BalanceListBox"].SelectedIndex

    $SelectedLeaveBalance = $Script:LeaveBalances[$SelectedIndex]

    $SelectedProjectedLeaveDetails = $Script:ProjectedLeave[$Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"].SelectedIndex]

    $PreviousName = $SelectedLeaveBalance.Name

    $Script:LeaveBalances.RemoveAt($SelectedIndex)

    $ChangesMade           = $False
    $LeaveNameCount        = 0
    $ProjectedLeaveDeleted = $False

    foreach($LeaveBalance in $Script:LeaveBalances)
    {
        if($LeaveBalance.Name -eq $PreviousName)
        {
            $LeaveNameCount++
        }
    }

    if($LeaveNameCount -eq 0)
    {
        $Index        = 0
        $DeletedCount = 0

        while($Index -lt $Script:ProjectedLeave.Count)
        {
            if($PreviousName -eq $Script:ProjectedLeave[$Index].LeaveBank)
            {
                $Script:ProjectedLeave.Remove($Script:ProjectedLeave[$Index])

                $ChangesMade = $True

                $ProjectedLeaveDeleted = $True

                $DeletedCount++
            }

            else
            {
                $Index++
            }
        }
    }

    else
    {
        $ChangesMade = UpdateProjectedLeaveDatesIfLeaveBalanceExpiresOrChangesOrDeleted -LeaveBankName $PreviousName
    }

    if($ChangesMade -eq $True)
    {
        if($Script:ProjectedLeave.Count -gt 1) #Only sort if there's more than one item.
        {
            $Script:ProjectedLeave = [System.Collections.Generic.List[PSCustomObject]] ($Script:ProjectedLeave | Sort-Object -Property "StartDate", "EndDate", "LeaveBank")
        }

        PopulateProjectedLeaveListBox

        if($Script:ProjectedLeave.IndexOf($SelectedProjectedLeaveDetails) -ge 0)
        {
            $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"].SelectedIndex = $Script:ProjectedLeave.IndexOf($SelectedProjectedLeaveDetails)
        }

        else
        {
            $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"].SelectedIndex = 0
        }

        if($ProjectedLeaveDeleted -eq $True)
        {
            $LeaveType = $PreviousName
            $EntryPlurality = "entry"

            if($LeaveType.ToLower().Contains("leave") -eq $False)
            {
                $LeaveType = $LeaveType.Trim() + " Leave"
            }

            if($DeletedCount -ne 1)
            {
                $EntryPlurality = "entries"
            }
            
            ShowMessageBox -Text "$DeletedCount $LeaveType $EntryPlurality removed from Projected Leave." -Caption "Projected Leave Removed" -Buttons "OK" -Icon "Information"
        }

        else
        {
            ShowMessageBox -Text "Projected Leave contained entries that happened after the latest possible expiration date. These entries have been updated to not happen after the expiration date. Please verify these are correct." -Caption "Projected Leave Updated" -Buttons "OK" -Icon "Information"
        }
        
    }

    PopulateLeaveBalanceListBox

    if($SelectedIndex -lt $Script:LeaveBalances.Count)
    {
        $Script:MainForm.Controls["LeavePanel"].Controls["BalanceListBox"].SelectedIndex = $SelectedIndex
    }

    else
    {
        $Script:MainForm.Controls["LeavePanel"].Controls["BalanceListBox"].SelectedIndex = $SelectedIndex - 1
    }
}

function MainFormProjectedLeaveListBoxDoubleClick
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)][System.EventArgs] $EventArguments
    )
    
    $DoubleClickedIndex = $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"].IndexFromPoint($EventArguments.Location)

    if($DoubleClickedIndex -ne -1)
    {
        MainFormProjectedEditButtonClick
    }
}

function MainFormProjectedAddButtonClick
{
    $Script:UnsavedProjectedLeave = $True
    
    $NewProjectedLeave = [PSCustomObject] @{
        LeaveBank      = [String]"Annual"
        StartDate      = [DateTime]$Script:CurrentDate
        EndDate        = [DateTime]$Script:CurrentDate
        HoursHashTable = @{
        $Script:CurrentDate.ToString("MM/dd/yyyy") = GetHoursForWorkDay -Day $Script:CurrentDate
        }
    }

    $Script:ProjectedLeave.Add($NewProjectedLeave)

    $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedEditButton"].Enabled   = $True
    $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedDeleteButton"].Enabled = $True

    PopulateProjectedLeaveListBox

    $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"].SelectedIndex = $Script:ProjectedLeave.Count - 1

    MainFormProjectedEditButtonClick
}

function MainFormProjectedEditButtonClick
{
    BuildEditProjectedLeaveForm

    $Script:EditProjectedLeaveForm.ShowDialog()
}

function MainFormProjectedDeleteButtonClick
{
    $SelectedIndex = $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"].SelectedIndex

    $Script:ProjectedLeave.RemoveAt($SelectedIndex)

    PopulateProjectedLeaveListBox

    if($SelectedIndex -lt $Script:ProjectedLeave.Count)
    {
        $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"].SelectedIndex = $SelectedIndex
    }

    else
    {
        $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"].SelectedIndex = $SelectedIndex - 1
    }

    if($Script:ProjectedLeave.Count -eq 0)
    {
        $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedEditButton"].Enabled   = $False
        $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedDeleteButton"].Enabled = $False
    }
}

function MainFormProjectToDateValueChanged
{
    $Script:ProjectToDate = $Script:MainForm.Controls["ReportPanel"].Controls["ProjectToDateDateTimePicker"].Value.Date
}

function MainFormAnnualGoalValueChanged
{
    $Script:AnnualGoal = $Script:MainForm.Controls["ReportPanel"].Controls["AnnualGoalNumericUpDown"].Value
}

function MainFormSickGoalValueChanged
{
    $Script:SickGoal = $Script:MainForm.Controls["ReportPanel"].Controls["SickGoalNumericUpDown"].Value
}

function MainFormEveryLeaveCheckBoxClicked
{
    $Script:DisplayAfterEachLeave = $Script:MainForm.Controls["SettingsPanel"].Controls["DisplayBalanceEveryLeaveCheckBox"].Checked
}

function MainFormEveryPPCheckBoxClicked
{
    $Script:DisplayAfterEachPP = $Script:MainForm.Controls["SettingsPanel"].Controls["DisplayBalanceEveryPayPeriodEnd"].Checked
}

function MainFormDisplayHighsLowsClicked
{
    $Script:DisplayHighsAndLows = $Script:MainForm.Controls["SettingsPanel"].Controls["DisplayLeaveHighsAndLows"].Checked
}

#endregion Main Form

#region Employee Info Form

function EmployeeInfoFormClosing
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)][System.EventArgs] $EventArguments
    )

    $ChangesMade = $False

    if($Script:EmployeeType -eq "Full-Time" -and
       $Script:EmployeeInfoForm.Controls["EmployeeTypePanel"].Controls["FullTimeRadioButton"].Checked -eq $False)
    {
        $ChangesMade = $True
    }

    elseif($Script:EmployeeType -eq "Part-Time" -and
           $Script:EmployeeInfoForm.Controls["EmployeeTypePanel"].Controls["PartTimeRadioButton"].Checked -eq $False)
    {
        $ChangesMade = $True
    }

    elseif($Script:EmployeeType -eq "SES" -and
           $Script:EmployeeInfoForm.Controls["EmployeeTypePanel"].Controls["SESRadioButton"].Checked -eq $False)
    {
        $ChangesMade = $True
    }

    elseif($Script:LeaveCeiling -eq 240 -and
           $Script:EmployeeInfoForm.Controls["LeaveCeilingPanel"].Controls["CONUSRadioButton"].Checked -eq $False)
    {
        $ChangesMade = $True
    }

    elseif($Script:LeaveCeiling -eq 360 -and
           $Script:EmployeeInfoForm.Controls["LeaveCeilingPanel"].Controls["OCONUSRadioButton"].Checked -eq $False)
    {
        $ChangesMade = $True
    }

    elseif($Script:LeaveCeiling -eq 720 -and
           $Script:EmployeeInfoForm.Controls["LeaveCeilingPanel"].Controls["SESCeilingRadioButton"].Checked -eq $False)
    {
        $ChangesMade = $True
    }

    elseif($Script:InaugurationHoliday -ne $Script:EmployeeInfoForm.Controls["InaugurationDayHolidayCheckBox"].Checked)
    {
        $ChangesMade = $True
    }

    elseif($Script:WorkSchedule["PayPeriodDay1"] -ne $Script:EmployeeInfoForm.Controls["Day1NumericUpDown"].Value)
    {
        $ChangesMade = $True
    }

    elseif($Script:WorkSchedule["PayPeriodDay2"] -ne $Script:EmployeeInfoForm.Controls["Day2NumericUpDown"].Value)
    {
        $ChangesMade = $True
    }

    elseif($Script:WorkSchedule["PayPeriodDay3"] -ne $Script:EmployeeInfoForm.Controls["Day3NumericUpDown"].Value)
    {
        $ChangesMade = $True
    }

    elseif($Script:WorkSchedule["PayPeriodDay4"] -ne $Script:EmployeeInfoForm.Controls["Day4NumericUpDown"].Value)
    {
        $ChangesMade = $True
    }

    elseif($Script:WorkSchedule["PayPeriodDay5"] -ne $Script:EmployeeInfoForm.Controls["Day5NumericUpDown"].Value)
    {
        $ChangesMade = $True
    }

    elseif($Script:WorkSchedule["PayPeriodDay6"] -ne $Script:EmployeeInfoForm.Controls["Day6NumericUpDown"].Value)
    {
        $ChangesMade = $True
    }

    elseif($Script:WorkSchedule["PayPeriodDay7"] -ne $Script:EmployeeInfoForm.Controls["Day7NumericUpDown"].Value)
    {
        $ChangesMade = $True
    }

    elseif($Script:WorkSchedule["PayPeriodDay8"] -ne $Script:EmployeeInfoForm.Controls["Day8NumericUpDown"].Value)
    {
        $ChangesMade = $True
    }

    elseif($Script:WorkSchedule["PayPeriodDay9"] -ne $Script:EmployeeInfoForm.Controls["Day9NumericUpDown"].Value)
    {
        $ChangesMade = $True
    }

    elseif($Script:WorkSchedule["PayPeriodDay10"] -ne $Script:EmployeeInfoForm.Controls["Day10NumericUpDown"].Value)
    {
        $ChangesMade = $True
    }

    elseif($Script:WorkSchedule["PayPeriodDay11"] -ne $Script:EmployeeInfoForm.Controls["Day11NumericUpDown"].Value)
    {
        $ChangesMade = $True
    }

    elseif($Script:WorkSchedule["PayPeriodDay12"] -ne $Script:EmployeeInfoForm.Controls["Day12NumericUpDown"].Value)
    {
        $ChangesMade = $True
    }

    elseif($Script:WorkSchedule["PayPeriodDay13"] -ne $Script:EmployeeInfoForm.Controls["Day13NumericUpDown"].Value)
    {
        $ChangesMade = $True
    }

    elseif($Script:WorkSchedule["PayPeriodDay14"] -ne $Script:EmployeeInfoForm.Controls["Day14NumericUpDown"].Value)
    {
        $ChangesMade = $True
    }

    $Response = ""

    if($ChangesMade -eq $True)
    {
        $Response = ShowMessageBox -Text "You have made changes to the Employee Information. Would you like to discard these changes?" -Caption "Profile Changed" -Buttons "YesNo" -Icon "Exclamation"
    }

    if($ChangesMade -eq $True -and $Response -eq "No")
    {
        $EventArguments.Cancel = $True
    }
}

function EmployeeInfoFormOkButtonClick
{
    $ScheduledHoursChanged      = $False
    $InaugurationHolidayChanged = $False
    
    if($Script:EmployeeInfoForm.Controls["EmployeeTypePanel"].Controls["FullTimeRadioButton"].Checked -eq $True)
    {
        $Script:EmployeeType = "Full-Time"
    }

    elseif($Script:EmployeeInfoForm.Controls["EmployeeTypePanel"].Controls["PartTimeRadioButton"].Checked -eq $True)
    {
        $Script:EmployeeType = "Part-Time"
    }

    else
    {
        $Script:EmployeeType = "SES"
    }

    if($Script:EmployeeInfoForm.Controls["LeaveCeilingPanel"].Controls["CONUSRadioButton"].Checked -eq $True)
    {
        $Script:LeaveCeiling = 240
    }

    elseif($Script:EmployeeInfoForm.Controls["LeaveCeilingPanel"].Controls["OCONUSRadioButton"].Checked -eq $True)
    {
        $Script:LeaveCeiling = 360
    }

    else
    {
        $Script:LeaveCeiling = 720
    }

    if($Script:InaugurationHoliday -ne $Script:EmployeeInfoForm.Controls["InaugurationDayHolidayCheckBox"].Checked)
    {
        $Script:InaugurationHoliday = $Script:EmployeeInfoForm.Controls["InaugurationDayHolidayCheckBox"].Checked
        
        $InaugurationHolidayChanged = $True
    }
    
    if($Script:WorkSchedule["PayPeriodDay1"] -ne $Script:EmployeeInfoForm.Controls["Day1NumericUpDown"].Value -or
       $Script:WorkSchedule["PayPeriodDay2"] -ne $Script:EmployeeInfoForm.Controls["Day2NumericUpDown"].Value -or
       $Script:WorkSchedule["PayPeriodDay3"] -ne $Script:EmployeeInfoForm.Controls["Day3NumericUpDown"].Value -or
       $Script:WorkSchedule["PayPeriodDay4"] -ne $Script:EmployeeInfoForm.Controls["Day4NumericUpDown"].Value -or
       $Script:WorkSchedule["PayPeriodDay5"] -ne $Script:EmployeeInfoForm.Controls["Day5NumericUpDown"].Value -or
       $Script:WorkSchedule["PayPeriodDay6"] -ne $Script:EmployeeInfoForm.Controls["Day6NumericUpDown"].Value -or
       $Script:WorkSchedule["PayPeriodDay7"] -ne $Script:EmployeeInfoForm.Controls["Day7NumericUpDown"].Value -or
       $Script:WorkSchedule["PayPeriodDay8"] -ne $Script:EmployeeInfoForm.Controls["Day8NumericUpDown"].Value -or
       $Script:WorkSchedule["PayPeriodDay9"] -ne $Script:EmployeeInfoForm.Controls["Day9NumericUpDown"].Value -or
       $Script:WorkSchedule["PayPeriodDay10"] -ne $Script:EmployeeInfoForm.Controls["Day10NumericUpDown"].Value -or
       $Script:WorkSchedule["PayPeriodDay11"] -ne $Script:EmployeeInfoForm.Controls["Day11NumericUpDown"].Value -or
       $Script:WorkSchedule["PayPeriodDay12"] -ne $Script:EmployeeInfoForm.Controls["Day12NumericUpDown"].Value -or
       $Script:WorkSchedule["PayPeriodDay13"] -ne $Script:EmployeeInfoForm.Controls["Day13NumericUpDown"].Value -or
       $Script:WorkSchedule["PayPeriodDay14"] -ne $Script:EmployeeInfoForm.Controls["Day14NumericUpDown"].Value)
    {
        $ScheduledHoursChanged = $True
    }

    if($ScheduledHoursChanged -eq $True -or
       $InaugurationHolidayChanged -eq $True)
    {
        $InaugurationHoursModified = $False
        $WorkHoursModified         = $False
        
        if($InaugurationHolidayChanged -eq $True)
        {
            foreach($Leave in $Script:ProjectedLeave)
            {
                foreach($Date in $($Leave.HoursHashTable.Keys)) #Can't update a hashtable while you're iterating through it, so have to do some magic to make it actually work. That's why the extra paranthesis.
                {
                    if($Script:InaugurationDayHashTable.ContainsKey($Date) -eq $True)
                    {
                        $DayOfPayPeriod = ((($Date | Get-Date).Date - $Script:BeginningOfPayPeriod).Days % 14) + 1

                        $Hours = $Script:WorkSchedule.("PayPeriodDay" + $DayOfPayPeriod)
                        
                        if($Script:InaugurationHoliday -eq $False -and
                           $Leave.HoursHashTable[$Date] -eq 0)
                        {
                            $Leave.HoursHashTable[$Date] = $Hours

                            $InaugurationHoursModified = $True
                        }

                        elseif($Script:InaugurationHoliday -eq $True -and
                               $Leave.HoursHashTable[$Date] -eq $Hours)
                        {
                            $Leave.HoursHashTable[$Date] = 0

                            $InaugurationHoursModified = $True
                        }
                    }
                }
            }
        }

        if($ScheduledHoursChanged -eq $True)
        {
            foreach($Leave in $Script:ProjectedLeave)
            {
                foreach($Date in $($Leave.HoursHashTable.Keys)) #Can't update a hashtable while you're iterating through it, so have to do some magic to make it actually work. That's why the extra paranthesis.
                {
                    $DayOfPayPeriod = ((($Date | Get-Date).Date - $Script:BeginningOfPayPeriod).Days % 14) + 1

                    $WorkScheduleDayString  = "PayPeriodDay" + $DayOfPayPeriod
                    $NumericUpDownDayString = "Day" + $DayOfPayPeriod + "NumericUpDown"

                    if($Script:WorkSchedule[$WorkScheduleDayString] -ne $Script:EmployeeInfoForm.Controls[$NumericUpDownDayString].Value -and
                       $Leave.HoursHashTable[$Date] -eq $Script:WorkSchedule[$WorkScheduleDayString])
                    {
                        if($Script:HolidaysHashTable.ContainsKey($Date) -eq $True -or
                          ($Script:InaugurationHoliday -eq $True -and
                           $Script:InaugurationDayHashTable.ContainsKey($Date) -eq $True))
                        {
                            $Leave.HoursHashTable[$Date] = $Script:EmployeeInfoForm.Controls[$NumericUpDownDayString].Value

                            $WorkHoursModified = $True
                        }

                        else
                        {
                            $Leave.HoursHashTable[$Date] = $Script:EmployeeInfoForm.Controls[$NumericUpDownDayString].Value

                            $WorkHoursModified = $True
                        }
                    }
                }
            }
        }

        if($WorkHoursModified -eq $True -or
           $InaugurationHoursModified -eq $True)
        {
            PopulateProjectedLeaveListBox
        }

        if($InaugurationHoursModified -eq $True -and
           $WorkHoursModified -eq $True)
        {
            ShowMessageBox -Text "All whole day projected leave days have been updated including the Inauguration Day. Please verify your projected leave days are correct." -Caption "Projected Leave Updated" -Buttons "OK" -Icon "Information"
        }

        elseif($InaugurationHoursModified -eq $True)
        {
            ShowMessageBox -Text "One of your projected leave days included the Inauguration Day. It has been updated. Please verify it has the correct number of hours." -Caption "Projected Leave Updated" -Buttons "OK" -Icon "Information"
        }

        elseif($WorkHoursModified -eq $True)
        {
            ShowMessageBox -Text "All whole day projected leave days have been updated. Please verify your projected leave days are correct." -Caption "Projected Leave Updated" -Buttons "OK" -Icon "Information"
        }
    }

    $Script:WorkSchedule["PayPeriodDay1"] = $Script:EmployeeInfoForm.Controls["Day1NumericUpDown"].Value
    $Script:WorkSchedule["PayPeriodDay2"] = $Script:EmployeeInfoForm.Controls["Day2NumericUpDown"].Value
    $Script:WorkSchedule["PayPeriodDay3"] = $Script:EmployeeInfoForm.Controls["Day3NumericUpDown"].Value
    $Script:WorkSchedule["PayPeriodDay4"] = $Script:EmployeeInfoForm.Controls["Day4NumericUpDown"].Value
    $Script:WorkSchedule["PayPeriodDay5"] = $Script:EmployeeInfoForm.Controls["Day5NumericUpDown"].Value
    $Script:WorkSchedule["PayPeriodDay6"] = $Script:EmployeeInfoForm.Controls["Day6NumericUpDown"].Value
    $Script:WorkSchedule["PayPeriodDay7"] = $Script:EmployeeInfoForm.Controls["Day7NumericUpDown"].Value

    $Script:WorkSchedule["PayPeriodDay8"]  = $Script:EmployeeInfoForm.Controls["Day8NumericUpDown"].Value
    $Script:WorkSchedule["PayPeriodDay9"]  = $Script:EmployeeInfoForm.Controls["Day9NumericUpDown"].Value
    $Script:WorkSchedule["PayPeriodDay10"] = $Script:EmployeeInfoForm.Controls["Day10NumericUpDown"].Value
    $Script:WorkSchedule["PayPeriodDay11"] = $Script:EmployeeInfoForm.Controls["Day11NumericUpDown"].Value
    $Script:WorkSchedule["PayPeriodDay12"] = $Script:EmployeeInfoForm.Controls["Day12NumericUpDown"].Value
    $Script:WorkSchedule["PayPeriodDay13"] = $Script:EmployeeInfoForm.Controls["Day13NumericUpDown"].Value
    $Script:WorkSchedule["PayPeriodDay14"] = $Script:EmployeeInfoForm.Controls["Day14NumericUpDown"].Value

    if($ScheduledHoursChanged -eq $True)
    {
        SetWorkHoursPerPayPeriod
    }

    if($Script:EmployeeType -eq "SES" -and
       $Script:DisplayLengthOrTimeUntilMilestone -ne "LengthOfService")
    {
        $Script:DisplayLengthOrTimeUntilMilestone = "LengthOfService"

        $Script:MainForm.Controls["SettingsPanel"].Controls["LengthOfServiceTextBox"].Text = $Script:LengthOfServiceString
    }

    UpdateTypeOfEmployeeTextBoxString

    $Script:EmployeeInfoForm.Close()
}

function EmployeeInfoFormCancelButtonClick
{
    $Script:EmployeeInfoForm.Close()
}

function EmployeeInfoFormFullTimeRadioButtonClick
{
    if($Script:EmployeeInfoForm.Controls["LeaveCeilingPanel"].Controls["SESCeilingRadioButton"].Checked -eq $True)
    {
        $Script:EmployeeInfoForm.Controls["LeaveCeilingPanel"].Controls["CONUSRadioButton"].Checked = $True
    }
}

function EmployeeInfoFormPartTimeRadioButtonClick
{
    if($Script:EmployeeInfoForm.Controls["LeaveCeilingPanel"].Controls["SESCeilingRadioButton"].Checked -eq $True)
    {
        $Script:EmployeeInfoForm.Controls["LeaveCeilingPanel"].Controls["CONUSRadioButton"].Checked = $True
    }
}

function EmployeeInfoFormSesRadioButtonClick
{
    $Script:EmployeeInfoForm.Controls["LeaveCeilingPanel"].Controls["SESCeilingRadioButton"].Checked = $True
}

function EmployeeInfoHoursChanged
{
    $TotalHours = 0
    
    for($Day = 1; $Day -le 14; $Day++)
    {
        $ControlName = "Day" + $Day + "NumericUpDown"

        $TotalHours += $Script:EmployeeInfoForm.Controls[$ControlName].Value
    }
    
    $Script:EmployeeInfoForm.Controls["HoursWorkedLabel"].Text = "Hours Per Pay Period: " + $TotalHours

    if($Script:EmployeeInfoForm.Controls["EmployeeTypePanel"].Controls["PartTimeRadioButton"].Checked -eq $False -and
       $TotalHours -ne 80)
    {
        $Script:EmployeeInfoForm.Controls["UnusualHoursLabel"].Visible = $True
    }

    elseif($Script:EmployeeInfoForm.Controls["EmployeeTypePanel"].Controls["PartTimeRadioButton"].Checked -eq $True -and
          ($TotalHours -lt 32 -or
           $TotalHours -gt 64))
    {
        $Script:EmployeeInfoForm.Controls["UnusualHoursLabel"].Visible = $True
    }
    
    else
    {
        $Script:EmployeeInfoForm.Controls["UnusualHoursLabel"].Visible = $False
    }
}

#endregion Employee Info Form

#region Edit Leave Balance Form

function EditLeaveBalanceFormClosing
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)][System.EventArgs] $EventArguments
    )

    $SelectedLeaveBalance = $Script:LeaveBalances[$Script:MainForm.Controls["LeavePanel"].Controls["BalanceListBox"].SelectedIndex]
    
    $ChangesMade = $False

    if($SelectedLeaveBalance.Name -cne $Script:EditLeaveBalanceForm.Controls["LeaveBalanceNameTextBox"].Text.Trim()) #Have to use the case sensitive operator -cne instead of -ne
    {
        $ChangesMade = $True
    }

    elseif($SelectedLeaveBalance.Balance -ne $Script:EditLeaveBalanceForm.Controls["BalanceNumericUpDown"].Value)
    {
        $ChangesMade = $True
    }

    if($SelectedLeaveBalance.Name -eq "Annual" -or
        $SelectedLeaveBalance.Name -eq "Sick")
    {
        if($SelectedLeaveBalance.Threshold -ne $Script:EditLeaveBalanceForm.Controls["ThresholdNumericUpDown"].Value)
        {
            $ChangesMade = $True
        }
    }

    else
    {
        if($SelectedLeaveBalance.Expires -ne $Script:EditLeaveBalanceForm.Controls["LeaveExpiresCheckBox"].Checked)
        {
            $ChangesMade = $True
        }

        elseif($SelectedLeaveBalance.ExpiresOn -ne $Script:EditLeaveBalanceForm.Controls["LeaveExpiresOnDateTimePicker"].Value.Date)
        {
            $ChangesMade = $True
        }
    }

    $Response = ""

    if($ChangesMade -eq $True)
    {
        $LeaveNameString = $Script:EditLeaveBalanceForm.Controls["LeaveBalanceNameTextBox"].Text

        if($LeaveNameString.ToLower().Contains("leave") -eq $False)
        {
            $LeaveNameString = $LeaveNameString.Trim() + " Leave"
        }
        
        $Response = ShowMessageBox -Text "You have made changes to the $LeaveNameString Balance. Would you like to discard these changes?" -Caption "Leave Changed" -Buttons "YesNo" -Icon "Exclamation"
    }

    if($ChangesMade -eq $False -or $Response -eq "Yes")
    {
        if($SelectedLeaveBalance.Name -eq "")
        {
            MainFormBalanceDeleteButtonClick
        }
    }

    else
    {
        $EventArguments.Cancel = $True
    }
}

function EditLeaveBalanceFormNameOrDateChanged
{
    $Valid = $True
    
    $Text = $Script:EditLeaveBalanceForm.Controls["LeaveBalanceNameTextBox"].Text.Trim().ToLower()
    
    if($Text -eq "annual")
    {
        $Valid = $False

        $Script:EditLeaveBalanceForm.Controls["WarningLabel"].Text = "Annual is a reserved name."
    }

    elseif($Text -eq "sick")
    {
        $Valid = $False

        $Script:EditLeaveBalanceForm.Controls["WarningLabel"].Text = "Sick is a reserved name."
    }

    elseif($Text -eq "")
    {
        $Valid = $False

        $Script:EditLeaveBalanceForm.Controls["WarningLabel"].Text = "Name cannot be blank."
    }

    #Regex matches a set [           ]
    #Exludes this set     ^
    #Capital letters       A-Z
    #Lowercase letters        a-z
    #Digits                      0-9
    #Also a space                   <space>
    elseif($Text -Match "[^A-Za-z0-9 ]" -eq $True)
    {
        $Valid = $False

        $Script:EditLeaveBalanceForm.Controls["WarningLabel"].Text = "Only letters/numbers/spaces permitted."
    }
    
    else
    {
        $SelectedIndex = $Script:MainForm.Controls["LeavePanel"].Controls["BalanceListBox"].SelectedIndex
        $SelectedLeaveBalance = $Script:LeaveBalances[$SelectedIndex]
        $DuplicateName        = $False
        $Index                = 2 #0 and 1 are Annual and Sick, skip them.
        
        while($Index -lt $Script:LeaveBalances.Count -and
              $DuplicateName -eq $False)
        {
            if($Index -ne $Script:MainForm.Controls["LeavePanel"].Controls["BalanceListBox"].SelectedIndex -and
               $Text -eq $Script:LeaveBalances[$Index].Name.ToLower())
            {
                $DuplicateName = $True
            }
            
            else
            {
                $Index++
            }
        }
        
        if($DuplicateName -eq $True -and
            $Text -eq $Script:LeaveBalances[$Index - 1].Name.ToLower()) #This is to account for if the first item with duplicate names is the one being edited, subtract the index by 1 to include it.
        {
            $Index--
        }

        if($DuplicateName -eq $True)
        {
            $StartingIndex = $Index
            $EndingIndex   = $Index
            
            while($EndingIndex -lt $Script:LeaveBalances.Count -and
                  $Text -eq $Script:LeaveBalances[$EndingIndex].Name.ToLower())
            {
                $EndingIndex++
            }
            
            $EndingIndex-- #Subtract one from the final loop for the end result.

            for($Index = $StartingIndex; $Index -le $EndingIndex; $Index++)
            {
                if($Script:EditLeaveBalanceForm.Controls["LeaveExpiresCheckBox"].Checked -eq $False -and
                   $Script:LeaveBalances[$Index].Expires -eq $False)
                {
                    $Valid = $False

                    $Script:EditLeaveBalanceForm.Controls["WarningLabel"].Text = "Duplicate of non-expiring leave."
                }
                
                elseif($Index -ne $Script:MainForm.Controls["LeavePanel"].Controls["BalanceListBox"].SelectedIndex -and
                       $Script:LeaveBalances[$Index].Expires -eq $True -and
                       $Script:EditLeaveBalanceForm.Controls["LeaveExpiresCheckBox"].Checked -eq $True -and
                       $Script:LeaveBalances[$Index].ExpiresOn -eq $Script:EditLeaveBalanceForm.Controls["LeaveExpiresOnDateTimePicker"].Value.Date)
                {
                    $Valid = $False
                    
                    $Script:EditLeaveBalanceForm.Controls["WarningLabel"].Text = "Duplicate of expiring leave."
                }
            }
        }
    }

    if($Valid -eq $True)
    {
        $Script:EditLeaveBalanceForm.Controls["WarningLabel"].Visible      = $False
        $Script:EditLeaveBalanceForm.Controls["EditLeaveOkButton"].Enabled = $True
    }

    else
    {
        $Script:EditLeaveBalanceForm.Controls["WarningLabel"].Visible      = $True
        $Script:EditLeaveBalanceForm.Controls["EditLeaveOkButton"].Enabled = $False
    }
}

function EditLeaveBalanceFormLeaveExpiresCheckBoxClick
{
    if($Script:EditLeaveBalanceForm.Controls["LeaveExpiresCheckBox"].Checked -eq $True)
    {
        $Script:EditLeaveBalanceForm.Controls["LeaveExpiresOnLabel"].Visible = $True

        $Script:EditLeaveBalanceForm.Controls["LeaveExpiresOnDateTimePicker"].Enabled = $True
        $Script:EditLeaveBalanceForm.Controls["LeaveExpiresOnDateTimePicker"].Visible = $True
    }

    else
    {
        $Script:EditLeaveBalanceForm.Controls["LeaveExpiresOnLabel"].Visible = $False

        $Script:EditLeaveBalanceForm.Controls["LeaveExpiresOnDateTimePicker"].Enabled = $False
        $Script:EditLeaveBalanceForm.Controls["LeaveExpiresOnDateTimePicker"].Visible = $False
    }
}

function EditLeaveBalanceFormOkButtonClick
{
    $SelectedIndex = $Script:MainForm.Controls["LeavePanel"].Controls["BalanceListBox"].SelectedIndex
    
    $SelectedLeaveBalance = $Script:LeaveBalances[$SelectedIndex]

    $NameChanged      = $False
    $BalanceChanged   = $False
    $ThresholdChanged = $False
    $ExpiresChanged   = $False
    $DateChanged      = $False

    $ProjectedLeaveUpdated    = $False
    $CheckProjectedLeaveDates = $False
    $PreviousName             = ""
    $OldName                  = $SelectedLeaveBalance.Name
    $NamesUpdated             = $False
    $ChangesMade              = $False

    if($SelectedLeaveBalance.Name -ne $Script:EditLeaveBalanceForm.Controls["LeaveBalanceNameTextBox"].Text)
    {
        $NameChanged = $True

        $LeaveNameCount = 0

        foreach($LeaveBalance in $Script:LeaveBalances)
        {
            if($LeaveBalance.Name -eq $SelectedLeaveBalance.Name)
            {
                $LeaveNameCount++
            }
        }

        if($LeaveNameCount -eq 1)
        {
            foreach($Leave in $Script:ProjectedLeave)
            {
                if($SelectedLeaveBalance.Name -eq $Leave.LeaveBank)
                {
                    $Leave.LeaveBank = $Script:EditLeaveBalanceForm.Controls["LeaveBalanceNameTextBox"].Text
                    $NamesUpdated    = $True
                }
            }

            $ProjectedLeaveUpdated = $True
        }

        else
        {
            $PreviousName = $SelectedLeaveBalance.Name
            
            $CheckProjectedLeaveDates = $True
        }
    }

    if($SelectedLeaveBalance.Balance -ne $Script:EditLeaveBalanceForm.Controls["BalanceNumericUpDown"].Value)
    {
        $BalanceChanged = $True

        if($SelectedLeaveBalance.Name -eq "Annual")
        {
            $Script:AnnualDecimal = 0.0
        }

        if($SelectedLeaveBalance.Name -eq "Sick")
        {
            $Script:SickDecimal = 0.0
        }
    }
    
    $SelectedLeaveBalance.Name = $Script:EditLeaveBalanceForm.Controls["LeaveBalanceNameTextBox"].Text.Trim()
    $SelectedLeaveBalance.Balance = $Script:EditLeaveBalanceForm.Controls["BalanceNumericUpDown"].Value

    if($SelectedLeaveBalance.Name -eq "Annual" -or
        $SelectedLeaveBalance.Name -eq "Sick")
    {
        if($SelectedLeaveBalance.Threshold -ne $Script:EditLeaveBalanceForm.Controls["ThresholdNumericUpDown"].Value)
        {
            $ThresholdChanged = $True
        
            $SelectedLeaveBalance.Threshold = $Script:EditLeaveBalanceForm.Controls["ThresholdNumericUpDown"].Value
        }
    }

    else
    {
        if($SelectedLeaveBalance.Expires -ne $Script:EditLeaveBalanceForm.Controls["LeaveExpiresCheckBox"].Checked)
        {
            $ExpiresChanged = $True
        }
        
        if($SelectedLeaveBalance.ExpiresOn -ne $Script:EditLeaveBalanceForm.Controls["LeaveExpiresOnDateTimePicker"].Value.Date)
        {
            $DateChanged = $True
        }
        
        $SelectedLeaveBalance.Expires = $Script:EditLeaveBalanceForm.Controls["LeaveExpiresCheckBox"].Checked
        $SelectedLeaveBalance.ExpiresOn = $Script:EditLeaveBalanceForm.Controls["LeaveExpiresOnDateTimePicker"].Value.Date
    }

    if($ExpiresChanged -eq $True -or
       $DateChanged -eq $True -or
       $CheckProjectedLeaveDates -eq $True)
    {
        $LeaveToCheck = $SelectedLeaveBalance.Name

        if($CheckProjectedLeaveDates -eq $True)
        {
            $LeaveToCheck = $PreviousName
        }

        $ChangesMade = UpdateProjectedLeaveDatesIfLeaveBalanceExpiresOrChangesOrDeleted -LeaveBankName $LeaveToCheck

        if($ChangesMade -eq $True)
        {
            $ProjectedLeaveUpdated = $True
        }
    }
    
    if($NameChanged -eq $True -or
       $ExpiresChanged -eq $True -or
       $DateChanged -eq $True)
    {
        $Script:LeaveBalances = [System.Collections.Generic.List[PSCustomObject]] ($Script:LeaveBalances | Sort-Object -Property @{Expression = "Static"; Descending = $True},
                                                                                                                                 @{Expression = "Name"; Descending = $False},
                                                                                                                                 @{Expression = "Expires"; Descending = $True},
                                                                                                                                 @{Expression = "ExpiresOn"; Descending = $False})

        $SelectedIndex = $Script:LeaveBalances.IndexOf($SelectedLeaveBalance)
    }

    if($BalanceChanged -eq $True -or
       $NameChanged -eq $True -or
       $ThresholdChanged -eq $True -or
       $ExpiresChanged -eq $True -or
       $DateChanged -eq $True)
    {
        PopulateLeaveBalanceListBox
        
        $Script:MainForm.Controls["LeavePanel"].Controls["BalanceListBox"].SelectedIndex = $SelectedIndex
    }

    if($ProjectedLeaveUpdated -eq $True)
    {
        $SelectedLeaveDetails = $Script:ProjectedLeave[$Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"].SelectedIndex]
        
        if($Script:ProjectedLeave.Count -gt 1) #Only sort if there's more than one item.
        {
            $Script:ProjectedLeave = [System.Collections.Generic.List[PSCustomObject]] ($Script:ProjectedLeave | Sort-Object -Property "StartDate", "EndDate", "LeaveBank")
        }

        PopulateProjectedLeaveListBox

        if($ChangesMade -eq $True)
        {
            ShowMessageBox -Text "Projected Leave contained entries that happened after the latest possible expiration date. These entries have been updated to not happen after the expiration date. Please verify these are correct." -Caption "Projected Leave Updated" -Buttons "OK" -Icon "Information"
        }

        $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"].SelectedIndex = $Script:ProjectedLeave.IndexOf($SelectedLeaveDetails)
    }

    if($NamesUpdated -eq $True)
    {
        $NewName = $Script:EditLeaveBalanceForm.Controls["LeaveBalanceNameTextBox"].Text
            
        if($OldName.ToLower().Contains("leave") -eq $False)
        {
            $OldName = $OldName.Trim() + " Leave"
        }

        if($NewName.ToLower().Contains("leave") -eq $False)
        {
            $NewName = $NewName.Trim() + " Leave"
        }
            
        ShowMessageBox -Text "$OldName updated to $NewName in Projected Leave." -Caption "Projected Leave Updated" -Buttons "OK" -Icon "Information"
    }

    $Script:EditLeaveBalanceForm.Close()
}

function EditLeaveBalanceFormCancelButtonClick
{
    $Script:EditLeaveBalanceForm.Close()
}

#endregion Edit Leave Balance Form

#region Edit Projected Leave Form

function EditProjectedLeaveFormClosing
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)][System.EventArgs] $EventArguments
    )

    $SelectedLeaveDetails = $Script:ProjectedLeave[$Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"].SelectedIndex]
    
    $ChangesMade = $False

    $NewHashTable = @{}

    foreach($NumericUpDown in $Script:EditProjectedLeaveForm.Controls["LeaveDatesFlowLayoutPanel"].Controls)
    {
        $Date  = $NumericUpDown.Controls["HoursNumericUpDown"].Tag
        $Hours = $NumericUpDown.Controls["HoursNumericUpDown"].Value

        $NewHashTable.$Date = $Hours
    }

    if((Compare-Object @(($SelectedLeaveDetails.HoursHashTable).Keys) @($NewHashTable.Keys)) -ne $Null) #See if the hashtables have different keys.
    {
        $ChangesMade = $True
    }

    else #If the two hashtables have the same keys, check that the values are the same.
    {
        foreach($Key in $NewHashTable.Keys)
        {
            if($NewHashTable[$Key] -ne ($SelectedLeaveDetails.HoursHashTable)[$Key])
            {
                $ChangesMade = $True
            }
        }
    }

    if($SelectedLeaveDetails.LeaveBank -ne $Script:EditProjectedLeaveForm.Controls["LeaveBankComboBox"].Text)
    {
        $ChangesMade = $True
    }

    $Response = ""

    if($ChangesMade -eq $True)
    {
        $Response = ShowMessageBox -Text "You have made changes to the selected projected leave. Would you like to discard these changes?" -Caption "Projected Leave Changed" -Buttons "YesNo" -Icon "Exclamation"
    }

    if($ChangesMade -eq $False -or $Response -eq "Yes")
    {
        if($Script:UnsavedProjectedLeave -eq $True)
        {
            $Script:UnsavedProjectedLeave = $False

            MainFormProjectedDeleteButtonClick
        }
    }

    else
    {
        $EventArguments.Cancel = $True
    }
}

function EditProjectedLeaveOkButton
{
    $SelectedIndex = $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"].SelectedIndex
    
    $SelectedLeaveDetails = $Script:ProjectedLeave[$SelectedIndex]

    $OriginalHours = 0
    $NewHours = 0

    $LeaveBankChanged = $False
    $StartDateChanged = $False
    $EndDateChanged   = $False
    $HoursChanged     = $False

    if($SelectedLeaveDetails.LeaveBank -ne $Script:EditProjectedLeaveForm.Controls["LeaveBankComboBox"].Text)
    {
        $LeaveBankChanged = $True
    }

    if($SelectedLeaveDetails.StartDate -ne $Script:EditProjectedLeaveForm.Controls["LeaveStartDateTimePicker"].Value)
    {
        $StartDateChanged = $True
    }

    if($SelectedLeaveDetails.EndDate -ne $Script:EditProjectedLeaveForm.Controls["LeaveEndDateTimePicker"].Value)
    {
        $EndDateChanged = $True
    }

    foreach($Date in $SelectedLeaveDetails.HoursHashTable.Keys)
    {
        $OriginalHours += $SelectedLeaveDetails.HoursHashTable[$Date]
    }
    
    $SelectedLeaveDetails.LeaveBank = $Script:EditProjectedLeaveForm.Controls["LeaveBankComboBox"].Text
    $SelectedLeaveDetails.StartDate = $Script:EditProjectedLeaveForm.Controls["LeaveStartDateTimePicker"].Value
    $SelectedLeaveDetails.EndDate   = $Script:EditProjectedLeaveForm.Controls["LeaveEndDateTimePicker"].Value

    $SelectedLeaveDetails.HoursHashTable.Clear()
    
    foreach($NumericUpDown in $Script:EditProjectedLeaveForm.Controls["LeaveDatesFlowLayoutPanel"].Controls)
    {
        $Date  = $NumericUpDown.Controls["HoursNumericUpDown"].Tag
        $Hours = $NumericUpDown.Controls["HoursNumericUpDown"].Value

        $SelectedLeaveDetails.HoursHashTable.$Date = $Hours

        $NewHours += $Hours
    }
    
    if($OriginalHours -ne $NewHours)
    {
        $HoursChanged = $True
    }
    
    if($HoursChanged -eq $True -or
       $EndDateChanged -eq $True -or
       $StartDateChanged -eq $True -or
       $LeaveBankChanged -eq $True -or
       $Script:UnsavedProjectedLeave -eq $True)
    {
        if($Script:ProjectedLeave.Count -gt 1) #Only sort if there's more than one item.
        {
            $Script:ProjectedLeave = [System.Collections.Generic.List[PSCustomObject]] ($Script:ProjectedLeave | Sort-Object -Property "StartDate", "EndDate", "LeaveBank")
        }

        PopulateProjectedLeaveListBox

        $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"].SelectedIndex = $Script:ProjectedLeave.IndexOf($SelectedLeaveDetails)
    }

    $Script:UnsavedProjectedLeave = $False
    
    $Script:EditProjectedLeaveForm.Close()
}

function EditProjectedLeaveCancelButton
{
    $Script:EditProjectedLeaveForm.Close()
}

function EditProjectedLeaveBankSelectedIndexChanged
{
    $ExpirationDate = $Script:EditProjectedLeaveForm.Controls["LeaveBankComboBox"].Tag[$Script:EditProjectedLeaveForm.Controls["LeaveBankComboBox"].Text]

    $EndOfPayPeriod = GetEndingOfPayPeriodForDate -Date $Script:EditProjectedLeaveForm.Controls["LeaveStartDateTimePicker"].Value.Date

    $SomethingChanged = $False

    if($Script:EditProjectedLeaveForm.Controls["LeaveStartDateTimePicker"].Value -gt $ExpirationDate)
    {
        $Script:EditProjectedLeaveForm.Controls["LeaveStartDateTimePicker"].Value = $ExpirationDate

        $Script:EditProjectedLeaveForm.Controls["LeaveEndDateTimePicker"].MinDate = $ExpirationDate

        $Script:EditProjectedLeaveForm.Controls["LeaveStartDateTimePicker"].Tag = $Script:EditProjectedLeaveForm.Controls["LeaveStartDateTimePicker"].Value.Date

        $SomethingChanged = $True
    }
    
    if($Script:EditProjectedLeaveForm.Controls["LeaveEndDateTimePicker"].Value -gt $ExpirationDate)
    {
        $Script:EditProjectedLeaveForm.Controls["LeaveEndDateTimePicker"].Value = $ExpirationDate

        $SomethingChanged = $True
    }

    if($SomethingChanged -eq $True)
    {
        PopulateProjectedLeaveDays
    }
    
    $Script:EditProjectedLeaveForm.Controls["LeaveStartDateTimePicker"].MaxDate = $ExpirationDate
    $Script:EditProjectedLeaveForm.Controls["LeaveEndDateTimePicker"].MaxDate   = $ExpirationDate

    if($ExpirationDate -gt $EndOfPayPeriod)
    {
        $Script:EditProjectedLeaveForm.Controls["LeaveEndDateTimePicker"].MaxDate = $EndOfPayPeriod
    }
}

function HourNumericUpDownChanged
{
    $TotalHours = 0

    foreach($Panel in $Script:EditProjectedLeaveForm.Controls["LeaveDatesFlowLayoutPanel"].Controls)
    {
        $TotalHours += $Panel.Controls["HoursNumericUpDown"].Value
    }

    $Script:EditProjectedLeaveForm.Controls["HoursOfLeaveTaken"].Text = "Hours of Leave Taken: " + $TotalHours
}

function StartDateCalendarClosedUp
{
    if($Script:EditProjectedLeaveForm.Controls["LeaveStartDateTimePicker"].Tag -ne $Script:EditProjectedLeaveForm.Controls["LeaveStartDateTimePicker"].Value.Date)
    {
        $DayDifference = ($Script:EditProjectedLeaveForm.Controls["LeaveEndDateTimePicker"].Value.Date - $Script:EditProjectedLeaveForm.Controls["LeaveStartDateTimePicker"].Tag).Days
        
        $NewDate = $Script:EditProjectedLeaveForm.Controls["LeaveStartDateTimePicker"].Value.Date.AddDays($DayDifference)

        $EndOfPayPeriod = GetEndingOfPayPeriodForDate -Date $Script:EditProjectedLeaveForm.Controls["LeaveStartDateTimePicker"].Value.Date

        $ExpirationDate = $Script:EditProjectedLeaveForm.Controls["LeaveBankComboBox"].Tag[$Script:EditProjectedLeaveForm.Controls["LeaveBankComboBox"].Text]
        
        $Script:EditProjectedLeaveForm.Controls["LeaveEndDateTimePicker"].MaxDate = $Script:LastSelectableDate #Setting this to the max so we don't have an exception where the min is greater than the max.

        $Script:EditProjectedLeaveForm.Controls["LeaveEndDateTimePicker"].MinDate = $Script:EditProjectedLeaveForm.Controls["LeaveStartDateTimePicker"].Value.Date

        #Actually set the max.
        if($ExpirationDate -lt $EndOfPayPeriod)
        {
            $Script:EditProjectedLeaveForm.Controls["LeaveEndDateTimePicker"].MaxDate = $ExpirationDate
        }

        else
        {
            $Script:EditProjectedLeaveForm.Controls["LeaveEndDateTimePicker"].MaxDate = $EndOfPayPeriod
        }

        if($NewDate -gt $EndOfPayPeriod -or
           $NewDate -gt $ExpirationDate)
        {
            $EarlierDate = $EndOfPayPeriod
            
            if($EarlierDate -gt $ExpirationDate)
            {
                $EarlierDate = $ExpirationDate
            }
            
            $Script:EditProjectedLeaveForm.Controls["LeaveEndDateTimePicker"].Value = $EarlierDate
        }

        else
        {
            $Script:EditProjectedLeaveForm.Controls["LeaveEndDateTimePicker"].Value = $NewDate
        }

        $Script:EditProjectedLeaveForm.Controls["LeaveStartDateTimePicker"].Tag = $Script:EditProjectedLeaveForm.Controls["LeaveStartDateTimePicker"].Value.Date
        $Script:EditProjectedLeaveForm.Controls["LeaveEndDateTimePicker"].Tag   = $Script:EditProjectedLeaveForm.Controls["LeaveEndDateTimePicker"].Value.Date

        PopulateProjectedLeaveDays

        HourNumericUpDownChanged
    }
}

function EndDateCalendarClosedUp
{
    $Script:CalendarDroppedDown = $False

    if($Script:EditProjectedLeaveForm.Controls["LeaveEndDateTimePicker"].Tag -ne $Script:EditProjectedLeaveForm.Controls["LeaveEndDateTimePicker"].Value)
    {
        PopulateProjectedLeaveDays

        HourNumericUpDownChanged

        $Script:EditProjectedLeaveForm.Controls["LeaveEndDateTimePicker"].Tag = $Script:EditProjectedLeaveForm.Controls["LeaveEndDateTimePicker"].Value.Date
    }
}

#endregion Edit Projected Leave Form

#region Output Form

function OutputFormCopyButton
{
    Set-Clipboard $Script:OutputForm.Controls["OutputRichTextBox"].Text
}

function OutputFormCloseButton
{
    $Script:OutputForm.Close()
}

#endregion Output Form

#endregion Event Handlers

#region Form Building Functions

function BuildMainForm
{
    $Script:MainForm          = New-Object System.Windows.Forms.Form
    $MainForm.Name            = "MainForm"
    $MainForm.BackColor       = "WhiteSmoke"
    $MainForm.Font            = $Script:FormFont
    $MainForm.FormBorderStyle = "FixedSingle"
    $MainForm.MaximizeBox     = $False
    $MainForm.Size            = New-Object System.Drawing.Size(950, 550)
    $MainForm.Text            = "Federal Civilian Leave Calculator"
    $MainForm.WindowState     = "Normal"

    $SettingsPanel = New-Object System.Windows.Forms.Panel
    $SettingsPanel.Name = "SettingsPanel"
    $SettingsPanel.BackColor = "LightGray"
    $SettingsPanel.Dock = "Top"
    $SettingsPanel.Height = 70
    $SettingsPanel.TabIndex = 1

    $SCDLeaveDateLabel = New-Object System.Windows.Forms.Label
    $SCDLeaveDateLabel.Name = "SCDLeaveDateLabel"
    $SCDLeaveDateLabel.Height = 17
    $SCDLeaveDateLabel.Left = 20
    $SCDLeaveDateLabel.Text = "SCD Leave Date:"
    $SCDLeaveDateLabel.Top = 8
    $SCDLeaveDateLabel.Width = 92

    $SCDLeaveDateDateTimePicker = New-Object System.Windows.Forms.DateTimePicker
    $SCDLeaveDateDateTimePicker.Name = "SCDLeaveDateDateTimePicker"
    $SCDLeaveDateDateTimePicker.Format = "Short"
    $SCDLeaveDateDateTimePicker.Left = 119
    $SCDLeaveDateDateTimePicker.MaxDate = $Script:CurrentDate
    $SCDLeaveDateDateTimePicker.Top = 8
    $SCDLeaveDateDateTimePicker.Width = 115

    $UpdateInfoButton = New-Object System.Windows.Forms.Button
    $UpdateInfoButton.Name = "UpdateInfoButton"
    $UpdateInfoButton.Height = 24
    $UpdateInfoButton.Left = 288
    $UpdateInfoButton.Text = "Update Employee Info"
    $UpdateInfoButton.Top = 8
    $UpdateInfoButton.Width = 138
    
    $LengthOfServiceTextBox = New-Object System.Windows.Forms.TextBox
    $LengthOfServiceTextBox.Name = "LengthOfServiceTextBox"
    $LengthOfServiceTextBox.Height = 20
    $LengthOfServiceTextBox.Left = 20
    $LengthOfServiceTextBox.ReadOnly = $True
    $LengthOfServiceTextBox.TabStop = $False
    $LengthOfServiceTextBox.Top = 38
    $LengthOfServiceTextBox.Width = 250
    
    $EmployeeTypeTextBox = New-Object System.Windows.Forms.TextBox
    $EmployeeTypeTextBox.Name = "EmployeeTypeTextBox"
    $EmployeeTypeTextBox.Height = 20
    $EmployeeTypeTextBox.Left = 288
    $EmployeeTypeTextBox.ReadOnly = $True
    $EmployeeTypeTextBox.TabStop = $False
    $EmployeeTypeTextBox.Top = 38
    $EmployeeTypeTextBox.Width = 138

    $DisplayBalanceEveryLeaveCheckBox = New-Object System.Windows.Forms.CheckBox
    $DisplayBalanceEveryLeaveCheckBox.Name = "DisplayBalanceEveryLeaveCheckBox"
    $DisplayBalanceEveryLeaveCheckBox.Left = 450
    $DisplayBalanceEveryLeaveCheckBox.Text = "Display Balance After Each Day of Leave"
    $DisplayBalanceEveryLeaveCheckBox.Top = 4
    $DisplayBalanceEveryLeaveCheckBox.Width = 230

    $DisplayBalanceEveryPayPeriodEnd = New-Object System.Windows.Forms.CheckBox
    $DisplayBalanceEveryPayPeriodEnd.Name = "DisplayBalanceEveryPayPeriodEnd"
    $DisplayBalanceEveryPayPeriodEnd.Left = 450
    $DisplayBalanceEveryPayPeriodEnd.Text = "Display Balance After Each Pay Period Ends"
    $DisplayBalanceEveryPayPeriodEnd.Top = 26
    $DisplayBalanceEveryPayPeriodEnd.Width = 247

    $DisplayLeaveHighsAndLows = New-Object System.Windows.Forms.CheckBox
    $DisplayLeaveHighsAndLows.Name = "DisplayLeaveHighsAndLows"
    $DisplayLeaveHighsAndLows.Left = 450
    $DisplayLeaveHighsAndLows.Text = "Display Annual/Sick Leave Highs/Lows"
    $DisplayLeaveHighsAndLows.Top = 48
    $DisplayLeaveHighsAndLows.Width = 247

    $ReportPanel = New-Object System.Windows.Forms.Panel
    $ReportPanel.Name = "ReportPanel"
    $ReportPanel.BackColor = "LightGray"
    $ReportPanel.Dock = "Bottom"
    $ReportPanel.Height = 70
    $ReportPanel.TabIndex = 3

    $ProjectBalanceRadioButton = New-Object System.Windows.Forms.RadioButton
    $ProjectBalanceRadioButton.Name = "ProjectBalanceRadioButton"
    $ProjectBalanceRadioButton.Left = 10
    $ProjectBalanceRadioButton.Top = 10
    $ProjectBalanceRadioButton.Width = 24
    
    $ReachGoalRadioButton = New-Object System.Windows.Forms.RadioButton
    $ReachGoalRadioButton.Name = "ReachGoalRadioButton"
    $ReachGoalRadioButton.Left = 10
    $ReachGoalRadioButton.Top = 40
    $ReachGoalRadioButton.Width = 24
    
    $ProjectToLabel = New-Object System.Windows.Forms.Label
    $ProjectToLabel.Name = "ProjectToLabel"
    $ProjectToLabel.Height = 17
    $ProjectToLabel.Left = 42
    $ProjectToLabel.Text = "Project to Date:"
    $ProjectToLabel.Top = 10
    $ProjectToLabel.Width = 82

    $ProjectToDateDateTimePicker = New-Object System.Windows.Forms.DateTimePicker
    $ProjectToDateDateTimePicker.Name = "ProjectToDateDateTimePicker"
    $ProjectToDateDateTimePicker.Format = "Short"
    $ProjectToDateDateTimePicker.Left = 130
    $ProjectToDateDateTimePicker.MinDate = $Script:BeginningOfPayPeriod
    $ProjectToDateDateTimePicker.MaxDate = $Script:LastSelectableDate
    $ProjectToDateDateTimePicker.Top = 6
    $ProjectToDateDateTimePicker.Width = 115

    $ReachGoalLabel = New-Object System.Windows.Forms.Label
    $ReachGoalLabel.Name = "ReachGoalLabel"
    $ReachGoalLabel.Height = 17
    $ReachGoalLabel.Left = 42
    $ReachGoalLabel.Text = "Reach Goal"
    $ReachGoalLabel.Top = 42
    $ReachGoalLabel.Width = 64

    $AnnualGoalLabel = New-Object System.Windows.Forms.Label
    $AnnualGoalLabel.Name = "AnnualGoalLabel"
    $AnnualGoalLabel.Height = 17
    $AnnualGoalLabel.Left = 129
    $AnnualGoalLabel.Text = "Annual Leave:"
    $AnnualGoalLabel.Top = 42
    $AnnualGoalLabel.Width = 77

    $AnnualGoalNumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $AnnualGoalNumericUpDown.Name = "AnnualGoalNumericUpDown"
    $AnnualGoalNumericUpDown.Height = 20
    $AnnualGoalNumericUpDown.Left = 207
    $AnnualGoalNumericUpDown.Maximum = $Script:MaximumAnnual
    $AnnualGoalNumericUpDown.Top = 40
    $AnnualGoalNumericUpDown.Width = 55

    $SickGoalLabel = New-Object System.Windows.Forms.Label
    $SickGoalLabel.Name = "SickGoalLabel"
    $SickGoalLabel.Height = 17
    $SickGoalLabel.Left = 280
    $SickGoalLabel.Text = "Sick Leave:"
    $SickGoalLabel.Top = 42
    $SickGoalLabel.Width = 63

    $SickGoalNumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $SickGoalNumericUpDown.Name = "SickGoalNumericUpDown"
    $SickGoalNumericUpDown.Height = 20
    $SickGoalNumericUpDown.Left = 345
    $SickGoalNumericUpDown.Maximum = $Script:MaximumSick
    $SickGoalNumericUpDown.Top = 40
    $SickGoalNumericUpDown.Width = 55

    $ProjectButton = New-Object System.Windows.Forms.Button
    $ProjectButton.Name = "ProjectButton"
    $ProjectButton.Height = 24
    $ProjectButton.Left = 501
    $ProjectButton.Text = "Run Projection"
    $ProjectButton.Top = 24
    $ProjectButton.Width = 363
    
    $LeavePanel = New-Object System.Windows.Forms.Panel
    $LeavePanel.Name = "LeavePanel"
    $LeavePanel.Dock = "Fill"
    $LeavePanel.TabIndex = 2

    $LeaveBalancesLabel = New-Object System.Windows.Forms.Label
    $LeaveBalancesLabel.Name = "LeaveBalancesLabel"
    $LeaveBalancesLabel.AutoSize = $True
    $LeaveBalancesLabel.Height = 17
    $LeaveBalancesLabel.Left = 60
    $LeaveBalancesLabel.Text = "Leave Balances as of Pay Period Ending: " + $Script:BeginningOfPayPeriod.AddDays(-1).ToString("MM/dd/yyyy") #Subtract one day so it's the end of the previous pay period matching what's in MyPay.
    $LeaveBalancesLabel.Top = 5
    $LeaveBalancesLabel.Width = 85

    $ProjectedLeaveLabel = New-Object System.Windows.Forms.Label
    $ProjectedLeaveLabel.Name = "ProjectedLeaveLabel"
    $ProjectedLeaveLabel.Height = 17
    $ProjectedLeaveLabel.Left = 609
    $ProjectedLeaveLabel.Text = "Projected Leave"
    $ProjectedLeaveLabel.Top = 5
    $ProjectedLeaveLabel.Width = 86

    $BalanceAddButton = New-Object System.Windows.Forms.Button
    $BalanceAddButton.Name = "BalanceAddButton"
    $BalanceAddButton.Font = $Script:IconsFont
    $BalanceAddButton.ForeColor = "Green"
    $BalanceAddButton.Height = 25
    $BalanceAddButton.Left = 127
    $BalanceAddButton.Text = [char]0xF8AA #+ Symbol
    $BalanceAddButton.Top = 325
    $BalanceAddButton.Width = 29

    $BalanceEditButton = New-Object System.Windows.Forms.Button
    $BalanceEditButton.Name = "BalanceEditButton"
    $BalanceEditButton.Font = $Script:IconsFont
    $BalanceEditButton.ForeColor = "Orange"
    $BalanceEditButton.Height = 25
    $BalanceEditButton.Left = 166
    $BalanceEditButton.Text = [char]0xE70F #Pencil/Edit Symbol
    $BalanceEditButton.Top = 325
    $BalanceEditButton.Width = 29
    
    $BalanceDeleteButton = New-Object System.Windows.Forms.Button
    $BalanceDeleteButton.Name = "BalanceDeleteButton"
    $BalanceDeleteButton.Font = $Script:IconsFont
    $BalanceDeleteButton.ForeColor = "Red"
    $BalanceDeleteButton.Height = 25
    $BalanceDeleteButton.Left = 206
    $BalanceDeleteButton.Text = [char]0xF78A #X Symbol
    $BalanceDeleteButton.Top = 325
    $BalanceDeleteButton.Width = 29
    
    $ProjectedAddButton = New-Object System.Windows.Forms.Button
    $ProjectedAddButton.Name = "ProjectedAddButton"
    $ProjectedAddButton.Font = $Script:IconsFont
    $ProjectedAddButton.ForeColor = "Green"
    $ProjectedAddButton.Height = 25
    $ProjectedAddButton.Left = 604
    $ProjectedAddButton.Text = [char]0xF8AA #+ Symbol
    $ProjectedAddButton.Top = 325
    $ProjectedAddButton.Width = 29
    
    $ProjectedEditButton = New-Object System.Windows.Forms.Button
    $ProjectedEditButton.Name = "ProjectedEditButton"
    $ProjectedEditButton.Enabled = $False
    $ProjectedEditButton.Font = $Script:IconsFont
    $ProjectedEditButton.ForeColor = "Orange"
    $ProjectedEditButton.Height = 25
    $ProjectedEditButton.Left = 643
    $ProjectedEditButton.Text = [char]0xE70F #Pencil/Edit Symbol
    $ProjectedEditButton.Top = 325
    $ProjectedEditButton.Width = 29
    
    $ProjectedDeleteButton = New-Object System.Windows.Forms.Button
    $ProjectedDeleteButton.Name = "ProjectedDeleteButton"
    $ProjectedDeleteButton.Enabled = $False
    $ProjectedDeleteButton.Font = $Script:IconsFont
    $ProjectedDeleteButton.ForeColor = "Red"
    $ProjectedDeleteButton.Height = 25
    $ProjectedDeleteButton.Left = 689
    $ProjectedDeleteButton.Text = [char]0xF78A #X Symbol
    $ProjectedDeleteButton.Top = 325
    $ProjectedDeleteButton.Width = 29
    
    $BalanceListBox = New-Object System.Windows.Forms.ListBox
    $BalanceListBox.Name = "BalanceListBox"
    $BalanceListBox.Height = 277
    $BalanceListBox.Left = 75
    $BalanceListBox.Top = 30
    $BalanceListBox.Width = 210

    $ProjectedListBox = New-Object System.Windows.Forms.ListBox
    $ProjectedListBox.Name = "ProjectedListBox"
    $ProjectedListBox.Height = 277
    $ProjectedListBox.Left = 490
    $ProjectedListBox.Top = 30
    $ProjectedListBox.Width = 330

    $MainForm.Controls.AddRange(($LeavePanel, $SettingsPanel, $ReportPanel))
    $SettingsPanel.Controls.AddRange(($SCDLeaveDateLabel, $SCDLeaveDateDateTimePicker, $UpdateInfoButton, $LengthOfServiceTextBox, $EmployeeTypeTextBox, $DisplayBalanceEveryLeaveCheckBox, $DisplayBalanceEveryPayPeriodEnd, $DisplayLeaveHighsAndLows))
    $ReportPanel.Controls.AddRange(($ProjectBalanceRadioButton, $ReachGoalRadioButton, $ProjectToLabel, $ProjectToDateDateTimePicker, $ReachGoalLabel, $AnnualGoalLabel, $AnnualGoalNumericUpDown, $SickGoalLabel, $SickGoalNumericUpDown, $ProjectButton))
    $LeavePanel.Controls.AddRange(($LeaveBalancesLabel, $ProjectedLeaveLabel, $BalanceListBox, $BalanceAddButton, $BalanceEditButton, $BalanceDeleteButton, $ProjectedListBox, $ProjectedAddButton, $ProjectedEditButton, $ProjectedDeleteButton))
    
    #Select/Check the appropriate things.

    $MainForm.ActiveControl = $ProjectButton

    if($Script:SCDLeaveDate -gt $Script:CurrentDate)
    {
        $Script:SCDLeaveDate = $Script:CurrentDate
    }

    $SCDLeaveDateDateTimePicker.Value = $Script:SCDLeaveDate

    UpdateTypeOfEmployeeTextBoxString

    UpdateLengthOfServiceStrings

    if($Script:DisplayAfterEachLeave -eq $True)
    {
        $DisplayBalanceEveryLeaveCheckBox.Checked = $True
    }

    if($Script:DisplayAfterEachPP -eq $True)
    {
        $DisplayBalanceEveryPayPeriodEnd.Checked = $True
    }

    if($Script:DisplayHighsAndLows -eq $True)
    {
        $DisplayLeaveHighsAndLows.Checked = $True
    }

    if($Script:ProjectToDate -lt $Script:CurrentDate)
    {
        $Script:ProjectToDate = $Script:CurrentDate
    }

    $ProjectToDateDateTimePicker.Value = $Script:ProjectToDate

    $AnnualGoalNumericUpDown.Value = $Script:AnnualGoal
    $SickGoalNumericUpDown.Value   = $Script:SickGoal

    if($Script:ProjectOrGoal -eq "Project")
    {
        $ProjectBalanceRadioButton.Checked = $True

        $AnnualGoalNumericUpDown.Enabled = $False
        $SickGoalNumericUpDown.Enabled   = $False
    }

    elseif($Script:ProjectOrGoal -eq "Goal")
    {
        $ReachGoalRadioButton.Checked = $True

        $ProjectToDateDateTimePicker.Enabled = $False
    }

    PopulateLeaveBalanceListBox

    if($Script:ProjectedLeave.Count -gt 0)
    {
        PopulateProjectedLeaveListBox
    }

    if($ProjectedListBox.Items.Count -gt 0)
    {
        $ProjectedListBox.SelectedIndex = 0
        $ProjectedEditButton.Enabled    = $True
        $ProjectedDeleteButton.Enabled  = $True
    }

    $MainForm.Add_FormClosing({MainFormClosing -EventArguments $_})
    $UpdateInfoButton.Add_Click({MainFormUpdateInfoButtonClick})
    $LengthOfServiceTextBox.Add_Click({MainFormLengthOfServiceTextBoxClick})
    $ProjectBalanceRadioButton.Add_Click({MainFormProjectBalanceRadioButtonClick})
    $ReachGoalRadioButton.Add_Click({MainFormReachGoalRadioButtonClick})
    $ProjectButton.Add_Click({MainFormProjectButtonClick})
    $BalanceAddButton.Add_Click({MainFormBalanceAddButtonClick})
    $BalanceEditButton.Add_Click({MainFormBalanceEditButtonClick})
    $BalanceDeleteButton.Add_Click({MainFormBalanceDeleteButtonClick})
    $ProjectedAddButton.Add_Click({MainFormProjectedAddButtonClick})
    $ProjectedEditButton.Add_Click({MainFormProjectedEditButtonClick})
    $ProjectedDeleteButton.Add_Click({MainFormProjectedDeleteButtonClick})
    $SCDLeaveDateDateTimePicker.Add_ValueChanged({MainFormSCDLeaveDateTimePickerValueChanged})
    $BalanceListBox.Add_SelectedIndexChanged({MainFormLeaveBalanceListBoxIndexChanged})
    $BalanceListBox.Add_DoubleClick({MainFormLeaveBalanceListBoxDoubleClick -EventArguments $_})
    $ProjectedListBox.Add_DoubleClick({MainFormProjectedLeaveListBoxDoubleClick -EventArguments $_})
    $ProjectToDateDateTimePicker.Add_ValueChanged({MainFormProjectToDateValueChanged})
    $AnnualGoalNumericUpDown.Add_ValueChanged({MainFormAnnualGoalValueChanged})
    $SickGoalNumericUpDown.Add_ValueChanged({MainFormSickGoalValueChanged})
    $DisplayBalanceEveryLeaveCheckBox.Add_Click({MainFormEveryLeaveCheckBoxClicked})
    $DisplayBalanceEveryPayPeriodEnd.Add_Click({MainFormEveryPPCheckBoxClicked})
    $DisplayLeaveHighsAndLows.Add_Click({MainFormDisplayHighsLowsClicked})
    
    #This is done after the events because I do want the event to fire.
    $BalanceListBox.SelectedIndex = 0
}

function BuildEmployeeInfoForm
{
    $Script:EmployeeInfoForm          = New-Object System.Windows.Forms.Form
    $EmployeeInfoForm.Name            = "EmployeeInfoForm"
    $EmployeeInfoForm.BackColor       = "WhiteSmoke"
    $EmployeeInfoForm.Font            = $Script:FormFont
    $EmployeeInfoForm.FormBorderStyle = "FixedSingle"
    $EmployeeInfoForm.MaximizeBox     = $False
    $EmployeeInfoForm.Size            = New-Object System.Drawing.Size(291, 451)
    $EmployeeInfoForm.Text            = "Employee Information"
    $EmployeeInfoForm.WindowState     = "Normal"

    $EmployeeInfoOkButton = New-Object System.Windows.Forms.Button
    $EmployeeInfoOkButton.Name = "EmployeeInfoOkButton"
    $EmployeeInfoOkButton.Height = 24
    $EmployeeInfoOkButton.Left = 49
    $EmployeeInfoOkButton.Text = "OK"
    $EmployeeInfoOkButton.Top = 377
    $EmployeeInfoOkButton.Width = 75
    
    $EmployeeInfoCancelButton = New-Object System.Windows.Forms.Button
    $EmployeeInfoCancelButton.Name = "EmployeeInfoCancelButton"
    $EmployeeInfoCancelButton.Height = 24
    $EmployeeInfoCancelButton.Left = 162
    $EmployeeInfoCancelButton.Text = "Cancel"
    $EmployeeInfoCancelButton.Top = 377
    $EmployeeInfoCancelButton.Width = 75
    
    $EmployeeTypePanel = New-Object System.Windows.Forms.Panel
    $EmployeeTypePanel.Name = "EmployeeTypePanel"
    $EmployeeTypePanel.Height = 76
    $EmployeeTypePanel.Width = 123

    $FullTimeRadioButton = New-Object System.Windows.Forms.RadioButton
    $FullTimeRadioButton.Name = "FullTimeRadioButton"
    $FullTimeRadioButton.Text = "Full-Time"
    $FullTimeRadioButton.Top = 15
    $FullTimeRadioButton.Width = 70

    $PartTimeRadioButton = New-Object System.Windows.Forms.RadioButton
    $PartTimeRadioButton.Name = "PartTimeRadioButton"
    $PartTimeRadioButton.Text = "Part-Time"
    $PartTimeRadioButton.Top = 33
    $PartTimeRadioButton.Width = 72

    $SESRadioButton = New-Object System.Windows.Forms.RadioButton
    $SESRadioButton.Name = "SESRadioButton"
    $SESRadioButton.Text = "SES"
    $SESRadioButton.Top = 51
    $SESRadioButton.Width = 66

    $EmploymentTypeLabel = New-Object System.Windows.Forms.Label
    $EmploymentTypeLabel.Name = "EmploymentTypeLabel"
    $EmploymentTypeLabel.Height = 17
    $EmploymentTypeLabel.Text = "Employment Type:"
    $EmploymentTypeLabel.Width = 99

    $LeaveCeilingPanel = New-Object System.Windows.Forms.Panel
    $LeaveCeilingPanel.Name = "LeaveCeilingPanel"
    $LeaveCeilingPanel.Height = 76
    $LeaveCeilingPanel.Left = 130
    $LeaveCeilingPanel.Width = 135

    $LeaveCeilingLabel = New-Object System.Windows.Forms.Label
    $LeaveCeilingLabel.Name = "LeaveCeilingLabel"
    $LeaveCeilingLabel.Height = 17
    $LeaveCeilingLabel.Text = "Leave Ceiling:"
    $LeaveCeilingLabel.Width = 76

    $CONUSRadioButton = New-Object System.Windows.Forms.RadioButton
    $CONUSRadioButton.Name = "CONUSRadioButton"
    $CONUSRadioButton.Text = "CONUS (240 Hours)"
    $CONUSRadioButton.Top = 15
    $CONUSRadioButton.Width = 126

    $OCONUSRadioButton = New-Object System.Windows.Forms.RadioButton
    $OCONUSRadioButton.Name = "OCONUSRadioButton"
    $OCONUSRadioButton.Text = "OCONUS (360 Hours)"
    $OCONUSRadioButton.Top = 33
    $OCONUSRadioButton.Width = 135

    $SESCeilingRadioButton = New-Object System.Windows.Forms.RadioButton
    $SESCeilingRadioButton.Name = "SESCeilingRadioButton"
    $SESCeilingRadioButton.Text = "SES (720 Hours)"
    $SESCeilingRadioButton.Top = 51
    $SESCeilingRadioButton.Width = 108

    $InaugurationDayHolidayCheckBox = New-Object System.Windows.Forms.CheckBox
    $InaugurationDayHolidayCheckBox.Name = "InaugurationDayHolidayCheckBox"
    $InaugurationDayHolidayCheckBox.Left = 25
    $InaugurationDayHolidayCheckBox.Text = "Entitled to a Holiday on Inauguration Day"
    $InaugurationDayHolidayCheckBox.Top = 78
    $InaugurationDayHolidayCheckBox.Width = 229

    $Week1Label = New-Object System.Windows.Forms.Label
    $Week1Label.Name = "Week1Label"
    $Week1Label.Height = 17
    $Week1Label.Left = 72
    $Week1Label.Text = "Week 1"
    $Week1Label.Top = 118
    $Week1Label.Width = 42

    $Week2Label = New-Object System.Windows.Forms.Label
    $Week2Label.Name = "Week2Label"
    $Week2Label.Height = 17
    $Week2Label.Left = 175
    $Week2Label.Text = "Week 2"
    $Week2Label.Top = 118
    $Week2Label.Width = 42

    $SundayLabel = New-Object System.Windows.Forms.Label
    $SundayLabel.Name = "SundayLabel"
    $SundayLabel.Height = 17
    $SundayLabel.Text = "Sunday"
    $SundayLabel.Top = 138
    $SundayLabel.Width = 43

    $MondayLabel = New-Object System.Windows.Forms.Label
    $MondayLabel.Name = "MondayLabel"
    $MondayLabel.Height = 17
    $MondayLabel.Text = "Monday"
    $MondayLabel.Top = 168
    $MondayLabel.Width = 44

    $TuesdayLabel = New-Object System.Windows.Forms.Label
    $TuesdayLabel.Name = "TuesdayLabel"
    $TuesdayLabel.Height = 17
    $TuesdayLabel.Text = "Tuesday"
    $TuesdayLabel.Top = 195
    $TuesdayLabel.Width = 48

    $WednesdayLabel = New-Object System.Windows.Forms.Label
    $WednesdayLabel.Name = "WednesdayLabel"
    $WednesdayLabel.Height = 17
    $WednesdayLabel.Text = "Wednesday"
    $WednesdayLabel.Top = 225
    $WednesdayLabel.Width = 64

    $ThursdayLabel = New-Object System.Windows.Forms.Label
    $ThursdayLabel.Name = "ThursdayLabel"
    $ThursdayLabel.Height = 17
    $ThursdayLabel.Text = "Thursday"
    $ThursdayLabel.Top = 255
    $ThursdayLabel.Width = 51

    $FridayLabel = New-Object System.Windows.Forms.Label
    $FridayLabel.Name = "FridayLabel"
    $FridayLabel.Height = 17
    $FridayLabel.Text = "Friday"
    $FridayLabel.Top = 282
    $FridayLabel.Width = 36

    $SaturdayLabel = New-Object System.Windows.Forms.Label
    $SaturdayLabel.Name = "SaturdayLabel"
    $SaturdayLabel.Height = 17
    $SaturdayLabel.Text = "Saturday"
    $SaturdayLabel.Top = 309
    $SaturdayLabel.Width = 50

    $Day1NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day1NumericUpDown.Name = "Day1NumericUpDown"
    $Day1NumericUpDown.Height = 20
    $Day1NumericUpDown.Left = 73
    $Day1NumericUpDown.Maximum = 24
    $Day1NumericUpDown.Top = 135
    $Day1NumericUpDown.Width = 40

    $Day2NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day2NumericUpDown.Name = "Day2NumericUpDown"
    $Day2NumericUpDown.Height = 20
    $Day2NumericUpDown.Left = 73
    $Day2NumericUpDown.Maximum = 24
    $Day2NumericUpDown.Top = 166
    $Day2NumericUpDown.Width = 40

    $Day3NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day3NumericUpDown.Name = "Day3NumericUpDown"
    $Day3NumericUpDown.Height = 20
    $Day3NumericUpDown.Left = 73
    $Day3NumericUpDown.Maximum = 24
    $Day3NumericUpDown.Top = 194
    $Day3NumericUpDown.Width = 40

    $Day4NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day4NumericUpDown.Name = "Day4NumericUpDown"
    $Day4NumericUpDown.Height = 20
    $Day4NumericUpDown.Left = 73
    $Day4NumericUpDown.Maximum = 24
    $Day4NumericUpDown.Top = 221
    $Day4NumericUpDown.Width = 40

    $Day5NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day5NumericUpDown.Name = "Day5NumericUpDown"
    $Day5NumericUpDown.Height = 20
    $Day5NumericUpDown.Left = 73
    $Day5NumericUpDown.Maximum = 24
    $Day5NumericUpDown.Top = 250
    $Day5NumericUpDown.Width = 40

    $Day6NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day6NumericUpDown.Name = "Day6NumericUpDown"
    $Day6NumericUpDown.Height = 20
    $Day6NumericUpDown.Left = 73
    $Day6NumericUpDown.Maximum = 24
    $Day6NumericUpDown.Top = 277
    $Day6NumericUpDown.Width = 40

    $Day7NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day7NumericUpDown.Name = "Day7NumericUpDown"
    $Day7NumericUpDown.Height = 20
    $Day7NumericUpDown.Left = 73
    $Day7NumericUpDown.Maximum = 24
    $Day7NumericUpDown.Top = 305
    $Day7NumericUpDown.Width = 40
    
    $Day8NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day8NumericUpDown.Name = "Day8NumericUpDown"
    $Day8NumericUpDown.Height = 20
    $Day8NumericUpDown.Left = 177
    $Day8NumericUpDown.Maximum = 24
    $Day8NumericUpDown.Top = 135
    $Day8NumericUpDown.Width = 40

    $Day9NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day9NumericUpDown.Name = "Day9NumericUpDown"
    $Day9NumericUpDown.Height = 20
    $Day9NumericUpDown.Left = 177
    $Day9NumericUpDown.Maximum = 24
    $Day9NumericUpDown.Top = 166
    $Day9NumericUpDown.Width = 40

    $Day10NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day10NumericUpDown.Name = "Day10NumericUpDown"
    $Day10NumericUpDown.Height = 20
    $Day10NumericUpDown.Left = 177
    $Day10NumericUpDown.Maximum = 24
    $Day10NumericUpDown.Top = 194
    $Day10NumericUpDown.Width = 40

    $Day11NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day11NumericUpDown.Name = "Day11NumericUpDown"
    $Day11NumericUpDown.Height = 20
    $Day11NumericUpDown.Left = 177
    $Day11NumericUpDown.Maximum = 24
    $Day11NumericUpDown.Top = 221
    $Day11NumericUpDown.Width = 40

    $Day12NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day12NumericUpDown.Name = "Day12NumericUpDown"
    $Day12NumericUpDown.Height = 20
    $Day12NumericUpDown.Left = 177
    $Day12NumericUpDown.Maximum = 24
    $Day12NumericUpDown.Top = 250
    $Day12NumericUpDown.Width = 40

    $Day13NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day13NumericUpDown.Name = "Day13NumericUpDown"
    $Day13NumericUpDown.Height = 20
    $Day13NumericUpDown.Left = 177
    $Day13NumericUpDown.Maximum = 24
    $Day13NumericUpDown.Top = 277
    $Day13NumericUpDown.Width = 40

    $Day14NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day14NumericUpDown.Name = "Day14NumericUpDown"
    $Day14NumericUpDown.Height = 20
    $Day14NumericUpDown.Left = 177
    $Day14NumericUpDown.Maximum = 24
    $Day14NumericUpDown.Top = 305
    $Day14NumericUpDown.Width = 40

    $HoursWorkedLabel = New-Object System.Windows.Forms.Label
    $HoursWorkedLabel.Name = "HoursWorkedLabel"
    $HoursWorkedLabel.Height = 17
    $HoursWorkedLabel.Top = 335
    $HoursWorkedLabel.Width = 150
    $HoursWorkedLabel.Text = "Hours Per Pay Period: " + $Script:WorkHoursPerPayPeriod

    $UnusualHoursLabel = New-Object System.Windows.Forms.Label
    $UnusualHoursLabel.Name = "UnusualHoursLabel"
    $UnusualHoursLabel.ForeColor = "Blue"
    $UnusualHoursLabel.Height = 17
    $UnusualHoursLabel.Text = "Please validate unusual hours."
    $UnusualHoursLabel.Top = 355
    $UnusualHoursLabel.Visible = $False
    $UnusualHoursLabel.Width = 200

    $EmployeeInfoForm.Controls.AddRange(($EmployeeInfoOkButton, $EmployeeInfoCancelButton, $EmployeeTypePanel, $LeaveCeilingPanel, $InaugurationDayHolidayCheckBox, $Week1Label, $Week2Label, $SundayLabel, $MondayLabel, $TuesdayLabel, $WednesdayLabel, $ThursdayLabel, $FridayLabel, $SaturdayLabel, $Day1NumericUpDown, $Day2NumericUpDown, $Day3NumericUpDown, $Day4NumericUpDown, $Day5NumericUpDown, $Day6NumericUpDown, $Day7NumericUpDown, $Day8NumericUpDown, $Day9NumericUpDown, $Day10NumericUpDown, $Day11NumericUpDown, $Day12NumericUpDown, $Day13NumericUpDown, $Day14NumericUpDown, $HoursWorkedLabel, $UnusualHoursLabel))
    $EmployeeTypePanel.Controls.AddRange(($FullTimeRadioButton, $PartTimeRadioButton, $SESRadioButton, $EmploymentTypeLabel))
    $LeaveCeilingPanel.Controls.AddRange(($LeaveCeilingLabel, $CONUSRadioButton, $OCONUSRadioButton, $SESCeilingRadioButton))

    #Select/Check the appropriate things.

    if($Script:EmployeeType -eq "Full-Time")
    {
        $FullTimeRadioButton.Checked = $True
    }

    elseif($Script:EmployeeType -eq "Part-Time")
    {
        $PartTimeRadioButton.Checked = $True
    }

    elseif($Script:EmployeeType -eq "SES")
    {
        $SESRadioButton.Checked = $True
    }

    if($Script:LeaveCeiling -eq 240)
    {
        $CONUSRadioButton.Checked = $True
    }

    elseif($Script:LeaveCeiling -eq 360)
    {
        $OCONUSRadioButton.Checked = $True
    }

    elseif($Script:LeaveCeiling -eq 720)
    {
        $SESCeilingRadioButton.Checked = $True
    }

    if($Script:InaugurationHoliday -eq $True)
    {
        $InaugurationDayHolidayCheckBox.Checked = $True
    }

    $Day1NumericUpDown.Value = $Script:WorkSchedule["PayPeriodDay1"]
    $Day2NumericUpDown.Value = $Script:WorkSchedule["PayPeriodDay2"]
    $Day3NumericUpDown.Value = $Script:WorkSchedule["PayPeriodDay3"]
    $Day4NumericUpDown.Value = $Script:WorkSchedule["PayPeriodDay4"]
    $Day5NumericUpDown.Value = $Script:WorkSchedule["PayPeriodDay5"]
    $Day6NumericUpDown.Value = $Script:WorkSchedule["PayPeriodDay6"]
    $Day7NumericUpDown.Value = $Script:WorkSchedule["PayPeriodDay7"]

    $Day8NumericUpDown.Value = $Script:WorkSchedule["PayPeriodDay8"]
    $Day9NumericUpDown.Value = $Script:WorkSchedule["PayPeriodDay9"]
    $Day10NumericUpDown.Value = $Script:WorkSchedule["PayPeriodDay10"]
    $Day11NumericUpDown.Value = $Script:WorkSchedule["PayPeriodDay11"]
    $Day12NumericUpDown.Value = $Script:WorkSchedule["PayPeriodDay12"]
    $Day13NumericUpDown.Value = $Script:WorkSchedule["PayPeriodDay13"]
    $Day14NumericUpDown.Value = $Script:WorkSchedule["PayPeriodDay14"]
    
    $EmployeeInfoForm.Add_FormClosing({EmployeeInfoFormClosing -EventArguments $_})
    $EmployeeInfoOkButton.Add_Click({EmployeeInfoFormOkButtonClick})
    $EmployeeInfoCancelButton.Add_Click({EmployeeInfoFormCancelButtonClick})
    $FullTimeRadioButton.Add_Click({EmployeeInfoFormFullTimeRadioButtonClick; EmployeeInfoHoursChanged})
    $PartTimeRadioButton.Add_Click({EmployeeInfoFormPartTimeRadioButtonClick; EmployeeInfoHoursChanged})
    $SESRadioButton.Add_Click({EmployeeInfoFormSesRadioButtonClick; EmployeeInfoHoursChanged})

    $Day1NumericUpDown.Add_TextChanged({EmployeeInfoHoursChanged})
    $Day2NumericUpDown.Add_TextChanged({EmployeeInfoHoursChanged})
    $Day3NumericUpDown.Add_TextChanged({EmployeeInfoHoursChanged})
    $Day4NumericUpDown.Add_TextChanged({EmployeeInfoHoursChanged})
    $Day5NumericUpDown.Add_TextChanged({EmployeeInfoHoursChanged})
    $Day6NumericUpDown.Add_TextChanged({EmployeeInfoHoursChanged})
    $Day7NumericUpDown.Add_TextChanged({EmployeeInfoHoursChanged})
    $Day8NumericUpDown.Add_TextChanged({EmployeeInfoHoursChanged})
    $Day9NumericUpDown.Add_TextChanged({EmployeeInfoHoursChanged})
    $Day10NumericUpDown.Add_TextChanged({EmployeeInfoHoursChanged})
    $Day11NumericUpDown.Add_TextChanged({EmployeeInfoHoursChanged})
    $Day12NumericUpDown.Add_TextChanged({EmployeeInfoHoursChanged})
    $Day13NumericUpDown.Add_TextChanged({EmployeeInfoHoursChanged})
    $Day14NumericUpDown.Add_TextChanged({EmployeeInfoHoursChanged})
}

function BuildEditLeaveBalanceForm
{
    $SelectedLeaveBalance = $Script:LeaveBalances[$Script:MainForm.Controls["LeavePanel"].Controls["BalanceListBox"].SelectedIndex]
    
    $Script:EditLeaveBalanceForm          = New-Object System.Windows.Forms.Form
    $EditLeaveBalanceForm.Name            = "EditLeaveBalanceForm"
    $EditLeaveBalanceForm.BackColor       = "WhiteSmoke"
    $EditLeaveBalanceForm.Font            = $Script:FormFont
    $EditLeaveBalanceForm.FormBorderStyle = "FixedSingle"
    $EditLeaveBalanceForm.MaximizeBox     = $False
    $EditLeaveBalanceForm.Size            = New-Object System.Drawing.Size(283, 204)
    $EditLeaveBalanceForm.Text            = "Edit Leave"
    $EditLeaveBalanceForm.WindowState     = "Normal"

    $LeaveBalanceNameLabel = New-Object System.Windows.Forms.Label
    $LeaveBalanceNameLabel.Name = "LeaveBalanceNameLabel"
    $LeaveBalanceNameLabel.Height = 17
    $LeaveBalanceNameLabel.Left = 13
    $LeaveBalanceNameLabel.Text = "Type of Leave:"
    $LeaveBalanceNameLabel.Top = 14
    $LeaveBalanceNameLabel.Width = 79

    $LeaveBalanceNameTextBox = New-Object System.Windows.Forms.TextBox #Todo Might need to set a maximum length here.
    $LeaveBalanceNameTextBox.Name = "LeaveBalanceNameTextBox"
    $LeaveBalanceNameTextBox.Height = 20
    $LeaveBalanceNameTextBox.Left = 102
    $LeaveBalanceNameTextBox.Text = $SelectedLeaveBalance.Name
    $LeaveBalanceNameTextBox.Top = 11
    $LeaveBalanceNameTextBox.Width = 160

    $BalanceLabel = New-Object System.Windows.Forms.Label
    $BalanceLabel.Name = "BalanceLabel"
    $BalanceLabel.Height = 17
    $BalanceLabel.Left = 13
    $BalanceLabel.Text = "Balance:"
    $BalanceLabel.Top = 39
    $BalanceLabel.Width = 48

    $BalanceNumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $BalanceNumericUpDown.Name = "BalanceNumericUpDown"
    $BalanceNumericUpDown.Height = 20
    $BalanceNumericUpDown.Left = 102
    $BalanceNumericUpDown.Maximum = $Script:MaximumSick #Largest feasible amount for any type of leave, not just sick. Adjusted if type is Annual leave.
    $BalanceNumericUpDown.Top = 36
    $BalanceNumericUpDown.Value = $SelectedLeaveBalance.Balance
    $BalanceNumericUpDown.Width = 54

    $AlertThresholdLabel = New-Object System.Windows.Forms.Label
    $AlertThresholdLabel.Name = "AlertThresholdLabel"
    $AlertThresholdLabel.Height = 17
    $AlertThresholdLabel.Left = 13
    $AlertThresholdLabel.Text = "Alert Threshold:"
    $AlertThresholdLabel.Top = 65
    $AlertThresholdLabel.Width = 85

    $ThresholdNumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $ThresholdNumericUpDown.Name = "ThresholdNumericUpDown"
    $ThresholdNumericUpDown.Height = 20
    $ThresholdNumericUpDown.Left = 102
    $ThresholdNumericUpDown.Maximum = $Script:MaximumSick #Adjusted if type is Annual leave.
    $ThresholdNumericUpDown.Top = 65
    $ThresholdNumericUpDown.Width = 54

    $LeaveExpiresCheckBox = New-Object System.Windows.Forms.CheckBox
    $LeaveExpiresCheckBox.Name = "LeaveExpiresCheckBox"
    $LeaveExpiresCheckBox.Left = 13
    $LeaveExpiresCheckBox.Text = "Leave Expires"
    $LeaveExpiresCheckBox.Top = 57
    $LeaveExpiresCheckBox.Width = 104
    
    $LeaveExpiresOnLabel = New-Object System.Windows.Forms.Label
    $LeaveExpiresOnLabel.Name = "LeaveExpiresOnLabel"
    $LeaveExpiresOnLabel.Height = 17
    $LeaveExpiresOnLabel.Left = 13
    $LeaveExpiresOnLabel.Text = "Expires On:"
    $LeaveExpiresOnLabel.Top = 87
    $LeaveExpiresOnLabel.Visible = $False
    $LeaveExpiresOnLabel.Width = 63
    
    $LeaveExpiresOnDateTimePicker = New-Object System.Windows.Forms.DateTimePicker
    $LeaveExpiresOnDateTimePicker.Name = "LeaveExpiresOnDateTimePicker"
    $LeaveExpiresOnDateTimePicker.Enabled = $False
    $LeaveExpiresOnDateTimePicker.Format = "Short"
    $LeaveExpiresOnDateTimePicker.Left = 87
    $LeaveExpiresOnDateTimePicker.MinDate = $Script:BeginningOfPayPeriod
    $LeaveExpiresOnDateTimePicker.Top = 83
    $LeaveExpiresOnDateTimePicker.Visible = $False
    $LeaveExpiresOnDateTimePicker.Width = 115

    $WarningLabel = New-Object System.Windows.Forms.Label
    $WarningLabel.Name = "WarningLabel"
    $WarningLabel.AutoSize = $True
    $WarningLabel.ForeColor = "Red"
    $WarningLabel.Left = 13
    $WarningLabel.Top = 110
    $WarningLabel.Visible = $False

    $EditLeaveOkButton = New-Object System.Windows.Forms.Button
    $EditLeaveOkButton.Name = "EditLeaveOkButton"
    $EditLeaveOkButton.Height = 24
    $EditLeaveOkButton.Left = 46
    $EditLeaveOkButton.Text = "OK"
    $EditLeaveOkButton.Top = 138
    $EditLeaveOkButton.Width = 75
    
    $EditLeaveCancelButton = New-Object System.Windows.Forms.Button
    $EditLeaveCancelButton.Name = "EditLeaveCancelButton"
    $EditLeaveCancelButton.Height = 24
    $EditLeaveCancelButton.Left = 144
    $EditLeaveCancelButton.Text = "Cancel"
    $EditLeaveCancelButton.Top = 138
    $EditLeaveCancelButton.Width = 75
    
    $EditLeaveBalanceForm.Controls.AddRange(($LeaveBalanceNameLabel, $LeaveBalanceNameTextBox, $BalanceLabel, $BalanceNumericUpDown, $WarningLabel))

    if($SelectedLeaveBalance.Name -eq "Annual" -or
       $SelectedLeaveBalance.name -eq "Sick")
    {
        $EditLeaveBalanceForm.Controls.AddRange(($AlertThresholdLabel, $ThresholdNumericUpDown))

        if($SelectedLeaveBalance.Name -eq "Annual")
        {
            $BalanceNumericUpDown.Maximum   = $Script:MaximumAnnual
            $ThresholdNumericUpDown.Maximum = $Script:MaximumAnnual
        }

        $ThresholdNumericUpDown.Value = $SelectedLeaveBalance.Threshold
        
        $LeaveBalanceNameTextBox.ReadOnly = $True
    }

    else
    {
        $EditLeaveBalanceForm.Controls.AddRange(($LeaveExpiresCheckBox, $LeaveExpiresOnLabel, $LeaveExpiresOnDateTimePicker))

        if($SelectedLeaveBalance.ExpiresOn -lt $Script:CurrentDate)
        {
            $SelectedLeaveBalance.ExpiresOn = $Script:CurrentDate
        }

        $LeaveExpiresOnDateTimePicker.Value = $SelectedLeaveBalance.ExpiresOn.Date

        if($SelectedLeaveBalance.Expires -eq $True)
        {
            $LeaveExpiresCheckBox.Checked = $True
            
            $LeaveExpiresOnLabel.Visible = $True

            $LeaveExpiresOnDateTimePicker.Enabled = $True
            $LeaveExpiresOnDateTimePicker.Visible = $True
        }
    }

    $EditLeaveBalanceForm.Controls.AddRange(($EditLeaveOkButton, $EditLeaveCancelButton))

    if($LeaveBalanceNameTextBox.Text -eq "Annual" -or
       $LeaveBalanceNameTextBox.Text -eq "Sick")
    {
        $LeaveBalanceNameTextBox.TabStop = $False
    }

    if($LeaveBalanceNameTextBox.Text -ne "")
    {
        $EditLeaveBalanceForm.ActiveControl = $BalanceNumericUpDown
    }
    
    $EditLeaveBalanceForm.Add_FormClosing({EditLeaveBalanceFormClosing -EventArguments $_})
    $LeaveBalanceNameTextBox.Add_TextChanged({EditLeaveBalanceFormNameOrDateChanged})
    $LeaveExpiresCheckBox.Add_Click({EditLeaveBalanceFormLeaveExpiresCheckBoxClick; EditLeaveBalanceFormNameOrDateChanged})
    $LeaveExpiresOnDateTimePicker.Add_Valuechanged({EditLeaveBalanceFormNameOrDateChanged})
    $EditLeaveOkButton.Add_Click({EditLeaveBalanceFormOkButtonClick})
    $EditLeaveCancelButton.Add_Click({EditLeaveBalanceFormCancelButtonClick})
}

function BuildEditProjectedLeaveForm
{
    $SelectedLeaveDetails = $Script:ProjectedLeave[$Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"].SelectedIndex]
    $EndOfPayPeriod       = GetEndingOfPayPeriodForDate -Date $SelectedLeaveDetails.StartDate

    $Script:EditProjectedLeaveForm          = New-Object System.Windows.Forms.Form
    $EditProjectedLeaveForm.Name            = "EditProjectedLeaveForm"
    $EditProjectedLeaveForm.BackColor       = "WhiteSmoke"
    $EditProjectedLeaveForm.Font            = $Script:FormFont
    $EditProjectedLeaveForm.FormBorderStyle = "FixedSingle"
    $EditProjectedLeaveForm.MaximizeBox     = $False
    $EditProjectedLeaveForm.Size            = New-Object System.Drawing.Size(455, 470)
    $EditProjectedLeaveForm.Text            = "Edit Projected Leave"
    $EditProjectedLeaveForm.WindowState     = "Normal"

    $LeaveStartLabel = New-Object System.Windows.Forms.Label
    $LeaveStartLabel.Name = "LeaveStartLabel"
    $LeaveStartLabel.Height = 17
    $LeaveStartLabel.Left = 15
    $LeaveStartLabel.Text = "Leave Start:"
    $LeaveStartLabel.Top = 41
    $LeaveStartLabel.Width = 65

    $LeaveEndLabel = New-Object System.Windows.Forms.Label
    $LeaveEndLabel.Name = "LeaveEndLabel"
    $LeaveEndLabel.Height = 17
    $LeaveEndLabel.Left = 15
    $LeaveEndLabel.Text = "Leave End:"
    $LeaveEndLabel.Top = 78
    $LeaveEndLabel.Width = 61

    $LeaveStartDateTimePicker = New-Object System.Windows.Forms.DateTimePicker
    $LeaveStartDateTimePicker.Name = "LeaveStartDateTimePicker"
    $LeaveStartDateTimePicker.Format = "Short"
    $LeaveStartDateTimePicker.Left = 90
    $LeaveStartDateTimePicker.Top = 38
    $LeaveStartDateTimePicker.Width = 115

    $LeaveEndDateTimePicker = New-Object System.Windows.Forms.DateTimePicker
    $LeaveEndDateTimePicker.Name = "LeaveEndDateTimePicker"
    $LeaveEndDateTimePicker.Format = "Short"
    $LeaveEndDateTimePicker.Left = 90
    $LeaveEndDateTimePicker.Top = 76
    $LeaveEndDateTimePicker.Width = 115

    $LeaveBankLabel = New-Object System.Windows.Forms.Label
    $LeaveBankLabel.Name = "LeaveBankLabel"
    $LeaveBankLabel.Height = 17
    $LeaveBankLabel.Left = 15
    $LeaveBankLabel.Text = "Leave Bank:"
    $LeaveBankLabel.Top = 15
    $LeaveBankLabel.Width = 67

    $LeaveBankComboBox = New-Object System.Windows.Forms.ComboBox
    $LeaveBankComboBox.Name = "LeaveBankComboBox"
    $LeaveBankComboBox.DropDownStyle = "DropDownList"
    $LeaveBankComboBox.Tag = @{}
    $LeaveBankComboBox.Top = 12
    $LeaveBankComboBox.Left = 90
    $LeaveBankComboBox.Width = 115

    $LeaveDatesFlowLayoutPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $LeaveDatesFlowLayoutPanel.Name = "LeaveDatesFlowLayoutPanel"
    $LeaveDatesFlowLayoutPanel.AutoScroll = $True
    $LeaveDatesFlowLayoutPanel.FlowDirection = "TopDown"
    $LeaveDatesFlowLayoutPanel.Height = 250
    $LeaveDatesFlowLayoutPanel.Left = 15
    $LeaveDatesFlowLayoutPanel.Top = 110
    $LeaveDatesFlowLayoutPanel.Width = 420
    $LeaveDatesFlowLayoutPanel.WrapContents = $False

    $HoursOfLeaveTaken = New-Object System.Windows.Forms.Label
    $HoursOfLeaveTaken.Name = "HoursOfLeaveTaken"
    $HoursOfLeaveTaken.Left = 10
    $HoursOfLeaveTaken.Height = 17
    $HoursOfLeaveTaken.Top = 370
    $HoursOfLeaveTaken.Width = 150
    
    $EditProjectedOkButton = New-Object System.Windows.Forms.Button
    $EditProjectedOkButton.Name = "EditProjectedOkButton"
    $EditProjectedOkButton.Height = 24
    $EditProjectedOkButton.Left = 43
    $EditProjectedOkButton.Text = "OK"
    $EditProjectedOkButton.Top = 394
    $EditProjectedOkButton.Width = 75
    
    $EditProjectedCancelButton = New-Object System.Windows.Forms.Button
    $EditProjectedCancelButton.Name = "EditProjectedCancelButton"
    $EditProjectedCancelButton.Height = 24
    $EditProjectedCancelButton.Left = 157
    $EditProjectedCancelButton.Text = "Cancel"
    $EditProjectedCancelButton.Top = 394
    $EditProjectedCancelButton.Width = 75

    $EditProjectedLeaveForm.Controls.AddRange(($LeaveBankLabel, $LeaveStartLabel, $LeaveEndLabel, $LeaveBankComboBox, $LeaveStartDateTimePicker, $LeaveEndDateTimePicker, $LeaveDatesFlowLayoutPanel, $HoursOfLeaveTaken, $EditProjectedOkButton, $EditProjectedCancelButton))

    $TotalHours = 0
    
    foreach($Day in $SelectedLeaveDetails.HoursHashTable.Keys)
    {
        $TotalHours += $SelectedLeaveDetails.HoursHashTable[$Day]
    }

    $HoursOfLeaveTaken.Text = "Hours of Leave Taken: " + $TotalHours
    
    foreach($LeaveItem in $Script:LeaveBalances)
    {
        if($LeaveBankComboBox.Items.Contains($LeaveItem.Name) -eq $False)
        {
            $LeaveBankComboBox.Items.Add($LeaveItem.Name)

            if($LeaveItem.Name -eq "Annual" -or
               $LeaveItem.Name -eq "Sick")
            {
                $LeaveBankComboBox.Tag.($LeaveItem.Name) = $Script:LastSelectableDate
            }
        }

        #In the hashtable in the tag, calculate the last date that the leave is valid.
        if($LeaveItem.Name -ne "Annual" -and
           $LeaveItem.Name -ne "Sick")
        {
            if($LeaveItem.Expires -eq $True)
            {
                if($LeaveItem.ExpiresOn -gt $LeaveBankComboBox.Tag.($LeaveItem.Name))
                {
                    $LeaveBankComboBox.Tag.($LeaveItem.Name) = $LeaveItem.ExpiresOn
                }
            }

            else
            {
                $LeaveBankComboBox.Tag.($LeaveItem.Name) = $Script:LastSelectableDate
            }
        }
    }

    $LeaveBankComboBox.SelectedIndex = $LeaveBankComboBox.Items.IndexOf($SelectedLeaveDetails.LeaveBank)

    $LeaveStartDateTimePicker.Value = $SelectedLeaveDetails.StartDate
    $LeaveEndDateTimePicker.Value = $SelectedLeaveDetails.EndDate

    $LeaveStartDateTimePicker.MaxDate = $LeaveBankComboBox.Tag.($SelectedLeaveDetails.LeaveBank)
    $LeaveStartDateTimePicker.MinDate = $Script:BeginningOfPayPeriod

    $LeaveEndDateTimePicker.MinDate = $LeaveStartDateTimePicker.Value

    if($EndOfPayPeriod -lt $LeaveBankComboBox.Tag.($SelectedLeaveDetails.LeaveBank))
    {
        $LeaveEndDateTimePicker.MaxDate = $EndOfPayPeriod
    }

    else
    {
        $LeaveEndDateTimePicker.MaxDate = $LeaveBankComboBox.Tag.($SelectedLeaveDetails.LeaveBank)
    }

    if($Script:UnsavedProjectedLeave -eq $False)
    {
        $EditProjectedLeaveForm.ActiveControl = $EditProjectedOkButton
    }

    else
    {
        $EditProjectedLeaveForm.ActiveControl = $LeaveBankComboBox
    }

    $LeaveStartDateTimePicker.Tag = $LeaveStartDateTimePicker.Value
    $LeaveEndDateTimePicker.Tag   = $LeaveEndDateTimePicker.Value

    PopulateProjectedLeaveDays

    $LeaveStartDateTimePicker.Add_CloseUp({StartDateCalendarClosedUp})
    $LeaveEndDateTimePicker.Add_CloseUp({EndDateCalendarClosedUp})
    $EditProjectedLeaveForm.Add_FormClosing({EditProjectedLeaveFormClosing -EventArguments $_})
    $EditProjectedOkButton.Add_Click({EditProjectedLeaveOkButton})
    $EditProjectedCancelButton.Add_Click({EditProjectedLeaveCancelButton})
    $LeaveBankComboBox.Add_SelectedIndexChanged({EditProjectedLeaveBankSelectedIndexChanged})
    $LeaveStartDateTimePicker.Add_KeyDown({$_.Handled = $True; $_.SuppressKeyPress = $True})
    $LeaveEndDateTimePicker.Add_KeyDown({$_.Handled = $True; $_.SuppressKeyPress = $True})
}

function BuildOutputForm
{
    $Script:OutputForm          = New-Object System.Windows.Forms.Form
    $OutputForm.Name            = "OutputForm"
    $OutputForm.BackColor       = "WhiteSmoke"
    $OutputForm.Font            = $Script:FormFont
    $OutputForm.FormBorderStyle = "FixedSingle"
    $OutputForm.MaximizeBox     = $False
    $OutputForm.Size            = New-Object System.Drawing.Size(450, 643)
    $OutputForm.Text            = "Projection Results"
    $OutputForm.WindowState     = "Normal"

    $OutputRichTextBox = New-Object System.Windows.Forms.RichTextBox
    $OutputRichTextBox.Name = "OutputRichTextBox"
    $OutputRichTextBox.Dock = "Top"
    $OutputRichTextBox.Height = 540
    $OutputRichTextBox.Multiline = $True
    $OutputRichTextBox.ReadOnly = $True
    $OutputRichTextBox.TabStop = $False
    
    $CopyButton = New-Object System.Windows.Forms.Button
    $CopyButton.Name = "CopyButton"
    $CopyButton.Height = 24
    $CopyButton.Left = 57
    $CopyButton.Text = "Copy"
    $CopyButton.Top = 544
    $CopyButton.Width = 336
    
    $CloseButton = New-Object System.Windows.Forms.Button
    $CloseButton.Name = "CloseButton"
    $CloseButton.Height = 24
    $CloseButton.Left = 57
    $CloseButton.Text = "Close"
    $CloseButton.Top = 577
    $CloseButton.Width = 336
    
    $OutputForm.Controls.AddRange(($OutputRichTextBox, $CopyButton, $CalculateWhenToStartLeaveButton, $CloseButton))

    $OutputForm.ActiveControl = $CloseButton
    
    PopulateOutputFormRichTextBox

    $CopyButton.Add_Click({OutputFormCopyButton})
    $CloseButton.Add_Click({OutputFormCloseButton})
}

function ShowMessageBox
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$True)]  $Text,
        [Parameter(Mandatory=$False)] $Caption,
        [Parameter(Mandatory=$False)] $Buttons,
        [Parameter(Mandatory=$False)] $Icon,
        [Parameter(Mandatory=$False)] $DefaultButton
    )

    <#
    Button Options:
    and Responses (including what X returns)
        OK                - OK                   (X)
        OKCancel          - OK, Cancel           (X)
        AbortRetryIgnore  - Abort, Retry, Ignore (X disabled)
        YesNoCancel       - Yes, No, Cancel      (X)
        YesNo             - Yes, No              (X disabled)
        RetryCancel       - Retry, Cancel        (X)

    Icon Options:
        None
        Error       - Causes an audible ding from Windows
        Question
        Exclamation - Causes an audible ding from Windows (Different than Error ding)
        Information - Causes an audible ding from Windows (Same ding as Exclamation ding)

    DefaultButton Options:
        Button1
        Button2
        Button3
    #>

    $Response = ""

    if(($Buttons -ne $Null) -and ($Icon -ne $Null) -and ($DefaultButton -ne $Null))
    {
        $Buttons       = [System.Windows.Forms.MessageBoxButtons]::$Buttons
        $Icon          = [System.Windows.Forms.MessageBoxIcon]::$Icon
        $DefaultButton = [System.Windows.Forms.MessageBoxDefaultButton]::$DefaultButton

        $Response = [System.Windows.Forms.MessageBox]::Show($Text, $Caption, $Buttons, $Icon, $DefaultButton)
    }

    elseif(($Buttons -ne $Null) -and ($Icon -ne $Null))
    {
        $Buttons = [System.Windows.Forms.MessageBoxButtons]::$Buttons
        $Icon    = [System.Windows.Forms.MessageBoxIcon]::$Icon

        $Response = [System.Windows.Forms.MessageBox]::Show($Text, $Caption, $Buttons, $Icon)
    }

    elseif($Buttons -ne $Null)
    {
        $Buttons = [System.Windows.Forms.MessageBoxButtons]::$Buttons

        $Response = [System.Windows.Forms.MessageBox]::Show($Text, $Caption, $Buttons)
    }

    else
    {
        $Response = [System.Windows.Forms.MessageBox]::Show($Text, $Caption)
    }

    return $Response
}

#endregion Form Building Functions

#clear #Todo uncomment this

Main