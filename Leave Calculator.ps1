#Line I have commented through (need to finish): 1523

$Script:Author      = "Mikepedia"
#SCRIPT NAME:         Federal Civilian Leave Calculator
$Script:DateUpdated = "17 June 2024 15:37 CDT"
$Script:Version     = "beta 1.0"
#SCRIPT PURPOSE:      To assist users with managing and projecting future leave.

#region Change Log
<#
    Version beta 1.0: Initial Release.
#>
#endregion Change Log

#Need these for the GUI.
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#Enable nicer looking visual styles, especially on the DateTimePickers.
[System.Windows.Forms.Application]::EnableVisualStyles()

#region Script variables

$Script:FormFont     = New-Object System.Drawing.Font("Times New Roman", 10)
$Script:IconsFont    = New-Object System.Drawing.Font("Segoe MDL2 Assets", 10, [System.Drawing.FontStyle]::Bold)
$Script:HelpIconFont = New-Object System.Drawing.Font("Segoe MDL2 Assets", 18, [System.Drawing.FontStyle]::Bold)

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

$Script:MaximumAnnual = 936  #With SES ceiling of 720 and SES accrual and rare year with 27 pay periods.
$Script:MaximumSick   = 9999 #Assuming you work full time, every year has 27 pay periods, and you never use a single hour, it would take almost 93 years to reach this amount.

$Script:LeaveBalances  = New-Object System.Collections.Generic.List[PSCustomObject]
$Script:ProjectedLeave = New-Object System.Collections.Generic.List[PSCustomObject]

$Script:RepoWebsite              = "https://github.com/mikedelliot/FedCivLeaveCalculator/blob/main/Leave%20Calculator.ps1"
$Script:ConfigFile               = "$env:APPDATA\PowerShell Scripts\Federal Civilian Leave Calculator\Federal Civilian Leave Calculator.ini"
$Script:HolidayWebsite           = "https://www.opm.gov/policy-data-oversight/pay-leave/federal-holidays/"
$Script:HolidaysHashTable        = @{} #Where the dates and names of holidays are contained except for Inauguration Day.
$Script:InaugurationDayHashTable = @{} #Same as above, but only for Inauguration Day.

$Script:UnsavedProjectedLeave = $False #A way to keep track if the Projected Leave is new or not, which controls the Cancel button behavior.
$Script:DrawingListBox        = $False #This is so when drawing the CheckedListBox and checking the appropriate boxes that we don't fire events unnecessarily.

#endregion Script variables

#region Default Settings -- Changing values here only affects the first launch of the program. After that, it loads from a config file. To make changes, launch the script and change the settings there.

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
    Name      = "Annual"
    Balance   = 0
    Threshold = 0
    Static    = $True
}

$SickLeaveCustomObject = [PSCustomObject] @{
    Name      = "Sick"
    Balance   = 0
    Threshold = 0
    Static    = $True
}

$Script:LeaveBalances.Add($AnnualLeaveCustomObject)
$Script:LeaveBalances.Add($SickLeaveCustomObject)

#endregion Default Settings

#region Functions

#This function is called once as the last line of the script. This is what kicks off the program.
function Main
{
    $Script:BeginningOfPayPeriod = GetBeginningOfPayPeriodForDate -Date $Script:CurrentDate
    $Script:LastSelectableDate   = GetLeaveYearEndForDate -Date ((Get-Date -Year ($Script:BeginningOfPayPeriod.Year + 2) -Month $Script:BeginningOfPayPeriod.Month -Day $Script:BeginningOfPayPeriod.Day)).Date #Get the last day of the leave year for whatever year is the current year + 2.
    GetOpmHolidaysForYears                          #Get the OPM holidays for the year range of beginning of current pay period year to last selectable date year.
    LoadConfig                                      #Load saved data.
    SetWorkHoursPerPayPeriod                        #Calculate how many hours per pay period based on saved data. This is so leave accruals can be correct if part-time.
    GetAccrualRateDateChange                        #Get the date that the accrual rates change.
    UpdateExistingBalancesAndProjectedLeaveAtLaunch #Update projected leave and balances to the current date based on what was entered.
    GetGitHubVersion                                #Check for program updates.

    clear

    BuildMainForm #Create the GUI.
    
    [System.Windows.Forms.Application]::Run($Script:MainForm) #Display the GUI.
}

#This function adds 3 years and 15 years to the SCD date, finds the end of the pay period it happens in if not the beginning of a pay period, and then adds one day to get the start of the next pay period.
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

#Returns how many hours of annual leave will be accrued based on the pay period argument.
function GetAnnualLeaveAccrualHours
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)][DateTime] $PayPeriod
    )
    
    $AccruedHours = 0

    if($Script:EmployeeType -eq "SES") #SES always gets 8 hours, no matter SCD date.
    {
        $AccruedHours = 8
    }

    else
    {
        if($Script:EmployeeType -eq "Full-Time") #Full-time gets a specific amount every pay period, no calculations needed.
        {
            if($PayPeriod -ge $Script:FifteenYearMark)
            {
                $AccruedHours = 8
            }

            elseif($PayPeriod -ge $Script:ThreeYearMark) #6 hours per pay period except last pay period of year, which is 10 hours.
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

        else #Part-Time accured hours are based on hours worked per pay period, so a calculation is needed.
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

#Returns the beginning of a pay period for the date argument.
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

#Returns the last day of a pay period for the date argument.
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

#Returns how many hours are worked on the day provided in the argument based on the input work schedule and if it's a holiday or not.
function GetHoursForWorkDay
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)][DateTime] $Day
    )
    
    $Hours = 0
    
    if($Script:HolidaysHashTable.ContainsKey($Day.ToString("MM/dd/yyyy")) -eq $False) #Only continue if not a federal holiday.
    {
        if($Script:InaugurationHoliday -eq $False -or
           $Script:InaugurationDayHashTable.ContainsKey($Day.ToString("MM/dd/yyyy")) -eq $False) #Only continue if not entitled to a holiday on Inauguration Day or the day argument isn't an Inauguration Day.
        {
            $DayOfPayPeriod = (($Day - $Script:BeginningOfPayPeriod).Days % 14) + 1 #Figure out what day of the pay period it is. For example, the first Sunday in a pay period is day 1, the Saturday is day 7.

            $Hours = $Script:WorkSchedule.("PayPeriodDay" + $DayOfPayPeriod) #Access the hash table of the work schedule to get the number of hours worked that day.
        }
    }
    
    return $Hours
}

#Returns the leave year end date for a calendar year.
function GetLeaveYearEndForDate
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)] $Date
    )

    $Year = (GetBeginningOfPayPeriodForDate -Date $Date).Year #Get the year that the current pay period started on.
    
    return (GetEndingOfPayPeriodForDate -Date (Get-Date -Year $Year -Month 12 -Day 31).Date) #Get the ending of the pay period for the date on 31 Dec of the calculated year.
}

#Returns the last day to schedule annual leave to be eligible to have it restored if you forfeit it for specific reasons.
function GetLeaveYearScheduleDeadline
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)] $Date
    )

    return $Date.AddDays(-42) #42 Days is the day before the start of the third biweekly pay period prior to the end of the leave year.
}

#Checks for updates to this script by reading the HTML of the GitHub page and using RegEx to check the version number listed at the beginning of the script text. Done this way so you don't need any GitHub dependencies or packages installed.
function GetGitHubVersion
{
    try
    {
        Write-Host "Checking version against online repository."
        
        $GitHubContent = (Invoke-WebRequest -DisableKeepAlive -Uri $Script:RepoWebsite -UserAgent "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36").Content #Needs a custom user agent so GitHub thinks you're accessing it from a web browser, not command line. Command line requires using the API and needs special authentication.
        
        #RegEx Matches
        #Literal text              Script:Version
        #Matches as many spaces as possible       +=
        #Matches literal text and escaped chars      \\"
        #Capture group                                  (   )
        #Matches char as many times (not greedy)         .*?
        #Matches literal text and escaped chars              \\"\\r
        if(($GitHubContent -match 'Script:Version += \\"(.*?)\\"\\r') -eq $True)
        {
            if($Matches[1] -ne $Script:Version) #If the version found online doesn't match the version of this script.
            {
                $NewVersion = $Matches[1]
                
                $Response = ShowMessageBox -Text "There is a new version available. You are running version $Script:Version and the newest version is $NewVersion. Would you like to open your web browser so you can download the newest version?" -Caption "Updates Available" -Buttons "YesNo" -Icon "Exclamation"

                if($Response -eq "Yes")
                {
                    Start-Process $Script:RepoWebsite #Open the website.
                }
            }
        }
    }

    catch
    {
        $Response = ShowMessageBox -Text "There was an error checking for updates. You are running version $Script:Version. Would you like to open your web browser to manually check if there is an update?" -Caption "Unable to Check for Updates" -Buttons "YesNo" -Icon "Error"

        if($Response -eq "Yes")
        {
            Start-Process $Script:RepoWebsite #Open the website.
        }
    }
}

#Populates the two holiday hash tables with the dates and names of the holidays.
function GetOpmHolidaysForYears
{
    $BeginningYear = ($Script:BeginningOfPayPeriod).Year
    $EndingYear    = $Script:LastSelectableDate.Year

    $Failed = $False
    
    try
    {
        Write-Host "Getting holidays from OPM website."
        
        $HolidayContent = (Invoke-WebRequest -DisableKeepAlive -Uri $Script:HolidayWebsite).Content
    }

    catch
    {
        $Failed = $True
    }

    for($TargetYear = $BeginningYear; $TargetYear -le $EndingYear; $TargetYear++) #Loop through the years.
    {
        $YearBeginningIndex = $Null
        $YearEndingIndex    = $Null
        $YearContent        = $Null

        try
        {
            #This if else statement is doing the HTML parsing of the website content. It's just looking for common strings in the website design.

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

                if($YearEndingIndex -eq -1) #This is triggered on the oldest year since it doesn't contain a "Back to top" link.
                {
                    $YearEndingIndex = $HolidayContent.LastIndexOf("</p>") + 4 #The + 4 is so we get the "</p>" so every holiday content we get should be identical for further processing.
                }

                $YearContent = $HolidayContent.Substring($YearBeginningIndex, ($YearEndingIndex - $YearBeginningIndex)) #Get just the HTML content for the year we are currently on in the loop.
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
                if($HolidayList.IndexOf($Line) % 2 -eq 0) #So we only get the actual dates, not which holiday it is (we want the date as the hashtable key and the holiday name as the value)
                {
                    $ModifiedLine = $Line.ToString().Replace("<td>", "").Replace("</td>", "").Trim() #Get rid of the additional HTML tags

                    if($ModifiedLine.Contains("<") -eq $True)
                    {
                        $ModifiedLine = $ModifiedLine.Substring(0, ($ModifiedLine.IndexOf("<"))).Trim() #If there are any notes on the date marked with an *, clear them out
                    }

                    $ModifiedLine = $ModifiedLine.Substring($ModifiedLine.IndexOf(" ") + 1) #Gets rid of the day of the week in front. Don't need it for parsing the date.

                    if(($HolidayList.IndexOf($Line) -eq 0) -and ($Line.ToString().Contains("December"))) #Sometimes OPM includes New Years in the previous year because it falls on a Saturday, so the holiday is given on a Friday.
                    {
                        $PreviousYear = $TargetYear - 1
                    
                        $ModifiedLine += ", $PreviousYear" #Add the previous year at the end
                    }

                    else
                    {
                        $ModifiedLine += ", $TargetYear" #Add the year at the end
                    }

                    if($ModifiedLine -match "\w* \d{2}, \d{4}, \d{4}") #Some of the entries rarely have the year already listed in the date, so we need to strip that off, or else it'll have two years to try to parse.
                    {
                        $ModifiedLine = $ModifiedLine.Substring(0, $ModifiedLine.Length - 6) #Trim off the year, space, and comma.
                    }

                    $DateObject = Get-Date $ModifiedLine #Parse the date.

                    $DateString = $DateObject.ToString("MM/dd/yyyy") #Put it in the format the rest of the script uses.

                    if($HolidaysHashTable.Contains($DateString) -eq $False) #To prevent adding duplicate dates.
                    {
                        $HolidayNameString = $HolidayList[$HolidayList.IndexOf($Line) + 1].ToString().Replace("<td>", "").Replace("</td>", "").Trim() #NOW we get the holiday name.

                        if($HolidayNameString.ToLower().Contains("inauguration") -eq $True) #Separate out the Inauguration Holidays.
                        {
                            $InaugurationDayHashTable[$DateString] = $HolidayNameString #Add to the Inaugruation Day hashtable.
                        }

                        else
                        {
                            $HolidaysHashTable[$DateString] = $HolidayNameString #Add to the normal holiday hashtable.
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

#Returns the number of sick leave hours to be accrued per pay period.
function GetSickLeaveAccrualHours
{
    $AccruedHours = 4

    if($Script:EmployeeType -eq "Part-Time") #If you're part-time it's based on hours worked, otherwise it's just 4.
    {
        $AccruedHours = $Script:WorkHoursPerPayPeriod / 20
    }

    return $AccruedHours
}

#Sets the script variable of how many hours are worked per pay period based on what the user enters in settings.
function SetWorkHoursPerPayPeriod
{
    $Script:WorkHoursPerPayPeriod = 0

    foreach($Day in $Script:WorkSchedule.Keys)
    {
        $Script:WorkHoursPerPayPeriod += $Script:WorkSchedule[$Day]
    }
}

#This function loads the settings from the config file and validates them.
function LoadConfig
{
    if((Test-Path -Path $Script:ConfigFile) -eq $True) #First make sure the file exists.
    {
        $Errors = $False
            
        try
        {
            $LoadedConfig = Import-Clixml -Path $Script:ConfigFile

            $LeaveNameHashTable = @{} #Create a hash table for the names of leave so we can use them for the leave banks of the projected leave.

            if($LoadedConfig.GetType().Name -eq "Object[]") #Check that the file is an array.
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
                          ($LoadedConfig[6][$Key] -isnot [Int32] -and
                           $LoadedConfig[6][$Key] -isnot [Decimal]) -and
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

                if(($LoadedConfig[9] -is [Int32] -or
                    $LoadedConfig[9] -is [Decimal]) -and
                    $LoadedConfig[9] -ge 0 -and
                    $LoadedConfig[9] -le $Script:MaximumAnnual) #Load AnnualGoal Int32
                {
                    $Script:AnnualGoal = [Int32]$LoadedConfig[9]
                }

                else
                {
                    $Errors = $True
                }

                if(($LoadedConfig[10] -is [Int32] -or
                    $LoadedConfig[10] -is [Decimal]) -and
                    $LoadedConfig[10] -ge 0 -and
                    $LoadedConfig[10] -le $Script:MaximumSick) #Load SickGoal Int32
                {
                    $Script:SickGoal = [Int32]$LoadedConfig[10]
                }

                else
                {
                    $Errors = $True
                }

                if(($LoadedConfig[11] -is [Double] -or
                    $LoadedConfig[11] -is [Decimal]) -and
                   $LoadedConfig[11] -ge 0 -and
                   $LoadedConfig[11] -lt 1) #Load AnnualDecimal Double
                {
                    $Script:AnnualDecimal = $LoadedConfig[11]
                }

                else
                {
                    $Errors = $True
                }

                if(($LoadedConfig[12] -is [Double] -or
                    $LoadedConfig[12] -is [Decimal]) -and
                   $LoadedConfig[12] -ge 0 -and
                   $LoadedConfig[12] -lt 1) #Load SickDecimal Double
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
                        if(($LoadedConfig[$Index].Balance -is [Decimal] -or
                            $LoadedConfig[$Index].Balance -is [Int32]) -and
                            $LoadedConfig[$Index].Balance -ge 0 -and
                            $LoadedConfig[$Index].Balance -le $Script:MaximumAnnual)
                        {
                            $Script:LeaveBalances[0].Balance = [Int32]$LoadedConfig[$Index].Balance
                        }

                        else
                        {
                            $Errors = $True
                        }

                        if(($LoadedConfig[$Index].Threshold -is [Decimal] -or
                            $LoadedConfig[$Index].Threshold -is [Int32]) -and
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
                        if(($LoadedConfig[$Index].Balance -is [Decimal] -or
                            $LoadedConfig[$Index].Balance -is [Int32]) -and
                            $LoadedConfig[$Index].Balance -ge 0 -and
                            $LoadedConfig[$Index].Balance -le $Script:MaximumSick)
                        {
                            $Script:LeaveBalances[1].Balance = [Int32]$LoadedConfig[$Index].Balance
                        }

                        else
                        {
                            $Errors = $True
                        }

                        if(($LoadedConfig[$Index].Threshold -is [Decimal] -or
                            $LoadedConfig[$Index].Threshold -is [Int32]) -and
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
                        
                        if([System.Windows.Forms.TextRenderer]::MeasureText($LoadedConfig[$Index].Name.Trim(), $Script:FormFont).Width -le 85 -and
                           $LoadedConfig[$Index].Name.Trim() -match "[^A-Za-z0-9 ]" -eq $False -and
                           $LoadedConfig[$Index].Name.ToLower().Trim() -ne "annual" -and
                           $LoadedConfig[$Index].Name.ToLower().Trim() -ne "sick" -and
                          ($LoadedConfig[$Index].Balance -is [Decimal] -or
                           $LoadedConfig[$Index].Balance -is [Int32]) -and
                           $LoadedConfig[$Index].Balance -ge 0 -and
                           $LoadedConfig[$Index].Balance -le $Script:MaximumSick -and
                           $LoadedConfig[$Index].Expires -is [Boolean] -and
                           $LoadedConfig[$Index].ExpiresOn -is [DateTime])
                        {
                            $NewLeaveBalance = [PSCustomObject] @{
                                Name      = $LoadedConfig[$Index].Name.Trim()
                                Balance   = [Int32]$LoadedConfig[$Index].Balance
                                Expires   = $LoadedConfig[$Index].Expires
                                ExpiresOn = $LoadedConfig[$Index].ExpiresOn.Date
                                Static    = $False
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
                       $LoadedConfig[$Index].EndDate -ge $LoadedConfig[$Index].StartDate -and
                       $LoadedConfig[$Index].Included -is [Boolean])
                    {
                        $NewProjectedLeave = [PSCustomObject] @{
                            LeaveBank      = $LoadedConfig[$Index].LeaveBank.Trim()
                            StartDate      = $LoadedConfig[$Index].StartDate.Date
                            EndDate        = $LoadedConfig[$Index].EndDate.Date
                            HoursHashTable = @{}
                            Included       = $LoadedConfig[$Index].Included
                        }

                        #Populate the HoursHashTable.
                        for($Date = $NewProjectedLeave.StartDate; $Date -le $NewProjectedLeave.EndDate; $Date = $Date.AddDays(1))
                        {
                            if($LoadedConfig[$Index].HoursHashTable.ContainsKey($Date.ToString("MM/dd/yyyy")) -eq $True -and
                               $LoadedConfig[$Index].HoursHashTable[$Date.ToString("MM/dd/yyyy")] -ge 0 -and
                               $LoadedConfig[$Index].HoursHashTable[$Date.ToString("MM/dd/yyyy")] -le 24) #If the hashtable value exists and is valid.
                            {
                                $NewProjectedLeave.HoursHashTable[$Date.ToString("MM/dd/yyyy")] = [Int32]$LoadedConfig[$Index].HoursHashTable[$Date.ToString("MM/dd/yyyy")]
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

                #Sort the projected leave to prevent any funny business. Only sort if there is more than one entry.
                if($Script:ProjectedLeave.Count -gt 1)
                {
                    $Script:ProjectedLeave = [System.Collections.Generic.List[PSCustomObject]] ($Script:ProjectedLeave | Sort-Object -Property "StartDate", "EndDate", "LeaveBank")
                }
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

#Give it a string and a number. Returns the string with an S if the number provided isn't 1. For example, give it "day" and "2" and it'll return "days".
function NumberGetsLetterS($String, $Number)
{
    if($Number -ne 1)
    {
        $String += "s"
    }

    return $String
}

#This is a funcion that appends text to a RichTextBox based on the arguments passed to it.
function RichTextBoxAppendText
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)]  $RichTextBox,
        [parameter(Mandatory=$True)]  $Text,
        [parameter(Mandatory=$False)] $Alignment,
        [parameter(Mandatory=$False)] $Color,
        [parameter(Mandatory=$False)] $FontSize,
        [parameter(Mandatory=$False)] $FontStyle
    )

    $RichTextBox.SelectionStart  = $RichTextBox.TextLength #Put the cursor at the end of the current text.
    $RichTextBox.SelectionLength = 0 #Select 0 characters.

    #Set defaults.
    $RichTextBox.SelectionAlignment = "Left"
    $RichTextBox.SelectionColor     = "Black"

    $NewFontSize  = $RichTextBox.Font.Size
    $NewFontStyle = $RichTextBox.Font.Style

    #Depending what is passed, modify the text to be added.

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
       $FontStyle -ne $Null) #Handles changing fonts or sizes.
    {
        $RichTextBox.SelectionFont = New-Object System.Drawing.Font($RichTextBox.Font.Name, $NewFontSize, [System.Drawing.FontStyle]::$NewFontStyle)
    }

    $RichTextBox.AppendText($Text)
}

#This function puts the items in the Leave Balances list box so the user can view/select them.
function PopulateLeaveBalanceListBox
{
    $LeaveBalanceListBox = $Script:MainForm.Controls["LeavePanel"].Controls["BalanceListBox"]

    $LeaveBalanceListBox.BeginUpdate()

    $LeaveBalanceListBox.Items.Clear()
    
    foreach($LeaveItem in $Script:LeaveBalances)
    {
        $String = $LeaveItem.Name + "`t"

        if([System.Windows.Forms.TextRenderer]::MeasureText($LeaveItem.Name, $Script:FormFont).Width -le 60) #In order to line up the balance.
        {
            $String += "`t"
        }

        $String += $LeaveItem.Balance

        if(($LeaveItem.Name -eq "Annual" -or
            $LeaveItem.Name -eq "Sick") -and
            $LeaveItem.Threshold -ne 0) #Add the threshold if it's not 0.
        {
            $String += "`tThreshold: " + $LeaveItem.Threshold
        }

        if($LeaveItem.Name -ne "Annual" -and
           $LeaveItem.Name -ne "Sick" -and
           $LeaveItem.Expires -eq $True) #Add the expiration date if the leave expires.
        {
            $ExpireDate = $LeaveItem.ExpiresOn.ToString("MM/dd/yyyy")

            $String += "`tExpires: $ExpireDate"
        }

        $LeaveBalanceListBox.Items.Add($String) | Out-Null
    }

    $LeaveBalanceListBox.EndUpdate()
}

#This function handles all of the output of the projection.
function PopulateOutputFormRichTextBox
{
    $RichTextBox = $Script:OutputForm.Controls["OutputRichTextBox"]
    
    $LeaveBalancesCopy   = New-Object System.Collections.Generic.List[PSCustomObject] #Make an object to hold a copy of the leave balances.
    $LeaveExpiresOnList  = New-Object System.Collections.Generic.List[PSCustomObject] #Make an object to hold a list of all the dates that leave expires on.
    $ProjectedLeaveIndex = 0 #A counter to keep track of which projected leave items have already been accounted for while cycling through the days.
    $EndOfPayPeriod      = GetEndingOfPayPeriodForDate -Date $Script:BeginningOfPayPeriod
    $LeaveYearEnd        = GetLeaveYearEndForDate -Date $EndOfPayPeriod

    #Get the starting highs/lows of the annual and sick leaves.
    $AnnualHigh = $Script:LeaveBalances[0].Balance
    $AnnualLow  = $Script:LeaveBalances[0].Balance
    $SickHigh   = $Script:LeaveBalances[1].Balance
    $SickLow    = $Script:LeaveBalances[1].Balance

    #Get copies of the decimal values too.
    $ProjectionAnnualDecimal = $Script:AnnualDecimal
    $ProjectionSickDecimal   = $Script:SickDecimal

    #Set this variable for if it's in reach goal mode.
    $GoalsMet = $False

    foreach($LeaveBalance in $Script:LeaveBalances) #Copy the leave balances into the object we created and also copy the dates that leave expires.
    {
        $LeaveBalancesCopy.Add($LeaveBalance.PSObject.Copy())

        if($LeaveBalance.Expires -eq $True)
        {
            $LeaveExpiresOnList.Add($LeaveBalance.ExpiresOn)
        }
    }
    
    $StartDate = $Script:BeginningOfPayPeriod

    #This determines the end date to stop the loop depending on which mode is running.
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

    $TitleString += ":`n"

    #Add the title.
    RichTextBoxAppendText -RichTextBox $RichTextBox -Text $TitleString -Alignment "Center" -FontSize ($Script:OutputForm.Controls["OutputRichTextBox"].Font.Size + 2) -FontStyle "Bold"

    #Add the section header
    RichTextBoxAppendText -RichTextBox $RichTextBox -Text "`nStarting Balances:`n" -FontStyle "Bold"

    #Add Annual/Sick
    RichTextBoxAppendText -RichTextBox $RichTextBox -Text ($LeaveBalancesCopy[0].Name + ":`t`t" + $LeaveBalancesCopy[0].Balance)
    RichTextBoxAppendText -RichTextBox $RichTextBox -Text ("`n" + $LeaveBalancesCopy[1].Name + ":`t`t" + $LeaveBalancesCopy[1].Balance)

    #If projecting to date, append the rest of the balances and their expiration if they expire. Reach goal mode is only concerned with annual/sick.
    if($Script:ProjectOrGoal -eq "Project")
    {
        for($Index = 2; $Index -lt $LeaveBalancesCopy.Count; $Index++) #Starts at index 2 because annual and sick are indices 0 and 1.
        {
            $String = "`n" + $LeaveBalancesCopy[$Index].Name + ":`t"

            if([System.Windows.Forms.TextRenderer]::MeasureText($LeaveBalancesCopy[$Index].Name, $Script:FormFont).Width -le 60) #In order to line up the balances.
            {
                $String += "`t"
            }

            $String += $LeaveBalancesCopy[$Index].Balance

            if($LeaveBalancesCopy[$Index].Expires -eq $True)
            {
                $String += "`tExpires: " + $LeaveBalancesCopy[$Index].ExpiresOn.ToString("MM/dd/yyyy")
            }

            RichTextBoxAppendText -RichTextBox $RichTextBox -Text $String
        }
    }

    #Add in projected leave included in the report.
    if($Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"].CheckedItems.Count -gt 0)
    {
        RichTextBoxAppendText -RichTextBox $RichTextBox -Text "`n`nProjected Leave Included in Report:"

        foreach($LeaveItem in $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"].CheckedItems)
        {
            $LeaveBankName = $Script:ProjectedLeave[$Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"].Items.IndexOf($LeaveItem)].LeaveBank.ToLower().Trim() #Get the leave bank name based on what index is being worked with.
            
            if($LeaveBankName -eq "annual" -or
               $LeaveBankName -eq "sick" -or
               $Script:ProjectOrGoal -eq "Project") #Add annual/sick and only add the others if it's in Project mode.
            {
                RichTextBoxAppendText -RichTextBox $RichTextBox -Text ("`n" + $LeaveItem)
            }
        }
    }

    if($Script:ProjectOrGoal -eq "Goal" -and
        $LeaveBalancesCopy[0].Balance -ge $Script:AnnualGoal -and
        $LeaveBalancesCopy[1].Balance -ge $Script:SickGoal) #This is a silly check to make sure the balances aren't already above the goals set.
    {
        $GoalsMet = $True
    }

    #Primary Loop through dates and pay periods to determine when to accrue leave, subtract leave from hours, etc...
    while($EndOfPayPeriod -le $EndDate -and
          $GoalsMet -eq $False)
    {
        if($Script:EmployeeType -ne "SES") #Displays an alert if the accrual rate changes. Doesn't need the -ge or -le because we're dealing with the end of a pay period.
        {
            if($EndOfPayPeriod.AddDays(-14) -lt $Script:FifteenYearMark -and
                   $EndOfPayPeriod -gt $Script:FifteenYearMark)
            {
                $String = "`n`nAnnual Leave Accrual Rate Changed to the Greater Than 15 Years Category on " + $Script:FifteenYearMark.ToString("MM/dd/yyyy") + "."
                
                RichTextBoxAppendText -RichTextBox $RichTextBox -Text $String -Color "Green"
            }

            elseif($EndOfPayPeriod.AddDays(-14) -lt $Script:ThreeYearMark -and
               $EndOfPayPeriod -gt $Script:ThreeYearMark)
            {
                $String = "`n`nAnnual Leave Accrual Rate Changed to the 3 to 15 Years Category on " + $Script:ThreeYearMark.ToString("MM/dd/yyyy") + "."
                
                RichTextBoxAppendText -RichTextBox $RichTextBox -Text $String -Color "Green"
            }
        }
        
        #This determines if leave expires this pay period. We need to check this so we can alert if unused leave expires.
        $LeaveExpiresThisPayPeriod = $False

        #Loop through the list we made to check if any leave expires.
        foreach($Expire in $LeaveExpiresOnList)
        {
            if($Expire -gt $EndOfPayPeriod.AddDays(-14) -and
               $Expire -le $EndOfPayPeriod)
            {
                $LeaveExpiresThisPayPeriod = $True
            }
        }

        if($Script:ProjectedLeave[$ProjectedLeaveIndex].StartDate -le $EndOfPayPeriod -or
           $LeaveExpiresThisPayPeriod -eq $True) #Need to simulate progressing the days one at a time for two weeks since leave either is taken or expires.
        {
            $Date = GetBeginningOfPayPeriodForDate -Date $EndOfPayPeriod
            
            $ProjectedLeaveEndIndex = $ProjectedLeaveIndex #We need to get an ending location too, so set this variable and adjust it.

            while($ProjectedLeaveEndIndex -lt $Script:ProjectedLeave.Count -and
                  $Script:ProjectedLeave[$ProjectedLeaveEndIndex].StartDate -le $EndOfPayPeriod) #Find the last leave in projected leave that is inside this pay period.
            {
                $ProjectedLeaveEndIndex++
            }

            for($Day = 0; $Day -lt 14; $Day++) #Loop through the two weeks so we can get a snapshot of each day.
            {
                for($ProjectedIndex = $ProjectedLeaveIndex; $ProjectedIndex -lt $ProjectedLeaveEndIndex; $ProjectedIndex++) #Loop through each projected leave item that happens in this pay period.
                {
                    if($Script:ProjectedLeave[$ProjectedIndex].Included -eq $True) #Only count the projeted leave if it's included (checked).
                    {
                        if($Script:ProjectedLeave[$ProjectedIndex].HoursHashTable.ContainsKey($Date.ToString("MM/dd/yyyy")) -eq $True -and
                           $Script:ProjectedLeave[$ProjectedIndex].HoursHashTable[$Date.ToString("MM/dd/yyyy")] -gt 0) #This is so we can account for each day of leave in the projected leave. Only applies to items that have this day with more than 0 hours taken.
                        {
                            $BalanceFound = $False
                            $BalanceIndex = 0

                            while($BalanceIndex -lt $LeaveBalancesCopy.Count -and
                                  $LeaveBalancesCopy[$BalanceIndex].Name -ne $Script:ProjectedLeave[$ProjectedIndex].LeaveBank) #Find the leave balance that matches the leave bank.
                            {
                                $BalanceIndex++
                            }

                            if($BalanceIndex -lt $LeaveBalancesCopy.Count -and
                               $LeaveBalancesCopy[$BalanceIndex].Name -eq $Script:ProjectedLeave[$ProjectedIndex].LeaveBank) #Continuation of the above while loop. This is to ensure we actually found a leave balance that matches.
                            {
                                $BalanceFound = $True
                            }

                            if($BalanceFound -eq $True) #This is where the leave is subtracted off the balance.
                            {
                                $BalanceBeforeSubtraction = $LeaveBalancesCopy[$BalanceIndex].Balance #This is used later to alert when dropping below a threshold so you only get alerted if you were above it and then below it, not if you keep going lower.
                            
                                $LeaveBalancesCopy[$BalanceIndex].Balance -= $Script:ProjectedLeave[$ProjectedIndex].HoursHashTable[$Date.ToString("MM/dd/yyyy")] #Subtract the leave out of the copied leave balance list.

                                while($LeaveBalancesCopy[$BalanceIndex].Balance -lt 0 -and
                                     ($BalanceIndex + 1) -lt $LeaveBalancesCopy.Count -and
                                      $LeaveBalancesCopy[$BalanceIndex + 1].Name -eq $LeaveBalancesCopy[$BalanceIndex].Name) #This handles when there are multiple leave balances with the same name. They're already sorted, so this just subtracts off the one that expires first, then second, etc... Deletes ones that are now empty but leaves the last one, even if negative.
                                {
                                    $LeaveBalancesCopy[$BalanceIndex + 1].Balance += $LeaveBalancesCopy[$BalanceIndex].Balance  #Take that negative balance off of the next entry with the same name.
                                
                                    $Count = 0

                                    foreach($LeaveBalance in $LeaveBalancesCopy) #See how many leave balances match the expiration date of the balance index we found earlier.
                                    {
                                        if($LeaveBalance.Expires -eq $True -and
                                           $LeaveBalance.ExpiresOn -eq $LeaveBalancesCopy[$BalanceIndex].ExpiresOn)
                                        {
                                            $Count++
                                        }
                                    }

                                    if($Count -gt 1) #If more than one, remove that date from our expires list since that date is now "passed". It's more than one since there might have been multiples added if two types of leave had the same expiring date but different names.
                                    {
                                        $LeaveExpiresOnList.Remove($LeaveBalance.ExpiresOn.ToString("MM/dd/yyyy"))
                                    }

                                    $LeaveBalancesCopy.RemoveAt($BalanceIndex) #Delete the now 0 or negative balance out of the copy list.
                                }

                                if($Script:DisplayHighsAndLows -eq $True) #Keeps track of the lows here for annual/sick leave. Nothing fancy.
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
                                   $LeaveBalancesCopy[$BalanceIndex].Name -eq "Sick")) #This displays the amount of hours taken off the balance and the remaining balance. Only shows if in Project mode, or if type is annual/sick.
                                {
                                    $LeaveName = $LeaveBalancesCopy[$BalanceIndex].Name

                                    if($LeaveName.ToLower().Contains("leave") -eq $False) #Just adding in the word "Leave" to make the output easier to read.
                                    {
                                       $LeaveName += " Leave"
                                    }
                                
                                    $String  = "`n`n" + $Script:ProjectedLeave[$ProjectedIndex].HoursHashTable[$Date.ToString("MM/dd/yyyy")] + " Hours of $LeaveName Taken on " + $Date.ToString("MM/dd/yyyy") + "."
                                    $String += "`n$LeaveName Balance is now: " + [Math]::Floor($LeaveBalancesCopy[$BalanceIndex].Balance)

                                    RichTextBoxAppendText -RichTextBox $RichTextBox -Text $String
                                }

                                if($LeaveBalancesCopy[$BalanceIndex].Name -eq "Annual" -or
                                   $LeaveBalancesCopy[$BalanceIndex].Name -eq "Sick" -and
                                  ($LeaveBalancesCopy[$BalanceIndex].Threshold -gt 0 -and
                                   $BalanceBeforeSubtraction -ge $LeaveBalancesCopy[$BalanceIndex].Threshold -and
                                   $LeaveBalancesCopy[$BalanceIndex].Balance -lt $LeaveBalancesCopy[$BalanceIndex].Threshold)) #Threshold alerting section. Only alerts when balance was above threshold and then dropped below.
                                {
                                    $String = "`n`n" + $LeaveBalancesCopy[$BalanceIndex].Name + " Leave balance is " + $LeaveBalancesCopy[$BalanceIndex].Balance + " which is below the set threshold of " + $LeaveBalancesCopy[$BalanceIndex].Threshold + " after taking leave on " + $Date.ToString("MM/dd/yyyy") + "."
                                
                                    RichTextBoxAppendText -RichTextBox $RichTextBox -Text $String -Color "Blue"
                                }

                                if($LeaveBalancesCopy[$BalanceIndex].Balance -lt 0) #Warning if any balance drops below 0.
                                {
                                    $LeaveName = $LeaveBalancesCopy[$BalanceIndex].Name

                                    if($LeaveName.ToLower().Contains("leave") -eq $False)
                                    {
                                       $LeaveName += " Leave"
                                    }
                                
                                    $String = "`n`n$LeaveName balance is negative (" + $LeaveBalancesCopy[$BalanceIndex].Balance + ") after taking leave on " + $Date.ToString("MM/dd/yyyy") + "."

                                    RichTextBoxAppendText -RichTextBox $RichTextBox -Text $String -Color "Red"
                                }
                            }
                        }
                    }
                }

                if($LeaveExpiresThisPayPeriod -eq $True) #If one or more of the leave balances expires this pay period.
                {
                    $Index = 0
                    
                    while($Index -lt $LeaveBalancesCopy.Count)
                    {
                        if($LeaveBalancesCopy[$Index].Expires -eq $True -and
                           $LeaveBalancesCopy[$Index].Balance -gt 0 -and
                           $LeaveBalancesCopy[$Index].ExpiresOn -eq $Date) #Only print if the leave expires on the date in the loop and it has a balance.
                        {
                            $LeaveName = $LeaveBalancesCopy[$Index].Name

                            if($LeaveName.ToLower().Contains("leave") -eq $False)
                            {
                                $LeaveName += " Leave"
                            }
                            
                            $String = "`n`n" + $LeaveBalancesCopy[$Index].Balance + " hours of $LeaveName will expire on " + $Date.ToString("MM/dd/yyyy") + "."

                            RichTextBoxAppendText -RichTextBox $RichTextBox -Text $String -Color "Red"
                            
                            $LeaveBalancesCopy.RemoveAt($Index)
                        }
                        
                        else
                        {
                            $Index++
                        }
                    }
                }
                
                $Date = $Date.AddDays(1) #Increment the date in this counter.
            }

            $ProjectedLeaveIndex = $ProjectedLeaveEndIndex #Set the new beginning index for the next loop.
        }

        #Accrue the annual and sick leave hours.
        $LeaveBalancesCopy[0].Balance += (GetAnnualLeaveAccrualHours -PayPeriod $EndOfPayPeriod)
        $LeaveBalancesCopy[1].Balance += GetSickLeaveAccrualHours

        if($EndOfPayPeriod -eq $LeaveYearEnd) #If it's the leave year end, check stuff and set the variable to the next leave year end.
        {
            if($LeaveBalancesCopy[0].Balance -gt $Script:LeaveCeiling) #If over the use/lose, show a warning and set the balance and decimal value to the ceiling.
            {
                $String  = "`n`n" + ($LeaveBalancesCopy[0].Balance - $Script:LeaveCeiling) + " hours of Annual Leave will be forfeited on " + $LeaveYearEnd.ToString("MM/dd/yyyy") + " which is the Leave Year End. "
                $String += "The Leave Year Schedule Deadline for the " + (GetBeginningOfPayPeriodForDate -Date $LeaveYearEnd).Year + " Leave Year is " + (GetLeaveYearScheduleDeadline -Date $LeaveYearEnd).ToString("MM/dd/yyyy") + " in order for the leave to be eligible to be restored in certain circumstances."
                
                RichTextBoxAppendText -RichTextBox $RichTextBox -Text $String -Color "Red"

                $LeaveBalancesCopy[0].Balance = $Script:LeaveCeiling

                if($Script:EmployeeType -eq "Part-Time")
                {
                    $ProjectionAnnualDecimal = 0.0
                }
            }
            
            $LeaveYearEnd = GetLeaveYearEndForDate -Date $LeaveYearEnd.AddDays(1) #Get the next leave year end and store it. 
        }

        if($Script:DisplayHighsAndLows -eq $True) #Keep track of the annual/sick leave highs. Nothing fancy.
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
        
        #Check if the goals were met so we can exit the loop.
        if($Script:ProjectOrGoal -eq "Goal" -and
           $LeaveBalancesCopy[0].Balance -ge $Script:AnnualGoal -and
           $LeaveBalancesCopy[1].Balance -ge $Script:SickGoal)
        {
            $GoalsMet = $True
        }

        #If displaying after each pay period, print the balances.
        if($Script:DisplayAfterEachPP -eq $True -and
           $EndOfPayPeriod -lt $EndDate -and
           $GoalsMet -eq $False)
        {
            #Add descriptive date
            RichTextBoxAppendText -RichTextBox $RichTextBox -Text ("`n`nBalances as of Pay Period Ending " + $EndOfPayPeriod.ToString("MM/dd/yyyy") + ":`n")
            
            #Add Annual with color
            $AnnualString = $LeaveBalancesCopy[0].Name + ":`t`t" + [Math]::Floor($LeaveBalancesCopy[0].Balance)

            if($LeaveBalancesCopy[0].Balance -lt 0) #Annual balance negative warning.
            {
                RichTextBoxAppendText -RichTextBox $RichTextBox -Text $AnnualString -Color "Red"
            }

            elseif($LeaveBalancesCopy[0].Balance -gt $Script:LeaveCeiling) #Annual balance above ceiling warning.
            {
                $AnnualString += "`tBalance is greater than your Annual Leave ceiling."
        
                RichTextBoxAppendText -RichTextBox $RichTextBox -Text $AnnualString -Color "Blue"
            }

            else #Print the annual balance.
            {
                RichTextBoxAppendText -RichTextBox $RichTextBox -Text $AnnualString
            }

            #Add Sick with color
            $SickString = "`n" + $LeaveBalancesCopy[1].Name + ":`t`t" + [Math]::Floor($LeaveBalancesCopy[1].Balance)

            if($LeaveBalancesCopy[1].Balance -lt 0) #Sick balance negative warning.
            {
                RichTextBoxAppendText -RichTextBox $RichTextBox -Text $SickString -Color "Red"
            }

            else #Print the sick balance.
            {
                RichTextBoxAppendText -RichTextBox $RichTextBox -Text $SickString
            }
    
            #If projecting to date, append the rest of the balances and their expiration if they expire.
            if($Script:ProjectOrGoal -eq "Project")
            {
                for($Index = 2; $Index -lt $LeaveBalancesCopy.Count; $Index++) #Start at 2 because 0 and 1 are already printed (annual and sick).
                {
                    if($LeaveBalancesCopy[$Index].Expires -eq $False -or
                      ($LeaveBalancesCopy[$Index].Expires -eq $True -and
                       $LeaveBalancesCopy[$Index].ExpiresOn -ge $EndOfPayPeriod))
                    {
                        $String = "`n" + $LeaveBalancesCopy[$Index].Name + ":`t"

                        if([System.Windows.Forms.TextRenderer]::MeasureText($LeaveBalancesCopy[$Index].Name, $Script:FormFont).Width -le 60) #In order to line up the balances.
                        {
                            $String += "`t"
                        }

                        $String += $LeaveBalancesCopy[$Index].Balance

                        if($LeaveBalancesCopy[$Index].Expires -eq $True)
                        {
                            $String += "`tExpires: " + $LeaveBalancesCopy[$Index].ExpiresOn.ToString("MM/dd/yyyy")
                        }

                        if($LeaveBalancesCopy[$Index].Balance -lt 0) #Warning if negative
                        {
                            RichTextBoxAppendText -RichTextBox $RichTextBox -Text $String -Color "Red"
                        }

                        else #Print it if not negative.
                        {
                            RichTextBoxAppendText -RichTextBox $RichTextBox -Text $String
                        }
                    }
                }
            }
        }

        $EndOfPayPeriod = $EndOfPayPeriod.AddDays(14) #Move on to the next pay period!
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
        $String = "`n`nAnnual Leave High:`t" + [Math]::Floor($AnnualHigh) + "`nSick Leave High:`t`t" + [Math]::Floor($SickHigh) + "`nAnnual Leave Low:`t" + [Math]::Floor($AnnualLow) + "`nSick Leave Low:`t`t" + [Math]::Floor($SickLow)
        
        RichTextBoxAppendText -RichTextBox $RichTextBox -Text $String
    }

    if($Script:ProjectOrGoal -eq "Goal")
    {
        if($GoalsMet -eq $True)
        {
            $String = "`n`nAnnual Leave and Sick Leave Goals Achieved After Pay Period Ending " + $EndOfPayPeriod.ToString("MM/dd/yyyy") + "."

            RichTextBoxAppendText -RichTextBox $RichTextBox -Text $String
        }

        else
        {
            $String = "`n`nGoals not met by " + $EndOfPayPeriod.ToString("MM/dd/yyyy") + "."

            RichTextBoxAppendText -RichTextBox $RichTextBox -Text $String
        }

        $String = "`n`nGoals:`nAnnual:`t" + $Script:AnnualGoal + "`nSick:`t" + $Script:SickGoal

        RichTextBoxAppendText -RichTextBox $RichTextBox -Text $String
    }

    RichTextBoxAppendText -RichTextBox $RichTextBox -Text "`n"

    #Add the section header
    RichTextBoxAppendText -RichTextBox $RichTextBox -Text ("`nEnding Balances After Pay Period Ending " + $EndOfPayPeriod.ToString("MM/dd/yyyy") + ":`n") -Alignment "Center" -FontSize ($Script:OutputForm.Controls["OutputRichTextBox"].Font.Size + 2) -FontStyle "Bold"

    #Add Annual with color
    $AnnualString = $LeaveBalancesCopy[0].Name + ":`t`t" + $LeaveBalancesCopy[0].Balance

    if($LeaveBalancesCopy[0].Balance -lt 0)
    {
        RichTextBoxAppendText -RichTextBox $RichTextBox -Text $AnnualString -Color "Red"
    }

    elseif($LeaveBalancesCopy[0].Balance -gt $Script:LeaveCeiling)
    {
        $AnnualString += "`tBalance is greater than your Annual Leave ceiling."
        
        RichTextBoxAppendText -RichTextBox $RichTextBox -Text $AnnualString -Color "Blue"
    }

    else
    {
        RichTextBoxAppendText -RichTextBox $RichTextBox -Text $AnnualString
    }

    #Add Sick with color
    $SickString = "`n" + $LeaveBalancesCopy[1].Name + ":`t`t" + $LeaveBalancesCopy[1].Balance

    if($LeaveBalancesCopy[1].Balance -lt 0)
    {
        RichTextBoxAppendText -RichTextBox $RichTextBox -Text $SickString -Color "Red"
    }

    else
    {
        RichTextBoxAppendText -RichTextBox $RichTextBox -Text $SickString
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
                $String = "`n" + $LeaveBalancesCopy[$Index].Name + ":`t"

                if([System.Windows.Forms.TextRenderer]::MeasureText($LeaveBalancesCopy[$Index].Name, $Script:FormFont).Width -le 60) #In order to line up the balances.
                {
                    $String += "`t"
                }

                $String += $LeaveBalancesCopy[$Index].Balance

                if($LeaveBalancesCopy[$Index].Expires -eq $True)
                {
                    $String += "`tExpires: " + $LeaveBalancesCopy[$Index].ExpiresOn.ToString("MM/dd/yyyy")
                }

                if($LeaveBalancesCopy[$Index].Balance -lt 0)
                {
                    RichTextBoxAppendText -RichTextBox $RichTextBox -Text $String -Color "Red"
                }

                else
                {
                    RichTextBoxAppendText -RichTextBox $RichTextBox -Text $String
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
               $Leave.Included -eq $True -and
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
            $EndOfPayPeriod = $EndOfPayPeriod.AddDays(14) #Need to add the days before accruing hours in this instance.
            
            $LeaveBalancesCopy[0].Balance += (GetAnnualLeaveAccrualHours -PayPeriod $EndOfPayPeriod)
        }

        if($LeaveBalancesCopy[0].Balance -gt $Script:LeaveCeiling)
        {
            $String  = "`n`nYour Annual Leave Balance is expected to be " + $LeaveBalancesCopy[0].Balance + " hours at the Leave Year End on " + $LeaveYearEnd.ToString("MM/dd/yyyy") + ". "
            $String += "This is " + ($LeaveBalancesCopy[0].Balance - $Script:LeaveCeiling) + " hours over your leave ceiling of " + $Script:LeaveCeiling + " hours. "
            $String += "If you schedule no additional Annual Leave before the Leave Year End, you will forfeit these hours."
            $String += "`n`nIf possible, you should submit the leave requests on or before " + (GetLeaveYearScheduleDeadline -Date $LeaveYearEnd).ToString("MM/dd/yyyy") + " so "
            $String += "these hours are eligible to be restored if you are unable to take your scheduled Annual Leave for a few specific reasons which causes you to forfeit these hours."

            RichTextBoxAppendText -RichTextBox $RichTextBox -Text $String -Color "Blue"
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
        $NewPanel.Height = 22
        $NewPanel.Width = 415

        $NewNumericUpDown = New-Object System.Windows.Forms.NumericUpDown
        $NewNumericUpDown.Name = "HoursNumericUpDown"
        $NewNumericUpDown.Height = 20
        $NewNumericUpDown.Maximum = 24
        $NewNumericUpDown.Tag = $DateString
        $NewNumericUpDown.Width = 40

        $DayOfWeekLabel = New-Object System.Windows.Forms.Label
        $DayOfWeekLabel.AutoSize = $True
        $DayOfWeekLabel.Left = 45
        $DayOfWeekLabel.Text = $Date.DayOfWeek.ToString()
        $DayOfWeekLabel.Top = 3

        $DateLabel = New-Object System.Windows.Forms.Label
        $DateLabel.AutoSize = $True
        $DateLabel.Left = 115
        $DateLabel.Text = $DateString
        $DateLabel.Top = 3
        
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
            $DateLabel.Text = $DateString + "  -  " + $Script:HolidaysHashTable[$DateString]
        }

        elseif($InaugurationHoliday -eq $True -and
               $Script:InaugurationDayHashTable.ContainsKey($DateString) -eq $True)
        {
            $DateLabel.Text = $DateString + "  -  " + $Script:InaugurationDayHashTable[$DateString]
        }

        $NewPanel.Controls.AddRange(($NewNumericUpDown, $DayOfWeekLabel, $DateLabel))
        $Script:EditProjectedLeaveForm.Controls["LeaveDatesFlowLayoutPanel"].Controls.Add($NewPanel)

        $NewNumericUpDown.Add_TextChanged({HourNumericUpDownChanged})
        $NewNumericUpDown.Add_MouseClick({NumericUpDownMouseClick -Sender $NewNumericUpDown -EventArguments $_}.GetNewClosure())
        $NewNumericUpDown.Add_Enter({NumericUpDownEnter -Sender $NewNumericUpDown -EventArguments $_}.GetNewClosure())
        
        $Date = $Date.AddDays(1)
    }

    $Script:EditProjectedLeaveForm.Controls["LeaveDatesFlowLayoutPanel"].Visible = $True
}

function PopulateProjectedLeaveListBox
{
    $ProjectedLeaveListBox = $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"]

    $SelectedItem = $Script:ProjectedLeave[$ProjectedLeaveListBox.SelectedIndex]

    $Script:DrawingListBox = $True

    $ProjectedLeaveListBox.BeginUpdate()

    $ProjectedLeaveListBox.Items.Clear()
    
    foreach($LeaveItem in $Script:ProjectedLeave)
    {
        $DateString = ""
        $TotalHours = 0

        if($LeaveItem.StartDate -eq $LeaveItem.EndDate)
        {
            $DateString = $LeaveItem.StartDate.ToString("MM/dd/yyyy")
        }

        else
        {
            $DateString = $LeaveItem.StartDate.ToString("MM/dd/yyyy") + "  to  " + $LeaveItem.EndDate.ToString("MM/dd/yyyy")
        }

        foreach($Day in $LeaveItem.HoursHashTable.Keys)
        {
            $TotalHours += $LeaveItem.HoursHashTable[$Day]
        }
        
        $String = $LeaveItem.LeaveBank + "`t"

        if([System.Windows.Forms.TextRenderer]::MeasureText($LeaveItem.LeaveBank, $Script:FormFont).Width -le 60) #Line up the date after the name if it's long.
        {
            $String += "`t"
        }

        $String += $DateString + "`t"

        if($DateString.Contains("to") -eq $False) #Line up the balance if it's not a date range.
        {
            $String += "`t"
        }

        $String += "$TotalHours "

        if($TotalHours -eq 1)
        {
            $String += "Hour"
        }

        else
        {
            $String += "Hours"
        }

        $ProjectedLeaveListBox.Items.Add($String) | Out-Null
        
        $ProjectedLeaveListBox.SetItemChecked($ProjectedLeaveListBox.Items.Count - 1, $LeaveItem.Included)
    }

    $ProjectedLeaveListBox.EndUpdate()

    $Script:DrawingListBox = $False

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

        $Script:TimeUntilMilestone = "Accrue Rate Changes in: " + $YearDifference + " " + (NumberGetsLetters -String "Year" -Number $YearDifference) + ", " + $MonthDifference + " " + (NumberGetsLetters -String "Month" -Number $MonthDifference) + ", " + $DayDifference + " " + (NumberGetsLetters -String "Day" -Number $DayDifference)
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

function MainFormKeyDown
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)][System.EventArgs] $EventArguments
    )
    
    if($EventArguments.KeyCode -eq "F1")
    {
        MainFormHelpButtonClick
    }

    elseif($EventArguments.KeyCode -eq "F6")
    {
        MainFormUpdateInfoButtonClick
    }

    elseif($EventArguments.KeyCode -eq "F7")
    {
        $MainForm.ActiveControl = $Script:MainForm.Controls["LeavePanel"].Controls["BalanceListBox"]
        
        MainFormBalanceEditButtonClick
    }

    elseif($EventArguments.KeyCode -eq "F8" -and
           $Script:ProjectedLeave.Count -gt 0)
    {
        $MainForm.ActiveControl = $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"]
        
        MainFormProjectedEditButtonClick
    }

    elseif($EventArguments.KeyCode -eq "F5")
    {
        MainFormProjectButtonClick
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

    $Script:MainForm.Controls["ReportPanel"].Controls["ProjectToEndOfPayPeriodLabel"].Visible = $True

    $Script:MainForm.Controls["ReportPanel"].Controls["AnnualGoalNumericUpDown"].Enabled = $False
    $Script:MainForm.Controls["ReportPanel"].Controls["SickGoalNumericUpDown"].Enabled   = $False
}

function MainFormReachGoalRadioButtonClick
{
    $Script:ProjectOrGoal = "Goal"
    
    $Script:MainForm.Controls["ReportPanel"].Controls["ProjectToDateDateTimePicker"].Enabled = $False

    $Script:MainForm.Controls["ReportPanel"].Controls["ProjectToEndOfPayPeriodLabel"].Visible = $False

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

function MainFormBalanceListBoxKeyDown
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)][System.EventArgs] $EventArguments
    )

    if($EventArguments.KeyCode -eq "Return")
    {
        MainFormBalanceEditButtonClick
    }

    elseif($EventArguments.KeyCode -eq "Add" -or
           $EventArguments.KeyCode -eq "Oemplus")
    {
        MainFormBalanceAddButtonClick
    }

    elseif($EventArguments.KeyCode -eq "Delete")
    {
        $SelectedLeaveBalance = $Script:LeaveBalances[$Script:MainForm.Controls["LeavePanel"].Controls["BalanceListBox"].SelectedIndex]
        
        if($SelectedLeaveBalance.Name.ToLower() -ne "annual" -and
           $SelectedLeaveBalance.Name.ToLower() -ne "sick")
        {
            MainFormBalanceDeleteButtonClick
        }
    }
}

function MainFormBalanceAddButtonClick
{
    $NewLeaveBalance = [PSCustomObject] @{
    Name      = "New Leave"
    Balance   = 0
    Expires   = $False
    ExpiresOn = $Script:CurrentDate
    Static    = $False
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

function MainFormProjectedLeaveListBoxClick
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)][System.EventArgs] $EventArguments
    )

    $ClickedIndex   = $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"].IndexFromPoint($EventArguments.Location)
    $CheckedListBox = $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"]
    
    if($ClickedIndex -ne -1 -and
       $EventArguments.Button -eq "Left" -and
       $EventArguments.X -gt 13) #13 is the far right edge of the checkbox. So any clicks greater than that are not checkbox clicks.
    {
        $CheckedListBox.SetItemChecked($ClickedIndex, -not $CheckedListBox.GetItemChecked($ClickedIndex))
    }

    if($ClickedIndex -eq -1 -and
       $Script:ProjectedLeave.Count -gt 0)
    {
        $SelectedIndex = $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"].SelectedIndex
        
        $CheckedListBox.SetItemChecked($SelectedIndex, -not $CheckedListBox.GetItemChecked($SelectedIndex))
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
    $CheckedListBox     = $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"]

    if($DoubleClickedIndex -ne -1 -and
       $EventArguments.X -gt 13) #13 is the far right edge of the checkbox. So any clicks greater than that are not checkbox clicks.)
    {
        $CheckedListBox.SetItemChecked($DoubleClickedIndex, -not $CheckedListBox.GetItemChecked($DoubleClickedIndex))
        
        MainFormProjectedEditButtonClick
    }

    if($DoubleClickedIndex -eq -1)
    {
        $SelectedIndex = $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"].SelectedIndex
        
        $CheckedListBox.SetItemChecked($SelectedIndex, -not $CheckedListBox.GetItemChecked($SelectedIndex))
    }
}

function MainFormProjectedListBoxKeyDown
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)][System.EventArgs] $EventArguments
    )

    if($EventArguments.KeyCode -eq "Add" -or
       $EventArguments.KeyCode -eq "Oemplus")
    {
        MainFormProjectedAddButtonClick
    }

    elseif($Script:ProjectedLeave.Count -gt 0)
    {
        if($EventArguments.KeyCode -eq "Return")
        {
            MainFormProjectedEditButtonClick
        }

        elseif($EventArguments.KeyCode -eq "Delete")
        {
            MainFormProjectedDeleteButtonClick
        }
    }
}

function MainFormProjectedLeaveListBoxItemCheck
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)][System.EventArgs] $EventArguments
    )

    if($Script:DrawingListBox -eq $False)
    {
        $SelectedIndex  = $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"].SelectedIndex
        $CheckedListBox = $Script:MainForm.Controls["LeavePanel"].Controls["ProjectedListBox"]
    
        $SelectedLeaveDetails = $Script:ProjectedLeave[$SelectedIndex]

        $SelectedLeaveDetails.Included = -not $CheckedListBox.GetItemChecked($SelectedIndex) #This event fires before the check change actually happens, so we need the opposite value.
    }
}

function MainFormProjectedAddButtonClick
{
    $Script:UnsavedProjectedLeave = $True
    
    $NewProjectedLeave = [PSCustomObject] @{
        LeaveBank      = "Annual"
        StartDate      = $Script:CurrentDate
        EndDate        = $Script:CurrentDate
        HoursHashTable = @{
        $Script:CurrentDate.ToString("MM/dd/yyyy") = GetHoursForWorkDay -Day $Script:CurrentDate
        }
        Included       = $True
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

    $Script:MainForm.Controls["ReportPanel"].Controls["ProjectToEndOfPayPeriodlabel"].Text = "Will Project Through Pay Period Ending: " + (GetEndingOfPayPeriodForDate -Date $Script:ProjectToDate).ToString("MM/dd/yyyy")
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

function MainFormHelpButtonClick
{
    BuildHelpForm

    $Script:HelpForm.ShowDialog()
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

function EmployeeInfoFormKeyDown
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)][System.EventArgs] $EventArguments
    )

    if($EventArguments.KeyCode -eq "Return")
    {
        EmployeeInfoFormOkButtonClick
    }
    
    elseif($EventArguments.KeyCode -eq "Escape")
    {
        EmployeeInfoFormCancelButtonClick
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

function EditLeaveBalanceFormCheckNameLength
{
    if([System.Windows.Forms.TextRenderer]::MeasureText($Script:EditLeaveBalanceForm.Controls["LeaveBalanceNameTextBox"].Text, $Script:FormFont).Width -gt 85) #Check the width of the name string
    {
        $Text = $Script:EditLeaveBalanceForm.Controls["LeaveBalanceNameTextBox"].Text

        $Script:EditLeaveBalanceForm.Controls["LeaveBalanceNameTextBox"].Text = $Text.Substring(0, $Text.Length - 1)

        $Script:EditLeaveBalanceForm.Controls["LeaveBalanceNameTextBox"].SelectionStart = $Text.Length - 1
    }
}

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

function EditLeaveBalanceFormKeyDown
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)][System.EventArgs] $EventArguments
    )

    if($EventArguments.KeyCode -eq "Return" -and
       $Script:EditLeaveBalanceForm.Controls["EditLeaveOkButton"].Enabled -eq $True)
    {
        EditLeaveBalanceFormOkButtonClick
    }
    
    elseif($EventArguments.KeyCode -eq "Escape")
    {
        EditLeaveBalanceFormCancelButtonClick
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

    if($ProjectedLeaveUpdated -eq $True -and
       $Script:ProjectedLeave.Count -gt 0)
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

function EditProjectedLeaveFormKeyDown
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)][System.EventArgs] $EventArguments
    )

    if($EventArguments.KeyCode -eq "Return")
    {
        EditProjectedLeaveOkButton
    }
    
    elseif($EventArguments.KeyCode -eq "Escape")
    {
        EditProjectedLeaveCancelButton
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

function OutputFormKeyDown
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)][System.EventArgs] $EventArguments
    )

    if($EventArguments.KeyCode -eq "Escape")
    {
        OutputFormCloseButton
    }
}

function OutputFormCopyButton
{
    #Do a few simple RegEx replacements to replace all tab characters with a space and then replace any consecutive spaces with only a single space.
    Set-Clipboard ($Script:OutputForm.Controls["OutputRichTextBox"].Text -replace "`t+", " " -replace " {2,}", " ")
}

function OutputFormCloseButton
{
    $Script:OutputForm.Close()
}

#endregion Output Form

#region Help Form

function HelpFormKeyDown
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)][System.EventArgs] $EventArguments
    )

    if($EventArguments.KeyCode -eq "Escape")
    {
        HelpFormCloseButton
    }
}

function HelpFormCloseButton
{
    $Script:HelpForm.Close()
}

#endregion Help Form

function Global:NumericUpDownMouseClick
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)] $Sender,
        [parameter(Mandatory=$True)] $EventArguments
    )

    $Sender.Select(0, $Sender.Text.Length)
}

function Global:NumericUpDownEnter
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$True)] $Sender,
        [parameter(Mandatory=$True)] $EventArguments
    )

    $Sender.Select(0, $Sender.Text.Length)
}

#endregion Event Handlers

#region Form Building Functions

function BuildMainForm
{
    $Script:MainForm          = New-Object System.Windows.Forms.Form
    $MainForm.Name            = "MainForm"
    $MainForm.BackColor       = "WhiteSmoke"
    $MainForm.Font            = $Script:FormFont
    $MainForm.FormBorderStyle = "FixedSingle"
    $MainForm.KeyPreview      = $True
    $MainForm.MaximizeBox     = $False
    $MainForm.Size            = New-Object System.Drawing.Size(957, 438)
    $MainForm.StartPosition   = "CenterScreen"
    $MainForm.Text            = "Federal Civilian Leave Calculator"
    $MainForm.WindowState     = "Normal"
    
    $SettingsPanel = New-Object System.Windows.Forms.Panel
    $SettingsPanel.Name = "SettingsPanel"
    $SettingsPanel.BackColor = "LightGray"
    $SettingsPanel.Dock = "Top"
    $SettingsPanel.Height = 71
    $SettingsPanel.TabIndex = 1

    $SCDLeaveDateLabel = New-Object System.Windows.Forms.Label
    $SCDLeaveDateLabel.Name = "SCDLeaveDateLabel"
    $SCDLeaveDateLabel.AutoSize = $True
    $SCDLeaveDateLabel.Left = 20
    $SCDLeaveDateLabel.Text = "SCD Leave Date:"
    $SCDLeaveDateLabel.Top = 11

    $SCDLeaveDateDateTimePicker = New-Object System.Windows.Forms.DateTimePicker
    $SCDLeaveDateDateTimePicker.Name = "SCDLeaveDateDateTimePicker"
    $SCDLeaveDateDateTimePicker.Format = "Short"
    $SCDLeaveDateDateTimePicker.Left = 125
    $SCDLeaveDateDateTimePicker.MaxDate = $Script:CurrentDate
    $SCDLeaveDateDateTimePicker.Top = 8
    $SCDLeaveDateDateTimePicker.Width = 200

    $UpdateInfoButton = New-Object System.Windows.Forms.Button
    $UpdateInfoButton.Name = "UpdateInfoButton"
    $UpdateInfoButton.Left = 345
    $UpdateInfoButton.Text = "Update Employee Info"
    $UpdateInfoButton.Top = 8
    $UpdateInfoButton.Width = 300
    
    $LengthOfServiceTextBox = New-Object System.Windows.Forms.TextBox
    $LengthOfServiceTextBox.Name = "LengthOfServiceTextBox"
    $LengthOfServiceTextBox.Left = 20
    $LengthOfServiceTextBox.ReadOnly = $True
    $LengthOfServiceTextBox.TabStop = $False
    $LengthOfServiceTextBox.Top = 38
    $LengthOfServiceTextBox.Width = 305
    
    $EmployeeTypeTextBox = New-Object System.Windows.Forms.TextBox
    $EmployeeTypeTextBox.Name = "EmployeeTypeTextBox"
    $EmployeeTypeTextBox.Left = 345
    $EmployeeTypeTextBox.ReadOnly = $True
    $EmployeeTypeTextBox.TabStop = $False
    $EmployeeTypeTextBox.Top = 38
    $EmployeeTypeTextBox.Width = 300

    $DisplayBalanceEveryLeaveCheckBox = New-Object System.Windows.Forms.CheckBox
    $DisplayBalanceEveryLeaveCheckBox.Name = "DisplayBalanceEveryLeaveCheckBox"
    $DisplayBalanceEveryLeaveCheckBox.AutoSize = $True
    $DisplayBalanceEveryLeaveCheckBox.Left = 665
    $DisplayBalanceEveryLeaveCheckBox.Text = "Display Balance After Each Day of Leave"
    $DisplayBalanceEveryLeaveCheckBox.Top = 2

    $DisplayBalanceEveryPayPeriodEnd = New-Object System.Windows.Forms.CheckBox
    $DisplayBalanceEveryPayPeriodEnd.Name = "DisplayBalanceEveryPayPeriodEnd"
    $DisplayBalanceEveryPayPeriodEnd.AutoSize = $True
    $DisplayBalanceEveryPayPeriodEnd.Left = 665
    $DisplayBalanceEveryPayPeriodEnd.Text = "Display Balance After Each Pay Period Ends"
    $DisplayBalanceEveryPayPeriodEnd.Top = 25

    $DisplayLeaveHighsAndLows = New-Object System.Windows.Forms.CheckBox
    $DisplayLeaveHighsAndLows.Name = "DisplayLeaveHighsAndLows"
    $DisplayLeaveHighsAndLows.AutoSize = $True
    $DisplayLeaveHighsAndLows.Left = 665
    $DisplayLeaveHighsAndLows.Text = "Display Annual/Sick Leave Highs/Lows"
    $DisplayLeaveHighsAndLows.Top = 48

    $ReportPanel = New-Object System.Windows.Forms.Panel
    $ReportPanel.Name = "ReportPanel"
    $ReportPanel.BackColor = "LightGray"
    $ReportPanel.Dock = "Bottom"
    $ReportPanel.Height = 70
    $ReportPanel.TabIndex = 3

    $ProjectBalanceRadioButton = New-Object System.Windows.Forms.RadioButton
    $ProjectBalanceRadioButton.Name = "ProjectBalanceRadioButton"
    $ProjectBalanceRadioButton.AutoSize = $True
    $ProjectBalanceRadioButton.Left = 10
    $ProjectBalanceRadioButton.Text = "Project to Date:"
    $ProjectBalanceRadioButton.Top = 10
    
    $ReachGoalRadioButton = New-Object System.Windows.Forms.RadioButton
    $ReachGoalRadioButton.Name = "ReachGoalRadioButton"
    $ReachGoalRadioButton.AutoSize = $True
    $ReachGoalRadioButton.Left = 10
    $ReachGoalRadioButton.Text = "Reach Goal:"
    $ReachGoalRadioButton.Top = 40
    
    $ProjectToDateDateTimePicker = New-Object System.Windows.Forms.DateTimePicker
    $ProjectToDateDateTimePicker.Name = "ProjectToDateDateTimePicker"
    $ProjectToDateDateTimePicker.Format = "Short"
    $ProjectToDateDateTimePicker.Left = 120
    $ProjectToDateDateTimePicker.MinDate = $Script:BeginningOfPayPeriod
    $ProjectToDateDateTimePicker.MaxDate = $Script:LastSelectableDate
    $ProjectToDateDateTimePicker.Top = 8
    $ProjectToDateDateTimePicker.Width = 281

    $AnnualGoalLabel = New-Object System.Windows.Forms.Label
    $AnnualGoalLabel.Name = "AnnualGoalLabel"
    $AnnualGoalLabel.AutoSize = $True
    $AnnualGoalLabel.Left = 120
    $AnnualGoalLabel.Text = "Annual Leave:"
    $AnnualGoalLabel.Top = 42
    
    $AnnualGoalNumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $AnnualGoalNumericUpDown.Name = "AnnualGoalNumericUpDown"
    $AnnualGoalNumericUpDown.Left = 207
    $AnnualGoalNumericUpDown.Maximum = $Script:MaximumAnnual
    $AnnualGoalNumericUpDown.Top = 40
    $AnnualGoalNumericUpDown.Width = 50

    $SickGoalLabel = New-Object System.Windows.Forms.Label
    $SickGoalLabel.Name = "SickGoalLabel"
    $SickGoalLabel.AutoSize = $True
    $SickGoalLabel.Left = 280
    $SickGoalLabel.Text = "Sick Leave:"
    $SickGoalLabel.Top = 42

    $SickGoalNumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $SickGoalNumericUpDown.Name = "SickGoalNumericUpDown"
    $SickGoalNumericUpDown.Left = 351
    $SickGoalNumericUpDown.Maximum = $Script:MaximumSick
    $SickGoalNumericUpDown.Top = 40
    $SickGoalNumericUpDown.Width = 50

    $ProjectToEndOfPayPeriodLabel = New-Object System.Windows.Forms.Label
    $ProjectToEndOfPayPeriodLabel.Name = "ProjectToEndOfPayPeriodLabel"
    $ProjectToEndOfPayPeriodLabel.AutoSize = $True
    $ProjectToEndOfPayPeriodLabel.Left = 501
    $ProjectToEndOfPayPeriodLabel.Top = 10

    $ProjectButton = New-Object System.Windows.Forms.Button
    $ProjectButton.Name = "ProjectButton"
    $ProjectButton.Left = 450
    $ProjectButton.Text = "Run Projection"
    $ProjectButton.Top = 40
    $ProjectButton.Width = 400
    
    $HelpButton = New-Object System.Windows.Forms.Button
    $HelpButton.Name = "HelpButton"
    $HelpButton.Font = $Script:HelpIconFont
    $HelpButton.ForeColor = "Blue"
    $HelpButton.Height = 45
    $HelpButton.Left = 880
    $HelpButton.Text = [char]0xE9CE #? in Circle Symbol
    $HelpButton.Top = 13
    $HelpButton.Width = 45
    
    $LeavePanel = New-Object System.Windows.Forms.Panel
    $LeavePanel.Name = "LeavePanel"
    $LeavePanel.Dock = "Fill"
    $LeavePanel.TabIndex = 2
    
    $DataAsOfDateLabel = New-Object System.Windows.Forms.Label
    $DataAsOfDateLabel.Name = "DataAsOfDateLabel"
    $DataAsOfDateLabel.AutoSize = $True
    $DataAsOfDateLabel.Font = New-Object System.Drawing.Font($Script:FormFont.Name, ($Script:FormFont.Size + 2), [System.Drawing.FontStyle]::Bold)
    $DataAsOfDateLabel.Left = 308
    $DataAsOfDateLabel.Text = "Input Data as of Pay Period Ending: " + $Script:BeginningOfPayPeriod.AddDays(-1).ToString("MM/dd/yyyy") #Subtract one day so it's the end of the previous pay period matching what's in MyPay.
    $DataAsOfDateLabel.Top = 10

    $LeaveBalancesLabel = New-Object System.Windows.Forms.Label
    $LeaveBalancesLabel.Name = "LeaveBalancesLabel"
    $LeaveBalancesLabel.AutoSize = $True
    $LeaveBalancesLabel.Font = New-Object System.Drawing.Font($Script:FormFont.Name, $Script:FormFont.Size, [System.Drawing.FontStyle]::Bold)
    $LeaveBalancesLabel.Left = 204
    $LeaveBalancesLabel.Text = "Leave Balances"
    $LeaveBalancesLabel.Top = 40

    $ProjectedLeaveLabel = New-Object System.Windows.Forms.Label
    $ProjectedLeaveLabel.Name = "ProjectedLeaveLabel"
    $ProjectedLeaveLabel.AutoSize = $True
    $ProjectedLeaveLabel.Font = New-Object System.Drawing.Font($Script:FormFont.Name, $Script:FormFont.Size, [System.Drawing.FontStyle]::Bold)
    $ProjectedLeaveLabel.Left = 640
    $ProjectedLeaveLabel.Text = "Projected Leave"
    $ProjectedLeaveLabel.Top = 40

    $BalanceAddButton = New-Object System.Windows.Forms.Button
    $BalanceAddButton.Name = "BalanceAddButton"
    $BalanceAddButton.AutoSize = $True
    $BalanceAddButton.Font = $Script:IconsFont
    $BalanceAddButton.ForeColor = "Green"
    $BalanceAddButton.Left = 139
    $BalanceAddButton.Text = [char]0xF8AA #+ Symbol
    $BalanceAddButton.Top = 225

    $BalanceEditButton = New-Object System.Windows.Forms.Button
    $BalanceEditButton.Name = "BalanceEditButton"
    $BalanceEditButton.AutoSize = $True
    $BalanceEditButton.Font = $Script:IconsFont
    $BalanceEditButton.ForeColor = "Orange"
    $BalanceEditButton.Left = 214
    $BalanceEditButton.Text = [char]0xE70F #Pencil/Edit Symbol
    $BalanceEditButton.Top = 225
    
    $BalanceDeleteButton = New-Object System.Windows.Forms.Button
    $BalanceDeleteButton.Name = "BalanceDeleteButton"
    $BalanceDeleteButton.AutoSize = $True
    $BalanceDeleteButton.Font = $Script:IconsFont
    $BalanceDeleteButton.ForeColor = "Red"
    $BalanceDeleteButton.Left = 289
    $BalanceDeleteButton.Text = [char]0xF78A #X Symbol
    $BalanceDeleteButton.Top = 225
    
    $ProjectedAddButton = New-Object System.Windows.Forms.Button
    $ProjectedAddButton.Name = "ProjectedAddButton"
    $ProjectedAddButton.AutoSize = $True
    $ProjectedAddButton.Font = $Script:IconsFont
    $ProjectedAddButton.ForeColor = "Green"
    $ProjectedAddButton.Left = 578
    $ProjectedAddButton.Text = [char]0xF8AA #+ Symbol
    $ProjectedAddButton.Top = 225
    
    $ProjectedEditButton = New-Object System.Windows.Forms.Button
    $ProjectedEditButton.Name = "ProjectedEditButton"
    $ProjectedEditButton.AutoSize = $True
    $ProjectedEditButton.Enabled = $False
    $ProjectedEditButton.Font = $Script:IconsFont
    $ProjectedEditButton.ForeColor = "Orange"
    $ProjectedEditButton.Left = 653
    $ProjectedEditButton.Text = [char]0xE70F #Pencil/Edit Symbol
    $ProjectedEditButton.Top = 225
    
    $ProjectedDeleteButton = New-Object System.Windows.Forms.Button
    $ProjectedDeleteButton.Name = "ProjectedDeleteButton"
    $ProjectedDeleteButton.AutoSize = $True
    $ProjectedDeleteButton.Enabled = $False
    $ProjectedDeleteButton.Font = $Script:IconsFont
    $ProjectedDeleteButton.ForeColor = "Red"
    $ProjectedDeleteButton.Left = 728
    $ProjectedDeleteButton.Text = [char]0xF78A #X Symbol
    $ProjectedDeleteButton.Top = 225
    
    $BalanceListBox = New-Object System.Windows.Forms.ListBox
    $BalanceListBox.Name = "BalanceListBox"
    $BalanceListBox.Height = 155
    $BalanceListBox.IntegralHeight = $False
    $BalanceListBox.Left = 64
    $BalanceListBox.Top = 60
    $BalanceListBox.Width = 375

    $ProjectedListBox = New-Object System.Windows.Forms.CheckedListBox
    $ProjectedListBox.Name = "ProjectedListBox"
    $ProjectedListBox.CheckOnClick = $True
    $ProjectedListBox.Height = 155
    $ProjectedListBox.IntegralHeight = $False
    $ProjectedListBox.Left = 503
    $ProjectedListBox.Top = 60
    $ProjectedListBox.Width = 375

    $MainForm.Controls.AddRange(($LeavePanel, $SettingsPanel, $ReportPanel))
    $SettingsPanel.Controls.AddRange(($SCDLeaveDateLabel, $SCDLeaveDateDateTimePicker, $UpdateInfoButton, $LengthOfServiceTextBox, $EmployeeTypeTextBox, $DisplayBalanceEveryLeaveCheckBox, $DisplayBalanceEveryPayPeriodEnd, $DisplayLeaveHighsAndLows))
    $ReportPanel.Controls.AddRange(($ProjectBalanceRadioButton, $ReachGoalRadioButton, $ProjectToDateDateTimePicker, $AnnualGoalLabel, $AnnualGoalNumericUpDown, $SickGoalLabel, $SickGoalNumericUpDown, $ProjectToEndOfPayPeriodLabel, $ProjectButton, $HelpButton))
    $LeavePanel.Controls.AddRange(($DataAsOfDateLabel, $LeaveBalancesLabel, $ProjectedLeaveLabel, $BalanceListBox, $BalanceAddButton, $BalanceEditButton, $BalanceDeleteButton, $ProjectedListBox, $ProjectedAddButton, $ProjectedEditButton, $ProjectedDeleteButton))
    
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

    $ProjectToEndOfPayPeriodLabel.Text = "Will Project Through Pay Period Ending: " + (GetEndingOfPayPeriodForDate -Date $Script:ProjectToDate).ToString("MM/dd/yyyy")
    
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
    $MainForm.Add_KeyDown({MainFormKeyDown -EventArguments $_})
    $UpdateInfoButton.Add_Click({MainFormUpdateInfoButtonClick})
    $LengthOfServiceTextBox.Add_Click({MainFormLengthOfServiceTextBoxClick})
    $ProjectBalanceRadioButton.Add_Click({MainFormProjectBalanceRadioButtonClick})
    $ReachGoalRadioButton.Add_Click({MainFormReachGoalRadioButtonClick})
    $ProjectButton.Add_Click({MainFormProjectButtonClick})
    $HelpButton.Add_Click({MainFormHelpButtonClick})
    $BalanceAddButton.Add_Click({MainFormBalanceAddButtonClick})
    $BalanceEditButton.Add_Click({MainFormBalanceEditButtonClick})
    $BalanceDeleteButton.Add_Click({MainFormBalanceDeleteButtonClick})
    $ProjectedAddButton.Add_Click({MainFormProjectedAddButtonClick})
    $ProjectedEditButton.Add_Click({MainFormProjectedEditButtonClick})
    $ProjectedDeleteButton.Add_Click({MainFormProjectedDeleteButtonClick})
    $SCDLeaveDateDateTimePicker.Add_ValueChanged({MainFormSCDLeaveDateTimePickerValueChanged})
    $BalanceListBox.Add_SelectedIndexChanged({MainFormLeaveBalanceListBoxIndexChanged})
    $BalanceListBox.Add_DoubleClick({MainFormLeaveBalanceListBoxDoubleClick -EventArguments $_})
    $BalanceListBox.Add_KeyDown({MainFormBalanceListBoxKeyDown -EventArguments $_})
    $ProjectedListBox.Add_Click({MainFormProjectedLeaveListBoxClick -EventArguments $_})
    $ProjectedListBox.Add_DoubleClick({MainFormProjectedLeaveListBoxDoubleClick -EventArguments $_})
    $ProjectedListBox.Add_ItemCheck({MainFormProjectedLeaveListBoxItemCheck -EventArguments $_})
    $ProjectedListBox.Add_KeyDown({MainFormProjectedListBoxKeyDown -EventArguments $_})
    $ProjectToDateDateTimePicker.Add_ValueChanged({MainFormProjectToDateValueChanged})
    $AnnualGoalNumericUpDown.Add_ValueChanged({MainFormAnnualGoalValueChanged})
    $SickGoalNumericUpDown.Add_ValueChanged({MainFormSickGoalValueChanged})
    $DisplayBalanceEveryLeaveCheckBox.Add_Click({MainFormEveryLeaveCheckBoxClicked})
    $DisplayBalanceEveryPayPeriodEnd.Add_Click({MainFormEveryPPCheckBoxClicked})
    $DisplayLeaveHighsAndLows.Add_Click({MainFormDisplayHighsLowsClicked})
    $AnnualGoalNumericUpDown.Add_MouseClick({NumericUpDownMouseClick -Sender $AnnualGoalNumericUpDown -EventArguments $_}.GetNewClosure())
    $AnnualGoalNumericUpDown.Add_Enter({NumericUpDownEnter -Sender $AnnualGoalNumericUpDown -EventArguments $_}.GetNewClosure())
    $SickGoalNumericUpDown.Add_MouseClick({NumericUpDownMouseClick -Sender $SickGoalNumericUpDown -EventArguments $_}.GetNewClosure())
    $SickGoalNumericUpDown.Add_Enter({NumericUpDownEnter -Sender $SickGoalNumericUpDown -EventArguments $_}.GetNewClosure())
    
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
    $EmployeeInfoForm.KeyPreview      = $True
    $EmployeeInfoForm.MaximizeBox     = $False
    $EmployeeInfoForm.Size            = New-Object System.Drawing.Size(305, 471)
    $EmployeeInfoForm.StartPosition   = "CenterParent"
    $EmployeeInfoForm.Text            = "Employee Information"
    $EmployeeInfoForm.WindowState     = "Normal"

    $EmployeeInfoOkButton = New-Object System.Windows.Forms.Button
    $EmployeeInfoOkButton.Name = "EmployeeInfoOkButton"
    $EmployeeInfoOkButton.Left = 49
    $EmployeeInfoOkButton.Text = "OK"
    $EmployeeInfoOkButton.Top = 397
    
    $EmployeeInfoCancelButton = New-Object System.Windows.Forms.Button
    $EmployeeInfoCancelButton.Name = "EmployeeInfoCancelButton"
    $EmployeeInfoCancelButton.Left = 162
    $EmployeeInfoCancelButton.Text = "Cancel"
    $EmployeeInfoCancelButton.Top = 397
    
    $EmployeeTypePanel = New-Object System.Windows.Forms.Panel
    $EmployeeTypePanel.Name = "EmployeeTypePanel"
    $EmployeeTypePanel.Height = 76
    $EmployeeTypePanel.Width = 110

    $FullTimeRadioButton = New-Object System.Windows.Forms.RadioButton
    $FullTimeRadioButton.Name = "FullTimeRadioButton"
    $FullTimeRadioButton.AutoSize = $True
    $FullTimeRadioButton.Text = "Full-Time"
    $FullTimeRadioButton.Top = 15

    $PartTimeRadioButton = New-Object System.Windows.Forms.RadioButton
    $PartTimeRadioButton.Name = "PartTimeRadioButton"
    $PartTimeRadioButton.AutoSize = $True
    $PartTimeRadioButton.Text = "Part-Time"
    $PartTimeRadioButton.Top = 33

    $SESRadioButton = New-Object System.Windows.Forms.RadioButton
    $SESRadioButton.Name = "SESRadioButton"
    $SESRadioButton.AutoSize = $True
    $SESRadioButton.Text = "SES"
    $SESRadioButton.Top = 51

    $EmploymentTypeLabel = New-Object System.Windows.Forms.Label
    $EmploymentTypeLabel.Name = "EmploymentTypeLabel"
    $EmploymentTypeLabel.AutoSize = $True
    $EmploymentTypeLabel.Text = "Employment Type:"

    $LeaveCeilingPanel = New-Object System.Windows.Forms.Panel
    $LeaveCeilingPanel.Name = "LeaveCeilingPanel"
    $LeaveCeilingPanel.Height = 76
    $LeaveCeilingPanel.Left = 130
    $LeaveCeilingPanel.Width = 140

    $LeaveCeilingLabel = New-Object System.Windows.Forms.Label
    $LeaveCeilingLabel.Name = "LeaveCeilingLabel"
    $LeaveCeilingLabel.AutoSize = $True
    $LeaveCeilingLabel.Text = "Leave Ceiling:"

    $CONUSRadioButton = New-Object System.Windows.Forms.RadioButton
    $CONUSRadioButton.Name = "CONUSRadioButton"
    $CONUSRadioButton.AutoSize = $True
    $CONUSRadioButton.Text = "CONUS    (240 Hours)"
    $CONUSRadioButton.Top = 15

    $OCONUSRadioButton = New-Object System.Windows.Forms.RadioButton
    $OCONUSRadioButton.Name = "OCONUSRadioButton"
    $OCONUSRadioButton.AutoSize = $True
    $OCONUSRadioButton.Text = "OCONUS (360 Hours)"
    $OCONUSRadioButton.Top = 33

    $SESCeilingRadioButton = New-Object System.Windows.Forms.RadioButton
    $SESCeilingRadioButton.Name = "SESCeilingRadioButton"
    $SESCeilingRadioButton.AutoSize = $True
    $SESCeilingRadioButton.Text = "SES           (720 Hours)"
    $SESCeilingRadioButton.Top = 51

    $InaugurationDayHolidayCheckBox = New-Object System.Windows.Forms.CheckBox
    $InaugurationDayHolidayCheckBox.Name = "InaugurationDayHolidayCheckBox"
    $InaugurationDayHolidayCheckBox.AutoSize = $True
    $InaugurationDayHolidayCheckBox.Left = 20
    $InaugurationDayHolidayCheckBox.Text = "Entitled to a Holiday on Inauguration Day"
    $InaugurationDayHolidayCheckBox.Top = 78
    
    $Week1Label = New-Object System.Windows.Forms.Label
    $Week1Label.Name = "Week1Label"
    $Week1Label.AutoSize = $True
    $Week1Label.Left = 72
    $Week1Label.Text = "Week 1:"
    $Week1Label.Top = 115

    $Week2Label = New-Object System.Windows.Forms.Label
    $Week2Label.Name = "Week2Label"
    $Week2Label.AutoSize = $True
    $Week2Label.Left = 175
    $Week2Label.Text = "Week 2:"
    $Week2Label.Top = 115

    $SundayLabel = New-Object System.Windows.Forms.Label
    $SundayLabel.Name = "SundayLabel"
    $SundayLabel.AutoSize = $True
    $SundayLabel.Text = "Sunday:"
    $SundayLabel.Top = 138

    $MondayLabel = New-Object System.Windows.Forms.Label
    $MondayLabel.Name = "MondayLabel"
    $MondayLabel.AutoSize = $True
    $MondayLabel.Text = "Monday:"
    $MondayLabel.Top = 169

    $TuesdayLabel = New-Object System.Windows.Forms.Label
    $TuesdayLabel.Name = "TuesdayLabel"
    $TuesdayLabel.AutoSize = $True
    $TuesdayLabel.Text = "Tuesday:"
    $TuesdayLabel.Top = 200

    $WednesdayLabel = New-Object System.Windows.Forms.Label
    $WednesdayLabel.Name = "WednesdayLabel"
    $WednesdayLabel.AutoSize = $True
    $WednesdayLabel.Text = "Wednesday:"
    $WednesdayLabel.Top = 231

    $ThursdayLabel = New-Object System.Windows.Forms.Label
    $ThursdayLabel.Name = "ThursdayLabel"
    $ThursdayLabel.AutoSize = $True
    $ThursdayLabel.Text = "Thursday:"
    $ThursdayLabel.Top = 262

    $FridayLabel = New-Object System.Windows.Forms.Label
    $FridayLabel.Name = "FridayLabel"
    $FridayLabel.AutoSize = $True
    $FridayLabel.Text = "Friday:"
    $FridayLabel.Top = 293

    $SaturdayLabel = New-Object System.Windows.Forms.Label
    $SaturdayLabel.Name = "SaturdayLabel"
    $SaturdayLabel.AutoSize = $True
    $SaturdayLabel.Text = "Saturday:"
    $SaturdayLabel.Top = 324

    $Day1NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day1NumericUpDown.Name = "Day1NumericUpDown"
    $Day1NumericUpDown.Height = 20
    $Day1NumericUpDown.Left = 76
    $Day1NumericUpDown.Maximum = 24
    $Day1NumericUpDown.Top = 135
    $Day1NumericUpDown.Width = 40

    $Day2NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day2NumericUpDown.Name = "Day2NumericUpDown"
    $Day2NumericUpDown.Height = 20
    $Day2NumericUpDown.Left = 76
    $Day2NumericUpDown.Maximum = 24
    $Day2NumericUpDown.Top = 166
    $Day2NumericUpDown.Width = 40

    $Day3NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day3NumericUpDown.Name = "Day3NumericUpDown"
    $Day3NumericUpDown.Height = 20
    $Day3NumericUpDown.Left = 76
    $Day3NumericUpDown.Maximum = 24
    $Day3NumericUpDown.Top = 197
    $Day3NumericUpDown.Width = 40

    $Day4NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day4NumericUpDown.Name = "Day4NumericUpDown"
    $Day4NumericUpDown.Height = 20
    $Day4NumericUpDown.Left = 76
    $Day4NumericUpDown.Maximum = 24
    $Day4NumericUpDown.Top = 228
    $Day4NumericUpDown.Width = 40

    $Day5NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day5NumericUpDown.Name = "Day5NumericUpDown"
    $Day5NumericUpDown.Height = 20
    $Day5NumericUpDown.Left = 76
    $Day5NumericUpDown.Maximum = 24
    $Day5NumericUpDown.Top = 259
    $Day5NumericUpDown.Width = 40

    $Day6NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day6NumericUpDown.Name = "Day6NumericUpDown"
    $Day6NumericUpDown.Height = 20
    $Day6NumericUpDown.Left = 76
    $Day6NumericUpDown.Maximum = 24
    $Day6NumericUpDown.Top = 290
    $Day6NumericUpDown.Width = 40

    $Day7NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day7NumericUpDown.Name = "Day7NumericUpDown"
    $Day7NumericUpDown.Height = 20
    $Day7NumericUpDown.Left = 76
    $Day7NumericUpDown.Maximum = 24
    $Day7NumericUpDown.Top = 321
    $Day7NumericUpDown.Width = 40
    
    $Day8NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day8NumericUpDown.Name = "Day8NumericUpDown"
    $Day8NumericUpDown.Height = 20
    $Day8NumericUpDown.Left = 180
    $Day8NumericUpDown.Maximum = 24
    $Day8NumericUpDown.Top = 135
    $Day8NumericUpDown.Width = 40

    $Day9NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day9NumericUpDown.Name = "Day9NumericUpDown"
    $Day9NumericUpDown.Height = 20
    $Day9NumericUpDown.Left = 180
    $Day9NumericUpDown.Maximum = 24
    $Day9NumericUpDown.Top = 166
    $Day9NumericUpDown.Width = 40

    $Day10NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day10NumericUpDown.Name = "Day10NumericUpDown"
    $Day10NumericUpDown.Height = 20
    $Day10NumericUpDown.Left = 180
    $Day10NumericUpDown.Maximum = 24
    $Day10NumericUpDown.Top = 197
    $Day10NumericUpDown.Width = 40

    $Day11NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day11NumericUpDown.Name = "Day11NumericUpDown"
    $Day11NumericUpDown.Height = 20
    $Day11NumericUpDown.Left = 180
    $Day11NumericUpDown.Maximum = 24
    $Day11NumericUpDown.Top = 228
    $Day11NumericUpDown.Width = 40

    $Day12NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day12NumericUpDown.Name = "Day12NumericUpDown"
    $Day12NumericUpDown.Height = 20
    $Day12NumericUpDown.Left = 180
    $Day12NumericUpDown.Maximum = 24
    $Day12NumericUpDown.Top = 259
    $Day12NumericUpDown.Width = 40

    $Day13NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day13NumericUpDown.Name = "Day13NumericUpDown"
    $Day13NumericUpDown.Height = 20
    $Day13NumericUpDown.Left = 180
    $Day13NumericUpDown.Maximum = 24
    $Day13NumericUpDown.Top = 290
    $Day13NumericUpDown.Width = 40

    $Day14NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $Day14NumericUpDown.Name = "Day14NumericUpDown"
    $Day14NumericUpDown.Height = 20
    $Day14NumericUpDown.Left = 180
    $Day14NumericUpDown.Maximum = 24
    $Day14NumericUpDown.Top = 321
    $Day14NumericUpDown.Width = 40

    $HoursWorkedLabel = New-Object System.Windows.Forms.Label
    $HoursWorkedLabel.Name = "HoursWorkedLabel"
    $HoursWorkedLabel.AutoSize = $True
    $HoursWorkedLabel.Top = 355
    $HoursWorkedLabel.Text = "Hours Per Pay Period: " + $Script:WorkHoursPerPayPeriod

    $UnusualHoursLabel = New-Object System.Windows.Forms.Label
    $UnusualHoursLabel.Name = "UnusualHoursLabel"
    $UnusualHoursLabel.AutoSize = $True
    $UnusualHoursLabel.ForeColor = "Blue"
    $UnusualHoursLabel.Text = "Please validate unusual hours."
    $UnusualHoursLabel.Top = 375
    $UnusualHoursLabel.Visible = $False

    $EmployeeInfoForm.Controls.AddRange(($EmployeeInfoOkButton, $EmployeeInfoCancelButton, $EmployeeTypePanel, $LeaveCeilingPanel, $InaugurationDayHolidayCheckBox, $Week1Label, $Week2Label, $SundayLabel, $MondayLabel, $TuesdayLabel, $WednesdayLabel, $ThursdayLabel, $FridayLabel, $SaturdayLabel, $Day1NumericUpDown, $Day2NumericUpDown, $Day3NumericUpDown, $Day4NumericUpDown, $Day5NumericUpDown, $Day6NumericUpDown, $Day7NumericUpDown, $Day8NumericUpDown, $Day9NumericUpDown, $Day10NumericUpDown, $Day11NumericUpDown, $Day12NumericUpDown, $Day13NumericUpDown, $Day14NumericUpDown, $HoursWorkedLabel, $UnusualHoursLabel))
    $EmployeeTypePanel.Controls.AddRange(($FullTimeRadioButton, $PartTimeRadioButton, $SESRadioButton, $EmploymentTypeLabel))
    $LeaveCeilingPanel.Controls.AddRange(($CONUSRadioButton, $OCONUSRadioButton, $SESCeilingRadioButton, $LeaveCeilingLabel))

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
    $EmployeeInfoForm.Add_KeyDown({EmployeeInfoFormKeyDown -EventArguments $_})
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

    $Day1NumericUpDown.Add_MouseClick({NumericUpDownMouseClick -Sender $Day1NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day1NumericUpDown.Add_Enter({NumericUpDownEnter -Sender $Day1NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day2NumericUpDown.Add_MouseClick({NumericUpDownMouseClick -Sender $Day2NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day2NumericUpDown.Add_Enter({NumericUpDownEnter -Sender $Day2NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day3NumericUpDown.Add_MouseClick({NumericUpDownMouseClick -Sender $Day3NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day3NumericUpDown.Add_Enter({NumericUpDownEnter -Sender $Day3NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day4NumericUpDown.Add_MouseClick({NumericUpDownMouseClick -Sender $Day4NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day4NumericUpDown.Add_Enter({NumericUpDownEnter -Sender $Day4NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day5NumericUpDown.Add_MouseClick({NumericUpDownMouseClick -Sender $Day5NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day5NumericUpDown.Add_Enter({NumericUpDownEnter -Sender $Day5NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day6NumericUpDown.Add_MouseClick({NumericUpDownMouseClick -Sender $Day6NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day6NumericUpDown.Add_Enter({NumericUpDownEnter -Sender $Day6NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day7NumericUpDown.Add_MouseClick({NumericUpDownMouseClick -Sender $Day7NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day7NumericUpDown.Add_Enter({NumericUpDownEnter -Sender $Day7NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day8NumericUpDown.Add_MouseClick({NumericUpDownMouseClick -Sender $Day8NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day8NumericUpDown.Add_Enter({NumericUpDownEnter -Sender $Day8NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day9NumericUpDown.Add_MouseClick({NumericUpDownMouseClick -Sender $Day9NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day9NumericUpDown.Add_Enter({NumericUpDownEnter -Sender $Day9NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day10NumericUpDown.Add_MouseClick({NumericUpDownMouseClick -Sender $Day10NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day10NumericUpDown.Add_Enter({NumericUpDownEnter -Sender $Day10NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day11NumericUpDown.Add_MouseClick({NumericUpDownMouseClick -Sender $Day11NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day11NumericUpDown.Add_Enter({NumericUpDownEnter -Sender $Day11NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day12NumericUpDown.Add_MouseClick({NumericUpDownMouseClick -Sender $Day12NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day12NumericUpDown.Add_Enter({NumericUpDownEnter -Sender $Day12NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day13NumericUpDown.Add_MouseClick({NumericUpDownMouseClick -Sender $Day13NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day13NumericUpDown.Add_Enter({NumericUpDownEnter -Sender $Day13NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day14NumericUpDown.Add_MouseClick({NumericUpDownMouseClick -Sender $Day14NumericUpDown -EventArguments $_}.GetNewClosure())
    $Day14NumericUpDown.Add_Enter({NumericUpDownEnter -Sender $Day14NumericUpDown -EventArguments $_}.GetNewClosure())
}

function BuildEditLeaveBalanceForm
{
    $SelectedLeaveBalance = $Script:LeaveBalances[$Script:MainForm.Controls["LeavePanel"].Controls["BalanceListBox"].SelectedIndex]
    
    $Script:EditLeaveBalanceForm          = New-Object System.Windows.Forms.Form
    $EditLeaveBalanceForm.Name            = "EditLeaveBalanceForm"
    $EditLeaveBalanceForm.BackColor       = "WhiteSmoke"
    $EditLeaveBalanceForm.Font            = $Script:FormFont
    $EditLeaveBalanceForm.FormBorderStyle = "FixedSingle"
    $EditLeaveBalanceForm.KeyPreview      = $True
    $EditLeaveBalanceForm.MaximizeBox     = $False
    $EditLeaveBalanceForm.Size            = New-Object System.Drawing.Size(260, 204)
    $EditLeaveBalanceForm.StartPosition   = "CenterParent"
    $EditLeaveBalanceForm.Text            = "Edit Leave"
    $EditLeaveBalanceForm.WindowState     = "Normal"

    $LeaveBalanceNameLabel = New-Object System.Windows.Forms.Label
    $LeaveBalanceNameLabel.Name = "LeaveBalanceNameLabel"
    $LeaveBalanceNameLabel.AutoSize = $True
    $LeaveBalanceNameLabel.Left = 13
    $LeaveBalanceNameLabel.Text = "Type of Leave:"
    $LeaveBalanceNameLabel.Top = 14

    $LeaveBalanceNameTextBox = New-Object System.Windows.Forms.TextBox
    $LeaveBalanceNameTextBox.Name = "LeaveBalanceNameTextBox"
    $LeaveBalanceNameTextBox.Height = 20
    $LeaveBalanceNameTextBox.Left = 110
    $LeaveBalanceNameTextBox.Text = $SelectedLeaveBalance.Name
    $LeaveBalanceNameTextBox.Top = 11
    $LeaveBalanceNameTextBox.Width = 100

    $BalanceLabel = New-Object System.Windows.Forms.Label
    $BalanceLabel.Name = "BalanceLabel"
    $BalanceLabel.AutoSize = $True
    $BalanceLabel.Left = 13
    $BalanceLabel.Text = "Balance:"
    $BalanceLabel.Top = 41

    $BalanceNumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $BalanceNumericUpDown.Name = "BalanceNumericUpDown"
    $BalanceNumericUpDown.Height = 20
    $BalanceNumericUpDown.Left = 110
    $BalanceNumericUpDown.Maximum = $Script:MaximumSick #Largest feasible amount for any type of leave, not just sick. Adjusted if type is Annual leave.
    $BalanceNumericUpDown.Top = 38
    $BalanceNumericUpDown.Value = $SelectedLeaveBalance.Balance
    $BalanceNumericUpDown.Width = 54

    $AlertThresholdLabel = New-Object System.Windows.Forms.Label
    $AlertThresholdLabel.Name = "AlertThresholdLabel"
    $AlertThresholdLabel.AutoSize = $True
    $AlertThresholdLabel.Left = 13
    $AlertThresholdLabel.Text = "Alert Threshold:"
    $AlertThresholdLabel.Top = 68

    $ThresholdNumericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $ThresholdNumericUpDown.Name = "ThresholdNumericUpDown"
    $ThresholdNumericUpDown.Height = 20
    $ThresholdNumericUpDown.Left = 110
    $ThresholdNumericUpDown.Maximum = $Script:MaximumSick #Adjusted if type is Annual leave.
    $ThresholdNumericUpDown.Top = 65
    $ThresholdNumericUpDown.Width = 54

    $LeaveExpiresCheckBox = New-Object System.Windows.Forms.CheckBox
    $LeaveExpiresCheckBox.Name = "LeaveExpiresCheckBox"
    $LeaveExpiresCheckBox.Left = 15
    $LeaveExpiresCheckBox.Text = "Leave Expires"
    $LeaveExpiresCheckBox.Top = 60
    $LeaveExpiresCheckBox.Width = 104
    
    $LeaveExpiresOnLabel = New-Object System.Windows.Forms.Label
    $LeaveExpiresOnLabel.Name = "LeaveExpiresOnLabel"
    $LeaveExpiresOnLabel.AutoSize = $True
    $LeaveExpiresOnLabel.Left = 13
    $LeaveExpiresOnLabel.Text = "Expires On:"
    $LeaveExpiresOnLabel.Top = 89
    $LeaveExpiresOnLabel.Visible = $False
    
    $LeaveExpiresOnDateTimePicker = New-Object System.Windows.Forms.DateTimePicker
    $LeaveExpiresOnDateTimePicker.Name = "LeaveExpiresOnDateTimePicker"
    $LeaveExpiresOnDateTimePicker.Enabled = $False
    $LeaveExpiresOnDateTimePicker.Format = "Short"
    $LeaveExpiresOnDateTimePicker.Left = 110
    $LeaveExpiresOnDateTimePicker.MinDate = $Script:BeginningOfPayPeriod
    $LeaveExpiresOnDateTimePicker.Top = 85
    $LeaveExpiresOnDateTimePicker.Visible = $False
    $LeaveExpiresOnDateTimePicker.Width = 100

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
    $EditLeaveOkButton.Left = 40
    $EditLeaveOkButton.Text = "OK"
    $EditLeaveOkButton.Top = 138
    $EditLeaveOkButton.Width = 75
    
    $EditLeaveCancelButton = New-Object System.Windows.Forms.Button
    $EditLeaveCancelButton.Name = "EditLeaveCancelButton"
    $EditLeaveCancelButton.Height = 24
    $EditLeaveCancelButton.Left = 130
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
    
    $EditLeaveBalanceForm.Add_FormClosing({EditLeaveBalanceFormClosing -EventArguments $_})
    $EditLeaveBalanceForm.Add_KeyDown({EditLeaveBalanceFormKeyDown -EventArguments $_})
    $LeaveBalanceNameTextBox.Add_TextChanged({EditLeaveBalanceFormCheckNameLength; EditLeaveBalanceFormNameOrDateChanged})
    $LeaveExpiresCheckBox.Add_Click({EditLeaveBalanceFormLeaveExpiresCheckBoxClick; EditLeaveBalanceFormNameOrDateChanged})
    $LeaveExpiresOnDateTimePicker.Add_Valuechanged({EditLeaveBalanceFormNameOrDateChanged})
    $EditLeaveOkButton.Add_Click({EditLeaveBalanceFormOkButtonClick})
    $EditLeaveCancelButton.Add_Click({EditLeaveBalanceFormCancelButtonClick})
    $BalanceNumericUpDown.Add_MouseClick({NumericUpDownMouseClick -Sender $BalanceNumericUpDown -EventArguments $_}.GetNewClosure())
    $BalanceNumericUpDown.Add_Enter({NumericUpDownEnter -Sender $BalanceNumericUpDown -EventArguments $_}.GetNewClosure())
    $ThresholdNumericUpDown.Add_MouseClick({NumericUpDownMouseClick -Sender $ThresholdNumericUpDown -EventArguments $_}.GetNewClosure())
    $ThresholdNumericUpDown.Add_Enter({NumericUpDownEnter -Sender $ThresholdNumericUpDown -EventArguments $_}.GetNewClosure())

    #This happens after the events so it'll fire.
    if($LeaveBalanceNameTextBox.Text -ne "")
    {
        $EditLeaveBalanceForm.ActiveControl = $BalanceNumericUpDown
    }
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
    $EditProjectedLeaveForm.KeyPreview      = $True
    $EditProjectedLeaveForm.MaximizeBox     = $False
    $EditProjectedLeaveForm.Size            = New-Object System.Drawing.Size(475, 400)
    $EditProjectedLeaveForm.StartPosition   = "CenterParent"
    $EditProjectedLeaveForm.Text            = "Edit Projected Leave"
    $EditProjectedLeaveForm.WindowState     = "Normal"

    $LeaveStartLabel = New-Object System.Windows.Forms.Label
    $LeaveStartLabel.Name = "LeaveStartLabel"
    $LeaveStartLabel.AutoSize = $True
    $LeaveStartLabel.Left = 15
    $LeaveStartLabel.Text = "Leave Start:"
    $LeaveStartLabel.Top = 41

    $LeaveEndLabel = New-Object System.Windows.Forms.Label
    $LeaveEndLabel.Name = "LeaveEndLabel"
    $LeaveEndLabel.AutoSize = $True
    $LeaveEndLabel.Left = 15
    $LeaveEndLabel.Text = "Leave End:"
    $LeaveEndLabel.Top = 67

    $LeaveStartDateTimePicker = New-Object System.Windows.Forms.DateTimePicker
    $LeaveStartDateTimePicker.Name = "LeaveStartDateTimePicker"
    $LeaveStartDateTimePicker.Format = "Short"
    $LeaveStartDateTimePicker.Left = 100
    $LeaveStartDateTimePicker.Top = 38
    $LeaveStartDateTimePicker.Width = 115

    $LeaveEndDateTimePicker = New-Object System.Windows.Forms.DateTimePicker
    $LeaveEndDateTimePicker.Name = "LeaveEndDateTimePicker"
    $LeaveEndDateTimePicker.Format = "Short"
    $LeaveEndDateTimePicker.Left = 100
    $LeaveEndDateTimePicker.Top = 64
    $LeaveEndDateTimePicker.Width = 115

    $LeaveBankLabel = New-Object System.Windows.Forms.Label
    $LeaveBankLabel.Name = "LeaveBankLabel"
    $LeaveBankLabel.AutoSize = $True
    $LeaveBankLabel.Left = 15
    $LeaveBankLabel.Text = "Leave Bank:"
    $LeaveBankLabel.Top = 15

    $LeaveBankComboBox = New-Object System.Windows.Forms.ComboBox
    $LeaveBankComboBox.Name = "LeaveBankComboBox"
    $LeaveBankComboBox.DropDownStyle = "DropDownList"
    $LeaveBankComboBox.Tag = @{}
    $LeaveBankComboBox.Top = 12
    $LeaveBankComboBox.Left = 100
    $LeaveBankComboBox.Width = 115

    $LeaveDatesFlowLayoutPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $LeaveDatesFlowLayoutPanel.Name = "LeaveDatesFlowLayoutPanel"
    $LeaveDatesFlowLayoutPanel.AutoScroll = $True
    $LeaveDatesFlowLayoutPanel.FlowDirection = "TopDown"
    $LeaveDatesFlowLayoutPanel.Height = 198
    $LeaveDatesFlowLayoutPanel.Left = 15
    $LeaveDatesFlowLayoutPanel.Top = 100
    $LeaveDatesFlowLayoutPanel.Width = 440
    $LeaveDatesFlowLayoutPanel.WrapContents = $False

    $HoursOfLeaveTaken = New-Object System.Windows.Forms.Label
    $HoursOfLeaveTaken.Name = "HoursOfLeaveTaken"
    $HoursOfLeaveTaken.AutoSize = $True
    $HoursOfLeaveTaken.Left = 15
    $HoursOfLeaveTaken.Top = 305
    
    $EditProjectedOkButton = New-Object System.Windows.Forms.Button
    $EditProjectedOkButton.Name = "EditProjectedOkButton"
    $EditProjectedOkButton.Height = 24
    $EditProjectedOkButton.Left = 130
    $EditProjectedOkButton.Text = "OK"
    $EditProjectedOkButton.Top = 330
    $EditProjectedOkButton.Width = 75
    
    $EditProjectedCancelButton = New-Object System.Windows.Forms.Button
    $EditProjectedCancelButton.Name = "EditProjectedCancelButton"
    $EditProjectedCancelButton.Height = 24
    $EditProjectedCancelButton.Left = 240
    $EditProjectedCancelButton.Text = "Cancel"
    $EditProjectedCancelButton.Top = 330
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
    $EditProjectedLeaveForm.Add_KeyDown({EditProjectedLeaveFormKeyDown -EventArguments $_})
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
    $OutputForm.KeyPreview      = $True
    $OutputForm.MaximizeBox     = $False
    $OutputForm.Size            = New-Object System.Drawing.Size(460, 503)
    $OutputForm.StartPosition   = "CenterParent"
    $OutputForm.Text            = "Projection Report"
    $OutputForm.WindowState     = "Normal"

    $OutputRichTextBox = New-Object System.Windows.Forms.RichTextBox
    $OutputRichTextBox.Name = "OutputRichTextBox"
    $OutputRichTextBox.Dock = "Top"
    $OutputRichTextBox.Height = 400
    $OutputRichTextBox.Multiline = $True
    $OutputRichTextBox.ReadOnly = $True
    $OutputRichTextBox.TabStop = $False
    
    $CopyButton = New-Object System.Windows.Forms.Button
    $CopyButton.Name = "CopyButton"
    $CopyButton.Height = 24
    $CopyButton.Left = 54
    $CopyButton.Text = "Copy"
    $CopyButton.Top = 404
    $CopyButton.Width = 336
    
    $CloseButton = New-Object System.Windows.Forms.Button
    $CloseButton.Name = "CloseButton"
    $CloseButton.Height = 24
    $CloseButton.Left = 54
    $CloseButton.Text = "Close"
    $CloseButton.Top = 437
    $CloseButton.Width = 336
    
    $OutputForm.Controls.AddRange(($OutputRichTextBox, $CopyButton, $CloseButton))
    
    $OutputForm.ActiveControl = $CloseButton
    
    PopulateOutputFormRichTextBox

    $OutputForm.Add_KeyDown({OutputFormKeyDown -EventArguments $_})
    $CopyButton.Add_Click({OutputFormCopyButton})
    $CloseButton.Add_Click({OutputFormCloseButton})
}

function BuildHelpForm
{
    $Script:HelpForm          = New-Object System.Windows.Forms.Form
    $HelpForm.Name            = "HelpForm"
    $HelpForm.BackColor       = "WhiteSmoke"
    $HelpForm.Font            = $Script:FormFont
    $HelpForm.FormBorderStyle = "FixedSingle"
    $HelpForm.KeyPreview      = $True
    $HelpForm.MaximizeBox     = $False
    $HelpForm.Size            = New-Object System.Drawing.Size(460, 503)
    $HelpForm.StartPosition   = "CenterParent"
    $HelpForm.Text            = "Help"
    $HelpForm.WindowState     = "Normal"

    $HelpRichTextBox = New-Object System.Windows.Forms.RichTextBox
    $HelpRichTextBox.Name = "HelpRichTextBox"
    $HelpRichTextBox.Dock = "Top"
    $HelpRichTextBox.Height = 430
    $HelpRichTextBox.Multiline = $True
    $HelpRichTextBox.ReadOnly = $True
    $HelpRichTextBox.TabStop = $False
    
    $CloseButton = New-Object System.Windows.Forms.Button
    $CloseButton.Name = "CloseButton"
    $CloseButton.Height = 24
    $CloseButton.Left = 54
    $CloseButton.Text = "Close"
    $CloseButton.Top = 437
    $CloseButton.Width = 336
    
    $HelpForm.Controls.AddRange(($HelpRichTextBox, $CloseButton))

    #Keyboard Shortcuts Section
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "Keyboard Shortcuts`n" -Alignment "Center" -FontSize 20 -FontStyle "Bold"
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n"
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "Enter – If in a sub-form and a button doesn’t have the focus, will generally hit the `“OK`” button."
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n`nEsc – If in a sub-form, will generally hit the `“Cancel`” button."
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n`nTab will cycle through the available controls. Shift-tab to go the opposite way."
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n`nArrow keys can be used to change some selections, and the space bar can be used to check/uncheck boxes."
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n`nF1: Show this help page."
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n`nF5: Run the projection with the current settings."
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n`nF6: Open Employee Info."
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n`nF7: Edit the selected Leave Balance item."
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n`nF8: Edit the selected Projected Leave item."
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n`nIf either the Leave Balance or Projected Leave list boxes have the focus, Enter will open the appropriate edit menu for the selected item, + will add a new item, and Del will delete the selected item. Double clicking an item in either list box will open the appropriate edit menu."
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n"

    #Program Features Section
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`nProgram Features`n" -Alignment "Center" -FontSize 20 -FontStyle "Bold"
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n"
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "Input your current leave balances, projected leave, the date you wish to project to or a goal you wish to reach, and the program will do the counting for you to help you manage and plan your leave. It allows you to keep track of when leave expires, as well as warns you if you’ll be over your Lose/Use ceiling at the end of the leave year or will forfeit leave. It has options to alert you if your annual or sick leave balance drops below a certain threshold. It also displays when your annual leave accrual rate will change if applicable. Supports Full-Time, Part-Time, and SES employees. The program also saves your data and automatically updates it based on the entered info."
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n"

    #Usage Section
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`nUsage`n" -Alignment "Center" -FontSize 20 -FontStyle "Bold"
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n"
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "This program does not submit leave requests on your behalf. It is a tool to help you easily calculate what your leave balances will be in the future after leave accrual and taking leave."
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n`nThis program is designed to do everything based on the pay period schedule. All information displayed on the main page is based on the data at the end of the last pay period and all output information shows to the end of the pay period. All dates are in the MM/DD/YYYY format. This program makes two web requests: One to the OPM Federal Holidays website (see Helpful Links below) to assist in filling out projected leave and the second to the GitHub repository for this program to check if there is an update."
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n`nUpon launching the program, you should configure your settings. Set your SCD Leave Date in the upper left (this can be found in your most recent SF-50 or LES), update your Employee Information by clicking the button. If your Employee Type is not SES and you are not in the 15+ year category, you can click the box showing your Length of Service to cycle through additional information. Your Employment Type can be found in Box 32 of your most recent SF-50. Your Leave Ceiling value can be found in your most recent LES under Max Leave Carry Over. See Helpful Links below to determine if you are entitled to a holiday on Inauguration Day. Fill out your current work schedule where Week 1 is the first week of a Pay Period and Week 2 is the second week."
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n`nCheck the appropriate boxes to have the program display the Leave Balances after each day that you have input that you are taking leave, after each Pay Period, or keep track of the highest and lowest value that your Annual and Sick leave reach during the projection period. These options can make the output quite verbose but give you a deeper understanding of your leave balances at any given time."
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n`nInput your Leave Balances as of the Pay Period ending listed in the program. You can find this from your most recent LES (you might need to wait a few days for this LES to be available). For Annual and Sick Leave, you may set an Alert Threshold. In the projection, if your balance drops below this value, it will alert you. Leave it set to 0 to disable this feature. This will work whether or not you have enabled displaying the Annual and Sick Leave highs and lows. For any other types of leave, select if the leave expires or not (e.g. award leave). If it expires, select the last date the leave is available for you to use."
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n`nInput your Projected Leave as of the Pay Period ending listed in the program. Select which Leave Bank you wish to use for this entry, the first date you will be using leave, and the last date you will be using leave. Please note, the Leave End date must be in the same pay period as the Leave Start date. If you expect to take leave across multiple pay periods, you must use multiple entries. After selecting your Leave Start/End dates, input the number of hours of leave you will take per day. These values will be automatically filled as whole days based on your work schedule input in the Employee Info. It will also automatically input 0 hours for holidays and show a note informing you of which holiday it is (It gets these holidays from the OPM website showing federal holidays. If this fails, it will warn you to manually verify the hours). You can always adjust this. By default, each entry will be checked. If you wish to quickly compare different projected leave schedules, you can uncheck the items you wish to be ignored in your calculations. Please remember, this is the input data as of the ending of the last pay period. Even if you have already taken leave since the end of the last pay period, include it here so the program is aware of it."
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n`nSelect which mode you wish the program to run in. Project to Date will calculate and show your balances through the end of the Pay Period for the date you select. Reach Goal will tell you what Pay Period your balance meets or exceeds the goals you set. Please note, in Reach Goal mode, the output will only include information about Annual and Sick Leave except for alerts such as going negative on a balance, falling below a set threshold, or forfeiting leave."
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n`nIf you are Part-Time, you might not accrue leave hours in whole numbers. The program keeps track of these partial hours and adds them to your balance, but it may not perfectly reflect your partial hour balance. To reset the partial hours, change your employee type to Full-Time or SES, close the program, and open it again."
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n"

    #Saving Your Data Section
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`nSaving Your Data`n" -Alignment "Center" -FontSize 20 -FontStyle "Bold"
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n"
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "If you close the program by clicking the X on the GUI, it will write a configuration file to `"$Script:ConfigFile`" which will save all your data. The next time you launch the program, it will calculate what your balance should be assuming you took all your projected leave (whether or not the projected leave item was checked) and update your balances and projected leave. If you do not wish for this configuration file to be written, close the program by clicking the X on the PowerShell window."
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n"

    #Helpful Links Section
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`nHelpful Links`n" -Alignment "Center" -FontSize 20 -FontStyle "Bold"
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n"
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "Annual Leave Fact Sheet: https://www.opm.gov/policy-data-oversight/pay-leave/leave-administration/fact-sheets/annual-leave/"
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n`nSick Leave Fact Sheet: https://www.opm.gov/policy-data-oversight/pay-leave/leave-administration/fact-sheets/sick-leave-general-information/"
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n`nLeave Year Beginning/Ending Dates: https://www.opm.gov/policy-data-oversight/pay-leave/leave-administration/fact-sheets/leave-year-beginning-and-ending-dates/"
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n`nHoliday Work Schedules and Pay (Are you entitled to a holiday on Inauguration Day?): https://www.opm.gov/policy-data-oversight/pay-leave/pay-administration/fact-sheets/holidays-work-schedules-and-pay/"
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n`nFederal Holidays: https://www.opm.gov/policy-data-oversight/pay-leave/federal-holidays/"
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n"

    #About Section
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`nAbout`n" -Alignment "Center" -FontSize 20 -FontStyle "Bold"
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n"
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "Author: $Script:Author"
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n`nVersion: $Script:Version"
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n`nLast Updated: $Script:DateUpdated"
    RichTextBoxAppendText -RichTextBox $HelpRichTextBox -Text "`n`nOnline Repository Link: $Script:RepoWebsite"
    
    $HelpForm.Add_KeyDown({HelpFormKeyDown -EventArguments $_})
    $HelpRichTextBox.Add_LinkClicked({Start-Process $_.LinkText}.GetNewClosure())
    $CloseButton.Add_Click({HelpFormCloseButton})
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

clear

Main
