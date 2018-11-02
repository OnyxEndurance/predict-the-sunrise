# ******************************************************
# *
# * Name:         grab-sun-rise-set-data.ps1
# *     
# * Design Phase:
# *     Author:   John Miner
# *     Date:     02-08-2018
# *     Purpose:  Grab sun rise and sun set data.
# *
# ******************************************************



#
#  1 - Get historical data
#


# Set working directory
Set-Location "c:\sunriseset"


# Grab 25 years of data
for($cnt = 0; $cnt -lt 25; $cnt++)
{

    # What year
    $year = 1993 + $cnt;

    # Output files
    $raw = 'c:\sunriseset\inbound\SUN-DATA-' + $year + '.TXT'
    $csv = 'c:\sunriseset\outbound\SUN-DATA-' + $year + '.CSV'


    #
    # Grab web data as html page
    #

    # Web site url
    $site = 'http://aa.usno.navy.mil/cgi-bin/aa_rstablew.pl?ID=AA&year=' + $Year + '&task=0&state=MA&place=Watertown'

    # Grab web page 
    $page = (New-Object System.Net.WebClient).DownloadString($site)
    

    #
    # Save data as space delimited
    #

    # Grab textual chart
    $start = $page.indexof('01  ')
    $end = $page.indexof('Add one hour')
    $len = $end-$start-5
    $data = $page.Substring($start, $len)
    
    # remove unwanted spaces
    $data = $data.Replace('      ', ' . ').replace('  ', ' ').replace('  ', ' ').replace('  ', ' ').replace('  ', ' ')

    # remove unwanted lines
    $output = ""
    ForEach ($line in $($data -split "`n"))
    {
        if ($line.Length -gt 2)
        {
            if ($line.Substring(0, 2) -match '\d+')
            {
                $output += $line + "`r`n"
            }
        }

    }

    # Make a header line
    $header = Get-Content -Path "c:\sunriseset\header-line.txt" 
    $header += "`r`n"

    # Write raw data file
    ($header + $output) | Out-File $raw -Force


    #
    # Pivot months into rows
    #

    # Read in delmited file
    $data = Import-Csv $raw -Delimiter ' '

    # Convert to array of objects
    $final = @()
    $data | ForEach-Object {

        # Jan
        $temp = $_ | Select-Object @{Name="month";Expression={[string]"01"}}, 
            @{Name="day";Expression={[string]$_.DAY}}, 
            @{Name="rise";Expression={[string]$_.JAN0}}, 
            @{Name="set";Expression={[string]$_.JAN1}};
        $final += $temp

        # Feb
        $temp = $_ | Select-Object @{Name="month";Expression={[string]"02"}}, 
            @{Name="day";Expression={[string]$_.DAY}}, 
            @{Name="rise";Expression={[string]$_.FEB0}}, 
            @{Name="set";Expression={[string]$_.FEB1}};
        $final += $temp

        # Mar
        $temp = $_ | Select-Object @{Name="month";Expression={[string]"04"}}, 
            @{Name="day";Expression={[string]$_.DAY}}, 
            @{Name="rise";Expression={[string]$_.MAR0}}, 
            @{Name="set";Expression={[string]$_.MAR1}};
        $final += $temp

        # Apr
        $temp = $_ | Select-Object @{Name="month";Expression={[string]"02"}}, 
            @{Name="day";Expression={[string]$_.DAY}}, 
            @{Name="rise";Expression={[string]$_.APR0}}, 
            @{Name="set";Expression={[string]$_.APR1}};
        $final += $temp

        # May
        $temp = $_ | Select-Object @{Name="month";Expression={[string]"05"}}, 
            @{Name="day";Expression={[string]$_.DAY}}, 
            @{Name="rise";Expression={[string]$_.MAY0}}, 
            @{Name="set";Expression={[string]$_.MAY1}};
        $final += $temp

        # Jun
        $temp = $_ | Select-Object @{Name="month";Expression={[string]"06"}}, 
            @{Name="day";Expression={[string]$_.DAY}}, 
            @{Name="rise";Expression={[string]$_.JUN0}}, 
            @{Name="set";Expression={[string]$_.JUN1}};
        $final += $temp

        # Jul
        $temp = $_ | Select-Object @{Name="month";Expression={[string]"07"}}, 
            @{Name="day";Expression={[string]$_.DAY}}, 
            @{Name="rise";Expression={[string]$_.JUL0}}, 
            @{Name="set";Expression={[string]$_.JUL1}};
        $final += $temp

        # Aug
        $temp = $_ | Select-Object @{Name="month";Expression={[string]"08"}}, 
            @{Name="day";Expression={[string]$_.DAY}}, 
            @{Name="rise";Expression={[string]$_.AUG0}}, 
            @{Name="set";Expression={[string]$_.AUG1}};
        $final += $temp

        # Sep
        $temp = $_ | Select-Object @{Name="month";Expression={[string]"09"}}, 
            @{Name="day";Expression={[string]$_.DAY}}, 
            @{Name="rise";Expression={[string]$_.SEP0}}, 
            @{Name="set";Expression={[string]$_.SEP1}};
        $final += $temp

        # Oct
        $temp = $_ | Select-Object @{Name="month";Expression={[string]"10"}}, 
            @{Name="day";Expression={[string]$_.DAY}}, 
            @{Name="rise";Expression={[string]$_.OCT0}}, 
            @{Name="set";Expression={[string]$_.OCT1}};
        $final += $temp

        # Nov
        $temp = $_ | Select-Object @{Name="month";Expression={[string]"11"}}, 
            @{Name="day";Expression={[string]$_.DAY}}, 
            @{Name="rise";Expression={[string]$_.NOV0}}, 
            @{Name="set";Expression={[string]$_.NOV1}};
        $final += $temp

        # Dec
        $temp = $_ | Select-Object @{Name="month";Expression={[string]"12"}}, 
            @{Name="day";Expression={[string]$_.DAY}}, 
            @{Name="rise";Expression={[string]$_.DEC0}}, 
            @{Name="set";Expression={[string]$_.DEC1}};
        $final += $temp

    }


    #
    # Write data as csv file
    #

    # Remove invalid data
    $final = $final | Where-Object { $_.rise -ne '.' };

   
    # Reformat data & sort by mon, day
    $output = $final | 

    Select-Object @{Name="month1";Expression={[string]$_.month}}, 
           @{Name="day1";Expression={[string]$_.day}},
           @{name="sunrise1";Expression={[string](($_.rise).SubString(0,2) + ":" + ($_.rise).SubString(2,2))}},
           @{name="sunset1";Expression={[string](($_.set).SubString(0,2) + ":" + ($_.set).SubString(2,2))}} |
    Sort-Object -Property month1, day1

    # Save to csv file
    $output | Export-Csv -Path $csv -NoTypeInformation
}


