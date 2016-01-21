# .SYNOPSIS
#    Good Times!
#
# .DESCRIPTION
#    Dieses Skript zeigt die Uptime-Zeiten der vergangenen Tage an, und berechnet
#    daraus die in der Zeitmanagement zu buchenden Zeiten, sowie Gleitzeit-Differenzen.
#
# .NOTES
#
#    Copyright 2015 Thomas Rosenau
#
#    Licensed under the Apache License, Version 2.0 (the "License");
#    you may not use this file except in compliance with the License.
#    You may obtain a copy of the License at
#
#        http://www.apache.org/licenses/LICENSE-2.0
#
#    Unless required by applicable law or agreed to in writing, software
#    distributed under the License is distributed on an "AS IS" BASIS,
#    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#    See the License for the specific language governing permissions and
#    limitations under the License.
#
# .PARAMETER  historyLength
#    Anzahl der angezeigten Tage in der Vergangenheit.
#    Standardwert: 30
#    Alias: -l
# .PARAMETER  workingHours
#    Anzahl der zu arbeitenden Stunden pro Tag.
#    Standardwert: 8
#    Alias: -h
# .PARAMETER  lunchBreak
#    Länge der Mittagspause in Stunden pro Tag.
#    Standardwert: 1
#    Alias: -b
# .PARAMETER  precision
#    Rundungspräzision in %, d.h. 1 = Rundung auf volle Stunde, 4 = Rundung auf 60/4=15 Minuten, …, 100 = keine Rundung
#    Standardwert: 4
#    Alias: -p
# .PARAMETER  dateFormat
#    Datumsformat gemäß https://msdn.microsoft.com/en-us/library/8kb3ddd4.aspx?cs-lang=vb#content
#    Standardwert: ddd dd/MM/yyyy
#    Alias: -d
#
# .INPUTS
#    Keine
# .OUTPUTS
#    Keine
#
# .EXAMPLE
#    .\goodTimes.ps1
#    (Aufruf mit Standardwerten)
# .EXAMPLE
#    .\goodTimes.ps1 -historyLength 30 -workingHours 8 -lunchBreak 1 -precision 4
#    (Aufruf mit explizit gesetzten Standardwerten)
#    (30 Tage anzeigen, Arbeitszeit 8 Stunden täglich, 1 Stunde Mittagspause, Rundung auf 15 (=60/4) Minuten)
# .EXAMPLE
#    .\goodTimes.ps1 30 -h 8 -b 1 -p 4
#    (Aufruf mit explizit gesetzten Standardwerten, Kurzschreibweise)
# .EXAMPLE
#    .\goodTimes.ps1 14 -h 7 -b .5 -p 6
#    (14 Tage anzeigen, Arbeitszeit 7 Stunden täglich, 30 Minuten Mittagspause, Rundung auf 10 (=60/6) Minuten)

param (
    [int]
    [parameter(Position=0)]
    [alias('l')]
        $historyLength = 30,
    [byte]
    [validateRange(0, 24)]
    [alias('h')]
        $workinghours = 40 / 5,
    [decimal]
    [validateRange(0, 24)]
    [alias('b')]
        $lunchbreak = 1,
    [byte]
    [validateRange(1, 100)]
    [alias('p')]
        $precision = 4,
    [string]
#    [ValidateScript({$_ -cmatch '\bd\b' -or ($_ -cmatch '\bdd\b' -and $_ -cmatch '\bM{1,4}\b')})]
    [ValidateScript({$_ -cnotmatch '[HhmsfFt]'})]
    [alias('d')]
        $dateFormat = 'ddd dd/MM/yyyy' # "/" ist Platzhalter für lokalisierten Trenner
)

# helper functions to calculate the required attributes

# total uptime
function getUptimeAttr($entry) {
    $result = New-TimeSpan
    foreach ($interval in $entry) {
        $result = $result.add($interval[1] - $interval[0])
    }
    write $result
}

# uptime intervals
function getIntervalAttr($entry) {
    $result = @();
    foreach ($interval in $entry) {
        $result += '{0:HH:mm}-{1:HH:mm}' -f $interval[0], $interval[1]
    }
    write ($result -join ', ')
}

# booking hours
function getBookingHoursAttr($interval) {
    $netTime = $interval.totalHours - $lunchbreak
    write ([math]::Round($netTime * $precision) / $precision)
}

# flex time delta
function getFlexTimeAttr($bookedHours) {
    $delta = $bookedHours - $workinghours
    $result = $delta.toString('+0.00;-0.00;?0.00')
    if ($delta -eq 0) {
        write $result, $null
    } elseif ($delta -gt 0) {
        write $result, 'darkgreen'
    } else {
        write $result, 'darkred'
    }
}
# end helper functions

# generate a hashmap of the abovementioned attributes
function getLogAttrs($entry) {
    $result = @{}
    $result.uptime = getUptimeAttr $entry
    $result.bookingHours = getBookingHoursAttr $result.uptime
    $result.flexTime = getFlexTimeAttr $result.bookingHours
    $result.intervals = getIntervalAttr $entry
    write $result
}

# convenience function to write to screen with or without color
function print($string, $color) {
    if ($color) {
        write-host -f $color -n $string
    } else {
        write-host -n $string
    }
}

# convenience function to write to screen with or without color
function println($string, $color) {
    print ($string + "`r`n") $color
}

# If running in the console, wait for input before continuing.
function wait() {
    if ($Host.Name -eq 'ConsoleHost') {
        Write-Host 'Press any key to continue...'
        $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyUp') > $null
    }
}

# helper to determine whether a given EventLogRecord is a boot or wakeup event
function isStartEvent($event) {
    return ($event.ProviderName -eq 'Microsoft-Windows-Kernel-General' -and $event.ID -eq 12) -or
            ($event.ProviderName -eq 'Microsoft-Windows-Power-Troubleshooter' -and $event.ID -eq 1)
}

# helper to determine whether a given EventLogRecord is a shutdown or suspend event
function isStopEvent($event) {
    return ($event.ProviderName -eq 'Microsoft-Windows-Kernel-General' -and $event.ID -eq 13) -or
            ($event.ProviderName -eq 'Microsoft-Windows-Kernel-Power' -and $event.ID -eq 42)
}




# create an array of filterHashTables that filter boot and shutdown events from the desired period
$startTime = (get-date).addDays(-$historyLength)
$filters = @(
    @{
        StartTime = $startTime
        LogName = 'System'
        ProviderName = 'Microsoft-Windows-Kernel-General'
        ID = 12, 13
    },
    @{
        StartTime = $startTime
        LogName = 'System'
        ProviderName = 'Microsoft-Windows-Kernel-Power'
        ID = 42 # what else?
    },
    @{
        StartTime = $startTime
        LogName = 'System'
        ProviderName = 'Microsoft-Windows-Power-Troubleshooter'
        ID = 1
    }
)

# get system log entries for boot/shutdown
$events = Get-WinEvent -FilterHashtable $filters | select ID, TimeCreated, ProviderName

# sort (reverse chronological order) and convert to ArrayList
[System.Collections.ArrayList]$events = $events | sort TimeCreated

# create an empty list, which will hold one entry per day
$log = New-Object System.Collections.ArrayList

# fill the $log list by searching for start/stop pairs
:outer while ($events.count -ge 2) {
    if ($log) {
        # find the latest stop event
        do {
            if ($events.count -lt 2) {
                # if there is only one stop event left, there can't be any more start event (e.g. when system log was cleared)
                break outer;
            }
            $end = $events[$events.count - 1]
            $events.remove($end)
        } while (-not (isStopEvent $end)) # consecutive start events. This may happen when the system crashes (power failure, etc.)
    } else {
        # add a fake shutdown event for this very moment
        $end = @{TimeCreated = get-date}
    }

    # find the corresponding start event
    do {
        if ($events.count -lt 1) {
            # no more events left
            break outer;
        }
        $start = $events[$events.count - 1]
        $events.remove($start)
    } while (-not (isStartEvent $start)) # not sure if ther can indeed be consecutive stop events, but let's better be safe than sorry

    # check if the current start/stop pair has occured on the same day as the previous one
    $last = $log[0]
    if ($last -and $start.TimeCreated.Date.equals($last[0][0].Date)) {
        # combine uptimes
        $log[0] = ,@($start.TimeCreated, $end.TimeCreated) + $last
    } else {
        # create new day
        $log.insert(0, @(,@($start.TimeCreated, $end.TimeCreated)))
    }

}

# colors
$oldfgColor= $host.UI.RawUI.ForegroundColor
$host.UI.RawUI.ForegroundColor = 'gray'
$oldbgColor = $host.UI.RawUI.BackgroundColor
$host.UI.RawUI.BackgroundColor = 'black'

# write the output
$screenwidth = $host.UI.RawUI.BufferSize.width

Write-Host ("{0,-$($screenwidth - 1)}" -f '    Datum     Buchen Gleitzeit  Uptime (incl. Pause)')
Write-Host ("{0,-$($screenwidth - 1)}" -f '------------- ------ ---------  --------------------')

foreach ($entry in $log) {
    $firstInterval = $entry[0]
    $day = $firstInterval[0].Date.toString($dateFormat)
    $dayOfWeek = ([int]$firstInterval[0].dayOfWeek + 6) % 7
    if ($dayOfWeek -lt $lastDayOfWeek) {
        println ("{0,-$($screenwidth - 1)}" -f '-------------')
    }
    $lastDayOfWeek = $dayOfWeek
    if ($dayOfWeek -ge 5) {
        Write-Host $day -n -backgroundColor darkred -foregroundcolor gray
    } else {
        print $day
    }
    $attrs = getLogAttrs($entry)
    print (' {0,5}     ' -f $attrs.bookingHours.toString('#0.00', [System.Globalization.CultureInfo]::getCultureInfo('de-DE'))) cyan
    print $attrs.flexTime[0] $attrs.flexTime[1]
    print ("{0,6:#0}:{1:00} | {2,-$($screenwidth - 42)}" -f $attrs.uptime.hours, [math]::round($attrs.uptime.minutes + $attrs.uptime.seconds / 60), $attrs.intervals) DarkGray
    Write-Host
}

# restore previous colors
$host.UI.RawUI.BackgroundColor = $oldbgColor
$host.UI.RawUI.ForegroundColor = $oldfgColor

wait
