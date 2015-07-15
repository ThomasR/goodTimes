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

function getUptimeAttr($entry) {
    $result = New-TimeSpan
    foreach ($interval in $entry) {
        $result = $result.add($interval[1] - $interval[0])
    }
    write $result
}

function getIntervalAttr($entry) {
    $result = @();
    foreach ($interval in $entry) {
        $result += '{0:HH:mm}-{1:HH:mm}' -f $interval[0], $interval[1]
    }
    write ($result -join ', ')
}

function getBookingHoursAttr($interval) {
    $netTime = $interval.totalHours - $lunchbreak
    write ([math]::Round($netTime * $precision) / $precision)
}

function getFlexTimeAttr($bookedHours) {
    $delta = $bookedHours - $workinghours
    $result = $delta.toString('+0.00;-0.00;±0.00')
    if ($delta -eq 0) {
        write $result, $null
    } elseif ($delta -gt 0) {
        write $result, 'darkgreen'
    } else {
        write $result, 'darkred'
    }
}

function getLogAttrs($entry) {
    $result = @{}
    $result.uptime = getUptimeAttr $entry
    $result.bookingHours = getBookingHoursAttr $result.uptime
    $result.flexTime = getFlexTimeAttr $result.bookingHours
    $result.intervals = getIntervalAttr $entry
    write $result
}

function print($string, $color) {
    if ($color) {
        write-host -f $color -n $string
    } else {
        write-host -n $string
    }
}

function println($string, $color) {
    print ($string + "`r`n") $color
}

function wait() {
    # If running in the console, wait for input before continuing.
    if ($Host.Name -eq 'ConsoleHost') {
        Write-Host 'Press any key to continue...'
        $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyUp') > $null
    }
}
$filter = @{
    LogName = 'System'
    ProviderName = 'Microsoft-Windows-Kernel-General'
    ID = 12, 13
    StartTime = (get-date).addDays(-$historyLength)
}

[System.Collections.ArrayList]$events = Get-WinEvent -FilterHashtable $filter | select ID, TimeCreated | sort TimeCreated

$events += New-Object PSCustomObject -Property @{
    ID = 13
    TimeCreated = get-date
}

$log = New-Object System.Collections.ArrayList

write-host $log

:outer while ($events.count -ge 2) {
    do {
        if ($events.count -lt 2) {
            break outer;
        }
        $end = $events[$events.count - 1]
        $events.remove($end)
    } while ($end.ID -ne 13)
    do {
        if ($events.count -lt 1) {
            break outer;
        }
        $start = $events[$events.count - 1]
        $events.remove($start)
    } while ($start.ID -ne 12)

    $last = $log[0]
    if ($last -and $start.TimeCreated.Date.equals($last[0][0].Date)) {
        $log[0] = ,@($start.TimeCreated, $end.TimeCreated) + $last
    } else {
        $log.insert(0, @(,@($start.TimeCreated, $end.TimeCreated)))
    }

}

$oldfgColor= $host.UI.RawUI.ForegroundColor
$host.UI.RawUI.ForegroundColor = 'gray'
$oldbgColor = $host.UI.RawUI.BackgroundColor
$host.UI.RawUI.BackgroundColor = 'black'
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
$host.UI.RawUI.BackgroundColor = $oldbgColor
$host.UI.RawUI.ForegroundColor = $oldfgColor

wait
