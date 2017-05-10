#Requires -Version 5.0

#Set-StrictMode -Version latest

New-Variable -Name ISYSEttingsXMLFileName -Value 'ISYSettings.xml'                           -Scope script -Option ReadOnly
New-Variable -Name ISYSettingsFileVersion -Value ([version]('1.0.0.0'))                      -Scope script -Option ReadOnly
New-Variable -Name ISYDBXMLFileName       -Value 'ISYDB.xml'                                 -Scope script -Option ReadOnly
New-Variable -Name ISYProgramsXMLFileName -Value 'ISYPrograms.xml'                           -Scope script -Option ReadOnly
New-Variable -Name ISYDefaultsFilePath    -Value (Join-Path $PSScriptRoot 'ISYDefaults.xml') -Scope script -Option ReadOnly

# used by set-x10device function. Defined here so it can be used in param checking
$x10decoder = @{
"AllLightsOff" = "1"
"On" = "3"
"AllLightsOn" = "5"
"Bright" = "7"
"Off" = "11"
"AllUnitsOff" = "13"
"Dim" = "15"

"ExtendedCode" = "9"
"HailRequest" = "14"
"HailAcknowledge" = "6"
"PreSetDim" = "12"
"StatusIsOn" = "8"
"StatusIsOff" = "2"
"StatusRequest" = "10"}

# TODO use enum instead of $x10decoder hash table.
<#
Enum X10Commands
{
    AllLightsOff = 1
    On = 3
    AllLightsOn = 5
    Bright = 7
    Off = 11
    AllUnitsOff = 13
    Dim = 15
    ExtendedCode = 9
    HailRequest = 14
    HailAcknowledge = 6
    PreSetDim = 12
    StatusIsOn = 8
    StatusIsOff = 2
    StatusRequest = 10
}
#>

<#
.Synopsis
   Converts control, action and insteon address values to friendly names. 
.DESCRIPTION
   This function acts on the xmlelements output from Convert-ISYEventToXmlElem by
   changing the control and action properties to friendly names. It also replaces 
   insteon addresses in the node and eventinfo properties to the names of these devices.
.EXAMPLE
   $evtfriendly = Get-Content .\textisyevents.txt | Convert-ISYEventToXmlElem | Convert-ISYEventToFriendlyName | Group-Object -Property control

    PS C:\>$evtfriendly

    Count Name                      Group
    ----- ----                      -----
        4 SystemConfigUpdatedEvent  {Event, Event, Event, Event}
      110 Trigger                   {Event, Event, Event, Event...}
       19 Heartbeat                 {Event, Event, Event, Event...}
        2 ELKEvents                 {Event, Event}
       19 RR                        {Event, Event, Event, Event...}
       19 OL                        {Event, Event, Event, Event...}
       45 ST                        {Event, Event, Event, Event...}
        1 BillingEvents             {Event}
        4 DON                       {Event, Event, Event, Event}
      106 SystemStatusEvent         {Event, Event, Event, Event...}
        3 DOF                       {Event, Event, Event}

   This example uses a file of captured ISY event strings to show what Convert-ISYEventToFriendlyNames does.
   First the strings are sent to Convert-ISYEventToXmlElem, then to Convert-ISYEventToFriendlyName, grouped by 
   the control property and then captured in variable $evtfriendly. $evtfriendly is then displayed at the console.
.EXAMPLE
   $evtfriendly | where name -eq DOF | Select-Object -Property group | Format-Custom

    class GroupInfo
    {
      Group =
        [
          class XmlElement
          {
            TimeStamp = 20160924T0519368639
            seqnum = 616
            sid = uuid:83
            control = DOF
            action = 0
            node = DriveWayAlert-Sensor-XD1
            eventInfo =
          }
          class XmlElement
          {
            TimeStamp = 20160924T0648150737
            seqnum = 858
            sid = uuid:83
            control = DOF
            action = 0
            node = DriveWayAlert-Sensor-XD1
            eventInfo =
          }
          class XmlElement
          {
            TimeStamp = 20160924T1350001950
            seqnum = 1995
            sid = uuid:83
            control = DOF
            action = 0
            node = DriveWayAlert-Sensor-XD1
            eventInfo =
          }
        ]

    }

    Using Format-Custom we see that the 3 DOF events are the DriveWayAlert-Sensor-XD1 turning off.
#>
function Convert-ISYEventToFriendlyName {
[CmdletBinding()]
[OutputType([System.Xml.XmlElement[]])]
param(
# event xmlelement object from the Convert-ISYEventToXmlElem function output
[parameter(ValueFromPipeline=$true)]
[System.Xml.XmlElement[]]$event)

BEGIN
{
    $instcmdsrgx = '^DON$|^DOF$|^DFON$|^DFOF$|^BMAN$|^SMAN$|^OL$|^RR$|^ST$|^BEEP$'
    $FileDateTime = 'yyyyMMddTHHmmssffff'
    # $dtobject = [datetime]::ParseExact($event.timestamp,$FileDateTime,$null)
    # TODO replace values in control:SystemProgressEvent action:ProgressUpdatedEvent with names
    <#
    $truncatedaddrregx = '([1-9A-F][0-9A-F] |[1-9A-F] ){2}[1-9A-F][0-9A-F]|[1-9A-F]'
    (([regex]::Matches('[Std-Cleanup ] 16.CD.8B-->ISY/PLM Group=1, Max Hops=1, Hops Left=0'.replace('.',' '), $truncatedaddrregx)) | where {$_.value.length -ge 5}).value
    $truncated = @()
    (Get-InsteonDevice -dbmap | sort address).ForEach({$truncated += [pscustomobject]@{name = $_.name;
                                                                      shortaddr = [regex]::match($_.address,$truncatedaddrregx).value}})
    ($truncated | where shortaddr -eq '9 14 25' | select -First 1).name
    #>

    #region hashtables

    $ControlHash = @{
                    _0 = 'Heartbeat'
                    _1 = 'Trigger'
                    _2 = 'ProtocolSpecificEvent'
                    _3 = 'NodesUpdatedEvent'
                    _4 = 'SystemConfigUpdatedEvent'
                    _5 = 'SystemStatusEvent'
                    _6 = 'InternerAccessEvent'
                    _7 = 'SystemProgressEvent'
                    _8 = 'SecuritySystemEvent'
                    _9 = 'SystemAlertEvent'
                    _10 = 'OpenADREvent'
                    _11 = 'ClimateEvent'
                    _12 = 'AMIMeterEvent'
                    _13 = 'ElectricityMonitorEvent'
                    _14 = 'UPBLinkerEvent'
                    _15 = 'UPBDeviceAdderState'
                    _16 = 'UPBStatusEvent'
                    _17 = 'GasMeterEvent'
                    _18 = 'ZigbeeEvent'
                    _19 = 'ELKEvents'
                    _20 = 'DeviceLinkerEvents'
                    _21 = 'Z-WaveEvents'
                    _22 = 'BillingEvents'
                    _23 = 'PortalEvents'}

    $SystemProgressActionHash = @{
                    '1' = 'ProgressUpdatedEvent'
                    '2.1' = 'UPBONLY'
                    '2.2' = 'DeviceAdderInfoEvent'
                    '2.3' = 'DeviceAdderWarnEvent'}

    $SystemStatusActionHash = @{
                    '0' = 'NotBusy'
                    '1' = 'Busy'
                    '2' = 'Completelyidle'
                    '3' = 'SafeMode'}

    $triggerActionHash = @{
                    '0' = 'Status'
                    '1' = 'ClientShouldGetStatus'
                    '2' = 'KeyChanged'
                    '3' = 'Information'
                    '4' = 'IRLearnMode'
                    '5' = 'ScheduleEvent'
                    '6' = 'VariableStatus'
                    '7' = 'VariableInitialized'
                    '8' = 'CurrentProgramKey'} 

    $SystemConfigActionHash = @{
                    '1' = 'TimeConfigurationUpdated'
                    '2' = 'NTPSettingsUpdated'
                    '3' = 'NotificationsSettingsUpdated'
                    '4' = 'NTPServerCommError'
                    '5' = 'BatchModeChanged'
                    '6' = 'BatteryDeviceWriteModeChanged'}

    $NodeUpdatedActionHash = @{
                    'NN' = 'NodeRenamed'
                    'NR' = 'NodeRemoved'
                    'ND' = 'NodeAdded'
                    'NE' = 'NodeError'
                    'CE' = 'NodeErrorCleared'
                    'EN' = 'NodeEnabled'
                    'PC' = 'NodesParentChanged'
                    'GN' = 'GroupRenamed'
                    'GR' = 'GroupRemoved'
                    'GD' = 'GroupAdded'
                    'FD' = 'FolderAdded'
                    'FN' = 'FolderRenamed'
                    'FR' = 'FolderRemoved'
                    'MV' = 'NodeMovedintoGroup'
                    'RG' = 'NodeRemovedfromGroup'
                    'CL' = 'NodeLinkChanged'
                    'SN' = 'DiscoveringNodes'
                    'SC' = 'StoppedLinking'
                    'PI' = 'PowerInfoChanged'
                    'WR' = 'NetworkRenamed'
                    'WH' = 'PendingDeviceWrites'
                    'WD' = 'WritingToDevice'}

    $InternetAccessActionHash = @{
                    '0' = 'Disabled'
                    '1' = 'Enabled'
                    '2' = 'Failed'}

    $SecuritySystemActionHash = @{
                    '0' = 'Disconnected'
                    '1' = 'Connected'
                    'DA' = 'Disarmed'
                    'AW' = 'ArmedAway'
                    'AS' = 'ArmedStay'
                    'ASI' = 'ArmedStayInstant'
                    'AN' = 'ArmedNight'
                    'ANI' = 'ArmedNightInstant'
                    'AV' = 'ArmedVacation'}
    $ELKActionHash = @{
                    "1" = "TopologyChange"
                    "2" = "AreaEvent"
                    "3" = "ZoneEvent"
                    "4" = "KeypadEvent"
                    "5" = "OutputEvent"
                    "6" = "SystemEvent"
                    "7" = "ThermostatEvent"
                    }

    $actionhash = @{
                    'Trigger' = $triggerActionHash
                    'SystemConfigUpdatedEvent'= $SystemConfigActionHash 
                    'NodesUpdatedEvent' = $NodeUpdatedActionHash
                    'SystemStatusEvent' = $SystemStatusActionHash
                    'SystemProgressEvent' = $SystemProgressActionHash
                    'InternerAccessEvent' = $InternetAccessActionHash
                    'SecuritySystemEvent' = $SecuritySystemActionHash
                    'ELKEvents' = $ELKActionHash}

    #endregion

    function GetLatestISYDB {
    [CmdletBinding()]
    param([switch]$noupdate)

    if (!$noupdate)
    {
        Update-ISYDBXMLFile -refreshnow
    }
    $script:dbmap = Get-InsteonDevice -dbmap
    $script:addrregx = $dbmap.ForEach('address') -join '|'
    $script:programlist = Get-ISYProgramList |  
        Select-Object -Property @{label = 'Progid'; expression = {'{0:X1}' -f [convert]::ToInt16($_.id,16)}}, name
    }

    # ISYDB was updated when the module was imported above. No need to do it again.
    GetLatestISYDB -noupdate

    function ConvertToFriendlyNames {
    [CmdletBinding()]
    param([System.Xml.XmlElement]$xmlelem)

    # friendify control and action
    if ($controlhash[$xmlelem.control])
    {
        $xmlelem.control = $controlhash[$xmlelem.control]

        if ($actionhash[$xmlelem.control])
        {
            if ($actionhash[$xmlelem.control][$xmlelem.Action])
            {
                $xmlelem.Action = $actionhash[$xmlelem.control][$xmlelem.Action]
            }
        }
    }

    # examine $xmlelem and look for add/remove/rename on devices or programs
    # if found call GetLatestISYDB
    # TODO only the information used here is subject to change in real time.
    # Maybe this should be a separate function 

    # replace insteon address with name in node property
    if ($xmlelem.control -match $instcmdsrgx)
    {
        $devname = ($script:dbmap | where address -eq $xmlelem.node).name
        if ($devname)
        {
            $xmlelem.node = $devname
        }
    }

    # replace insteon address with name in eventinfo property
    if ($xmlelem.eventinfo -match $script:addrregx)
    {
        $devname = ($dbmap | where address -eq $matches.0).name
        $xmlelem.eventinfo = $xmlelem.eventinfo.replace($matches.0, $devname)
        $xmlelem.eventinfo = $xmlelem.eventinfo -replace '\[ {1,}', "["
    }

    # replace program id with name in eventinfo.id property
    if ($xmlelem.control -eq 'trigger' -and  $xmlelem.action -eq 'status')
    {
        if ($xmlelem.eventInfo.id)
        {
            $xmlelem.eventInfo.id = ($script:programlist | where progid -eq $xmlelem.eventInfo.id).name
        }
    }

    $xmlelem

    }

}

PROCESS
{
    ConvertToFriendlyNames -xmlelem $_
}

END
{
    if (!($PSCmdlet.MyInvocation.ExpectingInput)) 
    {
        foreach ($xml in $xmlelem)
        {
            ConvertToFriendlyNames -xmlelem $xml
        }
    }
}


}
#end of function Convert-ISYEventToFriendlyName

<#
.Synopsis
   Converts the event strings from Register-ForISYEvent to xmlelement objects
.DESCRIPTION
   The xml formatted strings ouput from the Register-ForISYEvent function are converted to xmlelement objects.
   Each event has the following properties:
   TimeStamp : timestamp of when the event occurred in 'yyyyMMddTHHmmssffff' format where HH is 24 hour and ffff is milliseconds
        Note1: To covert to datetime object use [datetime]::ParseExact("20160913T1833374716",'yyyyMMddTHHmmssffff', $null)
        Note2: TimeStamp will NOT be present if the Register-ForISYEvent function was called with the -raw switch
   seqnum    : Unique message sequence number incremented with each message for this subscription 
   sid       : Subscription ID assigned by the ISY (example uuid:67)
   control   : event controls
   action    : event actions
   node      : node (if any) affected by the event
   eventInfo : more info (if any) about the event

   Normally consecutive heartbeat events are compressed and replaced with one instance with the eventinfo property
   set to '(Repeated n times)' where n is the number of consecutive heartbeat events. 
   The -nodedup switch parameter when present stops this behaviour and outputs all the heartbeat events.
.EXAMPLE
   $evtxml = Get-Content .\textisyevents.txt | Convert-ISYEventToXmlElem | Group-Object -Property control

   PS C:\>$evtxml

    Count Name                      Group
    ----- ----                      -----
        4 _4                        {Event, Event, Event, Event}
      110 _1                        {Event, Event, Event, Event...}
       19 _0                        {Event, Event, Event, Event...}
        2 _19                       {Event, Event}
       19 RR                        {Event, Event, Event, Event...}
       19 OL                        {Event, Event, Event, Event...}
       45 ST                        {Event, Event, Event, Event...}
        1 _22                       {Event}
        4 DON                       {Event, Event, Event, Event}
      106 _5                        {Event, Event, Event, Event...}
        3 DOF                       {Event, Event, Event}
    
    This example uses a file of captured ISY event strings to show what Convert-ISYEventToXmlElem does.
    First the strings are sent to Convert-ISYEventToXmlElem, grouped by the control property and then 
    captured in variable $evtxml. $evtxml is then displayed at the console.

.EXAMPLE
    $evtxlm | where name -eq _19 | Select-Object -Property group | Format-Custom

    class GroupInfo
    {
      Group =
        [
          class XmlElement
          {
            TimeStamp = 20160924T0209033885
            seqnum = 4
            sid = uuid:83
            control = _19
            action = 6
            node =
            eventInfo =
              class XmlElement
              {
                se =
                  class XmlElement
                  {
                    type = 156
                    val = 0
                  }
              }
          }
          class XmlElement
          {
            TimeStamp = 20160924T0209034085
            seqnum = 5
            sid = uuid:83
            control = _19
            action = 6
            node =
            eventInfo =
              class XmlElement
              {
                se =
                  class XmlElement
                  {
                    type = 157
                    val = 0
                  }
              }
          }
        ]

    }

    Examining the 2 events with control values of _19 in $evtxml using format-custom shows the details of these events.  
#>
function Convert-ISYEventToXmlElem {
[cmdletbinding()]
[outputtype([System.Xml.XmlElement[]])]
param(
    # Event strings to be converted which are output from the Register-ForISYEvent function
    [parameter(ValueFromPipeline=$true)]
    [string[]]$line,

    # if present consecutive heartbeat event(s) are not compressed. Normally these are replaced with 1 event displaying the number of repeats 
    [switch]$nodedup
    )

BEGIN
{
    $FileDateTime = 'yyyyMMddTHHmmssffff'
    #$dtobject = [datetime]::ParseExact($event.timestamp,$FileDateTime,$null)
    $heartbeat = '<control>_0</control>'
    $script:previous = ''
    $script:repeats = 0
    #$script:repmsg = ''
    $splat = @{nodedup = $nodedup.IsPresent}

    function ObjectifyLine {
    [cmdletbinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$text,

        [switch]$nodedup
        )

        Write-Verbose -Message "line = $text`n"

        if (!($nodedup) -and $text.Contains($heartbeat))
        {       
            $script:repeats++ 
            # TODO
            # to unconditionally emit 1st heartbeat (dont forget to subtract 1 from other path)
            # and emit when beat quota is reached
            <#
            if ($script:repeats -eq 1 -or $script:repeats -ge $quota)
            {
                try
                {
                    $event = [xml]$text
                }
                catch  [System.Management.Automation.RuntimeException]
                {
                    write-error -Message "cannot parse $text into an xmldoc" -Category InvalidArgument
                }
                $event

                if ($script:repeats -ge $quota )
                {
                     $script:repeats = 0
                }                
            }
            #> 
            # TODO
            # meaure time between heartbeats - if greater than 120 seconds unregister  
        }
        else
        { 
            write-verbose -Message "previous = $script:previous`n"
            
            if ($text -match '^\<' -and $script:repeats -and $script:previous.Contains($heartbeat))
            {
                if ($script:repeats -gt 1)
                {
                    # for unconditional 1st heartbeat emitted change to $script:repeats-1  
                    $repmsg = "(Repeated $script:repeats times)"
                    $pretext = $script:previous.replace('<eventInfo></eventInfo>',"<eventInfo>$repmsg</eventInfo>")
                } 
                else
                {
                    $pretext = $script:previous
                }                  
                
                try
                {
                    $event = [xml]$pretext
                }
                catch  [System.Management.Automation.RuntimeException]
                {
                    write-error -Message "cannot parse $text into an xmldoc" -Category InvalidArgument
                } 

                $event           
                #$script:heartbeatline = ''
                $script:repeats = 0
                
            }
        
            if ($text -match '^\<')
            {
                try
                {
                    $event = [xml]$_
                }
                catch  [System.Management.Automation.RuntimeException]
                {
                    write-error -Message "cannot parse $text into an xmldoc" -Category InvalidArgument
                }
                $event
            }
            elseif ($text -match '^\d{1,}')
            {
                write-verbose -Message "Content length = $text"
            }
            else
            {
                Throw('wtf! Text input is malformed.')
            }
        }

        if ($text -match '^\<')
        {
            $script:previous = $text
        }
    }
}
# end of BEGIN block

PROCESS
{
    $lineobj = ObjectifyLine -text ([string]($_)).trim([char]0) @splat
    Write-Verbose -Message "Control = $($($lineobj.event.control) -join ',')"

    if ($lineobj)
    {
        $lineobj.event
    }
}
# end of PROCESS block

END
{
    if (!($PSCmdlet.MyInvocation.ExpectingInput)) 
    {
        foreach ($l in $line) 
        {
            $lineobj = ObjectifyLine -text ([string]($l)).trim([char]0) @splat
            if ($lineobj)
            {
                $lineobj.event
            }
        }
    }
}
# end of END block
}
#end of function Convert-ISYEventToXmlElem

<#
.Synopsis
   Converts the text of the ISY events from the event viewer to an array of pscustomobjects. 
.DESCRIPTION
    This function parses the text from the Universal Devices event viewer that is launched from the Admin console.
    The user must copy this text to the clipboard or get the text from a saved file to use this function. 

   Each line of text represents 1 ISY event. Each line of text is converted to a  [pscustomobject].
   This object has the following properties:
    datetime: the time the event occured
    debuglevel: each event comes from a debug level from 1 to 3 
    Device/Code/Comm: depending upon the debug level the value is: 
        1: device status and operations 
        2: code representing the commands seen or sent by the ISY 
        3: Communication events
    action: the action performed 

    By default consecutive events which are duplicates are compressed into 1 event object with
    the text 'Repeated n times' appended to to this event. n is the number of messages not converted.
    For example the ISY will sometimes send many Time events in 1 second so they display as duplicates. 
    The -nocompress switch turns this off and produces a 1 to 1 line of text to object mapping.

.EXAMPLE
   'Wed 08/31/2016 22:31:39 : [D2D EVENT   ] Event [16 96 98 1] [ST] [255] uom=0 prec=-1' | Convert-ISYEventViewerStringToObject

   datetime              debuglevel Device/Code/Comm action
   --------              ---------- ---------------- ------
   8/31/2016 10:31:39 PM          2 D2D EVENT        Event [16 96 98 1] [ST] [255] uom=0 prec=-1

   This shows a debug level 2 event (more Info). 

.EXAMPLE
   'Wed 08/31/2016 22:31:39 : [  16 96 98 1]       ST 255' | Convert-ISYEventViewerStringToObject

    datetime              debuglevel Device/Code/Comm    action
    --------              ---------- ----------------    ------
    8/31/2016 10:31:39 PM          1 MyReadingDimmer-XB1 ST 255
    
    This shows a debug level 1 event (Status/Operational). The address in the raw text is resolved into the name of the device. 
    This event immediately followed the event in the example above.

.EXAMPLE
    'Wed 08/31/2016 22:31:39 : [INST-SRX    ] 02 50 16.96.98 1F.22.5F 23 11 FF    LTONRR (FF)' | Convert-ISYEventViewerStringToObject

    datetime              debuglevel Device/Code/Comm action
    --------              ---------- ---------------- ------
    8/31/2016 10:31:39 PM          3 INST-SRX         02 50 16.96.98 1F.22.5F 23 11 FF    LTONRR (FF)

    A level 3 event (Device communication events) which immediately followed the event shown above.

.EXAMPLE
    Get-Clipboard | Convert-ISYEventViewerStringToObject

    datetime              debuglevel Device/Code/Comm           action
    --------              ---------- ----------------           ------
    8/31/2016 10:55:25 PM          1 X10                        D5
    8/31/2016 10:55:26 PM          1 X10                        D5/Off (11)
    9/1/2016 1:52:36 AM            1 NTP                        Setting Time From NTP
    9/1/2016 2:17:45 AM            1 HallwayDimmer              DON   0
    9/1/2016 2:17:45 AM            1 HallwayDimmer              ST 255
    9/1/2016 2:25:21 AM            1 X10                        D5/All Units Off (13)
    9/1/2016 2:25:25 AM            1 HallwayDimmer              ST   0
    9/1/2016 2:25:26 AM            1 MasterBdrmCenterDimmer-XC2 ST   0
       :
       :

    This shows a way to display all the events from the ISY Event Viewer window. Just click the copy button in an open
    Event Viewer window and then use Get-Clipboard to send them to the function. The debug level in this case was set to 1
    so only those events were emitted.

.EXAMPLE
    Get-Clipboard | Convert-ISYEventViewerStringToObject

    datetime              debuglevel Device/Code/Comm           action
    --------              ---------- ----------------           ------
       :                               :                           :
       :                               :                           :
    9/1/2016 8:20:08 PM            3 INST-SRX                   02 50 16.96.98 00.00.01 C7 11 00    LTONRR (00)
    9/1/2016 8:20:08 PM            3 Std-Group                  16.96.98-->Group=1, Max Hops=3, Hops Left=1
    9/1/2016 8:20:08 PM            2 D2D EVENT                  Event [16 96 98 1] [DON] [0] uom=0 prec=-1
    9/1/2016 8:20:08 PM            1 MyReadingDimmer-XB1        DON   0
    9/1/2016 8:20:08 PM            2 D2D EVENT                  Event [16 96 98 1] [ST] [255] uom=0 prec=-1
    9/1/2016 8:20:08 PM            1 MyReadingDimmer-XB1        ST 255
    9/1/2016 8:20:09 PM            3 INST-SRX                   02 50 16.96.98 1F.22.5F 41 11 01    LTONRR (01)
    9/1/2016 8:20:09 PM            3 Std-Cleanup                16.96.98-->ISY/PLM Group=1, Max Hops=1, Hops Left=0
    9/1/2016 8:20:09 PM            3 INST-DUP                   Previous message ignored.
    9/1/2016 8:20:22 PM            3 INST-SRX                   02 50 16.96.98 00.00.01 C7 13 00    LTOFFRR(00)
    9/1/2016 8:20:22 PM            3 Std-Group                  16.96.98-->Group=1, Max Hops=3, Hops Left=1
    9/1/2016 8:20:22 PM            2 D2D EVENT                  Event [16 96 98 1] [DOF] [0] uom=0 prec=-1
    9/1/2016 8:20:22 PM            1 MyReadingDimmer-XB1        DOF   0
    9/1/2016 8:20:22 PM            2 D2D EVENT                  Event [16 96 98 1] [ST] [0] uom=0 prec=-1
    9/1/2016 8:20:22 PM            1 MyReadingDimmer-XB1        ST   0
    9/1/2016 8:20:22 PM            3 INST-SRX                   02 50 16.96.98 1F.22.5F 41 13 01    LTOFFRR(01)
    9/1/2016 8:20:22 PM            3 Std-Cleanup                16.96.98-->ISY/PLM Group=1, Max Hops=1, Hops Left=0
    9/1/2016 8:20:22 PM            3 INST-DUP                   Previous message ignored.

    With the debug level set to 3 this example shows the events produced by turning a single light (MyReadingDimmer-XB1)  
    on at full brightness (ST 255) and then off (ST   0) 13 seconds later using the physical dimmer switch.
#>
function Convert-ISYEventViewerStringToObject {
[cmdletbinding()]
[outputtype([pscustomobject[]])]
param(
# ISY event strings
[parameter(ValueFromPipeline=$true,Position=0)]
[string[]]$line,

# Do not compress repeated events with '(Repeated n times)' appended to event action
[switch]$nocompress)

begin
{
    $instdevrgx = ((Get-InsteonDevice -dbmap).ForEach('address') -join '|') 
    Write-Verbose -Message "Device regex = `n$instdevrgx"
    $instcmdsrgx = 'DON|DOF|DFON|DFOF|BMAN|SMAN|OL|RR|ST|BEEP'
    $script:previous = ''
    $script:repeats = 0
    $script:repmsg = ''
    $splat = @{nocompress = $nocompress.IsPresent}

    function Objectify-Line {
    [cmdletbinding()]
    param(
    [ValidateNotNullOrEmpty()]
    [string]$text,
    
    [switch]$nocompress)

    if (!($nocompress) -and ($text -eq $script:previous))
    {
        $script:repeats++
        $script:repmsg = " (Repeated $script:repeats times)"
    }
    else
    {
        $script:repeats = 0

        $first = ($text -split ' : ' ,2).ForEach('trim')
        # 'Setting Time From NTP' is a special case
        # add [NTP] to the message to fit the pattern
        if ($first[1] -eq 'Setting Time From NTP')
        {
            $first[1] = '[NTP] Setting Time From NTP'
        }

        try
        {
            $datetime = [datetime]($first[0])
        }
        catch
        {
            Write-Warning -Message "1st token is not a datetime." 
            $datetime = $first[0]
        }

        $second = (($first[1] -replace '^\[', '') -split '\]' ,2).ForEach('trim')
 
        switch -Regex ($second[0])
        {   
            $instdevrgx 
            {
                $devicename = Resolve-NodeAddress -address $second[0]
                Write-Verbose -Message "Devicename = $devicename"
            
                switch -Regex ($second[1]) 
                {
                    $instcmdsrgx 
                    {
                        $psobj = [pscustomobject]@{datetime=$datetime
                                                   debuglevel=1
                                                   'Device/Code/Comm'=$devicename
                                                   action=($second[1] + $script:repmsg)}
                    }

                    Default 
                    {
                        Write-Warning -Message "Unknown command in $($second[1])"
                        $psobj = [pscustomobject]@{datetime=$datetime
                                                   debuglevel=1
                                                   'Device/Code/Comm'=$devicename
                                                   action=($second[1] + $script:repmsg)}                 
                    }
                }        
            }

            '^X10$|^IR$|^NTP$' 
            {
                $psobj = [pscustomobject]@{datetime=$datetime
                                           debuglevel=1
                                           #INSTEONDevice=$($second[0])
                                           'Device/Code/Comm'=$($second[0])
                                           action=($second[1] + $script:repmsg)}
            }

            '^D2D|^TIME'
            {
                $psobj = [pscustomobject]@{datetime=$datetime
                                           debuglevel=2
                                           #INSTEONLevel2Code=$($second[0])
                                           'Device/Code/Comm'=$($second[0])
                                           action=($second[1] + $script:repmsg)}       
            }

            Default # 2 examples: '^INST-|^X10-'
            {
                $psobj = [pscustomobject]@{datetime=$datetime
                                           debuglevel=3
                                           #INSTEONCommunication=$($second[0])
                                           'Device/Code/Comm'=$($second[0])
                                           action=($second[1] + $script:repmsg)}            
            }

        } #end switch
        $script:repmsg = ''
        $psobj
    }

    $script:previous = $text

} #end function Objectify-Line

} #end begin block

process
{
    if (($PSCmdlet.MyInvocation.ExpectingInput))
    {
        if ($_ -ne '')
        {
            $propbag = Objectify-Line -text $_ @splat
            if ($propbag)
            {
                $propbag
            }
        }
    }
} #end process block

end
{
    #add code here if you need to use the function w/o the pipeline
    # if (!($PSCmdlet.MyInvocation.ExpectingInput)) {foreach $l in $line {objectify-line $l}}
} #end end block


} 
#end function Convert-ISYEventToObject

<#
.Synopsis
   ConvertX-ElemToXMLDOC

   Converts an xmlElement object into an xmlDocument object.
     
.DESCRIPTION
    xmlElement objects are convenient to deal with with regards to examining their content.
    There is not as much typing required to get out the info you need out. This is why most of
    the time the ISY functions operate on xmlElements. Elements are missing some very useful
    methods (save(), create() etc) however so at times it does become neccessary to convert them  
    into xmlDocuments in order to exploit these missing methods.

    The function requires an xmlElement object or an array of such objects.
    If the -rootname parameter is supplied this will be the rootname of the XML document created.
    If the -rooname parameter is NOT supplied the function will attempt to get the orginal rootname
    from the supplied xmlElement(s). If this comes up empty the string 'root' will be used.
    
.EXAMPLE
    Below the device list xmlElement (actually an array of xmlElements) is retrieved. Examining the type of the 1st element
    of $instdlist shows the type as XmlElement.

    > $instdlist = Get-InsteonDevice -listdb

    > $instdlist[0].GetType()

    IsPublic IsSerial Name                                     BaseType
    -------- -------- ----                                     --------
    True     False    XmlElement                               System.Xml.XmlLinkedNode

    After conversion we have an xmlDocument with many more useful methods.

    > $instdlistxdoc = Convert-XElemToXMLDOC $instdlist -rootname device
    > $instdlistxdoc.GetType()

    IsPublic IsSerial Name                                     BaseType
    -------- -------- ----                                     --------
    True     False    XmlDocument                              System.Xml.XmlNode

    After using these sundry methods to modify the information (add properties, elements etc) to get the array of xmlElements back
    simply use: $newxelems = $instdlistxdoc.device


#>
function Convert-XElemToXMLDOC { 
[CmdletBInding()]
[OutputType([System.Xml.XmlDocument])]
param(
[Parameter(mandatory=$true)]
[System.Xml.XmlElement[]]$xelem,
[string]$rootname)

    function Get-XMLElemRootName {
    [cmdletbinding()]
    param([System.Xml.XmlElement[]]$xelem)

    $xdoctop = $xelem.ownerdocument.documentelement.parentnode | select -First 1

    $propnames = ($xdoctop | Get-Member -MemberType Properties).Name

    $xdoctoppropname = $propnames | where {$xdoctop.$_.haschildnodes -ne $null -and $xdoctop.$_.haschildnodes }

    if (@($xdoctoppropname).count -ne 1)
    {
        Write-Error -Message 'wtf?'
        return
    }

    $xdoctoppropname
}

if (!$rootname)
{
    $rootname = Get-XMLElemRootName $xelem
    Write-Verbose -Message "rootname from Get-XMLElemRootName = $rootname"
}

if (!$rootname)
{
    $rootname = 'root'
    Write-Verbose -Message "Resorting to using 'root' for rootname"
}

$XmlDoc = New-Object System.XML.XMLDocument

[System.XML.XMLElement]$XMLRoot=$XmlDoc.CreateElement($rootname)

[void]$XmlDoc.AppendChild($XMLRoot)  

$XmlTarget=$XmlDoc.selectSingleNode($rootname)

$xelem | ForEach-Object {[void]$XmlTarget.AppendChild($XmlDoc.ImportNode($_, $true))}

return $XmlDoc
}
#end function Convert-XElemToXMLDOC

<#
.Synopsis
   Formats an XML object/file with indentation and outputs to a file and the console
.DESCRIPTION
   XML objects and files can be hard for ordinary humans to read.
   This function produces a string (and a file containing the output as well) of the xml, each line indented 
   by an amount based on its depth in the xml object. 

   The function can take a path to an xml file (-pathxmlfile), an object array of xmlelements or an xmldocument object (-xml) as input.
   The xml object(s) can be sent from the pipeline or the command line. The -pathxmlfile only accepts input from the command line.

   There are 3 other parameters:
   -destination is the fullname of the file to store the output. Default is $env:temp/out.xml
   -indent is the number of indent characters for each indent. Default is 1, minimum 1, maximum 5
   -indentchar is the character to use for indenting. Default is tab ("`t")

    Inspired from:
    http://www.powershellmagazine.com/2013/08/19/mastering-everyday-xml-tasks-in-powershell/
    https://blogs.msdn.microsoft.com/powershell/2008/01/18/format-xml/

.EXAMPLE
    $isydb.GetType().name
    XmlDocument

    PS C:\>$isydb | Format-XML | more
    <?xml version="1.0" encoding="UTF-8"?>
    <nodes>
            <root>ISYRoot</root>
            <folder flag="0" fullpath="ISYRoot/Outside">
                    <address>1635</address>
                    <name>Outside</name>
            </folder>
            <folder flag="0" fullpath="ISYRoot/Downstairs/Kitchen">
                    <address>24242</address>
                    <name>Kitchen</name>
                    <parent type="FOLDER">Downstairs</parent>
            </folder>
            <folder flag="0" fullpath="ISYRoot/Upstairs/MediaRoom">
                    <address>34455</address>
                    <name>MediaRoom</name>
                    <parent type="FOLDER">Upstairs</parent>
            </folder>
            <folder flag="0" fullpath="ISYRoot/Downstairs/MasterBedroom">
                    <address>5143</address>
                    <name>MasterBedroom</name>
                    <parent type="FOLDER">Downstairs</parent>
            </folder>
            <folder flag="0" fullpath="ISYRoot/Garage">
                    <address>53244</address>
                    <name>Garage</name>
            </folder>
            <folder flag="0" fullpath="ISYRoot/Upstairs">
                    <address>5677</address>
                    <name>Upstairs</name>
            </folder>
            <folder flag="0" fullpath="ISYRoot/Downstairs">
                    <address>8745</address>
                    <name>Downstairs</name>
            </folder>
            <folder flag="0" fullpath="ISYRoot/RemoteControls">
                    <address>9966</address>
                    <name>RemoteControls</name>
            </folder>
            <node flag="128" fullpath="ISYRoot/Outside">
               :
               :

        This shows an xmldocument object being piped into the function.
        The string output is always sent to the screen and to the a file specified by the -destination value. 
.EXAMPLE
    cat C:\isy\ISYDB.xml | more
    <?xml version="1.0" encoding="UTF-8"?><nodes><root>ISYRoot</root><folder flag="0" f...
    llpath="ISYRoot/Downstairs/Kitchen"><address>24242</address><name>Kitchen</name><pa
    oom"><address>34455</address><name>MediaRoom</name><parent type="FOLDER">Upstairs</...
    ss><name>MasterBedroom</name><parent type="FOLDER">Downstairs</parent></folder><fol...

    This shows an xml file with no indentation

    PS C:\>format-xml -pathxmlfile C:\isy\ISYDB.xml | more
    <?xml version="1.0" encoding="UTF-8"?>
    <nodes>
            <root>ISYRoot</root>
            <folder flag="0" fullpath="ISYRoot/Outside">
                    <address>1635</address>
                    <name>Outside</name>
            </folder>
            <folder flag="0" fullpath="ISYRoot/Downstairs/Kitchen">
                    <address>24242</address>
                    <name>Kitchen</name>
                    <parent type="FOLDER">Downstairs</parent>
            </folder>
            <folder flag="0" fullpath="ISYRoot/Upstairs/MediaRoom">
                    <address>34455</address>
                    <name>MediaRoom</name>
                    <parent type="FOLDER">Upstairs</parent>
            </folder>
              :
              :

    PS C:\>cat C:\Users\ralf\AppData\Local\Temp\out.xml | more
    <?xml version="1.0" encoding="UTF-8"?>
    <nodes>
            <root>ISYRoot</root>
            <folder flag="0" fullpath="ISYRoot/Outside">
                    <address>1635</address>
                    <name>Outside</name>
            </folder>
            <folder flag="0" fullpath="ISYRoot/Downstairs/Kitchen">
                    <address>24242</address>
                    <name>Kitchen</name>
                    <parent type="FOLDER">Downstairs</parent>
                      :
                      :

    After the function executes the indented file is saved in the default location.
#>
function Format-XML 
{
[cmdletbinding()]
param
(
# either an xmldocument object or an array of xmlelements
[parameter(ParameterSetName='object',Mandatory=$true,Position=0,ValueFromPipeline)]
[ValidateScript({$_.gettype().name -match 'xml(document|element)'})]
[object]$xml, 

# path to an xml file
[parameter(ParameterSetName='file',Mandatory=$true,Position=0)]
[ValidateScript({Test-Path -Path $_ -PathType Leaf })]
[string]$pathxmlfile,

# fullname of location to create a file containing the output
[string]$destination="$env:TEMP\out.xml",

# number of indent characters for each indent
[ValidateRange(1,5)]
[int]$indent=1, 

# character to use for indenting
[char]$indentchar="`t"
)

BEGIN 
{
    function Format-XDoc {
    [cmdletbinding()]
    param([xml]$xdoc)

    $StringWriter = New-Object System.IO.StringWriter
    $xwritersettings = @{Indent = $true
                        IndentChars = ([string]($indentchar)*$indent)
                        NewLineChars = "`r`n"
                        NewLineOnAttributes = $false
                        NewLineHandling = 'replace'
                        }
    $XmlWriter = [System.Xml.XmlWriter]::Create($StringWriter, $xwritersettings)

    $xdoc.Save($XmlWriter)

    $StringWriter.tostring()
    }    
}

END
{   
    if($PSCmdlet.ParameterSetName -eq 'file')
    {
        $xmldoc = New-Object -TypeName xml
        $xmldoc.load($pathxmlfile)   
    }
    else
    {
        if (($PSCmdlet.MyInvocation.ExpectingInput)) 
        {
            $xml = @($input)
            Write-Verbose -Message "Pipeline input detected"
        }

        if (($xml | Get-Member).typename -eq 'System.Xml.XmlElement')
        {
            $xmldoc = Convert-XElemToXMLDOC $xml -rootname root
            Write-Verbose -Message "Converting xmlelement(s) to xmldocument"
        }
        else
        {
            $xmldoc = @($xml)[0]
        }
    }

    $xstring = Format-XDoc -xdoc $xmldoc
    Set-Content -Value $xstring -Path $destination
    $xstring

<#
# cull the xml tags
$zerothpass = $$StringWriter.ToString() -replace '^\<\?xml version.*\>',''

$replacethese = [regex]::Matches($zerothpass,'\</\w+\>').value

$replregx = (($replacethese | sort -Unique) -join '|').replace('<','\<').replace('>','\>')

$firstpass = $zerothpass -replace $replregx,''

$secondpass = $firstpass -replace '(\<.+\=.+)\>','$1'

$thirdpass = $secondpass.replace('<','').replace('>',': ')

$fourthpass = $thirdpass -replace '(\w+)( \w+\=.+)', '$1:$2'

#>
}

}
#end of function Format-XML

<#
.Synopsis
   Get-Headers

   Returns the -Headers parameter to be used in the Invoke-Restmethod call. 
.DESCRIPTION
   This function takes the username and password to build the headers value.
#>
function Get-Headers {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
    [AllowEmptyString()]
    [string]$SOAPAction = '')

    Write-Verbose -Message ($MyInvocation.MyCommand.Name + ' called.') 

    $currentsettings = (Get-ISYSettings)

    $user = $currentsettings.username
    $pass = $currentsettings.GetNetworkCredential().Password
    
    $headers = @{
    Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($("{0}:{1}" -f $user, $pass)))
    }
    Write-Verbose "headers = user: $user, password: System.Security.SecureString" 
    if ($SOAPAction -eq '')
    {
        return $headers
    }
    $headers.add('SOAPAction',$SOAPAction)
    $headers
}
#end of function Get-Headers

<#
.Synopsis
   Get-Insteondevice

   Gets the status of an Insteon device or a list of Insteon devices.  
.DESCRIPTION
    Using the name parameter this function initiates a REST query to the insteon devices specified and returns the current 
    status as well as other properties of the devices to the caller.
    The name parameter is a list of strings that can be wildcards, or regular expressions if the -regx switch is used. 
    Each wildcard or regular expression string is used to derive a list of devices and these devices are unioned to produce 
    the final list of devices to get. Note that with no parameters all devices are returned (name = '*'). 
    
    There are 3 other switches supported. 
    -dbmap and -listdb are mutually exclusive with the name parameter and regx switch, as well as mutually exclusive with each other.
        
        -dbmap provides a cached (from memory) name to address map of the insteon devices. No REST query is perfomed.
        
        -listdb returns a cached (from disk) intsteon device list. Again no REST query is done. 
    
    -nocustomdefaultdisplayproperties leaves the default properties displayed unmodified.
        Unmodified properties are:
            flag,category,fullpath,address,name,parent,type,enabled,deviceClass,wattage,dcPeriod,pnode,ELK_ID,property
        and with the -listdb switch:
            OnLevel%,flag,category,fullpath,OnLevel,address,name,parent,type,enabled,deviceClass,wattage,dcPeriod,pnode,ELK_ID,property

        Without the -nocustomdefaultdisplayproperties switch the default displayed properties are:
            name,address,fullpath,onlevel%
        and with the -listdb switch
            name,address,fullpath,type 

        The -dbmap output is unaffected by the -nocustomdefaultdisplayproperties switch

.EXAMPLE
    Get-Insteondevice -name GuestReadingDimmer-XB3 -nocustomdefaultdisplayproperties | Format-Table -AutoSize

    OnLevel% flag category  fullpath                   OnLevel address    name                   parent type                                 enabled
    -------- ---- --------  --------                   ------- -------    ----                   ------ ----                                 -------
    51%      128  Dimmables ISYRoot/Upstairs/MediaRoom 128     16 96 86 1 GuestReadingDimmer-XB3 parent SWITCHLINC_DIMMER_W_SENSE_2476D v.38 true

    When we query the GuestReadingDimmer-XB3 we see it is on at a 51% level. OnLevel in this case is 128, Onlevel ranges from 0 to 255.

.EXAMPLE
    Get-InsteonDevice *dimmer*,*switch* | ft name,address,OnLevel%,OnLevel,fullpath,type -AutoSize

    name                        address    OnLevel% OnLevel fullpath                         type
    ----                        -------    -------- ------- --------                         ----
    AtticDimmer                 36 72 49 1 Off      0       ISYRoot/Upstairs                 SWITCHLINC_DIMMER_2WIRE_2474DWH v.42
    BackDoorDimmer-XD5          27 78 34 1 On       255     ISYRoot/Outside                  SWITCHLINC_DIMMER_2477D v.41
    CeciliaInlineDimmer-XC4     20 95 87 1 Off      0       ISYRoot/Downstairs/MasterBedroom INLINELINC_DIMMER_2475DA1 v.41
    CooktopDimmer               1 60 30 1  Off      0       ISYRoot/Downstairs/Kitchen       SWITCHLINC_V2_DIMMER_2476D v.27
    DresserInlineDimmer-XC5     20 95 49 1 Off      0       ISYRoot/Downstairs/MasterBedroom INLINELINC_DIMMER_2475DA1 v.41
    FrontPorchDimmer            27 78 45 1 31       78      ISYRoot/Outside                  SWITCHLINC_DIMMER_2477D v.41
    GarageBackDoorDimmer        16 96 86 1 Off      0       ISYRoot/Outside                  SWITCHLINC_DIMMER_W_SENSE_2476D v.38
    GarageCarDoorDimmer-XD3     17 6 82 1  Off      0       ISYRoot/Outside                  SWITCHLINC_DIMMER_2477D v.40
    GarageCeilingLightSwitch    33 85 34 1 Off      0       ISYRoot/Garage                   SWITCHLINC_RELAY_DUAL_BAND_2477S v.43
    GarageKitchenEntranceDimmer 24 84 71 1 Off      0       ISYRoot/Downstairs/Kitchen       KEYPAD_LINC_DIMMER_2334_2_5_BUTTON v.43
    GuestReadingDimmer-XB3      16 96 86 1 51       128     ISYRoot/Upstairs/MediaRoom       SWITCHLINC_DIMMER_W_SENSE_2476D v.38
    HallwayDimmer               27 78 40 1 Off      0       ISYRoot/Downstairs               SWITCHLINC_DIMMER_2477D v.41
    JimInlineDimmer-XC3         20 98 15 1 Off      0       ISYRoot/Downstairs/MasterBedroom INLINELINC_DIMMER_2475DA1 v.41
    KitchenTableDimmer          1 75 61 1  Off      0       ISYRoot/Downstairs/Kitchen       SWITCHLINC_V2_DIMMER_2476D v.27
    KitchenTableDimmerKTKP1     9 93 11 1  Off      0       ISYRoot/Downstairs/Kitchen       KEYPADLINC_DIMMER_2486D v.29
    MasterBdrmCenterDimmer-XC2  16 96 81 1 Off      0       ISYRoot/Downstairs/MasterBedroom SWITCHLINC_DIMMER_W_SENSE_2476D v.38
    MiddleReadingDimmer-XB2     16 96 87 1 Off      0       ISYRoot/Upstairs/MediaRoom       SWITCHLINC_DIMMER_W_SENSE_2476D v.38
    MyReadingDimmer-XB1         16 96 98 1 Off      0       ISYRoot/Upstairs/MediaRoom       SWITCHLINC_DIMMER_W_SENSE_2476D v.38
    SlidingDoorOutDimmer-XD6    16 98 41 1 Off      0       ISYRoot/Outside                  SWITCHLINC_DIMMER_W_SENSE_2476D v.38

    Finds all the devices with switch or dimmer in the name. Note "Get-InsteonDevice 'switch|dimmer' -regx" would give the same result.

.EXAMPLE
    Get-InsteonDevice -listdb -nocustomdefaultdisplayproperties | ft -autosize

    flag fullpath                         address    name                         parent type                                    enabled deviceClass wattage dcPeriod
    ---- --------                         -------    ----                         ------ ----                                    ------- ----------- ------- --------
    128  ISYRoot/Outside                  16 96 86 1 GarageBackDoorDimmer         parent SWITCHLINC_DIMMER_W_SENSE_2476D v.38    true    256         300     60
    128  ISYRoot/Upstairs/MediaRoom       16 96 87 1 MiddleReadingDimmer-XB2      parent SWITCHLINC_DIMMER_W_SENSE_2476D v.38    true    512         300     60
    128  ISYRoot/Downstairs/MasterBedroom 16 96 81 1 MasterBdrmCenterDimmer-XC2   parent SWITCHLINC_DIMMER_W_SENSE_2476D v.38    true    512         300     60
    128  ISYRoot/Upstairs/MediaRoom       16 96 86 1 GuestReadingDimmer-XB3       parent SWITCHLINC_DIMMER_W_SENSE_2476D v.38    true    512         300     60
    128  ISYRoot/Upstairs/MediaRoom       16 96 98 1 MyReadingDimmer-XB1          parent SWITCHLINC_DIMMER_W_SENSE_2476D v.38    true    512         300     60
    128  ISYRoot/Outside                  16 98 41 1 SlidingDoorOutDimmer-XD6     parent SWITCHLINC_DIMMER_W_SENSE_2476D v.38    true    256         300     60
    128  ISYRoot/Outside                  17 66 28 1 DriveWayAlert-Sensor-XD1     parent IO_LINC_2450 v.36                       true    0           0       0
    0    ISYRoot/Upstairs/MediaRoom       17 66 28 2 Unemployed-Relay             parent IO_LINC_2450 v.36                       true    0           0       0
    128  ISYRoot/Outside                  17 6 82 1  GarageCarDoorDimmer-XD3      parent SWITCHLINC_DIMMER_2477D v.40            true    256         300     60
    128  ISYRoot/Downstairs/MasterBedroom 20 95 49 1 DresserInlineDimmer-XC5      parent INLINELINC_DIMMER_2475DA1 v.41          true    512         300     60
    128  ISYRoot/Downstairs/MasterBedroom 20 95 87 1 CeciliaInlineDimmer-XC4      parent INLINELINC_DIMMER_2475DA1 v.41          true    512         300     60
    128  ISYRoot/Downstairs/MasterBedroom 20 98 15 1 JimInlineDimmer-XC3          parent INLINELINC_DIMMER_2475DA1 v.41          true    512         300     60
    128  ISYRoot/Outside                  27 78 34 1 BackDoorDimmer-XD5           parent SWITCHLINC_DIMMER_2477D v.41            true    512         300     60
    128  ISYRoot/Downstairs               27 78 40 1 HallwayDimmer                parent SWITCHLINC_DIMMER_2477D v.41            true    512         300     60
    128  ISYRoot/Outside                  27 78 45 1 FrontPorchDimmer             parent SWITCHLINC_DIMMER_2477D v.41            true    0           0       0
    128  ISYRoot/Downstairs/Kitchen       24 84 71 1 GarageKitchenEntranceDimmer  parent KEYPAD_LINC_DIMMER_2334_2_5_BUTTON v.43 true    0           0       0
    0    ISYRoot/Downstairs/Kitchen       24 84 71 3 GarageCeilingLightButtonGKKP parent KEYPAD_LINC_DIMMER_2334_2_5_BUTTON v.43 true    0           0       0
    0    ISYRoot/Downstairs/Kitchen       24 84 71 4 GarageKitchenKeypad-4        parent KEYPAD_LINC_DIMMER_2334_2_5_BUTTON v.43 true    0           0       0
    0    ISYRoot/Downstairs/Kitchen       24 84 71 5 GarageKitchenKeypad-5        parent KEYPAD_LINC_DIMMER_2334_2_5_BUTTON v.43 true    0           0       0
    0    ISYRoot/Downstairs/Kitchen       24 84 71 6 GarageKitchenKeypad-6        parent KEYPAD_LINC_DIMMER_2334_2_5_BUTTON v.43 true    0           0       0
    128  ISYRoot/Garage                   33 85 34 1 GarageCeilingLightSwitch     parent SWITCHLINC_RELAY_DUAL_BAND_2477S v.43   true    512         300     60
    128  ISYRoot/Upstairs                 36 72 49 1 AtticDimmer                  parent SWITCHLINC_DIMMER_2WIRE_2474DWH v.42    true    0           0       0
    128  ISYRoot/Downstairs/Kitchen       9 93 11 1  KitchenTableDimmerKTKP1      parent KEYPADLINC_DIMMER_2486D v.29            true    0           0       0
     :
     :
    128  ISYRoot/Downstairs/Kitchen       1 60 30 1  CooktopDimmer                parent SWITCHLINC_V2_DIMMER_2476D v.27         true    512         300     60
    128  ISYRoot/Downstairs/Kitchen       1 75 61 1  KitchenTableDimmer           parent SWITCHLINC_V2_DIMMER_2476D v.27         true    512         300     60

    Using the -listdb option returns a full unaltered device list but note the status is not current. The status is under the property named 'property'. 
    The status with the -listdb switch holds the status of the devices when the ISY db was saved to disk.

.EXAMPLE
    Get-InsteonDevice -dbmap | ft -AutoSize

    name                         LocalName address
    ----                         --------- -------
    GarageBackDoorDimmer         node      16 96 86 1
    MiddleReadingDimmer-XB2      node      16 96 87 1
    MasterBdrmCenterDimmer-XC2   node      16 96 81 1
    GuestReadingDimmer-XB3       node      16 96 86 1
    MyReadingDimmer-XB1          node      16 96 98 1
    SlidingDoorOutDimmer-XD6     node      16 98 41 1
    DriveWayAlert-Sensor-XD1     node      17 66 28 1
    Unemployed-Relay             node      17 66 28 2
    GarageCarDoorDimmer-XD3      node      17 6 82 1
    DresserInlineDimmer-XC5      node      20 95 49 1
    CeciliaInlineDimmer-XC4      node      20 95 87 1
    JimInlineDimmer-XC3          node      20 98 15 1
    BackDoorDimmer-XD5           node      27 78 34 1
    HallwayDimmer                node      27 78 40 1
    FrontPorchDimmer             node      27 78 45 1
    GarageKitchenEntranceDimmer  node      24 84 71 1
     :
     :
    CooktopDimmer                node      1 60 30 1
    KitchenTableDimmer           node      1 75 61 1

    THe -dbmap switch shows the name to address map of the insteon devices.
.EXAMPLE

    Get-InsteonDevice FrontPorchDimmer

    OnLevel% address    fullpath        name
    -------- -------    --------        ----
    Off      27 78 45 1 ISYRoot/Outside FrontPorchDimmer

    Here we see the default display properties using no format-* cmdlet.

#>
function Get-InsteonDevice {
[CmdletBinding(DefaultParameterSetName='filter')]
[OutputType([System.XML.XMLElement[]])]
param(
# an array of wildcard spec(s) or regular expression(s) representing the device name(s) to get
[Parameter(ParameterSetName='filter',Position=0,ValueFromPipelineByPropertyName,ValueFromPipeline)]
[object[]]$name = '*',

# specifies that the name string(s) are to be used as regular expression(s) not wildcards
[Parameter(ParameterSetName='filter')]
[switch]$regx,

# gets the name to address mapping from the cached data base - used mostly for tabexpansion
[Parameter(ParameterSetName='map')]
[switch]$dbmap,

# gets the cached list of devices with all properties from the database - no direct query to the device is performed 
[Parameter(ParameterSetName='list')]
[switch]$listdb,

# if set the default displayed properties are left unmodified.
[switch]$nocustomdefaultdisplayproperties
)

Begin
{
    Write-Verbose -Message ($MyInvocation.MyCommand.Name + ' called.') 
    
    if ($dbmap)
    {
        return ( $Script:list_of_insteon_devices.foreach({$_.PSObject.Copy()}) ) # send a copy so caller cannot change value
    }

    $devicelist = (Get-ISYDB).nodes.node

    if ($listdb)
    {
        if ($nocustomdefaultdisplayproperties)
        {
            return $devicelist
        }
        else 
        {
            return ($devicelist | Set-DefaultDisplayProperties -PropertyList name,address,fullpath,type)
        }
    }

    $status = @()
    $levelpct = [ScriptBlock]::Create('$this.property.formatted')
    $level = [ScriptBlock]::Create('$this.property.value')

    $Wildcardhash = @{$true = 'matching regular expressions '; $false = 'like wildcard(s) '}

    #TODO walk thru all name entries and remove redundancies for wildcards - ie if any one entry -eq '*' forget all the rest

    $deviceelems = @()
}

End
{
    if ($listdb -or $dbmap)
    {
        return
    }
    
    if ($PSCmdlet.MyInvocation.ExpectingInput)
    {
        $array = @($input)
        Write-Verbose -Message "Pipeline input found with $($array.count) device members"
    }
    else
    {
        $array = $name
        Write-Verbose -Message "Command line input found with $($array.count) device members"
    }

    if (($array | Get-Member -MemberType Properties).name -contains 'name')
    {
        $localname = ($array.localname | Sort-Object -Unique) 
        if ($localname -ne 'node')
        {
            Write-Error -Message "$localname object is not an Insteon device." -Category InvalidData
            return
        }
        $entries = $($array.name)
        Write-Verbose -Message "Objects have name property. $($entries.count) device names found."
    }
    else
    {
        $entries = $array
        Write-Verbose -Message "Objects have NO name property. $($entries.count) device names found."
    }

    foreach ($entry in $entries)
    {  
        if (!$regx)
        {
            $deviceelems  += @($devicelist | where name -like $entry)
        }
        else
        {
            try
            {
               [void]('ralf' -match $entry)
                $deviceelems += @($devicelist | where name -match $entry)
            }
            catch
            {
                Write-warning -Message "`'$entry`' is not a valid regular expression."
                continue
            }
        }
    }

    if (!$deviceelems)
    {
        Write-Warning -Message "No devices found $($Wildcardhash[$($regx.IsPresent)]) $($name -join ' or ')."
        return
    }

    $deviceelems = @($deviceelems | Sort-Object -Property name -Unique)

    foreach ($deviceelem in $deviceelems)
    {
        if ($deviceelems.count -le 4) # querying for 5 devices is about even with querying for all devices - verified using measure-command
        {
            $deviceelem  | ForEach-Object { $nodeid = $deviceelem.address
            $response = (Invoke-IsyRestMethod "/rest/status/$nodeid").properties
            $status += ($response | add-member -NotePropertyName Id -NotePropertyValue $nodeid -PassThru) }
        }
        else
        {
            if (!$fullstatus)
            {
                $fullstatus = Invoke-IsyRestMethod "/rest/status"
            }

            $status += $fullstatus.nodes.node | where id -eq $deviceelem.address

        }
    }

    $status | ForEach-Object { Add-Member -InputObject $_ -MemberType ScriptProperty -Name 'OnLevel%' -Value $levelpct -PassThru |
                             Add-Member -MemberType ScriptProperty -Name OnLevel -Value $level }

    foreach ($deviceelem in $deviceelems)
    {
        foreach ($stat in $status)
        {
            if ($stat.id -eq $deviceelem.address)
            {
                $deviceelem | Add-Member -NotePropertyName 'OnLevel%' -NotePropertyValue ($stat.'OnLevel%')
                $deviceelem.setattribute('OnLevel',$stat.OnLevel)
            }
        }
    }

    #$deviceelems.ForEach( { Add-Member -InputObject $_ -MemberType ScriptMethod -Name toString -Value {$this.name} -force } ) 
    if ($nocustomdefaultdisplayproperties)
    {
        return $deviceelems 
    }
    else
    {
        return ($deviceelems | Set-DefaultDisplayProperties -PropertyList name,address,onlevel%,fullpath)
    }
}

}
#end function Get-InsteonDevice

<#
.Synopsis
   Get-ISYConfig

   Retrieves all the configuration info of the ISY. 
.DESCRIPTION
   This function always gets the configuration info via a SOAP or REST call to the ISY.
   The default is to use SOAP but rge -REST switch will force a REST call instead.
.EXAMPLE
    Load the ISY configuration info into the $iconf variable. And examine some of the properties.

    $iconf = Get-ISYConfig

    $iconf

    deviceSpecs      : deviceSpecs
    upnpSpecs        : upnpSpecs
    controls         : controls
    driver_timestamp : 2014-10-30-08:07:30
    app              : Insteon_UD994
    app_version      : 4.2.18
    platform         : ISY-C-994
    build_timestamp  : 2014-10-30-08:07:30
    root             : root
    product          : product
    features         : features
    triggers         : true
    maxTriggers      : 2048
    variables        : true
    secsys           : secsys
    baseDriver       : baseDriver
    security         : security
    isDefaultCert    : true
    maxSSLStrength   : 2048

    $conf.deviceSpecs

    make                 : Universal Devices Inc.
    manufacturerURL      : http://www.universal-devices.com
    model                : ISY994i Series
    icon                 : /web/udlogo.jpg
    archive              : /web/insteon.jar
    chart                : /web/chart.jar
    queryOnInit          : true
    oneNodeAtATime       : true
    baseProtocolOptional : false

    $iconf.features.feature | ft -au

    id    desc                              isInstalled isAvailable
    --    ----                              ----------- -----------
    21010 OpenADR                           true        true
    21011 Electricity Monitor               false       true
    21012 Gas Meter                         false       true
    21013 Water Meter                       false       false
    21020 Weather Information               false       true
    21030 URL                               false       false
    21040 Networking Module                 false       true
    21050 AMI Electricity Meter             false       true
    21051 SEP ESP                           false       false
    21060 A10/X10 for INSTEON               false       true
    21070 Portal Integration - Check-it.ca  false       true
    21014 Current Cost Meter                false       false
    21080 Broadband SEP Device              false       true
    21071 Portal Integration - GreenNet.com false       true
    22000 RCS Zigbee Device Support         false       true
    23000 Irrigation/ETo Module             false       true
    21090 Elk Security System               false       true
    21072 Portal Integration - BestBuy.com  false       true
    24000 NorthWrite NOC Module             false       true
    21073 Portal Integration - MobiLinc     false       true
    21100 Z-Wave                            false       true
    25000 NCD Device Support                false       true
    21074 Portal Integration - VantagePoint false       true
    21075 Portal Integration - UDI          false       true
#>
function Get-ISYConfig {
    [CmdletBinding()]
    [OutputType([System.XML.XMLElement])]
    # If present config will be obtained via REST - default is SOAP
    param([switch]$rest)

    Write-Verbose -Message ($MyInvocation.MyCommand.Name + ' called.') 

    # verify isy settings before ISY access
    [void](Get-ISYSettings)

    # Why use SOAP as default? Format-XML cannot seem to properly format the xml from REST
    # Immediately following the driver_timestamp property there are no newlines. I gave up trying to debug this.
    # It looks like the REST xmldoc is already formatted with newlines up to and including driver_timestamp.
    # After this there are no newlines. This analysis is from examining the export-clixml files for REST vs SOAP xmldocument objects.   
    # The data from SOAP formats correctly using Format-XML.
    if ($rest)
    {
       return ((Invoke-IsyRestMethod "/rest/config").configuration)
    }

    return ((Invoke-ISYSOAPOperation -operation GetISYConfig).Envelope.Body.configuration)
}
#end function Get-ISYConfig

<#
.Synopsis
   Get-ISYDB

   ThIs function takes the ISYDB.xml file produced by the Update-ISYDBXMLFile function 
   and creates an xmlDocument object from it. It first calls Update-ISYDBXMLFile which
   if required will initiate a REST query to produce an updated ISYDB.xml file.
.DESCRIPTION
    Whenever a function requires static info about the nodes, scenes and folders it calls this 
    function and uses the  xmlDocument object produced.
.EXAMPLE
   The function has no parameters.
   Load the $isydb variable and display the folder info.
   $isydb = Get-ISYDB
   $isydb.nodes.folder | Format-Table -AutoSize

    flag address name
    ---- ------- ----
    0    1635    OutsideDevices
    0    24242   Kitchen Devices
    0    34455   MediaRoomDevices
    0    5143    MasterBedroomDevices
    0    9966    RemoteControls
#>
function Get-ISYDB {
    [CmdletBinding()]
    [OutputType([System.XML.XMLDocument])]
    param()

    Write-Verbose -Message ($MyInvocation.MyCommand.Name + ' called.') 

    $currentsettings = Get-ISYSettings
    Update-ISYDBXMLFile
    $isydb = New-Object -TypeName XML
    $isydb.Load((Join-Path $currentsettings.path $Script:ISYDBXMLFileName))
    return ($isydb) # | ForEach-Object{$_.PSObject.Copy()})
}
#end function Get-ISYDB

<#
.Synopsis
   Gets the current debug level of the ISY.
.DESCRIPTION
   Debug level will be either 1, 2 or 3. 
.EXAMPLE
    Get-ISYDebuglevel
    1
#>
function Get-ISYDebuglevel
{
    [CmdletBinding()]
    [Alias('GIDBG')]
    [OutputType([string])]
    Param()

    (Invoke-ISYSOAPOperation -operation GetDebugLevel).Envelope.Body.DBG.current

}
# end of function Get-ISYDebugLevel

<#
.Synopsis
   Get-ISYFolderList

   Retrieves all the folder info of the ISY. 
.DESCRIPTION
   This function always pulls the node info from via the Get-ISYDB function.
    
.EXAMPLE
    Retrieve the folder info and display it in table format. Note 
    Get-ISYFolderList | Format-Table -AutoSize

    flag address name
    ---- ------- ----
    0    1635    OutsideDevices
    0    24242   Kitchen Devices
    0    34455   MediaRoomDevices
    0    5143    MasterBedroomDevices
    0    9966    RemoteControls
#>
function Get-ISYFolderList {
    [CmdletBinding()]
    [OutputType([System.XML.XMLElement])]
    param()

    $folders = (Get-ISYDB).nodes.folder

    return ($folders) # | ForEach-Object{$_.PSObject.Copy()})

}
#end function Get-ISYFolderList

<#
.Synopsis
   Get-ISYNetwork

   Retrieves all the network info of the ISY. 
.DESCRIPTION
   This function always gets the network info via a REST call to the ISY.
.EXAMPLE
    $isynetwork = Get-ISYNetwork

    PS C:\>$isynetwork

    Interface WebServer ClientSecurity ServerSecurity
    --------- --------- -------------- --------------
    Interface WebServer ClientSecurity ServerSecurity

    PS C:\>$isynetwork.Interface

    isDHCP  : false
    isUPnP  : true
    ip      : 192.168.1.189
    mask    : 255.255.255.0
    gateway : 192.168.1.1
    dns     : 192.168.1.1


    PS C:\>$isynetwork.WebServer

    httpPort httpsPort
    -------- ---------
    80       443      

    PS C:\>$isynetwork.ClientSecurity

    protocol strength verifyPeer
    -------- -------- ----------
    3.3      3        false    

    First we assign the network config to the $isynetwork variable and then examine the data. 
#>
function Get-ISYNetwork {
    [CmdletBinding()]
    [OutputType([System.XML.XMLElement])]
    param()

    Write-Verbose -Message ($MyInvocation.MyCommand.Name + ' called.') 

    # verify isy settings before ISY access
    [void](Get-ISYSettings)


    return ((Invoke-IsyRestMethod "/rest/network").NetworkConfig)
    
}
#end function Get-ISYNetwork

<#
.Synopsis
   Get-ISYProgramList

   Retrieves the details of all of the ISY programs. 
.DESCRIPTION
   This function loads the ISYPrograms.xml file into an xml document object. If the -refreshnow switch is used,
   or the ISYPrograms.xml file does not exist or the ISYPrograms.xml file is older than ISYSettings.refreshtime
   the a REST query is performed to update the file. 
   Note program step details (actions in the if clause of a program for example) are not visible via the REST I/F (at least I cannot find them).
.EXAMPLE
    Place the node info in the $programs variable and then display it.
    $isyprograms = Get-ISYProgramList
    $isyprograms | Format-Table -AutoSize

    id   status folder name                    lastRunTime         lastFinishTime
    --   ------ ------ ----                    -----------         --------------
    0001 true   true   MyPrograms
    001C true   true   NIghtTime
    000D true   true   TODU
    001E true   false  HallwayLightOn          2015/03/09 19:00:44 2015/03/09 19:00:44
    0005 false  false  Turn Off Outside Lights 2015/03/09 18:45:45 2015/03/09 18:45:45
    0007 true   false  OutsideLightsOff        2015/03/09 18:36:05 2015/03/09 18:36:05
    001F true   false  HallwayLightOff         2015/03/09 03:00:00 2015/03/09 03:00:00
    0002 true   false  Query All               2015/03/09 03:00:00 2015/03/09 03:00:00
    0004 true   false  Kitchen Lights Off      2015/03/08 23:50:00 2015/03/09 00:02:01
    0003 true   false  Kitchen Lights On       2015/03/08 22:00:00 2015/03/08 22:24:31
    001A false  false  FloodlightsOff
    0019 false  false  MBRCenterLightOff
    0006 false  false  MBRCenterLightOn
    To get more fields use the Show-ISYObject function.
.EXAMPLE
    > Get-ISYProgramList | Show-ISYObject

    parent       name                                     folder id   status enabled runAtStartup running lastRunTime         lastFinishTime      nextScheduledRunTime
    ------       ----                                     ------ --   ------ ------- ------------ ------- -----------         --------------      --------------------
    MyPrograms   Turn Off Outside Lights                  false  0005 false  true    true         idle    2015/03/30 19:03:55 2015/03/30 19:03:55 2015/03/31 08:03:50
    MyPrograms   OutsideLightsOff                         false  0007 true   true    false        idle    2015/03/30 18:05:30 2015/03/30 18:05:30
    MyPrograms   Kitchen Lights On                        false  0003 true   true    true         idle    2015/03/30 22:00:00 2015/03/30 22:16:26 2015/03/31 22:00:00
    MyPrograms   HallwayLightOn                           false  001E true   true    false        idle    2015/03/30 19:18:54 2015/03/30 19:18:54 2015/03/31 19:19:44
    MyPrograms   MBRCenterLightOn                         false  0006 false  true    false        idle
    MyPrograms   MBRCenterLightOff                        false  0019 false  true    false        idle
    MyPrograms   Query All                                false  0002 true   true    false        idle    2015/03/30 03:00:01 2015/03/30 03:00:01 2015/03/31 03:00:00
    MyPrograms   FloodlightsOff                           false  001A false  true    false        idle
    MyPrograms   HallwayLightOff                          false  001F true   true    false        idle    2015/03/30 03:00:01 2015/03/30 03:00:01 2015/03/31 03:00:00
    MyPrograms   NightTime                                true   001C true                                                                        2015/03/31 07:33:50
    MyPrograms   Kitchen Lights Off                       false  0004 true   true    true         idle    2015/03/30 23:50:00 2015/03/31 00:05:51 2015/03/31 23:50:00
    MyPrograms   TODU                                     true   000D true
    NightTime    DriveWayAlertTurnOnGarageLight           false  001B false  true    false        idle    2015/03/30 22:44:49 2015/03/30 22:44:49
    NightTime    TurnOffDriveWayAlertActivatedGarageLight false  001D false  true    false        idle    2015/03/30 22:44:49 2015/03/30 22:44:49
    Summer       ControlDevicesSummer                     false  0008 false  false   false        idle
    TODU         Summer                                   true   000E false                                                                       2015/03/31 10:00:00
    TODU         TODUHolidays                             true   0016 true
    TODU         Winter                                   true   000F false                                                                       2015/03/31 06:00:00
    TODUHolidays New Years Day                            false  0014 false  false   false        idle                                            2016/01/01 00:00:00
    TODUHolidays Memorial Day                             false  0013 false  false   false        idle                                            2015/05/25 00:00:00
    TODUHolidays NotAnyHoliday                            false  0018 true   false   false        idle
    TODUHolidays Labor Day                                false  0012 false  false   false        idle                                            2015/09/07 00:00:00
    TODUHolidays Thanksgiving                             false  0015 false  false   false        idle                                            2015/11/26 00:00:00
    TODUHolidays Christmas Day                            false  0010 false  false   false        idle                                            2015/12/25 00:00:00
    TODUHolidays Independence Day                         false  0011 false  false   false        idle                                            2015/07/03 00:00:00
    TODUHolidays Good Friday                              false  0017 false  false   false        idle                                            2015/04/03 00:00:00
    Winter       ControlDevicesWinter                     false  0009 false  false   false        idle
                 MyPrograms                               true   0001 true 
#>
function Get-ISYProgramList {
[CmdletBinding()]
[OutputType([System.XML.XMLElement])]
param(
[switch]$refreshonly,
[switch]$refreshnow,
[switch]$fast)

Write-Verbose -Message ($MyInvocation.MyCommand.Name + ' called.') 

if ($fast) 
    {
    $fastsuccess = $false
    write-verbose "`n`$script:list_of_isy_programs exists = $(Test-Path variable:script:list_of_isy_programs)`n"
    if (Test-Path variable:script:list_of_isy_programs)
        {
            # have to psobject.copy() each element or only the reference is passed and thus variable $Script:list_of_* can be changed by caller
            $programs = $Script:list_of_isy_programs | ForEach-Object {$_.psobject.copy()}
            $fastsuccess = $true
            $refreshonly = $false
        }
    else 
        {
            $refreshnow = $true
            $refreshonly = $false
        }
    }

if (!$fastsuccess) 
{
    $currentsettings = Get-ISYSettings

    $refreshtime = $currentsettings.refreshtime

    $fullname = Join-Path $currentsettings.path $Script:ISYProgramsXMLFileName
    write-verbose "`n`$fullname = $fullname`n"

    if ($refreshnow -or $refreshonly -or !(Test-Path $fullname) -or (((Get-Date) - (Get-ChildItem $fullname).LastWriteTime) -gt [timespan]($refreshtime))) 
        {
            $isyprograms = (Invoke-IsyRestMethod -RestPath "/rest/programs?subfolders=true")

            $Script:list_of_isy_programs = $isyprograms.programs.program | Select-Object -Property id,localname,name
            write-verbose "`n`$script:list_of_isy_programs exists = $(Test-Path variable:script:list_of_isy_programs)`n"
            $idtoname = @{}
            $Script:list_of_isy_programs | ForEach-Object {$idtoname[$_.id] = $_.name}

            $isyprograms.programs.program | Where-Object {'parentid' -in ($_ | Get-Member -MemberType Properties).name} | 
                    ForEach-Object {$_.parentid = [string]($idtoname[$_.parentid])}

            $isyprograms.Save($fullname)
        }

    $isyprograms = New-Object -TypeName XML
    $isyprograms.Load((Join-Path $currentsettings.path $Script:ISYProgramsXMLFileName))

    $isyprograms.programs.program | Where-Object {'parentid' -in ($_ | Get-Member -MemberType Properties).name} |
                 Add-Member -MemberType AliasProperty -Name parent -Value parentid

}

if (!$refreshonly) 
    {
        if ($fast)
            { 
                if (!$fastsuccess)
                    {                      
                        # have to psobject.copy() each element or only the reference is passed and thus variable $Script:list_of_* can be changed by caller
                        $programs = $Script:list_of_isy_programs | ForEach-Object {$_.psobject.copy()}
                    } 
            }
        else 
            {
                $programs = $isyprograms.programs.program | ForEach-Object{$_.PSObject.Copy()}
            }
    }
else {return}


return $programs


}
#end function Get-ISYProgramList

<#
.Synopsis
   Get-ISYScene

   Gets the current status of all the devices in a an ISY scene or list of scenes. 
.DESCRIPTION
    Scenes themselves do not have a status (although this has been frequently debated) but the member devices do.
    This function gathers all the members from a list of scenes and initiates a REST status query for each of them. 
    This information is returned in an object from which the current status is taken and added to each scene object. 
    The status of the scene members (OnLevel% and OnLevel) is under the members.link property of the returned scene object(s).
    When the name parameter is used it is interpreted as an array of wildcard string(s) or regular expression(s) if 
    the -regx switch is present. Calling Get-ISYScene with no parameters gets all scenes (name = '*') except for the 
    predefined ones. (See -predefined switch below).

    The -dbmap switch is mutually exclusive to all other parameters and returns the cached (from memory) scene name to scene 
    address map. It is primarily used for tabexpansion.
    
    The -listdb switch returns a cached (from disk) full list of ISY scenes (although it also respects the -predefined switch 
    below). The status of each member device is not included.
    
    The -predefined switch causes the function to also return the predefined scenes (ADR scene and the scene that contains all 
    devices) for the -listdb and -name parameter sets. For the -name case this assumes the predefined names are included in the 
    wildcard or regx spec.  

    The -nocustomdefaultdisplayproperties switch will bypass the modification to the default display properties. By default the 
    default display properties are name,address,fullpath,membernames.

.EXAMPLE
    Get the status of the scene named cooktop and load the variable $status with this info. Because the status info is for 
    each of the devices this info is underneath the members property which contains a link object for each member. We see 3  
    scene members listed along with their status (2 are fully on, Remote1 button 2 has no status) as well as what role (Controller   
    or Responder) the device plays in the scene.

    > $status = Get-ISYScene -name cooktop
    > $status | ForEach-Object {$_.name; $_.members.link | ForEach-Object{$_} | Format-Table OnLevel%,OnLevel,MemberName,type -au}

    Cooktop

    membername         OnLevel% OnLevel type
    ----------         -------- ------- ----
    Remote1-2                           Controller
    CooktopButtonKTKP3 On       255     Controller
    CooktopDimmer      On       255     Controller

.EXAMPLE
    This example uses the Show-ISYObject function to simplify the syntax to get well formatted output.

    > Get-ISYScene -name Cooktop | Show-ISYObject

    ------------------------------------------------------------
    name        : Cooktop
    address     : 67679
    fullpath    : ISYRoot/Downstairs/Kitchen
    flag        : 132
    deviceGroup : 18
    ELK_ID      : B03

    MemberName         type       OnLevel% OnLevel
    ----------         ----       -------- -------
    Remote1-2          Controller
    CooktopButtonKTKP3 Controller Off      0
    CooktopDimmer      Controller Off      0

.EXAMPLE
    Get-ISYScene -dbmap | format-table -AutoSize

    name                   LocalName address
    ----                   --------- -------
    AllDevices             group     00:21:b9:00:f3:98
    MBRRalfReadingLight    group     73833
    MBRCeliaReadingLight   group     37733
    MBRDresserLights       group     47099
    KitchenTable           group     47750
    MBRCenterLight         group     57055
    MBROutsideLights       group     67764
    Cooktop                group     67679
    GarageCeilingLights    group     65477
    AutoDR                 group     ADR0001

    This shows what the -dbmap returns. Note the predfined scenes (in this case AllDevices and AutoDR) are always included as this 
    option does NOT respect the -predefined switch.

.EXAMPLE
    $scenelist = Get-ISYScene -listdb 
    PS C:\>$scenelist | format-table -AutoSize

    flag fullpath                         address name                   parent deviceGroup ELK_ID members
    ---- --------                         ------- ----                   ------ ----------- ------ -------
    132  ISYRoot/Downstairs/MasterBedroom 73833   MBRJimReadingLight     parent 21          C11    members
    132  ISYRoot/Downstairs/MasterBedroom 37733   MBRCeciliaReadingLight parent 22          C12    members
    132  ISYRoot/Downstairs/MasterBedroom 47099   MBRDresserLights       parent 23          C13    members
    132  ISYRoot/Downstairs/Kitchen       47750   KitchenTable           parent 17          A08    members
    132  ISYRoot/Downstairs/MasterBedroom 57055   MBRCenterLight         parent 19          C09    members
    132  ISYRoot/Outside                  67764   MBROutsideLights       parent 24          D01    members
    132  ISYRoot/Downstairs/Kitchen       67679   Cooktop                parent 18          B03    members
    132  ISYRoot/Garage                   65477   GarageCeilingLights    parent 20          D06    members

    PS C:\>$scenelist[3].members.link
    
    type       #text
    ----       -----
    Controller Remote1-1
    Controller KitchenTableDimmerKTKP1
    Responder  KitchenTableDimmer

    This example shows that the -listdb switch does not return the status of the scene members.

.Example
    (Get-ISYScene -name KitchenTable).members.link | ft -autosize

    OnLevel% type       SceneName    MemberName              OnLevel #text
    -------- ----       ---------    ----------              ------- -----
             Controller KitchenTable Remote1-1                       Remote1-1
    On       Controller KitchenTable KitchenTableDimmerKTKP1 255     KitchenTableDimmerKTKP1
    On       Responder  KitchenTable KitchenTableDimmer      255     KitchenTableDimmer

    This shows that using the name option does return the current status. Members KitchenTableDimmerKTKP1 and 
    KitchenTableDimmer are On and Remote1-1 has no status.
.EXAMPLE
    Get-ISYScene Cooktop

    address fullpath                   name    membernames
    ------- --------                   ----    -----------
    67679   ISYRoot/Downstairs/Kitchen Cooktop Remote1-2,CooktopButtonKTKP3,CooktopDimmer

    This shows what the default display properties are without any format-* or the show-isyobject cmdlets.
    Using the -nocustomdefaultdisplayproperties switch displays more properties in list format:

    PS C:\>Get-ISYScene Cooktop -nocustomdefaultdisplayproperties

    membernames : Remote1-2,CooktopButtonKTKP3,CooktopDimmer
    flag        : 132
    fullpath    : ISYRoot/Downstairs/Kitchen
    address     : 67679
    name        : Cooktop
    parent      : parent
    deviceGroup : 18
    ELK_ID      : B03
    members     : members

#>
function Get-ISYScene {
[CmdletBinding(DefaultParameterSetName='filter')]
[OutputType([System.XML.XMLElement[]])]
param(
# an array of wildcard spec(s) or regular expression(s) representing the scene name(s) to get
[Parameter(ParameterSetName='filter',Position=0,ValueFromPipelineByPropertyName,ValueFromPipeline)]
[object[]]$name = '*',

# specifies that the name string(s) are to be used as regular expression(s) not wildcards
[Parameter(ParameterSetName='filter')]
[switch]$regx,

# gets the scene name to address mapping from the cached data base - used mostly for tabexpansion
[Parameter(ParameterSetName='map')]
[switch]$dbmap,

# gets the cached list of scenes from the database - no direct query to the ISY is performed
[Parameter(ParameterSetName='list')]
[switch]$listdb,

# when present the predefined scenes (ADR scene and the scene that contains all devices) are included - these are excluded by default
[Parameter()]
[switch]$predefined,

# if set the default displayed properties are left unmodified.
[switch]$nocustomdefaultdisplayproperties
)

Begin
{
    Write-Verbose -Message ($MyInvocation.MyCommand.Name + ' called.') 
    
    $propsreturned = 'name','address','fullpath','flag','members','localname'
    $macregx = '^(([0-9a-f]{2}:)){5}[0-9a-f]{2}$'
    $Wildcardhash = @{$true = 'matching regular expressions '; $false = 'like wildcard(s) '}

    if ($dbmap)
    {
        return ( $Script:list_of_isy_scenes.foreach({$_.PSObject.Copy()}) ) # send a copy so caller cannot change value
    }

    $scenelist = (Get-ISYDB).nodes.group
    Write-Verbose -Message "There are $($scenelist.count) scenes in db"

    $scenelist | Add-Member -MemberType ScriptProperty -name membernames -Value {$this.members.link.'#text' -join ','}

    if (!$predefined)
    {
        $scenelist = $scenelist | where address -notmatch "^ADR0001$|$macregx"
    }

    if ($listdb)
    {
        if ($nocustomdefaultdisplayproperties)
        {
            return $scenelist 
        }
        else
        {
            return ($scenelist | Set-DefaultDisplayProperties -PropertyList name,address,fullpath,membernames)
        }
    }

    $sceneelems = @()

}

End
{
    if ($listdb -or $dbmap)
    {
        return
    }
    
    if ($PSCmdlet.MyInvocation.ExpectingInput)
    {
        $array = @($input)
        Write-Verbose -Message "Pipeline input found with $($array.count) scene members"
    }
    else
    {
        $array = $name
        Write-Verbose -Message "Command line input found with $($array.count) scene members"
    }

    if (($array | Get-Member -MemberType Properties).name -contains 'name')
    {
        $localname = ($array.localname | Sort-Object -Unique) 
        if ($localname -ne 'group')
        {
            Write-Error -Message "$localname object is not an ISY scene." -Category InvalidData
            return
        }
        $entries = $($array.name)
        Write-Verbose -Message "Objects have name property. $($entries.count) scene names found."
    }
    else
    {
        $entries = $array
        Write-Verbose -Message "Objects have NO name property. $($entries.count) scene names found."
    }

    foreach ($entry in $entries)
    {
        if (!$regx)
        {
            $sceneelems += @($scenelist | where name -like $entry)
        }
        else
        {
            try
            {
               [void]('ralf' -match $name)
                $sceneelems += @($scenelist | where name -match $entry)
            }
            catch
            {
                Write-Warning -Message "`'$entry`' is not a valid regular expression." 
                continue
            }
        }
    }
    if (!$sceneelems)
    {
        Write-Warning -Message "No scenes found $($Wildcardhash[$($regx.IsPresent)]) $($name -join ' or ')."
        return
    }

    $sceneelems = @($sceneelems | Sort-Object -Property name -Unique)

    $sceneelems.ForEach( { $scenename = $_.name 
                           $_.members.link.ForEach('setattribute','Scene_Name',$scenename) } )

    $memberslist = $sceneelems.ForEach({$_.membernames.split(',')}) | Sort-Object -Unique
    #$memberslist = $sceneelems.ForEach({$_.members.link.'#text'}) | Sort-Object -Unique

    $status = Get-InsteonDevice -name $memberslist 

    $sceneelems.members.link.ForEach({
                                    $current = $_
                                    $member = $status | where name -eq $current.'#text' # $current.membername works too
                                    $current.setattribute('OnLevel',$member.Onlevel)
                                    # cannot use setattribute with % in property name so use add-member.  
                                    # It will get dropped if saved as an xml doc but can be retrieved from OnLevel/255
                                    Add-Member -InputObject $current -NotePropertyName 'OnLevel%' -NotePropertyValue $member.'Onlevel%'
                                    $parentscene = $sceneelems | where name -eq $current.scene_name
                                    $propstoadd = ($parentscene | Get-Member -MemberType Properties | where definition -like string* | where name -ne name).name
                                    $propstoadd.foreach({$current.setattribute("Scene_$_",$parentscene.$_)})
                                    })

    #$sceneelems | where address -notmatch "^ADR0001$|$macregx" | 
     #                   Add-Member -MemberType ScriptProperty -name membernames -Value {$this.members.link.membername -join ','}

    
    if ($nocustomdefaultdisplayproperties)
    {
        return $sceneelems 
    }
    else
    {
        return ($sceneelems | Set-DefaultDisplayProperties -PropertyList name,address,fullpath,membernames)
    }

}

}
#end function Get-ISYScene

<#
.Synopsis
   Get-ISYSettings

   Retrieves $ISYSettings variable. 
.DESCRIPTION
   This function executes Test-ISYSettings and if sucessful returns the $ISYSettings variable.
    
.EXAMPLE
    Retrieve the $ISYSettings variable
    >Get-ISYSettings

    internalip  : 192.168.1.189
    externalip  : 172.19.13.67:650
    plmaddress  : 17 2 59
    refreshtime : 23:59:00
    path        : C:\isy
    tabexp      : True
    fileversion : 1.0.0.0
    UserName    : ralf
    Password    : System.Security.SecureString

#>
function Get-ISYSettings {
    [CmdletBinding()]
    [OutputType([pscredential])]
    param(
    [switch]$logintest,
    [switch]$notest
    )

$instruction = @"
With an appropiate xml file execute:
Import-ISYSettingsFromXMLFile -path <path to xml file> 
    or re-execute 
New-ISYSettings 
See the help for each command for more info.
"@



<#
    <#
    .Synopsis
       Test-ISYSettings

       This function tests the script $ISYSEttings variable. 
    .DESCRIPTION
       This function verifies the script $ISYsettings variable exists and is of the correct type.
       In addtion there is a check to see if the properties have been changed outside of New-ISYSettings.
       Note that powershell read-only variables only apply to scalars (this variable is of type pscredential).
       if declared as read-only it can still be changed.
       To work around this the  test-isysettings functioncompares the current running version of $isysettings 
       to the one reconstituted from the xml file that was created via the new-isysettings function.
     
       Returns $true if $ISYSettings is OK, $false if not.
    #>
        function Test-ISYSettings {
        [CmdletBinding()]
        [OutputType([boolean])]
        param([switch]$logintest)

        Write-Verbose -Message ('-'*80)
        Write-Verbose -Message ($MyInvocation.MyCommand.Name + ' called.')
        Write-Verbose -Message "logintest switch is set:  $($logintest.IsPresent)"

        $progress = @()

        if (!(Test-Path variable:script:isysettings))
        {
            Write-Error -Message "There is no `$ISYSettings var defined.`n$instruction" -Category InvalidResult
            return $false
        }
        $progress += "`$ISYSettings variable defined test: PASSED"

        if ($Script:ISYSettings.GetType().name -ne "PSCredential")
        {
            write-verbose ($progress -join "`n")
            Write-Error -Message "The `$ISYSettings variable is not of type PSCredential.`n$instruction" -Category InvalidType
            return $false
        } 
        $progress += "`$ISYSettings variable of type PSCredential test: PASSED"

        $currentfile = [System.IO.Path]::GetRandomFileName()
        $ISYSettings | Export-Clixml $currentfile
        $originalfile = join-path $Script:ISYSettings.path $ISYSEttingsXMLFileName

        try
        {
            # cannot use Import-ISYSettingsFromXMLFile because we need a 2nd opinion in the test function
            $varoriginal = Import-Clixml -path $originalfile
            $hashcurr = ($ISYSettings.GetNetworkCredential().password, (cat $currentfile | sls -not '<ss') -join '').GetHashCode()           
            $hashorig = ($varoriginal.GetNetworkCredential().password, (cat $originalfile | sls -not '<ss') -join '').GetHashCode()

            if ($hashorig -ne $hashcurr) 
            {
                write-verbose ($progress -join "`n")
                write-error "File $originalfile is likely bad. It seems to have been modified outside of the New-ISYSettings function.`n$instruction" -Category InvalidData
                return $false
            }
            $progress += "`$ISYSettings comparison between current and default file test: PASSED"

            if ($logintest)
            {
                write-verbose -Message "logintest being performed."
                $urlparms = @{
                intip = $Script:ISYSettings.internalip
                extip = $Script:ISYSettings.externalip
                }
                $url = (Select-IPToUse @urlparms) + '/rest/config'
                write-verbose -Message "URL = $url"

                $user = $Script:ISYSettings.UserName
                $pass = $Script:ISYSettings.GetNetworkCredential().password

                $headers = @{Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($("{0}:{1}" -f $user, $pass)))}

                $irmparms = @{Uri = $url
                              Method = 'Get'
                              Headers = $headers
                }
                $cfg = (Invoke-RestMethod @irmparms).configuration
                Write-Verbose -Message "config platform is $($cfg.platform)"

                if ($($cfg).platform -notlike '*isy*')
                {
                    write-verbose ($progress -join "`n")
                    write-error "Problem logging into ISY.`n$instruction" -Category ConnectionError
                }
                $progress += "Login to  ISY test: PASSED"
            }

        }
        catch
        {
            write-verbose ($progress -join "`n")
            return $false
        }

        finally
        {
           Remove-Item -Path $currentfile 
        }
        write-verbose ($progress -join "`n")
        return $true
        }
        #end function Test-ISYSettings


    if ($notest -or (Test-ISYSettings @PSBoundParameters))
    {
        $copy = $script:ISYSettings | ForEach-Object{$_.PSObject.Copy()}
        return ($copy)
    }

    throw "Unable to retreve ISYSettings."
}
#end function Get-ISYSettings

<#
.Synopsis
   Retrieves all of the current ISY subscriptions
.DESCRIPTION
   Each subscription has the the following properties:
       isExpired : yes or no
       isPortal : yes or no
       sid : string, can be used to unregister or re-subscribe.  Note expired subscriptions display as -1
       sock : socket number assigned by ISY
       isReusingSocket : yes or no
       isConnecting : yes or no
.EXAMPLE
    Get-ISYSubscriptions | format-table

    isExpired isPortal sid     sock isReusingSocket isConnecting
    --------- -------- ---     ---- --------------- ------------
    no        no       uuid:63 27   yes             no
    yes       no       -1      -1   no              no
    yes       no       -1      -1   no              no
    yes       no       -1      -1   no              no
    yes       no       -1      -1   no              no
    yes       no       -1      -1   no              no
    yes       no       -1      -1   no              no
    yes       no       -1      -1   no              no
    yes       no       -1      -1   no              no
    yes       no       -1      -1   no              no

    The admin console (sid uuid:63) subscribes to events to keep the GUI up to date in real time. 
#>
function Get-ISYSubscriptions
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([System.Xml.XmlElement[]])]
    Param
    ()

    $subscriptions = (Invoke-ISYRestMethod /rest/subscriptions).Subscriptions.Sub | Sort-Object -Property isexpired

    $subscriptions.where({$_.sid -ne -1}).foreach({$_.sid = ('uuid:' + $_.sid)})

    $subscriptions
}
# end of function Get-ISYSubscriptions

<#
.Synopsis
   Get-ISYTime

   Retrieves all the time info of the ISY and formats times to a readable format. 
.DESCRIPTION
   This function always gets the time info via a REST call to the ISY.
.EXAMPLE
    > Get-ISYTime
    
    NTP        : 03/11/2015 19:30:23
    TMZOffset  : -5 hours
    DST        : true
    Lat        : 35.897089
    Long       : 78.766519
    Sunrise    : 03/11/2015 07:33:05
    Sunset     : 03/11/2015 19:17:32
    IsMilitary : true
#>
function Get-ISYTime {
[CmdletBinding()]
[OutputType([System.XML.XMLElement])]
param()

# verify isy settings before ISY access
[void](Get-ISYSettings)

$isytime = Invoke-IsyRestMethod "/rest/time"

$item = Select-XML -Xml $isytime -XPath '//DT'
$item.node | 
ForEach-Object {
    # ISY only returns the integer part of the times - we add a dummy fractional part (multiply by 10000000)
    # Note: adding 1899 years because [datetime] was off due to epoch year of NTP formatted time being 2000. 
    # It seems as easy to use addyears(1899) as to add 599266080000000000 to the ISY (NTP time integer)*10000000.
    
    $plusfrac = $_.NTP + "0000000"
    $_.NTP = [string](([datetime]([long]$plusfrac)).AddYears(1899))

    $_.TMZOffset = [string]([int]($_.TMZOffset)/3600) + " hours"

    $plusfrac = $_.Sunrise + "0000000"
    $_.Sunrise = [string](([datetime]([long]$plusfrac)).AddYears(1899))

    $plusfrac = $_.Sunset + "0000000"
    $_.Sunset = [string](([datetime]([long]$plusfrac)).AddYears(1899))
}

return ($item.node) # | ForEach-Object{$_.PSObject.Copy()})

}
#end function Get-ISYTime

<#
.Synopsis
   Get-ISYUrl

   Returns the -Uri parameter to be used in the SOAP or REST call to the ISY. 
.DESCRIPTION
   This function takes the internal and external IPs and the restpath to build the Uri.
   If both internal and external IPs are -eq '' the function throws an exception.
   The internal IP is checked 1st and if it pings it is used.
   The external IP is assumed to be an https connection and is only used if the internal 
   IP is equal to '' or fails to ping.
   Note the IPs must include the port number if the standard ports (80/443) are not used.
   Also note when IPv6 is supported the address must be surrounded with [] even if the
   default ports are used.
#>
function Get-ISYUrl {
    [CmdletBinding()]
    [OutputType([string])]
    param(
    [parameter(mandatory=$true,Position=0)]
    [alias('restpath')]
    [string]$path)
    
    Write-Verbose -Message ($MyInvocation.MyCommand.Name + ' called.') 

    $currentsettings = (Get-ISYSettings)
    $intip = $currentsettings.internalip
    $extip = $currentsettings.externalip

    $isyendpoint = Select-IPToUse -intip $intip -extip $extip

    $url = "{0}{1}"-f $isyEndPoint, $path

    Write-Verbose "url = $url"
    $url
}
#end of function Get-ISYUrl

<#
.Synopsis
   Lists all the script scope ISY* variables in the ISY994i module.
.DESCRIPTION
   These are all the variables names and values defined at script scope beginning with names -like ISY*. 
   This command takes no parameters.
.EXAMPLE
   Get-ISYVars | Format-Table -Property Name,Value -AutoSize
    Name                   Value
    ----                   -----
    ISYDBXMLFileName       ISYDB.xml
    ISYDefaultsFilePath    C:\Users\ralf\Documents\WindowsPowerShell\Modules\ISY994i\ISYDefaults.xml
    ISYProgramsXMLFileName ISYPrograms.xml
    ISYSettingsFileVersion 1.0.0.0
    ISYSEttingsXMLFileName ISYSettings.xml

    Example shows the names and values of the ISY994i module's script scope variables.

#>
function Get-ISYVars {
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSVariable[]])]
    Param()

    Get-Variable -Name ISY* -Scope script 

}
#end of function Get-ISYVars

<#
.Synopsis
   Get-PLMAddress

   Gets the address of the PLM.
.DESCRIPTION
   Gets the Insteon address of the PLM.
.EXAMPLE
    No parameters and the function returns a string from the plmaddress field 
    ie $ISYSettings.plmaddress.path.

    > Get-PLMAddress
    1F 22 59
#>
function Get-PLMAddress {
[CmdletBinding()]
[OutputType([string])]
param()
# not sure how to query the ISY to get this - hardcoded for now in the settings var

Write-Verbose -Message ($MyInvocation.MyCommand.Name + ' called.') 

return ((Get-isysettings).plmaddress) # | ForEach-Object{$_.PSObject.Copy()})
}
#end function Get-PLMAddress

<#
.Synopsis
   Retrieves ISY settings from a settings file previously created.
.DESCRIPTION
   Following creation of an ISYSettings.xml file this function can read these into the $ISYSettings
   variable or return the $ISYSettings object to the caller via the -passthru param.

   Typically it would be used by a script as a 1st step to call any of the module commands without
   human intervention.

   The function takes a -path param which can be a directory containing the ISYSettings.xml file or the
   fully qualified path to an existing xml file. 
.EXAMPLE
   Import-ISYSettingsFromXMLFile -path c:/isy

   This is the typical use. There  must be a file called ISYSettings.xml in the directory supplied.
   The script itself shows no output but subsequent calls to Get-ISYsettings and Get-ISYTime show 
   ISYSettings are now in effect and working.  

   > Get-ISYSettings

    internalip  : 192.168.1.189
    externalip  : 11.12.13.14:9443
    plmaddress  : 1F 2 9F
    refreshtime : 23:59:00
    path        : C:\ISY\
    tabexp      : True
    fileversion : 1.0
    UserName    : admin
    Password    : System.Security.SecureString

   > Get-ISYTime

    NTP        : 10/03/2015 10:36:07
    TMZOffset  : -5 hours
    DST        : true
    Lat        : 35.783001
    Long       : 78.650002
    Sunrise    : 10/03/2015 07:10:08
    Sunset     : 10/03/2015 18:56:48
    IsMilitary : true

.EXAMPLE
    $isettings = Import-ISYSettingsFromXMLFile -path C:\isy\ISYSettings.xml -passthru

    This demonstrates the -passthru param. Instead of sending the data to teh $Script;ISYSettings
    the settings object is simply returned to the caller.

    > $isettings

    internalip  : 192.168.1.189
    externalip  : 11.12.13.14:9443
    plmaddress  : 1F 2 9F
    refreshtime : 23:59:00
    path        : C:\ISY\
    tabexp      : True
    fileversion : 1.0
    UserName    : admin
    Password    : System.Security.SecureString

    It also shows the path param being the fully qualified path to the settings xml file.
    This can be exploited to import a file not named ISYSettings.xml.

#>
function Import-ISYSettingsFromXMLFile {
[CmdletBinding()]
[OutputType([pscredential])]
param([Parameter(Mandatory=$true)]
      [String]
      $path,
      [switch]
      $passthru)
 
 $instruction = @"
With an appropiate xml file execute:
Import-ISYSettingsFromXMLFile -path <path to xml file> 
    or re-execute 
New-ISYSettings 
See the help for each command for more info.
"@
        
if (Test-Path $path)
{
    if (Test-Path $path -PathType Container)
    {
        $fqn = Join-Path $path $ISYSEttingsXMLFileName
        if (!(Test-Path -Path $fqn))
        {
            throw "No $ISYSEttingsXMLFileName found in $path. If name is different then pass in the fullname (FQN) of the file."
        }
        $pathtouse = Join-Path $path $ISYSEttingsXMLFileName 
    }
    else 
    {
        $pathtouse = $path
    }

}
else
{
    throw "No file or directory at $path" 
}


try
{
    $settings = Import-Clixml $pathtouse
}
catch
{
    if ($error[0].Exception -like '*key not valid*') 
    {
        throw "File $pathtouse is bad. It was likely not created by the current user.`n$instruction" 
    }
}


#$settings = Import-Clixml $pathtouse

if ($passthru)
{
    $settings
}

else
{
    $script:ISYSettings = $settings
}
Update-ISYDBXMLFile -refreshnow

}
#end of Import-ISYSettingsFromXMLFile

<#
.Synopsis
   Invoke-ISYProgram

   Executes one of (or part of one) the ISY programs. 
.DESCRIPTION
   This function accepts a program name and a command and executes all or part of that program.
   It can also stop,enable,disable,enableRunAtStartup,disableRunAtStartup a program.

   The -RC (return code) switch causes the response from the ISY to be sent back. Without this switch nothing is returned to the caller.

   Notes on errors:
     - If the program name is not valid you will get an error like:
     Invoke-ISYProgram : Cannot validate argument on parameter 'name'. The "$_ -in (Get-ISYProgramList -fast).name)" validation script for the argument with value "bogusprogram"

.EXAMPLE
     Run the then statements of the OutsideLightsOff program.

    > Invoke-ISYProgram -name OutsideLightsOff -command runThen
#>
function Invoke-ISYProgram {
[CmdletBinding()]
[OutputType([void])]
param(
[Parameter(mandatory=$true)]
[validatescript({$_ -in $($Script:list_of_isy_programs.name)})]
[string]$name,
[Parameter(mandatory=$true)]
[validateset("run","runThen","runElse","stop","enable","disable","enableRunAtStartup","disableRunAtStartup")]
[string]$command,
[switch]$RC)

Write-Verbose -Message ($MyInvocation.MyCommand.Name + ' called.') 

# verify isy settings before ISY access
[void](Get-ISYSettings)

$id = (Get-ISYProgramList | where name -eq $name).id

$IsyRestMethodparams = @{
RestPath  = "/rest/programs/$id/$command"
handlerc = !$RC}

Invoke-IsyRestMethod @IsyRestMethodparams

#Invoke-IsyRestMethod -RestPath "/rest/programs/$id/$command"

}
#end function Invoke-ISYProgram

<#
.Synopsis
   Invoke-ISYRestMethod

   Executes a GET method via the supplied restpath on the ISY REST interface. 
.DESCRIPTION
   This function can control Insteon devices, ISY scenes and programs. This is the 
   function that is called by all the other functions in this module that require interaction with the ISY.
   For more information on the ISY REST API see the ISY-WS-SDK-Manual.pdf document inside the docs folder in:
         http://www.universal-devices.com/developers/wsdk/4.2.18/ISY-WSDK-4.2.18.zip.
   This function courtesy of http://www.sytone.com/tag/powershell/

   When the -handleRC switch is present it causes the Invoke-ISYRestMethod function to handle the result code  
   returned from the ISY for the set actions. If not set, the RestResponse property of the RC is sent back to
   the caller.

.EXAMPLE
    Get all the nodes and their addresses in the ISY.
    (Invoke-IsyRestMethod "/rest/nodes").nodes.node | select name, address
.EXAMPLE
    Get information about a specific node.
    1st get the hallway light info and place this in the $hallwaylight variable.
    Then examine the info returned.

    $hallwaylight = Invoke-IsyRestMethod -RestPath "/rest/nodes/27 78 40 1"
    $hallwaylight.nodeInfo.node  | Format-Table -AutoSize

    flag address    name               type      enabled deviceClass wattage dcPeriod pnode      ELK_ID
    ---- -------    ----               ----      ------- ----------- ------- -------- -----      ------
    128  27 78 40 1 HallwayLightSwitch 1.32.65.0 true    0           0       0        27 78 40 1 B12
.EXAMPLE
    Turn the node on, in this case the kitchen table light.
    Invoke-IsyRestMethod -RestPath "/rest/nodes/F 5 D0 1/cmd/DON/"
.EXAMPLE
    Turn the node off, in this case the kitchen table light.
    Invoke-IsyRestMethod -RestPath "/rest/nodes/F 5 D0 1/cmd/DOF/"
#>
function Invoke-ISYRestMethod {
[CmdletBinding()]
[OutputType([void])]
param (
[Parameter(mandatory=$true)]
[string]$RestPath,
[switch]$HandleRC
)

Write-Verbose -Message ($MyInvocation.MyCommand.Name + ' called.') 

Write-Verbose "`nRESTPath = $restpath"

# verify isy settings before ISY access
$currentsettings = (Get-ISYSettings)

$url = Get-ISYUrl -restpath $RestPath

$headers = Get-Headers 
        
$response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers

if (!($response.RestResponse))
{
        return $response
    }

if (!$HandleRC)
{ 
        return $response.RestResponse
    }

# if you even get here...
$succeeded = $response.RestResponse.succeeded
$status = $response.RestResponse.status
Write-Verbose -message "Response from ISY is Succeeded: $succeeded, Status = $status"
if (!([bool]($succeeded)))
{
        Write-Error -Message "Bad response from ISY.`nResponse from ISY is Succeeded: $succeeded, Status = $status" -Category InvalidOperation
} 



}
#end function Invoke-ISYREstMethod

<#
.Synopsis
   Invokes one of the SOAP operations on the ISY.
.DESCRIPTION
   Currently all the ISY soap calls like Get* can be invoked using this script.

   If the soap call has no parameters the syntax is Invoke-ISYSOAPOperation -operation <soap call name>.
   Soap call name is one of GetDebugLevel,GetISYConfig,GetLastError,GetNodesConfig,GetProgramsDetail,
   GetSMTPConfig,GetStartupTime,GetSystemTime,GetSystemDateTime,GetSystemOptions or GetSystemStatus

   There are 5 Get operations that require parameters but GetVariable and GetVariables are both accessed by the GetVariables param:
   GetCurrentSystemStatus
    To call this use Invoke-ISYSOAPOperation -GetCurrentSystemStatus -SID <SID>

   GetSceneProfiles
    Invoke-ISYSOAPOperation -GetSceneProfiles -node <node> -controller <controller>

   GetVariables
    Invoke-ISYSOAPOperation -GetVariables -type <1|2> -id <id>
    If -id is not present then GetVariables is invoked, otherwise GetVariable is invoked.

   GetNodeProps
    Invoke-ISYSOAPOperation -GetNodeProps -address <node or group(scene) address>

   The -raw switch returns the full HtmlWebResponseObject object that was returned from the ISY.
   Normally what is returned is an XmlDocument object culled from the HtmlWebResponseObject representing the information associated with the particular soap call.

   Thanks to Bob Paauwe - his C# example event viewer code was what made this possible.
   http://www.bobshome.net/isy/ISYEventViewer.zip

.EXAMPLE
    Invoke-ISYSOAPOperation -operation GetISYConfig -raw

    StatusCode        : 200
    StatusDescription : OK
    Content           : <?xml version="1.0" encoding="UTF-8"?><s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope"><s:Body>
                            <configuration>
                                <deviceSpecs>
                                    <make>Universal Devices Inc.</make...
    RawContent        : HTTP/1.1 200 OK
                        Connection: Keep-Alive
                        EXT: UCoS, UPnP/1.0, UDI/1.0
                        Content-Length: 21030
                        Cache-Control: max-age=3600, must-revalidate
                        Content-Type: application/soap+xml; charset=UTF-8
                        Last-Modi...
    Forms             : {}
    Headers           : {[Connection, Keep-Alive], [EXT, UCoS, UPnP/1.0, UDI/1.0], [Content-Length, 21030], [Cache-Control, max-age=3600, must-revalidate]...}
    Images            : {}
    InputFields       : {}
    Links             : {}
    ParsedHtml        : mshtml.HTMLDocumentClass
    RawContentLength  : 21030

    This shows a call to get the ISY config using the -raw switch. The xml info describing the config is in the content property of
    the HtmlWebResponseObject returned.
.EXAMPLE
    $cooktopsceneprofile = Invoke-ISYSOAPOperation -GetSceneProfiles -node none -controller 42950 -Verbose
    VERBOSE: Body:
    <?xml version='1.0' encoding='utf-8'?><s:Envelope><s:Body><u:GetSceneProfiles
    xmlns:u='urn:udi-com:service:X_Insteon_Lighting_Service:1'><u:node>none</u:node><u:controller>42950</u:controller></u:GetSceneProfiles></s:Body></s:Envelope>

    VERBOSE: POST http://192.168.1.189/services with -1-byte payload
    VERBOSE: received 269-byte response of content type application/soap+xml; charset=UTF-8

    The GetSceneProfiles SOAP call is invoked getting the info for the Cooktop scene with address 42950 and the output
    is placed in the $cooktopsceneprofile variable. .
    Also the -verbose option caused various messages to display to the console. Normally this example would have no output to
    the console.

    Below we see what the particulars are for the cooktop scene (ie each members address, OL and RR) by examining the properties.
    PS C:\> $cooktopsceneprofile.Envelope.Body.SceneProfiles.sp

    node      OL  RR
    ----      --  --
    9 93 15 1 255 24
    9 75 87 1 188 24
.EXAMPLE
    $prgmdetail = Invoke-ISYSOAPOperation -operation GetProgramsDetail
    PS C:\>$prgmdetail.programdetails.triggers.d2d.trigger | format-table -autosize

    id name                                     parent if then else comment
    -- ----                                     ------ -- ---- ---- -------
    2  Query All                                1      if then      Factory Query Program
    1  MyPrograms                               0
    3  Kitchen Lights On                        1      if then
    4  Kitchen Lights Off                       1      if then
    5  Turn Off Outside Lights                  1      if then
    8  ControlDevicesSummer                     14     if then
    7  AllLLightsOff                            1      if then
    9  ControlDevicesWinter                     15     if then
    14 Summer                                   13     if
    10 OutsideLightsOff                         1         then
    13 TODU                                     1
    15 Winter                                   13     if
    6  MBRCenterLightOn                         1      if then
    17 Independence Day                         22     if
    16 Christmas Day                            22     if
    19 Memorial Day                             22     if
    18 Labor Day                                22     if
    21 Thanksgiving                             22     if
    20 New Years Day                            22     if
    22 TODUHolidays                             13
    23 Good Friday                              22     if
    24 NotAnyHoliday                            22     if
    25 MBRCenterLightOff                        1      if then
    26 FloodlightsOff                           1         then
    27 DriveWayAlertTurnOnGarageLight           28     if then
    28 NightTime                                1      if
    29 TurnOffDriveWayAlertActivatedGarageLight 28     if then
    30 HallwayLightOn                           1      if then
    31 HallwayLightOff                          1      if then

    Invoking the GetProgramsDetail SOAP call results in retrieval of all the program details.
    We see this under the $prgmdetail.programdetails.triggers.d2d.trigger properties. To
    see what each if, then or else contains we have to drill down more.

    PS C:\>$klo = $prgmdetail.programdetails.triggers.d2d.trigger | where name -eq 'Kitchen Lights On'
    PS C:\>$klo.then.device

    group control
    ----- -------
    60629 DON
    42950 DON

    We see that 2 scenes get turned ON in this program.

    For reference see latest WSDK (as of Sept 2016):
    http://www.universal-devices.com/developers/wsdk/4.4.6/ISY-WSDK-4.4.6.zip
    #>
function Invoke-ISYSOAPOperation {
[cmdletbinding()]
param(
# The operation (that requires no parameters) name
[parameter(ParameterSetName='noparms',Mandatory=$true,Position=0)]
[ValidateSet('GetDDNSHost','GetDebugLevel','GetISYConfig','GetLastError','GetNodesConfig','GetProgramsDetail',
             'GetSMTPConfig','GetStartupTime','GetSystemDateTime','GetSystemTime','GetSystemOptions','GetSystemStatus')]
[string]$operation,

# Returns the status of the ISY to the subscriber specified with $SID value.
[parameter(ParameterSetName='status',Mandatory=$true)]
[switch]$GetCurrentSystemStatus,
# The SID of the subscriber
[parameter(ParameterSetName='status',Mandatory=$true)]
[string]$SID,

# Gets scene ON Level (OL) and RampRate (RR) settings for devices in scenes.
[parameter(ParameterSetName='profile',Mandatory=$true)]
[switch]$GetSceneProfiles,
# 'none' or controllers address
[parameter(ParameterSetName='profile',Mandatory=$true)]
[string]$node,
# scene address
[parameter(ParameterSetName='profile',Mandatory=$true)]
[string]$controller,

# Gets all variables of a given type or one particular variable. Unfortunately the variable name is NOT returned. 
[parameter(ParameterSetName='variable',Mandatory=$true)]
[switch]$GetVariables,
# variable type 1 is a bool, type 2 an integer
[parameter(ParameterSetName='variable',Mandatory=$true)]
[ValidateSet('1','2')]
[string]$type,
# id of the variable
[parameter(ParameterSetName='variable')]
[string]$id,

# Returns the location, description and isLoad properties for a given address iff at least one of these properties is set for that address.
[parameter(ParameterSetName='props',Mandatory=$true)]
[switch]$GetNodeProps,
# The address of the node
[parameter(ParameterSetName='props',Mandatory=$true)]
[string]$address,

# returns the HtmlWebResponseObject object that was returned from the ISY instead of a xmldocument object
[switch]$raw
)

Write-Verbose -Message ($MyInvocation.MyCommand.Name + ' called.') 

$propercase = 'GetDDNSHost','GetDebugLevel','GetISYConfig','GetLastError','GetNodesConfig',
             'GetSMTPConfig','GetStartupTime','GetSystemOptions','GetSystemStatus'

switch ($PSCmdlet.ParameterSetName)
{
    'noparms'
    {
        $parms = ''
        if ($operation -eq 'GetProgramsDetail')
        {
            $cmd = 'GetAllD2D'
            Write-Verbose -Message "GetProgramsDetail aka GetAllD2D is not supported (more precisely 'Not implemented' in official documentation)."
            Write-Verbose -Message "GetAllD2D seems to 'work' as of build_timestamp 2015-10-13-11:35:11."
        }
        elseif ($operation -match 'GetSystemDateTime|GetSystemTime')
        {
            $cmd = 'GetSystemTime' # for some reason string in body is not equal to soap call name
        }
        else
        {
            $cmd = $propercase.where({$_ -match ('^' + $operation + '$')})
        }
    }

    'profile'
    {
        $cmd = 'GetSceneProfiles'
        $parms = "<u:node>$node</u:node><u:controller>$controller</u:controller>"
    }

    'status'
    {
        $cmd = 'GetCurrentSystemStatus'
        $parms = "<u:SID>$SID</u:SID>"
    }

    'variable'
    {
        if (!$id)
        {
            $cmd = 'GetVariables'
            $parms = "<u:type>$type</u:type>"
        }
        else
        {
            $cmd = 'GetVariable'
            $parms = "<u:type>$type</u:type><u:id>$id</u:id>"
        }
    }

    'props'
    {
        Write-Verbose -Message "GetNodeProps is deprecated."
        $cmd = 'GetNodeProps'
        $parms = "<id>$address</id>" # <u:id>$address</u:id> does not work?
    }

    Default {Write-Error -Message 'wtf?'}
}

$headers = Get-Headers -SOAPAction "urn:udi-com:device:X_Insteon_Lighting_Service:1#UDIService"
write-verbose -Message "auth = $($headers['Authorization'])"
Write-Verbose -Message "soapaction = $($headers['SOAPAction'])"

$body = "<?xml version='1.0' encoding='utf-8'?>"
$body += "<s:Envelope><s:Body><u:$cmd"
$body += " xmlns:u='urn:udi-com:service:X_Insteon_Lighting_Service:1'>"
$body += "$parms</u:$cmd></s:Body></s:Envelope>`r`n"

Write-Verbose -Message "Body:`n$body`n"

$uri = Get-ISYUrl -path '/services'
Write-Verbose -Message "the uri = $uri"

$splat = @{Uri = $uri
           DisableKeepAlive = $true
           Method = "POST"
           ContentType = 'text/xml; charset=utf-8'
           Headers = $headers
           Body = $body
            }

$response = Invoke-WebRequest @splat


if ($raw)
{
    return $response
}

if ($cmd -eq 'GetAllD2D')
{
    # content is malformed xml - no root placeholder
    [xml](($response.Content.replace('<?xml version="1.0"?>', '<?xml version="1.0"?><programdetails>')) + '</programdetails>')
}
else
{
    [xml]($response.Content)
}

}
#end of function Invoke-ISYSOAPOperation

<#
.Synopsis
   Creates an xml file containing all required info to access the ISY.
.DESCRIPTION
    There are 2 ways to enter this function, either via the command line or a form to be filled out.
    
    To use the form type 'New-ISYSettings -usegui'. This will then launch the Set-ISYDefaults function
    which presents a form with all the 6 params to be filled out. 

    To use the command line type the function name followed by all the params with the param names. Of course
    splatting can be used for all the paramters.

    The function then proceeds to prompt for credentials (username and password) of the ISY994i, wraps all this 
    information in a pscredential object and saves this in the directory specified in the -path parameter.    

    The 6 parameters (most of which are required to access the ISY) are:
    -internalip : internal IP (or valid DNS name) uses the http port (the port value can be set by appending :number).
    -externalip : external IP (or valid DNS name) uses the https port (the port value can be set by appending :number).
    -plmaddress : address of the PLM - used by scripts to resolve the PLM address
    -refreshtime : the time allowed before a refresh will be performed to the cached xml file holding the ISY node info      
    -path : path to where the xml files with cached info reside
    -tabexp : used to enable tab completion of names for the *Device, *Scene and Invoke-ISYProgram commands
    Note the parameters are not positional - the paramnames must be used unless the commandline defaults are used.

    Also note this function need only be called once as the info collected is saved in the ISYSettings.xml file.
    To access the ISY at a later time use Import-ISYSettingsFromXMLFile -path <path to ISYSettings.xml file>
    This will locate and load the saved ISYSettings.xml file and load it.

    Note the credentials stored can only be used by the windows user that created them (ie called New-ISYSettings).
    A different user would have to execute the command and choose a different directory to store the ISY xml files.

.EXAMPLE
    New-ISYSettings -usegui

    This will launch a form showing the 6 params that are required to be filled in. When the Submit button is pressed 
    a hashtable of these values is brought over and used to continue on, creating and storing the pscredential object.

.EXAMPLE
   New-ISYSettings -internalip 192.168.1.11 -externalip 47.47.47.47 -plmaddress 'A 55 F7' -refreshtime 08:00 -path $HOME -tabexp

   This example explicitly sets all 6 parameters using no default values. 
.EXAMPLE
   $paramhash = @{internalip='192.168.1.11';externalip='47.47.47.47';plmaddress='A 55 F7'
                refreshtime=[timespan]'08:00';path=$HOME;tabexp = $true}
   New-ISYSettings @paramhash

   This example shows the same values used but using splatting to pass the parameters to the function.
#>
function New-ISYSettings {
    [CmdletBinding(PositionalBinding=$false)]
    [OutputType([pscredential])]
    Param(

    # use the GUI to populate the default values for the 6 parameters below
    [Parameter(ParameterSetName='form',mandatory=$true)]
    [switch]
    $usegui,

    # this IP is usually behind a NAT and uses the http port
    [Parameter(ParameterSetName='cmdline')]
    [string]$internalip = "192.168.1.189",

    # the external IP uses the https port  
    [Parameter(ParameterSetName='cmdline')]
    [string]$externalip = "127.0.0.1:65092",

    # address of the PLM the ISY is connected to - only used to decode link tables by scripts 
    [Parameter(ParameterSetName='cmdline')]
    [string]$plmaddress = "14 22 54",

    # how long the system will wait between auto refreshes of the db
    [Parameter(ParameterSetName='cmdline')]
    [ValidateScript({$_ -ge [timespan]"00:30" -and $_ -le [timespan]"23:59"})]
    [timespan]$refreshtime = "23:59",
  
    # path used to store xml files used by the module
    [Parameter(ParameterSetName='cmdline')]
    [string]$path = "C:\ISY\",

    # whether tabexpansion for the applicable commands is on
    [Parameter(ParameterSetName='cmdline')]
    [switch]$tabexp
)

Write-Verbose -Message ($MyInvocation.MyCommand.Name + ' called.') 

if ($PSCmdlet.ParameterSetName -eq 'form')
{
    $defaults = Set-ISYDefaults -returnhash
    if ($defaults -eq $null) {write-verbose -message '$defaults is null.';return}

    $splathash = [ordered]@{'internalip' = $defaults['internalip']
                            'externalip' = $defaults['externalip']
                            'plmaddress' = $defaults['plmaddress']
                            'refreshtime' = $defaults['refreshtime']
                            'path'= $defaults['path']
                            'tabexp' = $defaults['tabexp']
                            'fileversion' = $ISYSettingsFileVersion}
}
else
{
    $splathash = [ordered]@{'internalip' = $internalip
                            'externalip' = $externalip
                            'plmaddress' = $plmaddress
                            'refreshtime' = [timespan]$refreshtime
                            'path'= Resolve-Path $path
                            'tabexp' = $tabexp.IsPresent
                            'fileversion' = $ISYSettingsFileVersion}
}

$credential = Get-Credential -Message "Input the username and password to be used to access the ISY." 
if (!$credential) {Write-Verbose -Message "Aborted. No action taken.";return}

$credential | Add-Member $splathash

$credential | Export-Clixml (Join-Path $path $ISYSEttingsXMLFileName)

$Script:ISYSettings = $credential

# Test/get the settings
Get-ISYSettings -logintest 

# TODO prompt for acceptance - maybe via GUI

# connect to the ISY and update db
Update-ISYDBXMLFile -refreshnow

}
#end function New-ISYSettings

<#
.Synopsis
   Registers to receive ISY events
.DESCRIPTION
   Once Register-ForISYEvent is called ISY event strings (xml format) are output to the pipeline.

   The -SID parameter can be used to re-subscribe to a previous event subscription. If left blank the ISY assigns a SID which is 
   included in all event strings. 

   By default a time stamp (format 'yyyyMMddTHHmmssffff' where HH is 24 hour and ffff is milliseconds) of when the event string was 
   output is added to the xml string.
   If this time stamp is not desired use the -raw switch parameter.
.EXAMPLE
    Register-ForISYEvent
    <?xml version="1.0"?><Event TimeStamp="20160927T1853037645" seqnum="0" sid="uuid:89"><control>_4</control><action>5</action><node></node><eventInfo><status>0</status></eventInfo></Event>
    <?xml version="1.0"?><Event TimeStamp="20160927T1853037715" seqnum="1" sid="uuid:89"><control>_4</control><action>6</action><node></node><eventInfo><status>1</status></eventInfo></Event>
    <?xml version="1.0"?><Event TimeStamp="20160927T1853037725" seqnum="2" sid="uuid:89"><control>_1</control><action>8</action><node></node><eventInfo>5C19A6.7D55E9</eventInfo></Event>
    <?xml version="1.0"?><Event TimeStamp="20160927T1853037840" seqnum="3" sid="uuid:89"><control>_0</control><action>120</action><node></node><eventInfo></eventInfo></Event>
        :
        :
    <?xml version="1.0"?><Event TimeStamp="20160927T1858155904" seqnum="110" sid="uuid:89"><control>_0</control><action>120</action><node></node><eventInfo></eventInfo></Event>
    <?xml version="1.0"?><Event TimeStamp="20160927T1858161577" seqnum="111" sid="uuid:89"><control>_5</control><action>1</action><node></node><eventInfo></eventInfo></Event>
    <?xml version="1.0"?><Event TimeStamp="20160927T1858161612" seqnum="112" sid="uuid:89"><control>_5</control><action>0</action><node></node><eventInfo></eventInfo></Event>
    <?xml version="1.0"?><Event TimeStamp="20160927T1858165306" seqnum="113" sid="uuid:89"><control>ST</control><action>255</action><node>16 96 98 1</node><eventInfo></eventInfo></Event>
    <?xml version="1.0"?><Event TimeStamp="20160927T1858165336" seqnum="114" sid="uuid:89"><control>_1</control><action>3</action><node></node><eventInfo>[  16 96 98 1]       ST 255</eventInfo></Event>
    <?xml version="1.0"?><Event TimeStamp="20160927T1858239497" seqnum="115" sid="uuid:89"><control>_5</control><action>1</action><node></node><eventInfo></eventInfo></Event>
    <?xml version="1.0"?><Event TimeStamp="20160927T1858239532" seqnum="116" sid="uuid:89"><control>_5</control><action>0</action><node></node><eventInfo></eventInfo></Event>
    <?xml version="1.0"?><Event TimeStamp="20160927T1858243227" seqnum="117" sid="uuid:89"><control>ST</control><action>0</action><node>16 96 98 1</node><eventInfo></eventInfo></Event>
    <?xml version="1.0"?><Event TimeStamp="20160927T1858243272" seqnum="118" sid="uuid:89"><control>_1</control><action>3</action><node></node><eventInfo>[  16 96 98 1]       ST   0</eventInfo></Event>
    <?xml version="1.0"?><Event TimeStamp="20160927T1858405882" seqnum="119" sid="uuid:89"><control>_0</control><action>120</action><node></node><eventInfo></eventInfo></Event>
    
    When Register-ForISYEvent is called the ISY sends event strings to the console. It will continue to do this until the 
    SID (uuid:89) is unregistered. This can be accomplished from another powershell console by executing 
    Unregister-FromISYEvent -SID "uuid:89". See help for Unregister-FromISYEvent for more info.
.EXAMPLE
   Register-ForISYEvent | Tee-Object -FilePath isyevents.txt | Convert-ISYEventToXmlElem | Convert-ISYEventToFriendlyName

   This example shows sending the event strings to a file as well as send them down the pipeline to Convert-ISYEventToXmlElem 
   and from that onto Convert-ISYEventToFriendlyNames, See the help for Convert-ISYEventToXmlElem and Convert-ISYEventToFriendlyName
   for more info.
#>
function Register-ForISYEvent {
[cmdletbinding()]
[OutputType([string[]])]
param(
# re-subscribe to this existing SID 
[string]$SID = '',

# if present then do NOT add time stamp
[switch]$raw
)

if ($SID -ne '') 
{
    if ($SID -notmatch '^uuid:')
    {
        $SID = 'uuid:' + $SID
    }
    $SIDXml = "<SID>$SID</SID>"
}

# username and password in base64 string 
$auth = (Get-Headers).Authorization
write-verbose -Message "auth = $auth)"

# create the body of soap request
$body = "<?xml version='1.0' encoding='utf-8'?>"
$body += "<s:Envelope><s:Body><u:Subscribe"
$body += " xmlns:u='urn:udi-com:service:X_Insteon_Lighting_Service:1'>"
$body += '<reportURL>REUSE_SOCKET</reportURL>'
$body += "<duration>infinite</duration>$SIDXml</u:Subscribe></s:Body></s:Envelope>`r`n" 

Write-Verbose -Message "Body:`n$body`n"

$uri = Get-ISYUrl -path '/dummy'
$ip_address = [regex]::Match($uri, 'http://(.*)/dummy').groups[1].value
Write-Verbose -Message "ip_address = $ip_address"

# create TCP socket to use for communication to ISY
$sock = new-object -typename System.Net.Sockets.TcpClient

# connect and send header and body 
$sock.Connect($ip_address, 80)

$writer = [System.IO.StreamWriter]::new($sock.GetStream())

$writer.WriteLine("POST /services HTTP/1.1")
$writer.WriteLine("ContentType: text/xml; charset=utf-8")
$writer.WriteLine("Authorization:$auth")
$writer.WriteLine("Content-Length:$($body.Length.ToString())")
$writer.WriteLine("SOAPAction: urn:udi-com:device:X_Insteon_Lighting_Service:1#Subscribe")
$writer.WriteLine()
$writer.Write($body)
$writer.Flush()

$sock.ReceiveTimeout = 120000

# create reader to retrive event strings
$reader = [System.IO.StreamReader]::new($sock.GetStream())

$gotlength = $false

while ($true) {

	$s = $reader.ReadLine()

    # If $s is null assume an unregister request came in
    if ($s -eq $null)
    {
        Write-Verbose -Message "Unregistered!"
        return
    }

    write-verbose -Message "Readline = $s"

    switch -Regex ($s)
    {   
        # get the length of the message
        'CONTENT-LENGTH'
        {
            $linelength = [int]::Parse($s.Substring(15))
            $gotlength = $true
        }
        
        # when $s = '' it is time to use readblock to retrieve event char array
        '^$'
        {
            if ($gotlength)
            {
                # create buffer of adequate size
                $buf = new-object System.Char[] ($linelength +2)

		        # read in chars and convert to string
                [void]$reader.ReadBlock($buf, 0, $linelength) 
                Write-Verbose -Message "line length = $linelength"

		        $str = $buf -join ''
                Write-Verbose -Message "Readblock = $str"

                # this line returns the SID assigned by the ISY
                if ($str.Contains('<SubscriptionResponse>'))
                {
                    # TODO Returned SID captured but not returned - return via [ref]?
                    $ReturnedSID = ([regex]::Match($str, '\<SID\>(uuid:\d{1,})\</SID\>')).groups[1].value
                    break
                }

                # add time stamp by default
                if (!$raw)
                {
                    # emit the event record with time stamp for the xml lines - which should be all lines
                    # the FileDateTime format is why this function requires version 5.0. If some other choice of format is used version 3.0 should work 
                    $output = $str.replace('<Event seqnum',('<Event TimeStamp="' + (get-date -Format FileDateTime) + '" seqnum'))
                    # Note: to convert FileDateTime string to datetime object: [datetime]::ParseExact("20160913T1833374716",'yyyyMMddTHHmmssffff', $null)
                }
                else
                {
                    $output = $str
                }
		
                # noticed lines ending in 00 using format-hex. For testing using clipboard (which worked) these were not present.
                # w/o this downstream Convert-ISYEventToXmlElem had trouble parsing text from a saved text file into [xml]
                if ($output -ne $null)
                {
                    $output.trim([char]0)
                }
                $gotlength = $false   
            }
        }

        Default {}
    }

}

}
# end of function Register-ForISYEvent

<#
.Synopsis
   Resolve-InsteonType

   This function decodes the type field in the object returned by Get-InsteonDevice to a readable format.
.DESCRIPTION
   To decode this information the latest 1_fam.xml file from Universal Devices should be placed in the $PSScriptRoot directory.
   This is where the module file (ISY994i.psm1) has been installed. (Normally c:\Users\ralf\WindowsPowershell\Modules\ISY9941)
   Note the function already has embedded in it the content of the 1_fam.xml from ISY-WSDK-4.4.6. The code will 1st use 
   the 1_fam.xml file and use the embedded info iff it cannot find the file.
   The function returns a psobject with 2 properties:
   category: this is the insteon device category such as DimmerDevs or SwitchedRelayDevs
   Type: This is a friendlyname describing the device along with the version of the device
.EXAMPLE
   Find what device 1.32.64.0 represents.

    > Resolve-InsteonType 1.32.64.0
    category     type
    --------     ----
    DimmableDevs SWITCHLINC_DIMMER_2477D v.40
.EXAMPLE
   Note this function is used so that the type field in the devicelist is in human readable in the ISYDB.xml file.
   So to get a list of all the unique Insteon device types present in your system, execute the following:
    
    > (Get-InsteonDevice -listdb).type | sort -Unique

    ICON_REMOTELINC_2843 v.0
    ICON_REMOTELINC_2843 v.7c
    ICON_SWITCH_DIMMER_2876DB v.39
    INLINELINC_DIMMER_2475DA1 v.41
    IO_LINC_2450 v.36
    KEYPADLINC_DIMMER_2486D v.29
    SWITCHLINC_DIMMER_2477D v.40
    SWITCHLINC_DIMMER_2477D v.41
    SWITCHLINC_DIMMER_W_SENSE_2476D v.38
    SWITCHLINC_V2_DIMMER_2476D v.27

    For reference see the latest ISY SDK. At the time of writing this was:
    http://www.universal-devices.com/developers/wsdk/4.4.6/ISY-WSDK-4.4.6.zip
#>
function Resolve-InsteonType {
[CmdletBinding()]
[OutputType([string])]
param(
[Parameter(mandatory=$true)]
[string]$type)

# The below $fam_1_content here-string has the content of 1_fam.xml from http://www.universal-devices.com/developers/wsdk/4.4.6/ISY-WSDK-4.4.6.zip.
$fam_1_content = @"
<?xml version="1.0" encoding="UTF-8"?>
<NodeFamily id="1" name="INSTEON">
	<description>
		Family of INSTEON Devices
	</description>
	
	<nodeCategory id="0" name="Controllers" >
		<nodeSubCategory id="0" name="DEV_SCAT_CONTROLINC_2430"/>
		<nodeSubCategory id="5" name="DEV_SCAT_ICON_REMOTELINC_2843"/>
		<nodeSubCategory id="6" name="DEV_SCAT_ICON_TABLETOP_2830"/>
		<nodeSubCategory id="9" name="DEV_SCAT_SIGNALINC_2442"/>
		<nodeSubCategory id="10" name="DEV_SCAT_POOLUX_LCD_CONTROLLER"/>
		<nodeSubCategory id="11" name="DEV_SCAT_ACCESSPOINT"/>
		<nodeSubCategory id="12" name="DEV_SCAT_IES_COLOR_TOUCHSCREEN"/>
		<nodeSubCategory id="16" name="DEV_SCAT_REMOTE_LINC_2_KEYPAD_4"/>
		<nodeSubCategory id="17" name="DEV_SCAT_REMOTE_LINC_2_SWITCH"/>
		<nodeSubCategory id="18" name="DEV_SCAT_REMOTE_LINC_2_KEYPAD_8"/>
	</nodeCategory>
	
	<nodeCategory id="1" name="Dimmable Devices">
		<nodeSubCategory id="0" name="DEV_SCAT_LAMPLINC_V2_2456D3"/>
		<nodeSubCategory id="1" name="DEV_SCAT_SWITCHLINC_V2_DIMMER_2476D"/>
		<nodeSubCategory id="2" name="DEV_SCAT_INLINE_DIMMABLE"/>
		<nodeSubCategory id="3" name="DEV_SCAT_ICON_SWITCH_DIMMER_2876D3"/>
		<nodeSubCategory id="4" name="DEV_SCAT_SWITCHLINK_V2_DIMMER_2476DH"/>
		<nodeSubCategory id="5" name="DEV_SCAT_KEYPADLINC_TIMER_2484DWH8"/>
		<nodeSubCategory id="6" name="DEV_SCAT_LAMPLINC_2_PIN"/>
		<nodeSubCategory id="7" name="DEV_SCAT_ICON_LAMPLINC_V2_2_PIN_2456D2"/>
		<nodeSubCategory id="9" name="DEV_SCAT_KEYPADLINC_DIMMER_2486D"/>
		<nodeSubCategory id="10" name="DEV_SCAT_ICON_INWALL_CONTROLLER_2886D"/>
		<nodeSubCategory id="11" name="DEV_SCAT_LAMPLINC_BI_PHY"/>
		<nodeSubCategory id="12" name="DEV_SCAT_KEYPADLINC_DIMMER_2486DWH8"/>
		<nodeSubCategory id="13" name="DEV_SCAT_SOCKETLINC_2454D"/>
		<nodeSubCategory id="14" name="DEV_CAT_BIPHY_LAMPLINC_B2457D2"/>
		<nodeSubCategory id="19" name="DEV_SCAT_ICON_SWITCHLINC_DIMMER_BELL_CANADA"/>
		<nodeSubCategory id="23" name="DEV_SCAT_TOGGLELINC_DIMMER_2466D"/>
		<nodeSubCategory id="24" name="DEV_SCAT_COMPANION_SWITCH_2474D"/>
		<nodeSubCategory id="25" name="DEV_SCAT_SWITCHLINC_DIMMER_W_SENSE_2476D"/>
		<nodeSubCategory id="26" name="DEV_SCAT_INLINELINC_DIMMER_2475D"/>
		<nodeSubCategory id="27" name="DEV_SCAT_KEYPAD_LINC_DIMMER_2486D_6"/>
		<nodeSubCategory id="28" name="DEV_SCAT_KEYPAD_LINC_DIMMER_2486D_8"/>
		<nodeSubCategory id="29" name="DEV_SCAT_SWITCH_LINC_DIMMER_2476DH"/>
		<nodeSubCategory id="30" name="DEV_SCAT_ICON_SWITCH_DIMMER_2876DB"/>
		<nodeSubCategory id="31" name="DEV_SCAT_TOGGLELINC_DIMMER_2466D_2"/>
		<nodeSubCategory id="32" name="DEV_SCAT_SWITCHLINC_DIMMER_2477D"/>
		<nodeSubCategory id="33" name="DEV_SCAT_OUTLETLINC_DIMMER_2472D_DUAL_BAND"/>
		<nodeSubCategory id="34" name="DEV_SCAT_LAMPLINC_2_PIN_DIMMER_2457D2X"/>
		<nodeSubCategory id="36" name="DEV_SCAT_SWITCHLINC_DIMMER_2WIRE_2474DWH"/>
		<nodeSubCategory id="37" name="DEV_SCAT_BALLAST_DIMMER_2475DA2"/>
		<nodeSubCategory id="39" name="DEV_SCAT_SWITCHLINC_DIMMER_2477D_SP"/>
		<nodeSubCategory id="41" name="DEV_SCAT_KEYPAD_LINC_DIMMER_2486D_8_SP"/>
		<nodeSubCategory id="42" name="DEV_SCAT_LAMPLINC_2_PIN_DIMMER_2457D2X_SP"/>
		<nodeSubCategory id="43" name="DEV_SCAT_SWITCHLINC_DIMMER_2477DH_SP"/>
		<nodeSubCategory id="44" name="DEV_SCAT_INLINELINC_DIMMER_2475D_SP"/>
		<nodeSubCategory id="45" name="DEV_SCAT_SWITCHLINC_DIMMER_2477DH"/>
		<nodeSubCategory id="46" name="DEV_SCAT_FANLINC_2475F"/>
		<nodeSubCategory id="48" name="DEV_SCAT_SWITCHLINC_DIMMER_2476D"/>
		<nodeSubCategory id="49" name="DEV_SCAT_SWITCHLINC_DIMMER_2478D"/>
		<nodeSubCategory id="50" name="DEV_SCAT_INLINELINC_DIMMER_2475DA1"/>
		<nodeSubCategory id="58" name="DEV_SCAT_LED_BULB_8WATT_2672_222"/>
		<nodeSubCategory id="73" name="DEV_SCAT_LED_BULB_12WATT_PAR38_2674_222"/>
		<!-- Global Line -->
		<nodeSubCategory id="52" name="DEV_SCAT_DIN_RAIL_DIMMER_2452_222"/>
		<nodeSubCategory id="54" name="DEV_SCAT_DIN_RAIL_DIMMER_2452_422"/>
		<nodeSubCategory id="55" name="DEV_SCAT_DIN_RAIL_DIMMER_2452_522"/>
		
		<nodeSubCategory id="53" name="DEV_SCAT_MICRO_MODULE_DIMMER_2442_222"/>
		<nodeSubCategory id="56" name="DEV_SCAT_MICRO_MODULE_DIMMER_2442_422"/>
		<nodeSubCategory id="57" name="DEV_SCAT_MICRO_MODULE_DIMMER_2442_522"/>
		
		<nodeSubCategory id="11" name="DEV_SCAT_PLUGIN_DIMMER_2632_422"/>
		<nodeSubCategory id="15" name="DEV_SCAT_PLUGIN_DIMMER_2632_432"/>
		<nodeSubCategory id="17" name="DEV_SCAT_PLUGIN_DIMMER_2632_442"/>
		<nodeSubCategory id="18" name="DEV_SCAT_PLUGIN_DIMMER_2632_522"/>
		
		<!--  U.S. -->
		<nodeSubCategory id="65" name="DEV_SCAT_KEYPAD_LINC_DIMMER_2334_2_8_BUTTON"/>
		<nodeSubCategory id="66" name="DEV_SCAT_KEYPAD_LINC_DIMMER_2334_2_5_BUTTON"/>
	</nodeCategory>
	
	<nodeCategory id="2" name="Switched/Relay Devices">
		<nodeSubCategory id="5" name="DEV_SCAT_KEYPADLINC_RELAY_2486SWH8"/>
		<nodeSubCategory id="6" name="DEV_SCAT_APPLIANCELINC_OUTDOOR_2456S3E"/>
		<nodeSubCategory id="7" name="DEV_SCAT_TIMERLINC_2456S3T"/>
		<nodeSubCategory id="8" name="DEV_SCAT_OUTLETLINC_2473"/>
		<nodeSubCategory id="9" name="DEV_SCAT_APPLIANCELINC_2456S3"/>
		<nodeSubCategory id="10" name="DEV_SCAT_SWITCHLINC_RELAY_2476S"/>
		<nodeSubCategory id="11" name="DEV_SCAT_ICON_ON_OFF_SWITCH_2876S"/>
		<nodeSubCategory id="12" name="DEV_SCAT_ICON_APPLIANCE_ADAPTER_2856S3"/>
		<nodeSubCategory id="13" name="DEV_SCAT_TOGGLELINC_RELAY_2466S"/>
		<nodeSubCategory id="14" name="DEV_SCAT_SWITCHLINC_RELAY_2476S_2"/>
		<nodeSubCategory id="15" name="DEV_SCAT_KEYPADLINC_RELAY_2486S"/>
		<nodeSubCategory id="16" name="DEV_SCAT_INLINE_RELAY"/>
		<nodeSubCategory id="17" name="DEV_SCAT_EZSWITCH_30"/>
		<nodeSubCategory id="18" name="DEV_SCAT_COMPANION_SWITCH_2474S"/>
		<nodeSubCategory id="19" name="DEV_SCAT_ICON_SWTICHLINC_RELAY_BELL_CANADA"/>
		<nodeSubCategory id="20" name="DEV_SCAT_INLINE_RELAY_WITH_SENSE"/>
		<nodeSubCategory id="21" name="DEV_SCAT_SWITCHLINC_RELAY_W_SENSE_2476S"/>
		<nodeSubCategory id="22" name="DEV_SCAT_ICON_RELAY_2876SB"/>
		<nodeSubCategory id="23" name="DEV_SCAT_ICON_APPLIANCELINC_2856S3B"/>
		<nodeSubCategory id="24" name="DEV_SCAT_SWITCHLINC_RELAY_2494S220"/>
		<nodeSubCategory id="25" name="DEV_SCAT_SWITCHLINC_RELAY_2494S220_B"/>
		<nodeSubCategory id="26" name="DEV_SCAT_TOGGLELINC_RELAY_2466S_2"/>
		<nodeSubCategory id="28" name="DEV_SCAT_SWITCHLINC_RELAY_REMOTE_CONTROL_2476S"/>
		<nodeSubCategory id="30" name="DEV_SCAT_KEYPADLINC_RELAY_2487S"/>
		<nodeSubCategory id="31" name="DEV_SCAT_INLINELINC_RELAY_DUALBAND_2475SDB"/>
		<nodeSubCategory id="37" name="DEV_SCAT_KEYPADLINC_TIMER_RELAY_2484SWH8"/>
		<nodeSubCategory id="41" name="DEV_SCAT_SWITCHLINC_RELAY_COUNTDOWN_TIMER_2476ST"/>
		<nodeSubCategory id="42" name="DEV_SCAT_SWITCHLINC_RELAY_DUAL_BAND_2477S"/>
		<nodeSubCategory id="44" name="DEV_SCAT_KEYPADLINC_DUAL_BAND_RELAY_2487S"/>
		<nodeSubCategory id="55" name="DEV_SCAT_ON_OFF_MODULE_2635_222"/>
		<nodeSubCategory id="57" name="DEV_SCAT_ON_OFF_2663_222"/>
		<nodeSubCategory id="59" name="DEV_SCAT_ON_OFF_OUTDOOR_MODULE_2634_222"/>
		
		<!-- Pro Series -->
		<nodeSubCategory id="32" name="DEV_SCAT_KEYPADLINC_RELAY_2486S_SP"/>
		<nodeSubCategory id="33" name="DEV_SCAT_OUTLETLINC_2473_SP"/>
		<nodeSubCategory id="34" name="DEV_SCAT_INLINE_RELAY_SP"/>
		<nodeSubCategory id="35" name="DEV_SCAT_SWITCHLINC_RELAY_2476S_SP"/>
		
		<!-- Global Line -->
		<nodeSubCategory id="46" name="DEV_SCAT_DIN_RAIL_RELAY_2453_222"/>
		<nodeSubCategory id="51" name="DEV_SCAT_DIN_RAIL_RELAY_2453_422"/>
		<nodeSubCategory id="52" name="DEV_SCAT_DIN_RAIL_RELAY_2453_522"/>
		
		<nodeSubCategory id="47" name="DEV_SCAT_MICRO_MODULE_RELAY_2443_222"/>
		<nodeSubCategory id="49" name="DEV_SCAT_MICRO_MODULE_RELAY_2443_422"/>
		<nodeSubCategory id="50" name="DEV_SCAT_MICRO_MODULE_RELAY_2443_522"/>
		
		<nodeSubCategory id="45" name="DEV_SCAT_PLUGIN_RELAY_2633_422"/>
		<nodeSubCategory id="48" name="DEV_SCAT_PLUGIN_RELAY_2633_432"/>
		<nodeSubCategory id="53" name="DEV_SCAT_PLUGIN_RELAY_2633_442"/>
		<nodeSubCategory id="54" name="DEV_SCAT_PLUGIN_RELAY_2633_522"/>
		
		<!-- End: Global Line -->
		<nodeSubCategory id="56" name="DEV_SCAT_ON_OFF_OUTDOOR_MODULE_2634_222"/>
		
	</nodeCategory>
	
	<nodeCategory id="3" name="Network Bridge">
		<nodeSubCategory id="1" name="DEV_SCAT_POWERLINC_SERIAL_2414S"/>
		<nodeSubCategory id="2" name="DEV_SCAT_POWERLINC_USB_2414U"/>
		<nodeSubCategory id="3" name="DEV_SCAT_ICON_POWERLINC_SERIAL_2814S"/>
		<nodeSubCategory id="4" name="DEV_SCAT_ICON_POWERLINC_USB_2814U"/>
		<nodeSubCategory id="5" name="DEV_SCAT_POWERLINE_MODEM"/>
		<nodeSubCategory id="6" name="DEV_SCAT_IRLINC"/>
		<nodeSubCategory id="7" name="DEV_SCAT_IRLINC_TX"/>
		<nodeSubCategory id="11" name="DEV_SCAT_POWERLINC_MODEM_USB"/>
		<nodeSubCategory id="13" name="DEV_SCAT_EZX10RF"/>
		<nodeSubCategory id="15" name="DEV_SCAT_EZX10IR"/>
	</nodeCategory>
	
	<nodeCategory id="4" name="Irrigation">
		<nodeSubCategory id="0" name="DEV_SCAT_COMPACTA_EZFLORA_SPRINKLER_CONTROLLER"/>
	</nodeCategory>
	
	<nodeCategory id="5" name="Climate">
		<nodeSubCategory id="0" name="DEV_SCAT_BROAN_SMSC080_EXHAUST_FAN"/>
		<nodeSubCategory id="1" name="DEV_SCAT_COMPACTA_EZTHERM"/>
		<nodeSubCategory id="2" name="DEV_SCAT_BROAN_SMSC110_EXHAUST_FAN"/>
		<nodeSubCategory id="3" name="DEV_SCAT_INSTEON_THERMOSTAT_ADAPTER"/>
		<nodeSubCategory id="4" name="DEV_SCAT_COMPACTA_EZTHERMX"/>
		<nodeSubCategory id="5" name="DEV_SCAT_BROAN_VENMAR_BEST"/>
		<nodeSubCategory id="9" name="DEV_SCAT_INSTEON_THERMOSTAT_ADAPTER_SP"/>
		<nodeSubCategory id="10" name="DEV_SCAT_INSTEON_THERMOSTAT_WIRELESS_2441ZTH"/>
		<nodeSubCategory id="11" name="DEV_SCAT_INSTEON_THERMOSTAT_TEMPLINC_2441TH"/>
		<nodeSubCategory id="14" name="DEV_SCAT_INSTEON_THERMOSTAT_ADAPTER_2491T"/>
		<nodeSubCategory id="19" name="DEV_SCAT_INSTEON_THERMOSTAT_TEMPLINC_2732_242"/>
		<nodeSubCategory id="20" name="DEV_SCAT_INSTEON_THERMOSTAT_TEMPLINC_2732_442"/>
		<nodeSubCategory id="21" name="DEV_SCAT_INSTEON_THERMOSTAT_TEMPLINC_2732_542"/>
		<nodeSubCategory id="22" name="DEV_SCAT_INSTEON_THERMOSTAT_TEMPLINC_2732_222_2"/>
		<nodeSubCategory id="23" name="DEV_SCAT_INSTEON_THERMOSTAT_TEMPLINC_2732_422_2"/>
		<nodeSubCategory id="24" name="DEV_SCAT_INSTEON_THERMOSTAT_TEMPLINC_2732_522_2"/>
	</nodeCategory>
	
	<nodeCategory id="6" name="Pool Control">
		<nodeSubCategory id="0" name="DEV_SCAT_COMPACTA_EZPOOL"/>
	</nodeCategory>
	
	<nodeCategory id="7" name="Sensors and Actuators">
		<nodeSubCategory id="0" name="DEV_SCAT_IO_LINC_2450"/>
		<nodeSubCategory id="1" name="DEV_SCAT_COMPACTA_EZSENSE"/>
		<nodeSubCategory id="2" name="DEV_SCAT_COMPACTA_EZIO_8T"/>
		<nodeSubCategory id="3" name="DEV_SCAT_COMPACTA_EZIO"/>
		<nodeSubCategory id="4" name="DEV_SCAT_COMPACTA_EZIO_8"/>
		<nodeSubCategory id="5" name="DEV_SCAT_COMPACTA_EZSNS_RF"/>
		<nodeSubCategory id="6" name="DEV_SCAT_COMPACTA_EZISNS_RF"/>
		<nodeSubCategory id="7" name="DEV_SCAT_COMPACTA_EZIO_6I"/>
		<nodeSubCategory id="8" name="DEV_SCAT_COMPACTA_EZIO_4O"/>
		<nodeSubCategory id="9" name="DEV_SCAT_SYNCHRO_LINC"/>
		<nodeSubCategory id="13" name="DEV_SCAT_REF_IO_LINC_2450"/>
	</nodeCategory>
	
	<nodeCategory id="9" name="Energy Management">
		<nodeSubCategory id="0" name="DEV_SCAT_EMETER_ZBPCM"/>
		<nodeSubCategory id="1" name="DEV_SCAT_ONSITE_PRO_LEAK_DETECTOR"/>
		<nodeSubCategory id="2" name="DEV_SCAT_ONSITE_PRO_CONTROL_VALVE"/>
		<nodeSubCategory id="7" name="DEV_SCAT_IMETER_SOLO"/>
		<nodeSubCategory id="10" name="DEV_SCAT_DUAL_BAND_NO_RELAY_240V_2477SA1"/>
		<nodeSubCategory id="11" name="DEV_SCAT_DUAL_BAND_NC_RELAY_240V_2477SA2"/>
		<nodeSubCategory id="13" name="DEV_SCAT_ENERGY_DISPLAY_2448A2"/>
	</nodeCategory>
	

	<nodeCategory id="14" name="Windows/Shades">
		<nodeSubCategory id="0" name="DEV_SCAT_SOMFY_DRAPE_CONTROLLER_RF"/>
		<nodeSubCategory id="1" name="DEV_SCAT_MICRO_MODULE_OPEN_CLOSE_2444_222"/>
		<nodeSubCategory id="2" name="DEV_SCAT_MICRO_MODULE_OPEN_CLOSE_2444_422"/>
		<nodeSubCategory id="3" name="DEV_SCAT_MICRO_MODULE_OPEN_CLOSE_2444_522"/>
	</nodeCategory>
	
	<nodeCategory id="15" name="Access Control/Doors/Locks">
		<nodeSubCategory id="0" name="DEV_SCAT_WEILAND_CENTRAL_DRIVE_CONTROLLER"/>
		<nodeSubCategory id="1" name="DEV_SCAT_WEILAND_SECONDARY_CENTRAL_DRIVE"/>
		<nodeSubCategory id="2" name="DEV_SCAT_WEILAND_ASSIST_DRIVE"/>
		<nodeSubCategory id="3" name="DEV_SCAT_WEILAND_ELEVATION_DRIVE"/>
		<nodeSubCategory id="6" name="DEV_SCAT_MORNING_LINC"/>
	</nodeCategory>
	
	<nodeCategory id="16" name="Security/Health/Safety">
		<nodeSubCategory id="1" name="DEV_SCAT_MOTION_SENSOR_2420M"/>
		<nodeSubCategory id="2" name="DEV_SCAT_TRIGGER_LINC_2421"/>
		<nodeSubCategory id="3" name="DEV_SCAT_MOTION_SENSOR_2420M_SP"/>
		<nodeSubCategory id="8" name="DEV_SCAT_LEAK_SENSOR_2852_222"/>
		<nodeSubCategory id="9" name="DEV_SCAT_OPEN_CLOSE_SENSOR_2843_222"/>
		<nodeSubCategory id="10" name="DEV_SCAT_SMOKE_SENSOR"/>
		<nodeSubCategory id="17" name="DEV_SCAT_DOOR_SENSOR_2845_222"/>
	</nodeCategory>
	
	<nodeCategory id="113" name="A10/X10 Nodes">
		<nodeSubCategory id="1" name="DEV_SCAT_X10"/>
		<nodeSubCategory id="2" name="DEV_SCAT_A10"/>
	</nodeCategory>
	
</NodeFamily>

"@

#Write-Verbose -Message ($MyInvocation.MyCommand.Name + ' called.') 

$fam_1_filename = '1_fam.xml'

$fullname = Join-Path $PSScriptRoot $fam_1_filename

if ( !(Test-Path $fullname) ) 
{
    $insteonfamily = [xml]$fam_1_content
}
else
{
    $insteonfamily = New-Object -TypeName XML
    $insteonfamily.Load($fullname)
}

$typearray = $type -split ".", 0, "simplematch"

$catid = $typearray[0]
$subcatid = $typearray[1]
$version = " v." +  [Convert]::ToString($typearray[2],16)

$current = $insteonfamily.NodeFamily.nodeCategory | where id -eq $catid

$repregex = '/| |and|device|Health|safety|access'

$catname = $current.name -replace $repregex

$typename = ($current | ForEach-Object {$_.nodeSubCategory | where id -eq $subcatid}).name.replace("DEV_SCAT_","")

$cattypeobj = Add-Member -InputObject (New-Object psobject) -NotePropertyName category -NotePropertyValue $catname -PassThru 
$cattypeobj | Add-Member -NotePropertyName type -NotePropertyValue ($typename + $version)

return ($cattypeobj)
}
#end function Resolve-InsteonType

<#
.Synopsis
   This function retrieves the ISY scene name or Insteon device name from the address provided.
.DESCRIPTION
   Used to get the name from the address shown in various files like logs, link tables or events.
.EXAMPLE
    Resolve-NodeAddress 67679

    Cooktop
#>
function Resolve-NodeAddress {
[CmdletBinding()]
[OutputType([string])]
param(
[Parameter(mandatory=$true)]
[validatescript({$addr = $_.trim()
                ($addr -in (Get-ISYScene -dbmap).foreach('address')) -or 
                ($addr.replace('.',' ') -in (Get-InsteonDevice -dbmap).foreach('address')) })]
[string]$address
)

if ($address.trim() -notmatch ' |\.')
{
    return (Get-ISYScene -dbmap | where address -eq $($address.trim())).name
}

return ((Get-InsteonDevice -dbmap) | where address -eq $($address.trim().replace('.',' '))).name

}
#end function Resolve-NodeAddress

<#
.Synopsis
   This function returns one URL from the 2 arguments -intip (internal IP address) and -extip (external IP address)
.DESCRIPTION
   This function will return a URL based upon one of 2 IP addresses passed to it.
   The 1st parameter is an internal IP address (inside NAT) and the 2nd is an external IP.
   If the internal IP is not '' and responds to a ping request then 'http://$intip' is returned.
   If the external IP is not '' it is also pinged but even if this fails 'https://$extip' is returned with a warning being issued as well.
   If there is no way to communcate to these IPs then an exception is thrown.
.EXAMPLE
    Select-IPToUse 192.168.1.89:77 177.19.235.67:6592
    http://192.168.1.89:77

    Since 192.168.89 responded to ping http://192.168.1.89:77 was returned.
.EXAMPLE
    Select-IPToUse '' 177.19.235.68:6592
    https://177.19.235.68:6592

    Internal IP is blank but the external IP is succesfully pinged so https://177.19.235.68:6592 is returned.
.EXAMPLE
    Select-IPToUse '' '177.19.235.67:6592'
    WARNING: Ping to 177.19.235.67 failed -- Status: TimedOut
    WARNING: External IP not responding to ping. This may be only because the router has ping response disabled.
    https://177.19.235.67:6592

    THe internal IP is empty so the external IP is pinged and fails but https://177.19.235.67:6592 is returned with a warning.
.EXAMPLE
    Select-IPToUse 192.168.1.88
    WARNING: Ping to 192.168.1.88 failed -- Status: TimedOut
    WARNING: Internal IP not responding to ping.
    Internal IP is either empty or not responding to ping and external IP is empty. No connection possible.
    At line:37 char:1
    + throw 'Internal IP is either empty or not responding to ping and exte ...
    + ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        + CategoryInfo          : OperationStopped: (Internal IP is ...ction possible.:String) [], RuntimeException
        + FullyQualifiedErrorId : Internal IP is either empty or not responding to ping and external IP is empty. No connection possible.

    The internal IP did not respond to ping and the external IP is empty. Therefore there is no communication possible internally or externally.
#>
function Select-IPToUse {
[cmdletbinding()]
param([string]$intip, [string]$extip)

$choices = $intip,$extip

if (($choices[0] -eq '') -and ($choices[1] -eq ''))
{
    throw "No connection possible to the ISY since both internal and external IP are blank.`nRe-execute New-ISYSettings and provide a value for either one or both of the IPs."
}

$sornot = @{0 = '';1 = 's'}


foreach ($i in 0,1)
{
    if ($choices[$i])
    {
        if (!(Test-Connection -ComputerName ($choices[$i].split(':'))[0] -count 1 -Quiet))
        {
            if (($i -eq 0))
            {
                $int = $false
                Write-Warning -Message "Internal IP not responding to ping."
                continue
            }
            else
            {
                Write-Warning -Message "External IP not responding to ping. This may be only because the router has ping response disabled."
            }
        }

        return "http$($sornot[$i])://$($choices[$i])"
    }
}

throw 'Internal IP is either empty or not responding to ping and external IP is empty. No connection possible.'

}
# end of function Select-IPToUse

<#
.Synopsis
   Sets the properties that will be displayed by default for an array of xml elements that are piped to it. 
.DESCRIPTION
   Sets the properties specified by the -PropertyList parameter that will be displayed by default for an array of xml elements
   specified by the -InputObject parameter. Any property in -PropertyList that is not in the list of properties of the xml 
   element is ignored.
   Note the xml element is left with all the properties, any downstream format-* or select-object sees all the properties. Only 
   the default display to the console host is affected.
   This functions' primary callers are other Get functions in this ISY994i module.
.EXAMPLE
    Get-InsteonDevice KitchenTableDimmer,MyReadingDimmer-XB1 

    OnLevel% address    fullpath                   name
    -------- -------    --------                   ----
    74       1 75 61 1  ISYRoot/Downstairs/Kitchen KitchenTableDimmer
    Off      16 96 98 1 ISYRoot/Upstairs/MediaRoom MyReadingDimmer-XB1

    Two insteon devices are retrieved by Get-InsteonDevice and the default properties name,address,onlevel%,fullpath are displayed.
    The Get-InsteonDevice cmdlet pipes its output to Set-DefaultDisplayProperties by default.
    By contrast the following output is seen when the setting of the default display properties are bypassed:

    PS C:\>Get-InsteonDevice KitchenTableDimmer,MyReadingDimmer-XB1 -nocustomdefaultdisplayproperties

    OnLevel%    : 74
    flag        : 128
    category    : Dimmables
    fullpath    : ISYRoot/Downstairs/Kitchen
    OnLevel     : 188
    address     : 1 75 61 1
    name        : KitchenTableDimmer
    parent      : parent
    type        : SWITCHLINC_V2_DIMMER_2476D v.27
    enabled     : true
    deviceClass : 512
    wattage     : 300
    dcPeriod    : 60
    pnode       : KitchenTableDimmer
    ELK_ID      : B01
    property    : property

    OnLevel%    : Off
    flag        : 128
    category    : Dimmables
    fullpath    : ISYRoot/Upstairs/MediaRoom
    OnLevel     : 0
    address     : 16 96 98 1
    name        : MyReadingDimmer-XB1
    parent      : parent
    type        : SWITCHLINC_DIMMER_W_SENSE_2476D v.38
    enabled     : true
    deviceClass : 512
    wattage     : 300
    dcPeriod    : 60
    pnode       : MyReadingDimmer-XB1
    ELK_ID      : A09
    property    : property

#>
function Set-DefaultDisplayProperties {
  [CmdletBinding(PositionalBinding=$false)]
  param
  (
    # the list of xml elements to set the default properties for
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [System.Xml.XmlElement[]]
    $InputObject,
    
    # the list of default display properties
    [Parameter(Mandatory=$true,Position = 0)]
    [string[]]
    $PropertyList
  )
  
  begin
  {
    
  }

  process
  {
     $entry = $_
     $propnames = ($entry | Get-Member -MemberType Properties).Name
     $defaultproperties = $propnames | Where-Object {$_ -in $PropertyList} 
     if (!$defaultproperties) 
     {
        Write-Warning -Message "No properties found with name(s) $($PropertyList -join ',')" 
        $entry
        return
     }
     $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultproperties)
     $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
     $entry | Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $PSStandardMembers -PassThru
  }

  end
  {

  }
  
}
# end function Set-DefaultDisplayProperties

<#
.Synopsis
   Set-Insteondevice

   Turns an Insteon device (or list of devices) on or off. The level parameter can be used with the DON command
   to set the level on a dimming device. Default level is fully on.  
.DESCRIPTION
   Only the following commands are supported: "ON,"DON","OFF","DOF","DFON","DFOF","BEEP"
   ON,DON: Turns device on. Level parameter from 1..255 - 255 is full brightness. To specify a percentage use a trailing % sign (1% to 100%).
   OFF,DOF: Turns device off. Level parameter is ignored.
   DFON: Turns device fast on. Level parameter is ignored.
   DFOF: Turns device fast off. Level parameter is ignored.
   BEEP: Device beeps if it is capable. Level parameter is ignored.

   The -passthru switch performs a Get-InsteonDevice on the devices that were operated on and returns the list of devices to the caller. 
   Normally nothing is returned to the caller. Can be used to check the results of the action performed on the devices.

   The -RC (result code) switch causes the response from the ISY to be sent back. Normally nothing is returned to the caller.

   Notes on errors:
     - For any device name that is not valid you will get an error like:
     "<name> is not an insteon device." 

     - If your level parameter is not in range (1-255 or 1% to 100%) you will get an error like:
     Set-Insteondevice : Cannot validate argument on parameter 'level'. The argument "150%" does not match the "((^100%$)|(^[.....
    
.EXAMPLE
    This turns the dimmer called MiddleReadingLightSwitch-XB2 on to a 50% level.
    Because of the -RC switch the command returns an acknowledgment. 

    Set-Insteondevice -name MiddleReadingLightSwitch-XB2 -command don -level 50% -rc

    xml                              RestResponse
    ---                              ------------
    version="1.0" encoding="UTF-8"   RestResponse
.EXAMPLE
    This turns the light off.
    To see what is in the response we examine the RestResponse property.

    (Set-Insteondevice -name MiddleReadingLightSwitch-XB2 -command dof -rc).RestResponse | ft -au

    succeeded status
    --------- ------
    true      200
.EXAMPLE
    The function can also accept input from the pipeline and a list of devices as the value of the -name parameter.

    > $kitchendimmers = Get-InsteonDevice -listdb | where fullpath -like *kitchen* | where name -like *dimmer* 

    > $kitchendimmers | select name, fullpath, type

    name                        fullpath                   type                                   
    ----                        --------                   ----                                   
    GarageKitchenEntranceDimmer ISYRoot/Downstairs/Kitchen KEYPAD_LINC_DIMMER_2334_2_5_BUTTON v.43
    KitchenTableDimmerKTKP1     ISYRoot/Downstairs/Kitchen KEYPADLINC_DIMMER_2486D v.29           
    CooktopDimmer               ISYRoot/Downstairs/Kitchen SWITCHLINC_V2_DIMMER_2476D v.27        
    KitchenTableDimmer          ISYRoot/Downstairs/Kitchen SWITCHLINC_V2_DIMMER_2476D v.27        

    > $kitchendimmers | Set-InsteonDevice -command on

    > Set-InsteonDevice $kitchendimmers off

#>
function Set-InsteonDevice {
[CmdletBinding()]
[OutputType([void])]
param(
# list of insteon device names  or list of insteon device xmlelements obtained from Get-Insteondevice
# ByPropertyName is redundant? http://powershell.com/cs/blogs/donjones/archive/2011/12/10/troubleshooting-pipeline-parameter-binding-by-peeking-inside.
[Parameter(mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true )]  
[object[]]$name,

# the command to be executed on the insteon device - either ON,OFF,DON,DOF,DFON,DFOF or BEEP
[Parameter(mandatory=$true,Position=1)]
[Alias('cmd')]
[validateset("ON","OFF","DON","DOF","DFON","DFOF","BEEP")]
[string]$command,

# the level of brightness to set for dimming devices. (Relay devices are full On or Off so this param is ignored). Can be 0..255 or specified as a % like 75%
[Parameter(Position=2)]
[string]$level = 255,

# if set then the function executes a Get-InsteonDevice on the devices that were operated on and returns the result
[Parameter()]
[switch]$passthru,

# if set the function receives the result code associated with the REST command sent to the ISY. Normally the RC is consumed by the Invoke-ISYRestMethod function.
[switch]$RC
)

Begin { 
    Write-Verbose -Message ($MyInvocation.MyCommand.Name + ' called.') 

    $devicesset = @()
    $delay = 3
    
    # verify isy settings before ISY access
    [void](Get-ISYSettings)

    $insteondevmap = (Get-InsteonDevice -dbmap)

    function InvokeISYRestMethod {
    [CmdletBinding()]
    [OutputType([void])]
    param
    (
        [string]$name
    )

        if ($name -in $insteondevmap.name)
        {
            $nodeid = ($insteondevmap | where name -eq $name).address

            #$levelrx = '(^1[0-9][0-9]$)|(^2[0-4][0-9]$)|(^25[0-5]$)|(^[1-9][0-9]$)|(^[0-9]$)'
            #$pctrx = '(^100%$)|(^[1-9][0-9]%$)|(^[0-9]%$)'

            $cmd = $command.ToUpper()

            switch ($command)
            {
                "ON"  {$cmd = "DON"}
                "OFF" {$cmd = "DOF"}
            }

            If ($level -match '%$') {$level = [string][int]([int]($level.trimend('%'))*255/100)}

            $IsyRestMethodparams = @{
            RestPath  = "/rest/nodes/$nodeid/cmd/$cmd/$level"
            handlerc = !$RC}
            Write-Verbose "Set-InsteonDevice restpath = $($IsyRestMethodparams.RestPath)"

            Invoke-IsyRestMethod @IsyRestMethodparams

        }
        else
        {
            Write-Warning -Message "$name is not an insteon device." 
        }
    }
}

Process { 

    if (($PSCmdlet.MyInvocation.ExpectingInput))
    {
        $devname = $_
        if ($_.name)
        {
            $devname = $_.name
        }
        InvokeISYRestMethod -name $devname

        $devicesset += $devname
    }
}

End {
        if (!($PSCmdlet.MyInvocation.ExpectingInput))
        {   
            $gm = $name | Get-Member

            if (($gm.typename | Sort-Object -Unique).count -ne 1)
            {
                Write-Error -Message "Object array with members having multiple types not supported." -Category InvalidData
                return                    
            }
            elseif ($gm.typename -eq 'system.string')
            {
                $namelist = $name 
            }
            elseif ($gm.name -contains 'name')
            {
                $namelist = $name.name
            }
            else 
            {
                Write-Error -Message "Cannot process object of type $($gm.typename)." -Category InvalidData
                return
            }
            
            foreach ($entry in $namelist)
            {
                InvokeISYRestMethod -name $entry

                $devicesset += $entry
            }
        }

        if ($passthru)
        {
            Start-Sleep -Seconds $delay
            Get-InsteonDevice -name $devicesset
        }
    }
}
#end function Set-InsteonDevice

<#
.Synopsis
   Sets the current debug level of the ISY.
.DESCRIPTION
   Debug level can be set to either 1, 2 or 3. Default is 1.
.EXAMPLE
    Set-ISYDebuglevel -level 2

    PS C:\>Get-ISYDEbugLevel 
    2

    First the ISY debug level is set to 2 and then Get-ISYDebugLevel confirms it is at 2.
#>
function Set-ISYDebugLevel {
[cmdletbinding()]
param(
# debug level either 1, 2 or 3 
[ValidateRange(1,3)]
[string]$level = 1,

# if set the function outputs the result code associated with the webrequest sent to the ISY. Normally a successful operation returns nothing to the caller.
[switch]$RC
)

$headers = Get-Headers -SOAPAction "urn:udi-com:device:X_Insteon_Lighting_Service:1#UDIService"
write-verbose -Message "auth = $($headers['Authorization'])"
Write-Verbose -Message "soapaction = $($headers['SOAPAction'])"

$body = "<?xml version='1.0' encoding='utf-8'?>"
$body += "<s:Envelope><s:Body><u:SetDebugLevel"
$body += " xmlns:u='urn:udi-com:service:X_Insteon_Lighting_Service:1'>"
$body += "<option>$level</option></u:SetDebugLevel></s:Body></s:Envelope>`r`n" 

Write-Verbose -Message "Body:`n$body`n"

$uri = Get-ISYUrl -path '/services'
Write-Verbose -Message "the uri = $uri"

$splat = @{Uri = $uri
           DisableKeepAlive = $true
           Method = "POST"
           ContentType = 'text/xml; charset=utf-8'
           Headers = $headers
           Body = $body
            }

$response = Invoke-WebRequest @splat

if (!$RC)
{
    $status = ([xml]($response.Content)).Envelope.Body.UDIDefaultResponse.Status
    if ($status -ne '200')
    {
        Write-Error -Message "Something went wrong. Status = $status" -Category InvalidOperation
    }
}
else
{
    ([xml]($response.Content)).Envelope.Body.UDIDefaultResponse
}
}
# end of function Set-ISYDebugLevel

<#
.Synopsis
   This function launches a GUI form for the user to fill in default values for the ISY.
.DESCRIPTION
   The fields datafilled are the IP addresses (internal and external), PLM address, the 
   time span between automatic db refreshes, the path to the db xml files and the boolean
   representing whether tab expansion for the ISY commands are in effect. The ISY commands 
   which tab expansion works for are the *Device, *Scene and Invoke-ISYProgram commands.

   Note the form presented to the user will display the text boxes with the values previously
   collected. If run for the 1st time these fields will be empty.
   
   The only parameter this function takes is the -returnhash switch. With this switch 
   present the function will set, save and return a hashtable of the the above fields and 
   then return. Without the -returnhash switch the function will go on to call New-ISYSettings     
   and thus launch the credential dialog window for the user to enter the username and password   
   to access the ISY. Then New-ISYSettings will continue on to create the full settings for the    
   ISY and login and refresh the db of Insteon devices, ISY scenes and ISY programs.
#>
function Set-ISYDefaults {
    [CmdletBinding()]
    [OutputType([hashtable])]
    Param
    (
        # do not continue on to gather credential info - only create the non credential field info.
        [Parameter()]
        [switch]
        $returnhash
    )

    function Resolve-Expression {
        [CmdletBinding()]
        [OutputType([string])]
        param([string]$expr)

<#
# quick success test for directories that already exist:
'$home/documents','$env:localappdata','(get-item env:/appdata).value', 'c:/isy' | % {Resolve-Path (Resolve-Expression -expr $_)}
#>   
 
        if ($expr -eq '' -or $expr -eq $null) {return ''}

        $ConfirmPreference = 'low'
    
        $result = ''

        try
        {
            $result = $ExecutionContext.InvokeCommand.InvokeScript($expr)
        }
        catch
        {
            write-verbose -Message "InvokeScript must have trapped, trying ExpandString."
            $result = $ExecutionContext.InvokeCommand.ExpandString($expr)
        }
        finally
        {
            $result
        }
    }

    function Test-ISYDefaults{
        [CmdletBinding()]
        [OutputType([void])]
        Param
        (
            # hashtable of default values
            [Parameter(Mandatory=$true, Position=0)]
            [hashtable]
            $defaults
        )

        $message = "" 

        # validate the path
        try
        {
            $evaledpath = Resolve-Expression $defaults.path
        }
        catch
        {
        $message += "`nCannot resolve $($defaults.path) to a directory. `nChange path setting to an existing directory or create directory $($defaults.path)"
        return $message
        }
        if (($evaledpath -eq '') -or !(Test-Path $evaledpath -PathType Container))
        {
            $message += "`n$($defaults.path) is not, or does not a evaluate to, an existing directory. `nChange path setting to an existing directory or create directory $($defaults.path)"
        }

        # validate refreshtime
        try
        {
            $timespan = [timespan]($defaults.refreshtime)
        }
        catch
        {
            $message += "`n$($defaults.refreshtime) is not a valid time span."
            return $message
        }
        if ($timespan -lt [timespan]"00:30" -or $timespan -gt [timespan]"23:59")
        {
            $message += "`n$timespan is not between 00:30 and 23:59"
        }

        # validate IPs
        if ($defaults.internalip -eq '' -and $defaults.externalip -eq '') 
        {
            $message += "`nBoth Internal IP and External IP cannot be blank."
        }

        if ($defaults.internalip -like '*:*:*') 
        {
            $message += "`nISY994i does not support IPv6 for internal IP."
        }

        if ($defaults.externalip -like '*:*:*')
        {
            $message += "`nISY994i does not support IPv6 for external IP."
        }
        return $message
    }

$inputxml = @"
<Window x:Name="ISYSettingsMainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="ISY Default Settings" Height="350" Width="525">
    <Grid Margin="0,2,-0.333,-2.333">
        <Label x:Name="InternalIPLabel" Content="Internal IP Address" HorizontalAlignment="Left" Margin="30,28,0,0" VerticalAlignment="Top"/>
        <Label x:Name="ExternalIPLabel" Content="External IP Address" HorizontalAlignment="Left" Margin="30,69,0,0" VerticalAlignment="Top"/>
        <Label x:Name="PLMAddressLabel" Content="PLM Address" HorizontalAlignment="Left" Margin="29,114,0,0" VerticalAlignment="Top"/>
        <Label x:Name="RefreshTimeLabel" Content="Refresh Time" HorizontalAlignment="Left" Margin="30,159,0,0" VerticalAlignment="Top"/>
        <Label x:Name="PathLabel" Content="Path" HorizontalAlignment="Left" Margin="30,205,0,0" VerticalAlignment="Top"/>
        <Label x:Name="TabExpansionLabel" Content="Tab Expansion" HorizontalAlignment="Left" Margin="30,249,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="InternalIPTextBox" HorizontalAlignment="Left" Height="23" Margin="207,31,0,0" TextWrapping="Wrap" ToolTip="Internal IP address or DNS name. Ex: 192.168.1.11" VerticalAlignment="Top" Width="251" FontFamily="Segoe UI Light" />
        <TextBox x:Name="ExternalIPTextBox" HorizontalAlignment="Left" Height="23" Margin="207,73,0,0" TextWrapping="Wrap" ToolTip="External IP address:port or DNS name:port. Ex: 11.12.13.14:9443" VerticalAlignment="Top" Width="251" FontFamily="Segoe UI Light"/>
        <TextBox x:Name="PLMAddressTextBox" HorizontalAlignment="Left" Height="23" Margin="207,118,0,0" TextWrapping="Wrap" ToolTip="Get from label on PLM. Ex: FA 9 4B" VerticalAlignment="Top" Width="251" FontFamily="Segoe UI Light"/>
        <TextBox x:Name="RefreshTimeTextBox" HorizontalAlignment="Left" Height="23" Margin="207,162,0,0" TextWrapping="Wrap" ToolTip="00:30 to 23:59" VerticalAlignment="Top" Width="251" FontFamily="Segoe UI Light"/>
        <TextBox x:Name="PathTextBox" HorizontalAlignment="Left" Height="23" Margin="207,209,0,0" TextWrapping="Wrap" ToolTip="Valid local directory. Ex: C:\ISY " VerticalAlignment="Top" Width="251" FontFamily="Segoe UI Light"/>
        <CheckBox x:Name="TabExpansionCheckBox" Content="For choices of -name param values in commands" HorizontalAlignment="Left" Margin="207,251,0,0" VerticalAlignment="Top" Height="26" FontFamily="Segoe UI Light" ToolTip="Commands affected: *Device|*Scene|Invoke-ISYProgram"/>
        <Button x:Name="SubmitButton" Content="Submit" HorizontalAlignment="Left" Margin="246,282,0,0" VerticalAlignment="Top" Width="75" />
        <Button x:Name="Cancelbutton" Content="Cancel" HorizontalAlignment="Left" Margin="387,282,0,0" VerticalAlignment="Top" Width="75"/>
    </Grid>
</Window>
"@

    $inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:Name",'Name'  -replace '^<Win.*', '<Window'
 
    $VarPrefix = 'ISYDefaultsForm'
    Write-Verbose "`$VarPrefix = $VarPrefix"

    $Script:formcancelled = $false

    Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,system.windows.forms
    [xml]$XAML = $inputXML

    #Read XAML
    $reader = (New-Object System.Xml.XmlNodeReader $xaml) 

    # watch for XamlParseException exceptions. Look for "Failed to create...from...'<whatever>'"
    # look for <whatever> in xaml and remove attribute. (ie TextChanged="RefreshText_TextChanged", Checked="checkBox_Checked")
    # These are the names of methods VS adds when placing objects on the designer screen.
    # sls -Path .\xaml-from-vs.xaml -Pattern ' .*=".*_.*"'
    $Form = [Windows.Markup.XamlReader]::Load( $reader )

    #===========================================================================
    # Load XAML Objects In PowerShell
    #===========================================================================
     Write-Verbose "Loading xaml objects"

    $xaml.SelectNodes("//*[@Name]") | ForEach-Object {Set-Variable -Name "$VarPrefix$($_.Name)" -Value $Form.FindName($_.Name)}
 
    Write-Verbose "Found the following interactable elements from form" 
    get-variable "$VarPrefix*" | % {write-verbose ("{0} = {1}" -f $_.name, $_.value)}

    if (Test-Path -Path $ISYDefaultsFilePath)
    {
        $ISYHash = Import-Clixml ($ISYDefaultsFilePath)
        $ISYDefaultsFormInternalIPTextBox.Text = $ISYHash.internalip
        $ISYDefaultsFormExternalIPTextBox.Text = $ISYHash.externalip
        $ISYDefaultsFormPLMAddressTextBox.Text = $ISYHash.plmaddress
        $ISYDefaultsFormRefreshTimeTextBox.Text = $ISYHash.refreshtime
        $ISYDefaultsFormPathTextBox.Text = $ISYHash.path
        $ISYDefaultsFormTabExpansionCheckBox.IsChecked = $ISYHash.tabexp
    }

    $ISYDefaultsFormCancelbutton.add_click({ $script:formcancelled = $true
                                             write-verbose "Cancelled no action taken."
                                             $form.close() })

    $ISYDefaultsFormSubmitButton.add_click( {$Script:ISYDefaultHash = @{internalip = $ISYDefaultsFormInternalIPTextBox.Text
                                                                     externalip = $ISYDefaultsFormExternalIPTextBox.Text
                                                                     plmaddress = $ISYDefaultsFormPLMAddressTextBox.Text
                                                                     refreshtime = $ISYDefaultsFormRefreshTimeTextBox.Text
                                                                     path = $ISYDefaultsFormPathTextBox.Text
                                                                     tabexp = $ISYDefaultsFormTabExpansionCheckBox.IsChecked  }
                                        $message = Test-ISYDefaults -defaults $ISYDefaultHash
                                        if ($message -ne '')
                                        { 
                                            $messpref = "After clicking OK, fix the issue(s) below and then press Submit again."
                                            [System.Windows.Forms.MessageBox]::Show($messpref + $message)
                                        }
                                        else
                                        {
                                            $ISYDefaultHash.refreshtime = [timespan]($ISYDefaultsFormRefreshTimeTextBox.Text)
                                            $ISYDefaultHash.path = Resolve-Path (Resolve-Expression $ISYDefaultHash.path)
                                            #$ISYDefaultHash | Export-Clixml $ISYDefaultsFilePath
                                            $form.close()}
                                        } )

    $Form.ShowDialog() | out-null

    if (!$script:formcancelled)
    {
        Write-Verbose "ISYDefaultHash.count = $($Script:ISYDefaultHash.count)"
        $Script:ISYDefaultHash.keys | % {write-verbose "$_ = $($Script:ISYDefaultHash[$_])"}

        $ISYDefaultHash | Export-Clixml $ISYDefaultsFilePath

        if ($returnhash.IsPresent) 
        {
            return $ISYDefaultHash
        }
        else
        {
            New-ISYSettings @ISYDefaultHash
            return
        }
    }
}
#end function Set-ISYDefaults

<#
.Synopsis
   Set-ISYScene

   Turns an ISY scene (or list of scenes) on or off. 
.DESCRIPTION
   Only the following commands are supported: "ON","DON","OFF","DOF","DFON","DFOF"
   ON,DON: Turns scene on. 
   OFF,DOF: Turns scene off. 
   DFON: Turns scene fast on. 
   DFOF: Turns scene fast off. 

   The -passthru switch performs a Get-ISYScene on the scenes that were operated on and returns the list of devices to the caller. 
   Normally nothing is returned to the caller. Can be used to check the results of the action performed on the scenes.

   The -RC (return code) switch causes the response from the ISY to be sent back. Normally nothing is returned to the caller. 

   Notes on errors:
     - If the scene name is not valid you will get an error like:
       "<name> is not an ISY scene." 

.EXAMPLE
    This turns the scene called Cooktop on.
    The -rc switch causes the function to return an acknowledgment. 

    Set-ISYScene -name Cooktop -command don -rc

    succeeded status
    --------- ------
    true      200
.EXAMPLE
    This turns the scene off.
    To see what is in the response we examine the RestResponse property.

    Set-ISYScene -name Cooktop -command don -rc | Format-Table -Autosize

    succeeded status
    --------- ------
    true      200
.EXAMPLE
    The function also accepts input from the pipeline and an array as input for the name parameter.
    
    > $kitchenscenes =  Get-ISYSceneList | where fullpath -like *kitchen*

    > $kitchenscenes | ft

    flag fullpath                   address name         parent deviceGroup ELK_ID members
    ---- --------                   ------- ----         ------ ----------- ------ -------
    132  ISYRoot/Downstairs/Kitchen 47750   KitchenTable parent 17          A08    members
    132  ISYRoot/Downstairs/Kitchen 67679   Cooktop      parent 18          B03    members


    > Set-ISYScene -name $kitchenscenes -command ON

    > Set-ISYScene  $kitchenscenes  Off    
#>
function Set-ISYScene {
[CmdletBinding()]
[OutputType([void])]
param(
# list of scene names or list of scene xmlelements obtained from Get-ISYScene
[Parameter(mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true )]
[object[]]$name,

# command to to executed on the scene(s) - either ON,OFF,DON,DOF,DFON or DFOF
[Parameter(mandatory=$true,Position=1)]
[Alias('cmd')]
[validateset("ON","OFF","DON","DOF","DFON","DFOF")]
[string]$command,

# if set then the function executes a Get-ISYScene on the scenes that were operated on and returns the result
[Parameter()]
[switch]$passthru,

# if set the function receives the result code associated with the REST command sent to the ISY. Normally the RC is consumed by the Invoke-ISYRestMethod function.
[switch]$RC
)

BEGIN {
    Write-Verbose -Message ($MyInvocation.MyCommand.Name + ' called.') 

    $scenesset = @()
    $delay = 3

    # verify isy settings before ISY access
    [void](Get-ISYSettings)

    $isyscenemap  = Get-ISYScene -dbmap

    function InvokeISYRestMethod {
    [CmdletBinding()]
    [OutputType([void])]
    param
    (
        [string]$name
    )

        if ($name -in $isyscenemap.name)
        {
        $nodeid = ($isyscenemap | where name -eq $name).address

        $cmd = $command.ToUpper()

        switch ($command)
        {
            "ON"  {$cmd = "DON"}
            "OFF" {$cmd = "DOF"}
        }

        $IsyRestMethodparams = @{
        RestPath  = "/rest/nodes/$nodeid/cmd/$cmd"
        handlerc = !$RC}

        Invoke-IsyRestMethod @IsyRestMethodparams

        #Invoke-IsyRestMethod "/rest/nodes/$nodeid/cmd/$cmd" 
        }
        else
        {
            Write-Warning -Message "$name is not an ISY scene." 
        }
    } 
}

PROCESS {
    if (($PSCmdlet.MyInvocation.ExpectingInput))
    {
        $scenename = $_
        if ($_.name)
        {
            $scenename = $_.name
        }

        $scenesset += $scenename

        InvokeISYRestMethod -name $scenename
    }
}

END {
    if (!($PSCmdlet.MyInvocation.ExpectingInput))
    {
        Write-Verbose -Message 'Executing on the command line object.'
            
        $gm = $name | Get-Member

        if (($gm.typename | Sort-Object -Unique).count -ne 1)
        {
            Write-Error -Message "Object array with members having multiple types not supported." -Category InvalidData
            return                    
        }
        elseif ($gm.typename -eq 'system.string')
        {
            $namelist = $name 
        }
        elseif ($gm.name -contains 'name')
        {
            $namelist = $name.name
        }
        else 
        {
            Write-Error -Message "Cannot process object of type $($gm.typename)." -Category InvalidData
            return
        }
            
        foreach ($entry in $namelist)
        {
            $scenesset += $entry
            
            InvokeISYRestMethod -name $entry
        }
    }

    if ($passthru)
    {
        Start-Sleep -Seconds $delay

        Get-ISYScene -name $scenesset
    }
  }
}
#end function Set-ISYScene

<#
.Synopsis
   Set-X10DEvice

   Sends X10 commands to housecodes and unitcodes. 
     
.DESCRIPTION
   The following commands are supported: 
   Standard codes:
   "AllUnitsOff","AllLightsOn","AllLightsOff","On","Off","Dim","Bright",
   Extended codes:
   "ExtendedCode","HailRequest","HailAcknowledge","PreSetDim","StatusIsOn","StatusIsOff","StatusRequest"

   The function has 2 parameter sets: x10codes and ISY names.
    - x10codes require a housecode and a unitcode (housecode only for allights, allunits etc)
    - ISYNames can map names to house and unit code iff the names of the devices end with -X<housecode><unitcode>

   The -RC (return code) switch causes the response from the ISY to be sent back. Without this switch nothing is returned to the caller. 

   Notes on errors:
     - If the device name is not valid you will get an error like:
     Set-X10Device : Cannot validate argument on parameter 'name'. The "$_ -in ((Get-InsteonDevice -dbmap).name | where name -Match "-x[...

     - If your houseandunitcode are out of range you will get an error like:
     Set-X10Device : Cannot validate argument on parameter 'houseandunitcode'. The argument "q3" does not match the "(^[A-P][1-...
    
.EXAMPLE
    This turns the X10 device with address D1 on

    > Set-X10DEvice -houseandunitcode d1 -command On

#>
function Set-X10Device {
[CmdletBinding(DefaultParameterSetName = 'X10COdes')]
[OutputType([void])]
param(
[Parameter(ParameterSetName = 'X10COdes',mandatory=$true,Position = 0)]
[validatepattern('(^[A-P][1-9][0-6]$)|(^[A-P][1-9])$|(^[A-P])$')]
[string]$houseandunitcode,

[Parameter(ParameterSetName = 'ISYNames',mandatory=$true,Position = 0)]
[validatescript({$_ -in ((Get-InsteonDevice -dbmap | where name -match "-X[A-P](([1-9]$)|(1[0-6]$))")).name})]
[string]$name,

[validateset("AllUnitsOff","AllLightsOn","AllLightsOff","On","Off","Dim","Bright",
             "ExtendedCode","HailRequest","HailAcknowledge","PreSetDim","StatusIsOn","StatusIsOff","StatusRequest")]
[string]$command,

[switch]$RC)

#[X10Commands]$command # if enum is used this is all that is required

Write-Verbose -Message ($MyInvocation.MyCommand.Name + ' called.') 

# verify isy settings before ISY access
[void](Get-ISYSettings)

Write-Verbose "`n"
Write-Verbose $houseandunitcode.ToUpper()
write-verbose $name
write-verbose $x10decoder[$($command)]
Write-Verbose "`n"

if ($PSCmdlet.ParameterSetName -eq 'ISYNames') 
{
    # my naming convention has any devices with x10 codes having names ending in -X<housecode><unitcode>
    $rgx = "-X[A-P](([1-9]$)|(1[0-6]$))"
    [void]($name -match $rgx) 
    $code = $matches[0].Replace('-X','')
    $houseandunitcode = $code
}
    
$ISYRestMethodparams = @{
                        RestPath = "/rest/X10/$($houseandunitcode.ToUpper())/$($x10decoder[$($command)])"
                        handlerc = !$RC}

Invoke-ISYRestMethod @ISYRestMethodparams

}
#end function Set-X10DEvice

<#
.Synopsis
   Show-ISYObject

   This function displays the content of the objects produced by the functions in this module. 
.DESCRIPTION
  The syntax to get some of the information from the various objects in a pleasing format can be quite complex and obscure.
  To help with this the Show-ISYObject function can be used to get what is thought to be the most useful data displayed on 
  the console in a friendly format. 
  The function supports input objects from the pipeline as well as the parameter method.
  Note this function only supports the output from the higher level functions like Get-ISYScene and Get-InsteonDevice. 
  The 'raw' data returned directly by Invoke-ISYRestMethod is not supported. The Invoke-ISYRestMethod output is what the
  higher level functions consume and modify.
  Also note that the output of this function is meant for display only. It does NOT output consumable objects since it uses
  Format-Table and Format-List. The -formatgrid param can be used to create objects (via export-csv for example) via the  
  ouputmode/passthru parameters. See 'Get-Help Out-GridView -examples' for more info on how to do this.
    
.EXAMPLE
    Get the list of devices and display the relevant fields on the console in table format. 
    
    > Get-InsteonDevice | Show-ISYObject -formatable

    name                         address    OnLevel% OnLevel type                                    fullpath                         enabled flag deviceClass dcPeriod wattage ELK_ID pnode
    ----                         -------    -------- ------- ----                                    --------                         ------- ---- ----------- -------- ------- ------ -----
    AtticDimmer                  36 72 49 1 Off      0       SWITCHLINC_DIMMER_2WIRE_2474DWH v.42    ISYRoot/Upstairs                 true    128  0           0        0       B04    AtticDimmer
    BackDoorDimmer-XD5           27 78 34 1 Off      0       SWITCHLINC_DIMMER_2477D v.41            ISYRoot/Outside                  true    128  512         60       300     C15    BackDoorDimmer-XD5
    CeciliaInlineDimmer-XC4      20 95 87 1 Off      0       INLINELINC_DIMMER_2475DA1 v.41          ISYRoot/Downstairs/MasterBedroom true    128  512         60       300     C01    CeciliaInlineDimmer-XC4
    CooktopButtonKTKP3           9 93 11 3  Off      0       KEYPADLINC_DIMMER_2486D v.29            ISYRoot/Downstairs/Kitchen       true    0    0           0        0       A13    KitchenTableDimmerKTKP1
    CooktopDimmer                1 60 30 1  Off      0       SWITCHLINC_V2_DIMMER_2476D v.27         ISYRoot/Downstairs/Kitchen       true    128  512         60       300     B02    CooktopDimmer
    DresserInlineDimmer-XC5      20 95 49 1 Off      0       INLINELINC_DIMMER_2475DA1 v.41          ISYRoot/Downstairs/MasterBedroom true    128  512         60       300     B15    DresserInlineDimmer-XC5
    DriveWayAlert-Sensor-XD1     17 66 28 1 Off      0       IO_LINC_2450 v.36                       ISYRoot/Outside                  true    128  0           0        0       B10    DriveWayAlert-Sensor-XD1
    FloodlightsButtonKTKP3-XD2   9 93 11 6  Off      0       KEYPADLINC_DIMMER_2486D v.29            ISYRoot/Downstairs/Kitchen       true    0    0           0        0       A16    KitchenTableDimmerKTKP1
    FrontPorchDimmer             27 78 45 1 Off      0       SWITCHLINC_DIMMER_2477D v.41            ISYRoot/Outside                  true    128  0           0        0       D10    FrontPorchDimmer
    GarageBackDoorDimmer         16 96 86 1 Off      0       SWITCHLINC_DIMMER_W_SENSE_2476D v.38    ISYRoot/Outside                  true    128  256         60       300     D05    GarageBackDoorDimmer
    GarageCarDoorDimmer-XD3      17 6 82 1  Off      0       SWITCHLINC_DIMMER_2477D v.40            ISYRoot/Outside                  true    128  256         60       300     C14    GarageCarDoorDimmer-XD3
    GarageCeilingLightButtonGKKP 24 84 71 3 Off      0       KEYPAD_LINC_DIMMER_2334_2_5_BUTTON v.43 ISYRoot/Downstairs/Kitchen       true    0    0           0        0       C10    GarageKitchenEntranceDimmer
    GarageCeilingLightSwitch     33 85 34 1 Off      0       SWITCHLINC_RELAY_DUAL_BAND_2477S v.43   ISYRoot/Garage                   true    128  512         60       300     B13    GarageCeilingLightSwitch
    GarageKitchenEntranceDimmer  24 84 71 1 Off      0       KEYPAD_LINC_DIMMER_2334_2_5_BUTTON v.43 ISYRoot/Downstairs/Kitchen       true    128  0           0        0       B14    GarageKitchenEntranceDimmer
    GarageKitchenKeypad-4        24 84 71 4 Off      0       KEYPAD_LINC_DIMMER_2334_2_5_BUTTON v.43 ISYRoot/Downstairs/Kitchen       true    0    0           0        0       D02    GarageKitchenEntranceDimmer
    GarageKitchenKeypad-5        24 84 71 5 Off      0       KEYPAD_LINC_DIMMER_2334_2_5_BUTTON v.43 ISYRoot/Downstairs/Kitchen       true    0    0           0        0       D03    GarageKitchenEntranceDimmer
    GarageKitchenKeypad-6        24 84 71 6 Off      0       KEYPAD_LINC_DIMMER_2334_2_5_BUTTON v.43 ISYRoot/Downstairs/Kitchen       true    0    0           0        0       D04    GarageKitchenEntranceDimmer
    GuestReadingDimmer-XB3       16 96 86 1 Off      0       SWITCHLINC_DIMMER_W_SENSE_2476D v.38    ISYRoot/Upstairs/MediaRoom       true    128  512         60       300     A11    GuestReadingDimmer-XB3
    HallwayDimmer                27 78 40 1 Off      0       SWITCHLINC_DIMMER_2477D v.41            ISYRoot/Downstairs               true    128  512         60       300     B12    HallwayDimmer
     :
     :
    SlidingDoorOutDimmer-XD6     16 98 41 1 91       231     SWITCHLINC_DIMMER_W_SENSE_2476D v.38    ISYRoot/Outside                  true    128  256         60       300     C16    SlidingDoorOutDimmer-XD6
    Unemployed-Relay             17 66 28 2 Off      0       IO_LINC_2450 v.36                       ISYRoot/Upstairs/MediaRoom       true    0    0           0        0       B11    DriveWayAlert-Sensor-XD1

.EXAMPLE
    This example shows that the parameter approach can be used as well. Note the default format is -formatlist

    > Show-ISYObject (Get-Insteondevice -name HallwayDimmer)

    name        : HallwayDimmer
    address     : 27 78 40 1
    OnLevel%    : Off
    OnLevel     : 0
    type        : SWITCHLINC_DIMMER_2477D v.41
    fullpath    : ISYRoot/Downstairs
    enabled     : true
    flag        : 128
    deviceClass : 512
    dcPeriod    : 60
    wattage     : 300
    ELK_ID      : B12
    pnode       : HallwayDimmer

.EXAMPLE
    Get the status of a scene in a variable and then display the info to the console.
    This shows the CooktopDimmer at 35% and the CooktopButtonKTKP3 at full on (OnLevel 255)

    > $scenestatus = Get-ISYScene -name cooktop
    > $scenestatus | Show-ISYObject
    ------------------------------------------------------------
    name        : Cooktop
    address     : 67679
    fullpath    : ISYRoot/Downstairs/Kitchen
    flag        : 132
    deviceGroup : 18
    ELK_ID      : B03

    MemberName         type       OnLevel% OnLevel
    ----------         ----       -------- -------
    Remote1-2          Controller Off      0
    CooktopButtonKTKP3 Controller On       255
    CooktopDimmer      Controller 35       89

.EXAMPLE
    The function also supports the -formatgrid switch which ultimately makes a call to Out-GridView which  
    opens an interactive window on the screen. See Get-Help Out-GridView for more information.

    Get-InsteonDevice | Show-ISYObject -formatgrid -title 'Insteon Devices'

    This opens an interactive gridview window with the title 'Insteon Devices'

#>
function Show-ISYObject {
[CmdletBinding(DefaultParameterSetName='list')]
[OutputType([object])]

param(
# the array of xmlelements produced by the various get cmdlets in the ISY994i module 
[Parameter(ValueFromPipeline=$true,Position=0,Mandatory=$true)] 
[System.XML.XMLElement[]]$objtoshow,

# uses the Format-List cmdlet to display the information. This is the default format.
[Parameter(ParameterSetName='list')]
[alias("fl")]
[switch]$formatlist,

# uses the Format-Table cmdlet to display the information.
[Parameter(ParameterSetName='table',Mandatory=$true)]
[alias("ft")]
[switch]$formatable,

# sends output to the Out-Gridview cmdlet which presents data in an interactive window
[Parameter(ParameterSetName='grid',Mandatory=$true)]
[alias("fg")]
[switch]$formatgrid,

# the title parameter passed to the Out-GridView cmdlet. See Get-Help Out-GridView -parameter title for more info.
[Parameter(ParameterSetName='grid')]
[string]$title,

# the wait parameter passed to the Out-GridView cmdlet. See Get-Help Out-GridView -parameter wait for more info.
[Parameter(ParameterSetName='grid')]
[switch]$wait,

# the outputmode parameter passed to the Out-GridView cmdlet. See Get-Help Out-GridView -parameter outputmode for more info.
[Parameter(ParameterSetName='grid')]
[ValidateSet('multiple','none','single')]
[string]$outputmode,

# the passthru parameter passed to the Out-GridView cmdlet. See Get-Help Out-GridView -parameter passthru for more info.
[Parameter(ParameterSetName='grid')]
[switch]$passthru
)

BEGIN 
{   
    $devpropnames = 'name','address','OnLevel%','OnLevel','category',
                    'fullpath','enabled','type','flag','deviceClass','dcPeriod',
                    'wattage','ELK_ID','pnode','parent','property'

    $scenepropnames = 'name','address','fullpath','flag','deviceGroup',
                      'ELK_ID','family','members','parent'

    $memberspropnames = 'Scene_Name','MemberName','type','OnLevel%','OnLevel',
                        'Scene_fullpath','Scene_address','Scene_flag','Scene_family'

    $programpropnames = 'parent','name','folder','id','status','enabled','runAtStartup',
                        'running','lastRunTime','lastFinishTime','nextScheduledRunTime'

    function Get-SortedProps {
    [cmdletbinding()]
    param([System.Xml.XmlElement[]]$xelem,
          [string[]]$sortedprops)

    $proplist = New-Object System.Collections.ArrayList
    $xelem | ForEach-Object {
            $currentxelem = $_
            $currentpropnames = ($currentxelem | Get-Member -MemberType Properties).Name 
            $culled = $sortedprops | where {$_ -in $currentpropnames} | 
                where { $($_ -ne $currentxelem.$_.tostring()) } | 
                    where { $($currentxelem.$_.tostring() -ne '') }
            foreach ($entry in $culled)
            {
                if ($proplist.Count -eq 0)
                {
                    [void]$proplist.Add($entry)
                    $nextinsert = 1
                }
                elseif (!$proplist.Contains($entry))
                {
                    $proplist.Insert($nextinsert, $entry)
                    $nextinsert++
                }
                else
                {
                    $nextinsert = ($proplist.IndexOf($entry) + 1)
                }
                
            }
            $nextinsert = 1
    
     }
    return $proplist

    }

}

END 
{ 
if ($PSCmdlet.MyInvocation.ExpectingInput) 
{
    $temparray = @($input)
}
else 
{
    $temparray = $objtoshow
}

    $allmethodnames = ($temparray | Get-Member -MemberType Method -Force).name
    $allpropnames   = ($temparray | Get-Member -MemberType Properties).name

    if (!(('get_LocalName' -in $allmethodnames) -or ('localname' -in $allpropnames)))  
        {write-error -Message "This Object is not supported" -Category InvalidArgument;return}

    $localname = $temparray.localname | Sort-Object -Unique
    if ($localname.count -ne 1) {write-error -Message "localname not unique." -Category InvalidArgument;return}

    if ($PSCmdlet.ParameterSetName -eq 'grid')
    {
        [void]$PSCmdlet.MyInvocation.BoundParameters.Remove('formatgrid')
        [void]$PSCmdlet.MyInvocation.BoundParameters.Remove('objtoshow')                                                      
        $gridsplathash = $PSCmdlet.MyInvocation.BoundParameters
        write-verbose -Message "gridsplathash below:`n$($gridsplathash.keys.ForEach({'{0} = {1}' -f $_,$gridsplathash[$_]}))"
    }

    write-verbose "`nLocalname = $localname`n"

    switch ($localname)
    {
        'node'    
        {
            $propnames = Get-SortedProps -xelem $temparray -sortedprops $devpropnames
                    
            if ($formatable)
            {
                $temparray | Format-Table $propnames -AutoSize
            }
            elseif ($formatgrid)
            { 
                $temparray | select-object -Property $propnames | Out-GridView @gridsplathash
            }
            else
            {
                $temparray | select-object -Property $devpropnames | format-list
            }
       }

        'group'   
        {
            $upperpropnames = Get-SortedProps -xelem $temparray -sortedprops $scenepropnames
            Write-Verbose -Message "upperpropnames = $($upperpropnames -join ',')"
            $lowerpropnames = Get-SortedProps -xelem $temparray.members.link -sortedprops $memberspropnames
            Write-Verbose -Message "lowerpropnames = $($lowerpropnames -join ',')"

            if ($formatable)
            {
                $temparray.members.link | Format-Table $lowerpropnames
            }
            elseif ($formatgrid)
            {
                $temparray.members.link | Select-Object -Property $lowerpropnames | Out-GridView @gridsplathash
            }
            else
            {
                $formatted = $temparray.foreach( {
                                '-'*60;$_ | Format-List -Property $upperpropnames
                                $_.members.link | Format-Table ($lowerpropnames | where {$_ -notmatch '^Scene_'}) -AutoSize
                                               } )
                ($formatted | out-string ) -split ([char]13 + "`n" ) | where {$_ -ne ''} | 
                        ForEach-Object {$_.replace("MemberName","`nMemberName")}                       
            } 
         }

        'folder'  
        {
            if ($formatable)
            {
                $temparray | Sort-Object -Property name | Format-Table -AutoSize
            }
            elseif ($formatgrid)
            { 
                $temparray | Sort-Object -Property name | Out-GridView @gridsplathash
            }
            else
            {
                $temparray | Sort-Object -Property name | format-list
            }
         }  

        'program' 
        {
            if ($formatable)
            {
                $temparray | Sort-Object -Property parent | Format-Table $programpropnames -AutoSize
            }
            elseif ($formatgrid)
            { 
                $temparray | Select-Object $programpropnames | Sort-Object -Property parent | Out-GridView @gridsplathash
            }
            else
            {
                $temparray | Select-Object $programpropnames | Sort-Object -Property parent | format-list
            }
        } 

        'configuration' 
        {
            if ($formatable)
            {
                $temparray | Format-Table -AutoSize
            }
            elseif ($formatgrid)
            { 
                $temparray | Out-GridView @gridsplathash
            }
            else
            {
                $temparray | format-list
            }
        }

        'DT' 
        {
            if ($formatable)
            {
                $temparray | Format-Table -AutoSize
            }
            elseif ($formatgrid)
            { 
                $temparray | Out-GridView @gridsplathash
            }
            else
            {
                $temparray | format-list
            }
        }

        'RestResponse' 
        {
            if ($formatable)
            {
                $temparray | Format-Table -AutoSize
            }
            elseif ($formatgrid)
            { 
                $temparray | Out-GridView @gridsplathash
            }
            else
            {
                $temparray | format-list
            }
        }

        default   
        {
            write-error -Message "localName type $localname is unknown." -Category NotImplemented
            return
        }
    }
}

}
#end function Show-ISYObject

<#
.Synopsis
   This function unsubscribes an event subscriber from the ISY. 
.DESCRIPTION
   Sends an unsubscribe SOAP request to the ISY for the subscriber specified by the -SID parameter.
   The SID can be obtained from the Get-ISYSubscription function or by examining the events SID property.
   
   The -RC parameter is used if the caller wants to receive the UDIDefaultResponse of a successful operation. 
   Without this switch parameter nothing is returned for a successful execution.
.EXAMPLE
    Unregister-FromISYEvent -SID uuid:65 -RC

    status info
    ------ ----
    200    n/a

    Shows the success response.  
#>
function Unregister-FromISYEvent {
[cmdletbinding()]
param(
# unsubscribe to this existing SID 
[Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true)]
[string]$SID,

# 0 to error if not subscribed, 1 to be quiet.
[ValidateRange(0,1)]
[int]$flag = 0,

# if set the function outputs the result code associated with the webrequest sent to the ISY. Normally a successful operation returns nothing to the caller.
[switch]$RC
)

if ($SID -notmatch '^uuid:')
{
    $SID = 'uuid:' + $SID
}

Write-Verbose -Message "SID = $SID"

$headers = Get-Headers -SOAPAction "urn:udi-com:device:X_Insteon_Lighting_Service:1#UDIService"
write-verbose -Message "auth = $($headers['Authorization'])"
Write-Verbose -Message "soapaction = $($headers['SOAPAction'])"

$body = "<?xml version='1.0' encoding='utf-8'?>"
$body += "<s:Envelope><s:Body><u:Unsubscribe"
$body += " xmlns:u='urn:udi-com:service:X_Insteon_Lighting_Service:1'>"
$body += "<SID>$SID</SID><flag>$flag</flag></u:Unsubscribe></s:Body></s:Envelope>`r`n" 

Write-Verbose -Message "Body:`n$body`n"

$uri = Get-ISYUrl -path '/services'
Write-Verbose -Message "the uri = $uri"

$splat = @{Uri = $uri
           DisableKeepAlive = $true
           Method = "POST"
           ContentType = 'text/xml; charset=utf-8'
           Headers = $headers
           Body = $body
            }

$response = Invoke-WebRequest @splat

if (!$RC)
{
    $status = ([xml]($response.Content)).Envelope.Body.UDIDefaultResponse.Status
    if ($status -ne '200')
    {
        Write-Error -Message "Something went wrong. Status = $status" -Category InvalidOperation
    }
}
else
{
    ([xml]($response.Content)).Envelope.Body.UDIDefaultResponse
}

}
# end of function Unregister-FromISYEvent

<#
.Synopsis
   Update-ISYDBXMLFile

   This function produces a new ISYDB.xml file when "required".
.DESCRIPTION
  If the -refreshnow pamameter is used or there is no ISYDB.xml or if the file is older than the 
  refreshtime time set in the ISYSettings.xml file then the a new ISYDB.xml file is created.
  Otherwise the ISYDB.xml file in the ISYSettings.path folder is left untouched.
.EXAMPLE
    The following executes a REST query to the ISY and writes the new ISYDB.xml file to the ISYSettings.path folder.
    
    > Update-ISYDBXMLFile -refreshnow
#>
function Update-ISYDBXMLFile {
    [CmdletBinding()]
    [OutputType([void])]
    Param([switch]$refreshnow)

    Write-Verbose -Message ($MyInvocation.MyCommand.Name + ' called.') 

    $currentsettings = Get-ISYSettings
    if (!(Test-Path $currentsettings.path)) {New-Item  -itemtype directory -path $currentsettings.path}
    $fullname = Join-Path $currentsettings.path $Script:ISYDBXMLFileName

#region create database and vars
    # can replace (Get-Date) - (Get-ChildItem $fullname).LastWriteTime) -gt [timespan]($currentsettings.refreshtime)
    #        with New-Timespan -Start (Get-ChildItem $fullname).LastWriteTime -End (get-date) in powershell 4.
    if ($refreshnow -or !(Test-Path $fullname) -or 
             (((Get-Date) - (Get-ChildItem $fullname).LastWriteTime) -gt [timespan]($currentsettings.refreshtime))) 
    {
        $isydb = Invoke-IsyRestMethod "/rest/nodes" -ErrorAction Stop
        #change all addresses and types to names for parent and member fields
        $typehash = @{'0' = '0';'1' = 'NODE'; '2' = 'GROUP';'3' = 'FOLDER';'32' = 'Responder';'16' = 'Controller'}
        #create a global address to name hash
        $faddrtoname = @{}
        $isydb.nodes.folder | select name,address | ForEach-Object {$faddrtoname[$_.address] = $_.name}

        $daddrtoname = @{}
        $isydb.nodes.node | select name,address | ForEach-Object {$daddrtoname[$_.address] = $_.name}

        $saddrtoname = @{}
        $isydb.nodes.group | select name,address | ForEach-Object {$saddrtoname[$_.address] = $_.name}

        $fdsaddrtoname = $faddrtoname + $daddrtoname + $saddrtoname  
         
        #scriptblock to make the changes to friendly values from addresses
        $changerblock = {if ($_.'#text') 
                            {$friendlyval = [string]$fdsaddrtoname[$_.'#text']
                            if ($friendlyval)
                                {$_.'#text' = $friendlyval}
                            }
                        if ($_.type)
                            {$friendlyval = [string]$typehash[$_.type]
                            if ($friendlyval)
                                {$_.type = $friendlyval}                                                 
                            }
                        if ($_.pnode)
                            {$friendlyval = [string]$fdsaddrtoname[$_.pnode]
                            if ($friendlyval)
                                {$_.pnode = $friendlyval}                                                 
                            }
                        }

        # update parent, type, members.link and pnode fields
        foreach ($region in '//folder', '//node', '//group')
        {
            $item = Select-XML -Xml $isydb -XPath $region
            $item.node.parent | ForEach-Object $changerblock

            if ($region -eq '//group')
                {
                    $item.node.members.link | ForEach-Object $changerblock

                    $item.node.members.link | ForEach-Object {$_.setattribute('MemberName', $_.'#text')} 
                }
            if ($region -eq '//node')
                {
                    $item.node | ForEach-Object $changerblock

                    $item.node | ForEach-Object {
                                    $cattype = Resolve-InsteonType -type $_.type -ErrorAction silentlycontinue
                                    $friendly_type = $cattype.type
                                    if ($friendly_type) {$_.type = $friendly_type;$_.setattribute('category',$cattype.category)}
                                }

                }
        }

        $rootdirname = $isydb.nodes.root
        # add the full folder path to each folder object
        foreach ($folder in $isydb.nodes.folder)
        {
            $temp  = $folder
            $patharray = New-Object System.Collections.ArrayList
            [void]$patharray.add($folder.name)
            while ($temp.parent)
            { 
                $patharray.Insert(0, $temp.parent.'#text')
                $temp = $isydb.nodes.folder  | where {$_.name -eq $temp.parent.'#text'}
            }
            $fullpath = $patharray -join '/'
            $folder.SetAttribute('fullpath', "$rootdirname/$fullpath")   
        }

        #Add fullpath field to devices and scenes
        foreach ($isyobj in $isydb.nodes.node, $isydb.nodes.group)
        {
            switch ($isyobj)
            {
                {$_.parent.type -eq 'folder'} 
                {
                    $current = $_
                    $fullpath = ($isydb.nodes.folder | where {$_.name -eq $current.parent.'#text'}).fullpath
                    $current.SetAttribute("fullpath", $fullpath)
                    continue
                }

                {$_.parent.type -eq 'node'}   
                {
                    $current = $_
                    # assumes a parent node must have folder as a parent (all node grandparents are folders)
                    $parentfoldername = ($isyobj | where {$_.name -eq $current.parent.'#text'}).parent.'#text'
                    $fullpath = ($isydb.nodes.folder | where {$_.name -eq $parentfoldername}).fullpath
                    $current.SetAttribute('fullpath', $fullpath)
                    continue
                }

                {(!($_.parent))}
                {   
                    $_.SetAttribute('fullpath', $rootdirname)
                    continue
                }

                Default 
                {
                    Write-Error -Message "wtf? - no such parent type $($_.parent.type) in $($_.name)" -Category InvalidType
                }
            }
        }

        $isydb.save($fullname)

        # this will create the $Script:list_of_isy_programs variable and update the ISYPrograms.xml file if required
        Get-ISYProgramList -fast | Out-Null

        $Script:list_of_insteon_devices = ($isydb.nodes.node | Select-Object -Property name,LocalName,address) + 
                                            [pscustomobject]@{name = 'PLM';LocalName = 'node';address = (Get-PLMAddress)}

        $Script:list_of_isy_scenes = $isydb.nodes.group | Select-Object -Property name,LocalName,address
#endregion create database and vars

#region tabexpansion

#region tabexpansion scriptblocks

        $Completion_DeviceName = {
 
        param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)

        function quote {param([string]$text) if ($text -like '* *') {$text = "'" + $text + "'"};return $text}

        Get-InsteonDevice -dbmap | Sort-Object name | where name -like $wordToComplete* |
            ForEach-Object {
                New-Object System.Management.Automation.CompletionResult (quote $_.name), (quote $_.name), 'ParameterValue', ('{0} ({1})' -f $_.name, $_.address)}
        }

        $Completion_SceneName = {
 
        param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)

        function quote {param([string]$text) if ($text -like '* *') {$text = "'" + $text + "'"};return $text}

        Get-ISYScene -dbmap | Sort-Object name | where name -like $wordToComplete* |
            ForEach-Object {
                New-Object System.Management.Automation.CompletionResult (quote $_.name), (quote $_.name), 'ParameterValue', ('{0} ({1})' -f $_.name, $_.address)}
        }

        $Completion_ProgramName = {
 
        param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)

        function quote {param([string]$text) if ($text -like '* *') {$text = "'" + $text + "'"};return $text}

        Get-ISYprogramList -fast | Sort-Object name | where name -like $wordToComplete* |
                ForEach-Object {
                    New-Object System.Management.Automation.CompletionResult (quote $_.name), (quote $_.name), 'ParameterValue', ('{0} ({1})' -f $_.name, $_.id)}
        }

        $Completion_X10Device = {
 
        param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)

        function quote {param([string]$text) if ($text -like '* *') {$text = "'" + $text + "'"};return $text}

        Get-InsteonDevice -dbmap | Sort-Object name | where name -like $wordToComplete* | where name  -match "-X[A-P](([1-9]$)|(1[0-6]$))" |
            ForEach-Object {
                New-Object System.Management.Automation.CompletionResult (quote $_.name), (quote $_.name), 'ParameterValue', ('{0} ({1})' -f $_.name, $_.address)}
        }

#endregion tabexpansion scriptblocks

        if ($currentsettings.tabexp -eq $true)
        {
            Register-ArgumentCompleter -CommandName Set-InsteonDevice,Get-InsteonDevice -ParameterName Name -ScriptBlock $Completion_DeviceName
            Register-ArgumentCompleter -CommandName Set-ISYScene,Get-ISYScene -ParameterName Name -ScriptBlock $Completion_SceneName
            Register-ArgumentCompleter -CommandName Invoke-ISYProgram -ParameterName Name -ScriptBlock $Completion_ProgramName
            Register-ArgumentCompleter -CommandName Set-X10Device -ParameterName Name -ScriptBlock $Completion_X10Device
        }
#endregion tabexpansion    

    }

}
#end function Update-ISYDBXMLFile
