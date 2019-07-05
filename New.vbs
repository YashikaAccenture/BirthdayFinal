strComputer = "."

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & _
        strComputer & "\root\cimv2")
Set colMonitoredEvents = objWMIService.ExecNotificationQuery _
    ("SELECT * FROM __InstanceCreationEvent WITHIN 10 WHERE " _
        & "Targetinstance ISA 'CIM_DirectoryContainsFile' and " _
            & "TargetInstance.GroupComponent= " _
                & "'Win32_Directory.Name=""c:\\\\Users\\\\asha.chauhan\\\\Desktop\\\\WishingTool\\\\Birthday""'")
Do
    Set objLatestEvent = colMonitoredEvents.NextEvent
    call DoSomething
	
Loop

Sub DoSomething
 dim shell
set shell=createobject("wscript.shell")
shell.run "C:\Users\asha.chauhan\Desktop\WishingTool\Birthday\gitpush.bat" 
set shell=nothing
end sub