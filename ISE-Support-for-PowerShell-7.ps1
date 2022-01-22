# Copy into ISE's Profile
# The Menu 'Add-ons' will add to options

#$ErrorActionPreference= 'silentlycontinue'
$psISE.CurrentPowerShellTab.AddOnsMenu.Submenus.Clear()
$psISE.CurrentPowerShellTab.AddOnsMenu.Submenus.Add("Switch to PowerShell 7", { 
    function New-OutOfProcRunspace {
        param($ProcessId)
        $ci = New-Object -TypeName System.Management.Automation.Runspaces.NamedPipeConnectionInfo -ArgumentList @($ProcessId)
        $tt = [System.Management.Automation.Runspaces.TypeTable]::LoadDefaultTypeFiles()
        $Runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($ci, $Host, $tt)
        $Runspace.Open()
        $Runspace
    }  
    $PowerShell = Start-Process PWSH -ArgumentList @("-NoExit") -PassThru -WindowStyle Hidden
    $Runspace = New-OutOfProcRunspace -ProcessId $PowerShell.Id
    $Host.PushRunspace($Runspace)
}, "ALT+F5") | Out-Null
  
$psISE.CurrentPowerShellTab.AddOnsMenu.Submenus.Add("Switch to PowerShell 5", { 
    $Host.PopRunspace()
    $ChildProc = Get-CimInstance -ClassName win32_process | where {$_.ParentProcessId -eq $Pid}
    $ChildProc | ForEach-Object { Stop-Process -Id $_.ProcessId }
}, "ALT+F6") | Out-Null

