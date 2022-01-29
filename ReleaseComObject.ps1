function Get-SubRange{
    [CmdletBinding(DefaultParameterSetName='ByName')]
    param(
        [Parameter(Mandatory=$true,Position=0)]
        $ParentRange,
        [Parameter(Mandatory=$true,Position=1,ParameterSetName='ByName')]
        [String]$RangeName,
        [Parameter(Mandatory=$true,Position=1,ParameterSetName='ByArea')]
        [int[]]$RangeArea,
        [Parameter(Mandatory=$true,Position=2)]
        [ref]$SubRange
    )
    if ($PSCmdlet.ParameterSetName -eq "ByName"){
        $SubRange.Value = $ParentRange.Range($RangeName)
    } else {
        $SubRange.Value = $ParentRange.Range(($ParentRange.Cells[$RangeArea[0],$RangeArea[1]]),($ParentRange.Cells[$RangeArea[2],$RangeArea[3]]))
    }
}

$xl = New-Object -ComObject Excel.Application
$wkb = $xl.workbooks.open((Get-Item ./test.xlsx).FullName)
$wks = $wkb.worksheets.item(1)
$rng = $null
Get-SubRange -ParentRange $wks -RangeArea (1,1,3,3) -SubRange ([ref]$rng)
Get-SubRange -ParentRange $wks -RangeName "data" -SubRange ([ref]$rng)

for ($i = 1; $i -le 3; $i++) {
    for ($j = 1; $j -le 3; $j++) {
        Write-Output ('{0},{1}: {2}' -f $i,$j,($rng.Cells[$i,$j]).Text)
    }
}

$wkb.close($false)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($rng) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wks) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wkb) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

#不可见窗口没有MainWindowsHandle
#$xl.Visible = $true
#$xlhwnd = $xl.hwnd
#$xlps = Get-Process excel | Where-Object {$_.MainWindowHandle -eq $xlhwnd}
#$xlpid = $xlps.Id
#Stop-Process -Id $xlpid

Get-Process -ErrorAction SilentlyContinue -Name Excel
