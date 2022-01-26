function script:moveData {
    param(
        $theTarget,
        $theSource,
        [int[]]$targetCols,
        [int[]]$sourceCols,
        [ref]$theRow
    )
    $r = $theRow.Value
    $start = $r
    for ($i = 1; $i -le $theSource.Rows.Count; $i++) {
        if ($theSource.Cells[$i,1].Text -ne "") {
            $theTarget.Cells[$r,1].Value = "$r"
            for ($j = 0;$j -lt $targetCols.Count; $j++) {
                $theTarget.Cells[$r,$targetCols[$j]].Value = $theSource.Cells[$i,$sourceCols[$j]].Text
            }
            $r++
        }
    }
    $theRow.Value = $r
    $end = $r - 1
    $info = ("  序号:{0}...{1}" -f $start, $end)
    Add-Content -Path info.txt -Value $info 
    Write-Host $info -ForegroundColor DarkYellow
}

function script:IsNameInWorkbook {
    param(
        [Parameter(Mandatory=$true,Position=0)]
            $Workbook,
        [string]$AreaName = "data"
    )
    $ret = $false
    $n = $Workbook.Names.Count
    if ($n -gt 0){
        for ($i = 1; $i -le $n; $i++) {
            if ($Workbook.Names[$i].Name -eq $AreaName){
                $ret = $true
                break
            }
        }
    }
    $ret
}

#铸件和自制件入库记录
function Get-InboundRecord {   
    param([string]$Type="")
    [string]$p="铸件"
    [string]$cangku="zhujianruku"
    [string]$PartNo=""
    [string]$sql=""

    if ($Type -eq 'zzj'){
        $cangku = "waixiezhujianruku"
        $p = "锻件"
    }

    while ($true)
    {
        Write-Host $p -NoNewline 
        $PartNo = Read-Host -Prompt "零件号"
        if ($PartNo -eq 'exit')
        {
            return
        }else
        {
            $PartNo = ($PartNo -replace '"','\"')
            $sql = "SELECT r.*, d.miao_shu_cn, d.cai_liao, d.ke_hu FROM $cangku r INNER JOIN lingjianqingdan d ON r.ling_jian_wu_hao = d.ling_jian_wu_hao WHERE r.ling_jian_wu_hao = '$PartNo' ORDER BY r.id DESC LIMIT 10;"
            mysql --login-path=hsw -e $sql hsw
        }
    }
}

function Merge-ShipList {
    if (-not (Test-Path .\list.txt)) {
        Write-Host -Object "当前目录下没有目录文件'list.txt'" -ForegroundColor Red
        return 
    }

    $files = Get-Content -Path .\list.txt
    if ($files.Length -gt 0) {
        if (-not ($files[0] -match "铸件领用.xls$")) {
            Write-Host -Object "第一个文件名不是'*铸件领用.xls'" -ForegroundColor Red
            return
        }
        if (-not (Test-Path $files[0])) {
            Write-Host -Object ("文件'{0}'不存在" -f $files[0]) -ForegroundColor Red
            return
        }
    }

    $wkbSource = $null
    $rngSource = $null
    $nextRow = 1
    $success = $true

    $xlAPP = New-Object -ComObject 'Excel.Application'
    $wkbTarget =  $xlAPP.WorkBooks.Open($files[0])
    $rngTarget = $wkbTarget.Worksheets[1].Cells[4,1]

    Clear-Content -ErrorAction SilentlyContinue info.txt

    for ($i = 1; $i -lt $files.Length; $i++) {
        $file = $files[$i]
        if (Test-Path $file) {
            $wkbSource =  $xlAPP.WorkBooks.Open($file)
            Add-Content -Path info.txt -Value $wkbSource.Name
            "打开文件:" + $wkbSource.Name
            if (IsNameInWorkbook $wkbSource) {
                $rngSource = $wkbSource.Worksheets[1].Range("data")
                script:moveData $rngTarget $rngSource (2,3,4) (1,4,6) ([ref]$nextRow)
                $wkbSource.Close($false)
            }else{
                Write-Host "----没有定义数据块'data'"
                $success = $false
                $wkbSource.Close($false)
                break
            }
        }else{
            Write-Host -Object "文件'$file'不存在" -ForegroundColor Red
            $success = $false
            break
        }
    }

    if ($success) {
        '共合并了 ' + ($nextRow - 1) + ' 条数据'
        $wkbTarget.Save()
    }else{
        "`n合并失败退出...`n"
    }
    $wkbTarget.Close($false)

    Get-Process -ErrorAction SilentlyContinue -Name Excel | Where-Object {$_.MainWindowHandle -eq 0} | Stop-Process
}

function Get-ShipList {
    if (-not (Test-Path "发货清单")) {
        Write-Host "没有文件夹'发货清单'"
        return
    }
    #Clear-Content -ErrorAction SilentlyContinue list.txt
    $content = $(
        (Get-ChildItem -ErrorAction SilentlyContinue -Path *铸件领用.xls).FullName
        (Get-ChildItem -ErrorAction SilentlyContinue -Path 发货清单/*).FullName
        (Get-ChildItem -ErrorAction SilentlyContinue -Path tmp/萧山发票.xlsx).FullName
    )
    Set-Content -Value $content -Path list.txt
}

function Fetch-ShipList {
    [string]$baseName = (Get-Item ..).BaseName
    if ($baseName -match '(\d{4})-(\d{1,2})') {
        $year = $Matches[1].Substring(2)
        [int]$month = $Matches[2]
        $sPath = "\\sunch\铸件领料单\{0}年{1}月" -f $year,$month
        if (-not (Test-Path $sPath)) {
            Write-Host "目标文件夹 '$sPath' 不存在." -ForegroundColor Yellow
            return
        }
    } else {
        Write-Host "当前文件夹名字不符合模式'1970-01'."
        return
    }

    $dPath = Get-Location
    Write-Output $sPath

    $sFiles = Get-ChildItem -Path $sPath
    if ($sFiles -eq $null) {
        Write-Host "源文件为空"
        return
    }

    $dFiles = Get-ChildItem -Path $dPath
    if ($dFiles -eq $null) {
        $sFiles | foreach {
            Copy-Item -Path $_.FullName -Destination $dPath
            Write-Host ('    {0}' -f $_.BaseName)
        }
        return
    }

    $diffs = Compare-Object -ReferenceObject $sFiles -DifferenceObject $dFiles -Property Name
    if ($diffs -eq $null) {
        return
    }

    $diffs | foreach {
        $copyParam = @{
            'Path' = '{0}\{1}' -f $sPath, $_.Name
        }
        if ($_.SideIndicator -eq '<=') {
            $copyParam.Destination = $dPath
            Copy-Item @copyParam
            Write-Host ('    {0}' -f (Split-Path $copyParam.Path -Leaf))
        }
    }
}
