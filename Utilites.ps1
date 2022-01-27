function Get-PixelCircle {
   param(
       [parameter(Mandatory=$true)]
        [int] $diameter)
    [double]$r = 0
    [double]$x = 0
    [int]$y = 0
    [double]$offSetX = 0
    [double]$offSetY = 0
    [int]$tmpY = 0
    [int]$nBlock =0
    if ($diameter % 2 -eq 0){
        $offSetX = 0.5
    }else{
        $offSetY = 0.5
    }
    $r = $diameter/2
    $tmpY = $r + $offSetY
    $x = $offSetX
    while ($x -le $r)
    {
        $y = [math]::Pow($r * $r - $x * $x,0.5) + $offSetY
        if ($tmpY -ne $y){
            $tmpY = $y
            Write-Output $nBlock
            $nBlock = 1
        }else{
            $nBlock ++
        }
        if($y -le $x){return}
        $x ++
    }
}

function Convert-StrToUnicode {
    param(
        [string]$str = ""
    )
    [string]$ustr = ""
    for($i = 0; $i -lt $str.Length; $i++) {
        $int = [int]$str[$i] 
        if($int -gt 32 -and $int -lt 127) {
            $ustr += $str[$i]
        }
    else{
        $ustr += ("\u{0:x4}" -f $int)
    }
}
    $ustr
}

#
function Get-WanIP {
    (Invoke-WebRequest -Uri 'http://www.net.cn/static/customercare/yourip.asp' `
        | Select-String -Pattern "<h2>(\d+\.\d+\.\d+\.\d+)</h2>" `
        | Select-Object -ExpandProperty Matches).Groups[1].Value
}

#
function Get-YoutubeThumbnail {
    param(
        [parameter(Mandatory=$true)]
        [string]$url)
    [string]$key = ""
    if (($url -match "watch\?v=(.{11})") -or ($url -match "be/(.{11})")) {
        $key = $Matches.item(1)
        $ret = "https://i.ytimg.com/vi/$key/maxresdefault.jpg"
        Set-Clipboard $ret
	return $ret
    }
}

function ipInfo($ip) {
    if ($ip -notmatch "^(?:\d{1,3}\.){3}\d{1,3}$") {
        Write-Output "ip格式不对"
        return
    }
    $cmd = "ipinfo.io/{0}{1}" -f $ip,"?token=the token"
    #curl -x $proxy $cmd
    curl $cmd
}

# Get directory or file size.
function du {
    param(
        [parameter(Mandatory=$true)]
	    [string]$path
    )
    if(-not (Test-Path $path)){
        Write-Host "Path not exits!" -ForegroundColor Yellow
        return
    }
    $thePath = Convert-Path $path
    $theItem = Get-Item $thePath
    $theSize = 0
    $unit = ''
    if($theItem.PSIsContainer){
        $theSize = (Get-ChildItem $thePath -Recurse | Measure-Object -Sum Length).Sum + 0
    }else {
        $theSize = $theItem.Length
    }
    if($theSize -lt 1mb){
        $theSize =$theSize/1kb
        $unit = 'kb'
    }elseif($theSize -lt 1gb){
        $theSize = $theSize/1mb
        $unit = 'Mb'
    }else{
        $theSize = $theSize/1gb
        $unit = 'Gb'
    }
    $theSize = [System.Math]::Round($theSize,2)
    return ('{0}{1}' -f $theSize,$unit)
}
