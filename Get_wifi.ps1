$profiles = netsh wlan show profiles
$wifiProfiles = @()

$i = 0
$total = $profiles -match "所有用户配置文件 : (.*)" | measure | select -expand Count

foreach ($profile in $profiles) {
    if ($profile -match "所有用户配置文件 : (.*)") {
        $wifiName = $matches[1]
        $profileDetail = netsh wlan show profile name="$wifiName" key=clear
        if ($profileDetail) {
            $match = $profileDetail | Select-String "关键内容"
            if ($match) {
                $password = $match.ToString().Split(':')[1].Trim()
                $wifiProfile = [PSCustomObject]@{
                    "Wi-Fi 名称" = $wifiName
                    "密码" = $password
                }
                $wifiProfiles += $wifiProfile
            } else {
                $wifiProfile = [PSCustomObject]@{
                    "Wi-Fi 名称" = $wifiName
                    "密码" = "未找到密码"
                }
                $wifiProfiles += $wifiProfile
            }
        }
    }
    $i++
    if ($total -eq 0) {
        Write-Progress -Activity "正在导出 Wi-Fi 配置信息" -Status "正在处理 $wifiName" -PercentComplete 0
    } else {
        $percentComplete = [Math]::Min([Math]::Floor(($i/$total)*100), 100)
        Write-Progress -Activity "正在导出 Wi-Fi 配置信息" -Status "正在处理 $wifiName" -PercentComplete $percentComplete
    }
}

$wifiProfiles | Export-Csv -Path "D:\wifi_profiles.csv" -Encoding UTF8 -NoTypeInformation

Write-Progress -Activity "正在导出 Wi-Fi 配置信息" -Completed