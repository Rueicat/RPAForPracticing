########Section 1
cls
Set-executionPolicy Unrestricted -Scope Process

$scriptPath = ".\"
$newFolderName = Read-Host "請輸入臨場日期(例如：0416)"
$newFolderPath = Join-Path -Path $scriptPath -ChildPath $newFolderName

if (-Not (Test-Path -Path $newFolderPath)){
   New-Item -Path $newFolderPath -ItemType Directory
}else {
    Write-Host "資料夾已存在於： $scriptPath"
}

#####Section 2
write-host ""
write-host ""
$storeCode1 = Read-Host "請輸入第一個店鋪代號(格式為 BXXX)"
$storeCode2 = Read-Host "請輸入第二個店鋪代號(格式為 BXXX)"

##############Progress bar, use "start-Job" technique, 新技術, 重要, 背景作業6.12.2024
###增添部分會註記

  
 

###From ur ShopCodeSearch
$API = "https://script.google.com/macros/s/AKfycbyDlYdGCRokpubbOgB4c_QAGJ_" `
+ "setByurself/exec"

$Body = @{
  storeCode1 = $storeCode1
  storeCode2 = $storeCode2
} | ConvertTo-Json


#背景作業進行API請求
$job = Start-Job -ScriptBlock {
    param($API, $Body)
    Invoke-RestMethod -Uri $API -Method `
    Post -Body $Body -ContentType "application/json"
} -ArgumentList $API, $Body


###initialize the progress bar
 # 初始化進度條
$activity = "正在向 API 發送請求"
$status = "請等待服務器回應..."
$percentComplete = 0

# 當背景作業運行時，更新進度條
while ($job.State -eq "Running") {
    Write-Progress -Activity $activity -Status $status -PercentComplete $percentComplete
    Start-Sleep -Seconds 1
    $percentComplete += 5
    if ($percentComplete -ge 100) {
        $percentComplete = 0
    }
}

# 背景作業完成後獲取結果
$response = Receive-Job -Job $job

# 清除進度條並顯示完成狀態
Write-Progress -Activity $activity -Status "已收到回應" -Completed

# 清理作業
Remove-Job -Job $job

### 這麼留著可以測試回應狀況
#### Write-Host "API Response: $($response | ConvertTo-Json)"


$storeName1 = $response.storeName1
$storeName2 = $response.storeName2
$downloadurl1 = $response.downloadurl1
$downloadurl2 = $response.downloadurl2

Write-Host "@
 **店鋪名稱查詢結果：
   $storeName1
    -改善單下載連結： $downloadurl1
   $storeName2
    -改善單下載連結： $downloadurl1
@"


########Section 3

$storeFolder1 = Join-Path -Path $newFolderPath -ChildPath $storeName1
$storeFolder2 = Join-Path -Path $newFolderPath -ChildPath $storeName2

New-Item -Path $storeFolder1 -ItemType directory -Force
New-Item -Path $storeFolder2 -ItemType directory -Force

$phtoFolder1 = Join-Path -Path $storeFolder1 -ChildPath ("$storeName1" + "_照片")
$phtoFolder2 = Join-Path -Path $storeFolder2 -ChildPath ("$storeName2" + "_照片")

New-Item -Path $phtoFolder1 -ItemType Directory -Force
New-Item -Path $phtoFolder2 -ItemType Directory -Force

Invoke-WebRequest -Uri $downloadurl1 -OutFile `
(Join-Path -Path $storeFolder1 -ChildPath "$storeName1-臨場服務改善提醒單.docx")
Invoke-WebRequest -Uri $downloadurl2 -OutFile `
(Join-Path -Path $storeFolder2 -ChildPath "$storeName2-臨場服務改善提醒單.docx")

Write-Host ""
Read-Host "資料套組已設置好, 請按任意鍵關閉"