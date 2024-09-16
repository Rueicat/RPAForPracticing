cls
Set-ExecutionPolicy Unrestricted -Scope Process

# Check ImportExcel has been installed or not
If ( -not (Get-Package -Name ImportExcel -ErrorAction SilentlyContinue)){
     Install-Module -Name ImportExcel -Scope CurrentUser
     Write-Host "ImportExcel has installed"
}else{
  Write-Host "ImportExcel has been installed already"
}

# 創建和配置OpenFileDialog
Add-Type -AssemblyName System.Windows.Forms

$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog `
   -Property @{
        InitialDirectory = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)
        Filter = "Excel files (*.xlsx)|*.xlsx"
        Title = "請選擇人事資料的excel檔案(不要選其他excel喔= =)"
   }


$result = $openFileDialog.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $filePath = $openFileDialog.FileName

    
    # 獲取新檔案名
    $directory = [System.IO.Path]::GetDirectoryName($filePath)
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($filePath)
    $newFilePath = [System.IO.Path]::Combine($directory, "$fileName-已修改.xlsx")
    
    

    # 嘗試打開並解鎖Excel文件, open-ExcelPackage 的方法是開啟excel 做處理, 
    # 下面會用Import-Excel 提取數據檔案, 容易轉換成JSON 格式傳送利用Webapp
    # 因為沒有GCP帳號, 不能直接上傳檔案到google drive, 所以透過提取資料傳輸(Using JSON convert the data) 

    #處理文件(Open-ExcelPackage)
  try {
        $excelPackage = Open-ExcelPackage -Path $filePath -Password "如果有密碼的話"

       
          $data = $excelPackage.Workbook.Worksheets["在職&留停"]
        #renew data
          $data.cells["P1"].Value = "年齡"
          $data.cells["X1"].Value = "生日"
          $data.cells["Y1"].Value = "年資"
          $data.cells["Z1"].Value = "到職日期"

      
         #-ne = "Not Equal"
        
        $row = 2
        while ($data.Cells["A$row"].Value -ne $null) {
             
          #Set formula first
          #更正, 不要用formula, 直接在powershell 內計算, 不透過excel處理-----好方法
          ###yyMMdd 有分大小寫, M means Month, m means minutes
####################################################################################
           ### 先轉換必要資料寫入
            $originalbirth = $data.Cells["N$row"].Value
            $originalgoMOS = $data.Cells["O$row"].Value
            #用powershell 解析格式化成日期
             $parsebirth = [datetime]::ParseExact($originalbirth, "yyyyMMdd", $null)
             $parsegoMOS = [datetime]::ParseExact($originalgoMOS, "yyyyMMdd", $null)

             $formattedbirth = $parsebirth.ToString("yyyy/MM/dd")
             $formattedgoMOS = $parsegoMOS.ToString("yyyy/MM/dd")

             $data.Cells["X$row"].Value = $formattedbirth
             $data.Cells["Z$row"].Value = $formattedgoMOS
           ### 利用上面資料接著下面的處理
             $today = [datetime]::Today
             $age = [math]::Truncate(($today - $parsebirth).TotalDays / 365.25)
             $data.cells["P$row"].Value = $age

            <# 有點複雜的年資計算, 
                年-年, 月減月
                當日小於到職日, 月份要-1
                月份差異負值的話, 年份-1, 月份+12
            #>
            $yearOfService = $today.Year - $parsegoMOS.Year
            $monthOfservice = $today.Month - $parsegoMOS.Month
            if ($today.Day -lt $parsegoMOS.Day){
                $monthOfservice--
               }
            if ($monthOfservice -lt 0){
                 $yearOfService--
                 $monthOfservice += 12            
               }

            $yearAndMonth = "${yearOfservice}年${monthOfservice}月"
            $data.Cells["Y$row"].Value = $yearAndMonth
######################################################################
           
            $row++
        }

###我是分隔線#####

        # 儲存並關閉Excel文件
        $excelPackage.SaveAs($newFilePath)
        $excelPackage.Dispose()
        Write-Host "Excel文件已成功開啟並處理，並另存為：$newFilePath"

      
    ###Convert to JASON data and upload to Webapp
        $finalData = Import-Excel -Path $newFilePath -Password "SM00" -WorksheetName '在職&留停'
        $leaveData = Import-Excel -Path $newFilePath -Password "SM00" -WorksheetName '請假單' -StartRow 4
        $Interval = Open-ExcelPackage -Path $newFilePath -Password "SM00"
        $Intervaldata = $Interval.Workbook.Worksheets['請假單'].Cells['A3'].Text
    ###-Depth 參數不需要, 我們的資料只有一層, 預設好像會抓到2層
        $JASONData = $finalData
        $JASONfileName = $fileName

    ###-Compress 很重要, 避免轉換到base64時錯誤
        $JasonObj = @{
            data = $JASONData
            fileName = $fileName
            leaveData = $leaveData
            Interval = $Intervaldata
          } | ConvertTo-Json -Compress

       
    ##轉換編碼, 中文字傳輸後顯示不出來..7/10
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($JasonObj)

    $base64EncodingData = [System.Convert]::ToBase64String($bytes)


  Write-Host ""
   
  #upload Data
   $webAppurl = "選擇已不屬的google drive API程式連結"
    
       try{
             $scriptBlock = {
                param($webAppurl, $base64EncodingData)
            Invoke-RestMethod -Uri $webAppurl -Method Post -Body $base64EncodingData -ContentType "application/json"
             }
      $job = Start-Job -ScriptBlock $scriptBlock -ArgumentList $webAppurl, $base64EncodingData
       $i = 0
         while ($job.State -eq "Running"){
             
             Write-Host "`r資料上傳中... $i 秒" -NoNewline          
             Start-Sleep -Seconds 1
             $i++
         }
           $response = Receive-Job -Job $job
           Write-Host ""
           Write-Host $response
           Remove-Job -Job $job
         
       }catch{
         Write-host "Error: $($_.Exception.Message)"
        }
    
    

    


  }catch {
        Write-Host "錯誤訊息：$($_.Exception.Message)"
    }
} else {
    Write-Host "沒有選擇文件。"
 }

Write-Host ""
Read-Host "請按任意鍵關閉視窗"