## 特色是遠端抓取資料庫, 設定相對應的檔案名稱寫入檔名
## 沒有安裝PDF也能使用
## 遠端資料庫的API要自己設計

cls
# Set the execution policy
Set-ExecutionPolicy Unrestricted -Scope Process

# Check for NuGet registered
if (-not (Get-PackageSource -ProviderName Nuget -ErrorAction SilentlyContinue)) {
    try {
        Register-PackageSource -Name MyNuGet -Location https://www.nuget.org/api/v2 -ProviderName NuGet -Trusted -ErrorAction Stop
        Write-Host "NuGet Source 'MyNuGet' registered successfully."
    } catch {
        Write-Host "Failed to registered 'MyNuGet': $_"
    }
} else {
    Write-Host "'MyNuGet' already exists, no need to register again."
    Write-Host ""
}

# Check for PDFsharp was installed or not
if (-not (Get-Package -Name PDFSharp-gdi -ErrorAction SilentlyContinue)) {
    try {
        Install-Package -Name PDFSharp-gdi -RequiredVersion 1.50.5147 -ProviderName NuGet -Scope CurrentUser -ErrorAction Stop
        Write-Host "PDFsharp installed successfully."
    } catch {
        Write-Host "Failed to install PDFsharp: $_"
    }
} else {
    Write-Host "PDFsharp is already installed."
    Write-Host ""
}

# Import PDFsharp
$pdfsharppackage = Get-Package -Name "PDFSharp-gdi" -ErrorAction Stop
$pdfSharpFolder = $pdfsharppackage.Source.Substring(0, $pdfsharppackage.Source.LastIndexOf("\")) 
$pdfsharppath = (ls -Path $pdfSharpFolder -Recurse -Include "PdfSharp-gdi.dll" -ErrorAction SilentlyContinue).FullName
Add-Type -LiteralPath $pdfsharppath
Write-Host "PDFSharp library successfully imported"
Write-Host ""

#Select shopCode
 $shopCode = Read-Host "請輸入店舖代號, 例如：B666"

#From destination search data dictionary(換行是因為用PS2exe轉換時, 程式碼太長好像會有問題?)
 $API = "https://script.google.com/macros/s/AKfycbypy6X" `
 + "asurIDCode/exec"

 $JSONBody = @{
       shopCode = $shopCode 
       } | ConvertTo-Json

 $scriptBlock = {
   param($API, $JSONBody)
   IRM -Uri $API -Method Post -Body $JSONBody -ContentType "application/json"
 }

 $Job = Start-Job -ScriptBlock $scriptBlock -ArgumentList $API, $JSONBody
    $i = 0
        while ($Job.State -eq "Running"){
           Write-Host "`r處理中...$i 秒" -NoNewline
           Start-Sleep -Seconds 1
           $i++
        } 

    $response = Receive-Job -Job $job
    Write-Host ""
    Write-Host $response
    Remove-Job -Job $Job
  Write-Host ""
  write-host "店鋪名稱查詢結果: $($response.storeName)"

  $storeName = $response.storeName



Write-Host ""

# Import System.Windows.Forms for file dialog
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# Open file dialog to select a PDF file
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "PDF Files (*.pdf)|*.pdf"
$openFileDialog.Title = "Select a PDF file"

# Show the file dialog
if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $selectedFile = $openFileDialog.FileName
    Write-Host "File selected: $selectedFile"
    Write-Host ""

    # Load the selected PDF file
    $pdfDocument = [PdfSharp.Pdf.IO.PdfReader]::Open($selectedFile, [PdfSharp.Pdf.IO.PdfDocumentOpenMode]::Import)
    Write-Host "PDF loaded successfully"
    Write-Host ""

    # Check if the PDF has at least two pages
    if ($pdfDocument.PageCount -lt 2) {
        Write-Host "The selected PDF does not have enough pages (minimum 2 required)."
    } else {
        # Create new PDF document for the first two pages
        $newPdfDocument = New-Object PdfSharp.Pdf.PdfDocument

        # Add the first two pages to the new document
        $newPdfDocument.Pages.Add($pdfDocument.Pages[0])
        $newPdfDocument.Pages.Add($pdfDocument.Pages[1])

        # Save the new PDF document with the first two pages
        $newFilePath = [System.IO.Path]::Combine($pwd, $storeName + "-臨場服務附表八(店鋪版).pdf")
        $newPdfDocument.Save($newFilePath)
        Write-Host "Pages 1 and 2 combined and saved as '$newFilePath'"

        # Copy the original PDF to a new location and rename it
        $fullCopyPath = [System.IO.Path]::Combine($pwd, $storeName + "-臨場服務附表八(完整版).pdf")
        Copy-Item -Path $selectedFile -Destination $fullCopyPath
        Write-Host "Original PDF copied and renamed to '$fullCopyPath'"
    }

    # Clean up
    $newPdfDocument.Close()
} else {
    Write-Host "No file selected."
}
Write-Host ""
Read-Host "資料處理完成, 請按任意鍵關閉視窗"
