#############################################################################
#
# Script created by Imnoss Limited
# Last Updated: 2021-10-27
# Please feel free to share while retaining attribution
#
# For full instructions on how to use please visit:
# https://imnoss.com/batch-protect-pdfs-with-different-passwords/
#
#############################################################################


# If you are viewing this in the PowerShell ISE, highlight the row below and press F8 to 
# enable the script to be run after saving.
Set-ExecutionPolicy Unrestricted -Scope Process


# clear the blue part of the screen (makes it easier to see what is output from the script and what is the script
Clear-Host

Write-Host @"

Script created by Imnoss Limited

please visit https://imnoss.com/batch-protect-pdfs-with-different-passwords/ for more information

Many thanks to empira Software GmbH (http://www.pdfsharp.net) who created the PdfSharp library and published it as Open Source software

Starting PDF Password protection script...

"@



# First try installing the PDFSharp library
try
{

    # Check for NuGet being registered as a package source, and register if not
    if ( -not (Get-PackageSource -ProviderName NuGet -ErrorAction SilentlyContinue))
    { Register-PackageSource -Name MyNuGet -Location https://www.nuget.org/api/v2 -ProviderName NuGet -Trusted -ErrorAction Stop}

    # Check for the PDFSharp library being installed, install if not
    if ( -not (Get-Package -Name "PDFSharp-gdi" -ErrorAction SilentlyContinue))
    { Install-Package -Name PDFSharp-gdi -RequiredVersion 1.50.5147 -ProviderName NuGet -Scope CurrentUser -ErrorAction Stop}

    # The package should be installed, so try grabbing it, then use the package file path to identify the desired DLL file
    $pdfSharpPackage = Get-Package -Name "PDFSharp-gdi" -ErrorAction Stop
    $pdfSharpFolder = $pdfSharpPackage.Source.Substring(0, $pdfSharpPackage.Source.LastIndexOf("\")) 
    $pdfSharpPath = (Get-ChildItem -Path $pdfSharpFolder -Recurse -Include "PdfSharp-gdi.dll" -ErrorAction SilentlyContinue).FullName

    Add-Type -LiteralPath $pdfSharpPath

    Write-Host "PDFSharp library successfully imported"
    Write-Host ""
    
}
catch
{
    $errorType = if ($_.InvocationInfo.InvocationName) {$_.InvocationInfo.InvocationName} else {$_.Exception.Message}
    $errorCode = switch($errorType)
                    {
                        "Register-PackageSource" {"Nuget could not be registered as a source"; Break}
                        "Install-Package" {"The PDFSharp package could not be installed"; Break}
                        "Get-Package" {"The PDFSharp package was not installed correctly"; Break}
                        "Cannot index into a null array." {"Could not find DLL library"; Break}
                        "Add-Type" {"The PDFSharp package could not be added as a library"; Break}
                        default {$_.Exception.Message}
                    }
    Write-Host "An error occurred: $errorCode"
	Read-Host -Prompt "PDF files cannot be password protected, press Enter to exit script"

    Break
}


try
{
    # add in the system open file dialog library so the dialog can be displayed
    Add-Type -AssemblyName System.Windows.Forms
    
    # set the initial directory for the open file dialog
    # "special folders" like the desktop or documnets folders can be specified, the full list of special folders
    # can be found here: https://docs.microsoft.com/en-us/dotnet/api/system.environment.specialfolder
    # alternatively this can be changed to a hard coded path, e.g. $InitialDirectory = "D:\Letters"
    $InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
    #InitialDirectory = $InitialDirectory 

    # set up the open file dialog with a 
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog `
            -Property @{ Title = "Please select your list of PDF files and passwords" 
                         Filter = "CSV/Text files (*.csv;*.txt)|*.csv;*.txt|All files (*.*)|*.*"
                         FilterIndex = 1 }
    $button = $FileBrowser.ShowDialog()

    if (-not ($FileBrowser.FileName)) {Throw "File dialog cancelled"}
    
###################################################################################
#這個程式留了一個難關, 以下補充將BIG5轉成UTF8
#因為本身程式沒辦法轉換,用 .NET Framework 這是微軟提供的開源資料庫, 可以直接使用, so eazy 
###################################################################################
    $originalFilePath = $FileBrowser.FileName

    #用.NET 上面的System.Text.Encoding 來讀取和轉換編碼
    $big5Encoding = [System.Text.Encoding]::GetEncoding("big5")
    $utf8ENcoding = [System.Text.Encoding]::UTF8

 $big5Content = [System.IO.File]::ReadAllText($originalFilePath, $big5Encoding)

 #讀取完後, 換成UTF8重新寫入原檔案
 [System.IO.File]::WriteAllText($originalFilePath, $big5Content, $utf8ENcoding)

 $renewfile = $originalFilePath
#####################################################################

    # read the CSV file into variable $csv
    $csv = Import-Csv -Path $renewfile -Delimiter "," -Encoding UTF8

    # check for the key columns of 'FilePath' and 'Password', if not found throw an error
    if(-not($csv | Get-Member -Name "FilePath")) {throw "FilePath"}
    if(-not($csv | Get-Member -Name "Password")) {throw "Password"}

    # create a file path for the results - this is the same as the input CSV file path, except that the file name ends with
    # 'Results' followed by the date and time the script is run - this should ensure a unique results file each time the
    # script is run, and avoid problems updating the input CSV which may occur if the file is open in e.g. Excel
    $outputCsv = "$($renewfile.Substring(0, $renewfile.LastIndexOf('.'))) Results $(Get-Date -Format "yyyy-MM-dd hh.mm.ss").csv"

    Write-Host "CSV file read"
    Write-Host ""
}

catch
{
    $errorType = if ($_.InvocationInfo.InvocationName) {$_.InvocationInfo.InvocationName} else {$_.Exception.Message}
    $errorCode = switch($errorType)
                    {
                        "Add-Type" {"Could not add FileDialog library"; Break}
                        "New-Object" {"Could not create a FileDialog to select CSV"; Break}
                        "File dialog cancelled" {"File dialog cancelled"; Break}
                        "Import-Csv" {"The CSV could not be imported, it may be locked"; Break}
                        "FilePath" {"Could not find the column 'FilePath' in the CSV file"; Break}
                        "Password" {"Could not find the column 'Password' in the CSV file"; Break}
                        default {$_.Exception.Message}                   
                    }
    #Write-Host "An error occurred while running $($_.InvocationInfo.InvocationName): $($_.Exception.Message)"
    Write-Host "An error occurred: $errorCode"
	Read-Host -Prompt "PDF files cannot be password protected, press Enter to exit script"
    Break
}




# initialise some tracking numbers
$fileCount = 0
$successCount = 0


# loop through each row ($r) in the csv.
foreach ($r in $csv)
{

    # increment the $fileCount
    $fileCount += 1

    # wrap everything in a try block
    # the PDFSharp library produces errors in its own format, so we will parse any error codes in the catch block and then output
    # to the console and also to the results file
    try
    {
        # check if the PDF file exists, throw an error if not
        if(-not(Test-Path $r.FilePath)) { throw "Error: `"PDF file does not exist`""}

        # open the PDF in PdfSharp
        $document = [PdfSharp.Pdf.IO.PdfReader]::Open($r.FilePath)
        $securitySettings = $document.SecuritySettings;

        # Set the UserPassword (password to open), it is possible to set the owner password which prevents e.g. editing, but we have not done that here
        $securitySettings.UserPassword = $r.Password

        # save the pdf file
        $document.Save($r.FilePath)

        # update $result to reflect success and increment the $successCount
        $result = "Success"
        $successCount += 1
    }

    catch
    {
        # extract from the error message from PDFSharp the bit after the colon
        # also clean up, removing the double quotes, new lines, and the full file path if present
        $result = "Error: " + $_.Exception.Message.Substring($_.Exception.Message.IndexOf(":") + 2).Replace("`n","").Replace("`"","").Replace(" '$($r.FilePath)'", '')
    }

    # write the row to $outputCsv file
    $r | Select-Object -Property "FilePath" | Add-Member -MemberType NoteProperty -Name "Result" -Value $result -PassThru | Export-Csv $outputCsv -Delimiter "," -Append -NoTypeInformation

    # Output the the result and a new line to the consol window
    # Output the file name of the PDF being protected to the consol window
    Write-Host "Protecting file: $($r.FilePath.Substring($r.FilePath.LastIndexOf('\') + 1))  |  $($result)"
}

Write-Host ""
Write-Host "$successCount of $fileCount files have been password protected"
Write-Host "The full results can be found here: $outputCsv"
Write-Host ""
Read-Host -Prompt "Press Enter to finish"


