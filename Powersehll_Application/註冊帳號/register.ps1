cls
Write-Host @"
-註冊帳號與密碼, 取得工具使用權
-Movement and script by Ruei 2024.5.17
-Upgraded to Version 2.0.0 on 2024.8.13
---------------------------------------
"@
Write-Host ""
Set-ExecutionPolicy Unrestricted -Scope Process

#SHA512 function
function Convert-secureStringToSHA512 {
  param (
    [parameter(Mandatory=$true)]
    [System.Security.SecureString]$SecurePassword
  )

  if ($saltedResponse) {
      $Marshal = [System.Runtime.InteropServices.Marshal]
      $BSTR = $Marshal::SecureStringToBSTR($passWord)
      $passwordBytes = [System.Text.Encoding]::UTF8.GetBytes($Marshal::PtrToStringAuto($BSTR))
      $Marshal::ZeroFreeBSTR($BSTR)

      #recover saltTobytes
      $saltbytes = [Convert]::FromBase64String($saltedResponse)

      $conbined = $passwordBytes + $saltbytes

      $sha512 = [System.Security.Cryptography.SHA512Managed]::new()
      $hashcompared = $sha512.ComputeHash($conbined)

      $hash64combined = [Convert]::ToBase64String($hashcompared)
      return $hash64combined

  } else {
    $Marshal = [System.Runtime.InteropServices.Marshal]

  $BSTR = $Marshal::SecureStringToBSTR($SecurePassword)
  $passwordBytes = [System.Text.Encoding]::UTF8.GetBytes($Marshal::PtrToStringAuto($BSTR))
  $Marshal::ZeroFreeBSTR($BSTR)

  $salt = [System.Text.Encoding]::UTF8.GetBytes([System.Guid]::NewGuid().Tostring())
  $saltBase64 = [Convert]::ToBase64String($salt)

  $saltedPasswordBytes = $passwordBytes + $salt

  $SHA512 = [System.Security.Cryptography.SHA512Managed]::new()
  $Hash = $SHA512.ComputeHash($saltedPasswordBytes)

  $HashBase64 = [Convert]::ToBase64String($Hash)

  $H512All = New-Object PSObject -Property @{
      HashBase64o = $HashBase64
      SaltBase64o = $saltBase64
  }
  return $H512All
  } 
}

###3 of functions as below

Function register {
$username = Read-Host "請輸入註冊帳號"
$SecurePassword = Read-Host "請輸入密碼" -AsSecureString
$HashResult = Convert-secureStringToSHA512 -SecurePassword $SecurePassword
$email = Read-Host "請輸入E-mail"

Write-Host ""
Write-Host "資料處理中..."
Write-Host ""

$type = $resultaction
  $body = @{
    type = $type
    username = $username
    password = $HashResult.HashBase64o
    salt = $HashResult.SaltBase64o
    useremail = $email
  } | ConvertTo-Json

$Uri = "改成你的部屬API"

$response = Invoke-RestMethod -Uri $Uri -Body $body -Method Post -ContentType "application/json"
Write-Host $response
Write-Host ""
Read-Host -Prompt "請按任意鍵結束程式"
}
#------------------------------------#
Function change {}
#------------------------------------#
Function validate {
$username = Read-Host "請輸入註冊帳號"
Write-Host ""
Write-Host "搜尋帳號..."

$UriFetchSalt = "改成你的部屬API"

$Body = @{
   type = "getSalt"
   username = $username
} | ConvertTo-Json

##想寫在同一個doPost裡
$saltedResponse = Invoke-RestMethod -Uri $UriFetchSalt -Body $Body -Method Post -ContentType "application/json"

  if ($saltedResponse -eq '未找到用戶') {
      Write-Host "鹽值: 未找到帳戶"
      Read-Host -Prompt "請按任意鍵關閉程式"
      return
  } else {
    Write-Host "鹽值:$saltedResponse"
    Write-Host ""
    $passWord = Read-Host "請輸入密碼" -AsSecureString
    $HashResult = Convert-secureStringToSHA512 -SecurePassword $passWord
    Write-Host "比對中..."

  $Cbody = @{
    type = "validate"
    username = $username
    password = $HashResult
 } | ConvertTo-Json

   $Ccomparekey = Invoke-RestMethod -Uri $UriFetchSalt -Body $Cbody -Method Post -ContentType "application/json"
        Write-Host $Ccomparekey
        Write-Host ""
       Read-Host -Prompt "請按任意鍵結束程式"
  }
}
#------------------------------------#

###select function
$action = $null
$resultaction = $null
$saltedResponse = $null

while ($action -ne '1' -and $action -ne '2' -and $action -ne '3'){
  $action = Read-Host @"
    請選擇操作功能:
      1- 註冊新帳號
      2- 更改密碼
      3- 驗證密碼
請輸入
"@
  Switch ($action){
    '1' {
      $resultaction = 'register'
      register
    }
    '2' {
      $resultaction = 'change'
      change
    }
    '3' {
      $resultaction = 'validate'
      validate
    }
    default{
      Write-Host "無效輸入, 請重新輸入"
      $action = $null
    }
  }

}