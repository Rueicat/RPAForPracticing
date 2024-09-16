
Sub 抓取最新檔案清單()
   
'異常處理
    On Error GoTo ErrorMsg
   
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("處理工作區(更改檔名)")
       ws.Cells.Clear
    Set ws = ThisWorkbook.Worksheets("檔案路徑(加密碼)")
       ws.Cells.Clear
'關閉畫面更新
    Application.ScreenUpdating = False
   
   Dim fd As FileDialog
   Dim folderPath As String
   Dim i As Integer
   Dim fileName As String
   Dim purefile As String
   Dim lastRow As Integer


   
'設定標題
 Worksheets("處理工作區(更改檔名)").Range("A1").Value = "更改檔名"
 Worksheets("處理工作區(更改檔名)").Range("B1").Value = "原始檔名"
 Worksheets("處理工作區(更改檔名)").Range("C1").Value = "實際資料夾顯示狀況"
 Worksheets("檔案路徑(加密碼)").Range("A1").Value = "FilePath"
 Worksheets("檔案路徑(加密碼)").Range("B1").Value = "Password"
     '顏色格式設定
         With Worksheets("處理工作區(更改檔名)").Range("A1:C1")
               .Interior.Color = RGB(255, 222, 213)
               .Font.Color = RGB(0, 0, 0)
               .Font.Name = "Microsoft JhengHei"
         End With
          With Worksheets("檔案路徑(加密碼)").Range("A1:B1")
               .Interior.Color = RGB(201, 203, 241)
               .Font.Color = RGB(0, 0, 0)
               .Font.Name = "Microsoft JhengHei"
         End With
        
   
'跳出視窗選擇資料夾
  Set fd = Application.FileDialog(msoFileDialogFolderPicker)
      With fd
            .Title = "請選擇, 想要處理的PDF, 所在的資料夾"
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
           Exit Sub
        End If
     End With
  
folderPath = folderPath & "\"

'找PDF file
   fileName = Dir(folderPath & "*.pdf")
   
     i = 1
     
       Do While fileName <> ""
                                purefile = Left(fileName, InStrRev(fileName, ".") - 1)
                                
           Worksheets("處理工作區(更改檔名)").Cells(i + 1, 3).Value = fileName
           Worksheets("檔案路徑(加密碼)").Cells(i + 1, 1).Value = folderPath & fileName
           Worksheets("處理工作區(更改檔名)").Cells(i + 1, 2).Value = purefile
           fileName = Dir()
           i = i + 1
        Loop
        
        If i = 1 Then
           MsgBox "在這個資料夾中沒有找到PDF檔案"
        End If
        
        
'設定column width
    Worksheets("處理工作區(更改檔名)").Columns("A:A").ColumnWidth = 30
    Worksheets("處理工作區(更改檔名)").Columns("B:C").AutoFit
    Worksheets("檔案路徑(加密碼)").Columns("A:B").AutoFit
    
'這個sheet的column C, 從下面數上來碰到
lastRow = Worksheets("處理工作區(更改檔名)").Cells(Worksheets("處理工作區(更改檔名)").Rows.Count, "C").End(xlUp).Row
    
    
'設定格線
    With Worksheets("處理工作區(更改檔名)").Range("A1:C" & lastRow)
               .Borders.LineStyle = xlContinuous
               .Borders.Weight = xlMedium
    End With
    
    With Worksheets("檔案路徑(加密碼)").Range("A1:B" & lastRow)
               .Borders.LineStyle = xlContinuous
               .Borders.Weight = xlMedium
    End With

'格行上色
    Dim colorStep As Integer
    
    For colorStep = 3 To lastRow Step 2
       Worksheets("處理工作區(更改檔名)").Range("A" & colorStep & ":C" & colorStep).Interior.Color = RGB(255, 247, 214)
       Worksheets("檔案路徑(加密碼)").Range("A" & colorStep & ":B" & colorStep).Interior.Color = RGB(255, 247, 214)
    Next
    
'開啟畫面更新
 Application.ScreenUpdating = True
    
    Exit Sub
ErrorMsg:
    MsgBox "發生錯誤, 請記下: 1.操作過程, 2.拍照顯示的錯誤訊息--給銳銳" & Err.Number & " 錯誤說明:" & Err.Description
    
    
End Sub


Sub 更改檔名()

'異常處理
    On Error GoTo ErrorMsg
    
'關閉畫面更新
    Application.ScreenUpdating = False
    
    Dim wsRename As Worksheet
    Dim wsPath As Worksheet
    Dim i As Long
    Dim oldPath As String
    Dim newPath As String
    Dim newFileName As String
    
Set wsRename = Worksheets("處理工作區(更改檔名)")
Set wsPath = Worksheets("檔案路徑(加密碼)")

lastRow = wsPath.Cells(wsPath.Rows.Count, "A").End(xlUp).Row

      For i = 2 To lastRow
        newFileName = wsRename.Cells(i, 1).Value    '新檔名
        
        oldPath = wsPath.Cells(i, 1).Value   '舊路徑
        newPath = Left(oldPath, InStrRev(oldPath, "\") - 1) & "\" & newFileName & ".pdf"
        
       
'同樣檔名不能執行, 會錯誤, 所以要跳過
      If Dir(newPath) = "" Then
          If newFileName <> "" And Dir(oldPath) <> "" Then
            Name oldPath As newPath
          End If
      Else
          MsgBox "檔案名稱 " & newFileName & " 已存在於目標資料夾中。", vbExclamation
      End If
      
    
      Next i

MsgBox "檔案重新命名完成~", vbInformation
    
   Exit Sub
ErrorMsg:
    MsgBox "發生錯誤, 請記下: 1.操作過程, 2.拍照顯示的錯誤訊息--給銳銳--" & Err.Number & " 錯誤說明:" & Err.Description
'開啟畫面更新
 Application.ScreenUpdating = True

End Sub

Sub 輸出CSV檔()

     Dim ws As Worksheet
     Dim tempWorkbook As Workbook
     Dim csvFilePath As String
     
Set ws = ThisWorkbook.Worksheets("檔案路徑(加密碼)")

csvFilePath = ThisWorkbook.Path & "\檔案路徑(加密碼)-請不要打開-編碼會錯誤.csv"


'CSV特殊格式設定
      ws.Copy
Set tempWorkbook = ActiveWorkbook

     Application.DisplayAlerts = False    '因為保存CSV問題會很多, 先關密警告
     
     'format 直接編碼UTF-8 只適用於2016以後的版本, 如果可以的話, 直接修改成即可FileFormat:=xlCSVUTF8
     tempWorkbook.SaveAs fileName:=csvFilePath, FileFormat:=xlCSV, local:=True
     tempWorkbook.Close False
     
     Application.DisplayAlerts = True


   MsgBox "CSV 文件已成功保存至: " & csvFilePath, vbInformation


End Sub

Sub 清除資料()
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("處理工作區(更改檔名)")
       ws.Cells.Clear
    Set ws = ThisWorkbook.Worksheets("檔案路徑(加密碼)")
       ws.Cells.Clear
End Sub
