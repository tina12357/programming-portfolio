Sub pdf轉檔()

    Dim source_path_name As String
    Dim target_path As String
    Dim 表 As Worksheet
    Dim sName As String
    Dim source_file_name As String

    ' Get file name without extension
    ' 取得檔名（不含副檔名）
    source_file_name = Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
    
    ' Get file path
    ' 取得檔案路徑
    source_path_name = ActiveWorkbook.Path
    
    ' Set output folder
    ' 設定輸出資料夾
    target_path = source_path_name & "\" & source_file_name
    
    ' Create folder if not exists
    ' 若資料夾不存在則建立
    On Error Resume Next
    MkDir target_path
    On Error GoTo 0

    ' Loop through all worksheets
    ' 逐一處理每個工作表
    For Each 表 In Worksheets
        sName = target_path & "\" & source_file_name & "_" & 表.Name & ".pdf"
        
        表.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=sName
    Next

    MsgBox "PDF export completed! / 已完成 PDF 匯出"

End Sub
