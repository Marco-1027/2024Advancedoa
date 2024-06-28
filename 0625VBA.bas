Attribute VB_Name = "Module1"
Sub FirstVBAHelp()
'
' FirstVBAHelp 巨集
'
' 快速鍵: Ctrl+q

   '點選A1儲存格
    Range("A1").Select
   '點選A1起始最末列CTRL+往下
    Selection.End(xlDown).Select
    '彈跳視窗文字要用雙引號包起來 Row的意思是列集合的索引
    MsgBox "報告主播,列數有" & Selection.End(xlDown).Rows.Row
    '點選最末欄CTRL+往右
    Selection.End(xlToRight).Select
    '彈跳視窗抓最右欄,欄索引 文字要用雙引號包起來 Couumns是欄集合的索引
    MsgBox "報告主播,欄數有" & Selection.End(xlToRight).Columns.Column
    '執行完畢,游標回到A1
    Range("A1").Select
    
End Sub
