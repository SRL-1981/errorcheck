Attribute VB_Name = "Module1"
Sub errorcheck()

  Dim path As String
  path = "C:\Users\User\Documents\Macro\Sales.xlsx"

  Dim wbthis As Workbook
  Set wbthis = ThisWorkbook
  
'異常データ出力用シートとして既存の「エラーデータ」シートを定義
  Dim wserror As Worksheet
  Set wserror = wbthis.Worksheets("エラーデータ")

'「エラーデータ」シートのシート内容をリセット
  wserror.Range("a2:c" & Rows.Count).ClearContents

'異常データ検出対象ファイルの存在確認
On Error Resume Next

  Dim wbsales As Workbook
  Set wbsales = Workbooks.Open(path)
  
On Error GoTo 0

If wbsales Is Nothing Then

  Debug.Print "対象ファイルが見つかりません。"
  Exit Sub
  
Else

  Debug.Print "「Sales」ファイル取得"
  
End If


On Error Resume Next

'異常データ検出対象シート
  Dim wsdata As Worksheet
  Set wsdata = wbsales.Worksheets("データ")
  
On Error GoTo 0

If wsdata Is Nothing Then

  Debug.Print "対象シートが見つかりません。"
  Exit Sub
  
Else

  Debug.Print "「データ」シート取得"
  
End If


'対象シート最終行取得
  Dim datalastrow As Long
  datalastrow = wsdata.Cells(wsdata.Rows.Count, 1).End(xlUp).Row
  
'対象シートのデータをヘッダーを除いたデータを配列へ格納
  Dim dataarr As Variant
  dataarr = wsdata.Range("a2:d" & datalastrow).Value
  
'辞書変数定義
  Dim iddict As Object
  Set iddict = CreateObject("scripting.dictionary")
  
  Dim i As Long
  
  Dim j As Long
  j = 2
  
  
'格納した配列からループで1行ずつ順番に取り出し
'エラーの有無をチェック～エラー該当行の出力
For i = 1 To UBound(dataarr, 1)

'定義した辞書を用いてID重複の有無を判定
  If Not iddict.exists(dataarr(i, 1)) Then
  
    iddict.Add dataarr(i, 1), True
    
  Else
  
    wserror.Cells(j, 1).Value = dataarr(i, 1)
    wserror.Cells(j, 2).Value = dataarr(i, 2)
    wserror.Cells(j, 3).Value = dataarr(i, 1)
    wserror.Cells(j, 4).Value = "ID重複"
    
    j = j + 1
    
  End If
  
  If dataarr(i, 3) <= 0 Or dataarr(i, 3) = "" Then
    
    wserror.Cells(j, 1).Value = dataarr(i, 1)
    wserror.Cells(j, 2).Value = dataarr(i, 2)
    wserror.Cells(j, 3).Value = dataarr(i, 3)
    wserror.Cells(j, 4).Value = "金額値異常"
    
    j = j + 1
      
  End If
  
  If IsDate(dataarr(i, 4)) = False Then
      
    wserror.Cells(j, 1).Value = dataarr(i, 1)
    wserror.Cells(j, 2).Value = dataarr(i, 2)
    wserror.Cells(j, 3).Value = dataarr(i, 4)
    wserror.Cells(j, 4).Value = "日付異常"

    j = j + 1
        
  End If
      

Next i


If Not wbsales Is Nothing Then

  wbsales.Close Savechanges = False
  
End If


End Sub
