Option Explicit
Option Base 1

Sub Main()

    Call startmacro
    Call deleteBlankCell
    Call chkString
    Call endmacro
    MsgBox "処理完了"

End Sub

Sub deleteBlankCell()
'空白セルの削除をしつつ、B列にすべての値を持っていく

    Dim aMaxRow As Long: aMaxRow = Cells(Rows.Count, 1).End(xlUp).Row
    Dim bMaxRow As Long: bMaxRow = Cells(Rows.Count, 2).End(xlUp).Row
    Dim MaxRow As Long
    Dim LoopColCnt As Long
    
    Worksheets(2).Activate
    
    '空白セルを削除する
    With Range(Cells(2, 1), Cells(aMaxRow, 11))
        .SpecialCells(xlCellTypeBlanks).Select
        Selection.Delete Shift:=xlUp
    End With
    
    Range("B1").Select
    bMaxRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    '一列にする
    For LoopColCnt = 3 To 11
        Cells(1, LoopColCnt).ClearContents '列名的なのはいらないので削除する
        
        MaxRow = Cells(Rows.Count, LoopColCnt).End(xlUp).Row
        Range(Cells(2, LoopColCnt), Cells(MaxRow, LoopColCnt)).Select
        Selection.Cut Destination:=Cells(bMaxRow + 1, 2) '切り取り
        
        bMaxRow = Cells(Rows.Count, 2).End(xlUp).Row 'B列の最終行再取得
    Next LoopColCnt
    
    Range("B1").Select '最後の切り取り範囲が選択されたままだと気持ち悪いので…

End Sub

Sub chkString()
'重複削除+文字列にスペースが有ればスペースを除去、ひらがなが含まれればC列に「ひらがなあるよ」を記載

    'Long
    Dim MaxRow As Long
    Dim LoopCnt As Long 'カウンタ変数
    Dim arrCnt As Long: arrCnt = 1 '配列用カウンタ変数
    
    'String
    Dim chkStr As String
    Dim msg As String
    
    'Array
    Dim chkArr() As Variant '判定した文字列を配列に入れる
    Dim hanteiArr() As Variant  '判定結果を配列に格納する
    
    MaxRow = Cells(Rows.Count, 2).End(xlUp).Row '最終行取得
    Range(Cells(1, 2), Cells(MaxRow, 2)).RemoveDuplicates Columns:=1, Header:=xlYes '重複削除
    

    MaxRow = Cells(Rows.Count, 2).End(xlUp).Row '最終行再取得
    chkArr() = Range(Cells(2, 2), Cells(MaxRow, 2)) '配列に入れる
    
    For LoopCnt = 1 To UBound(chkArr)
        chkStr = chkArr(LoopCnt, 1)
        If hasSpace(chkStr) = True Then 'スペースあるか判定
            chkStr = Replace(chkStr, " ", "") '半角スペースとる
            chkStr = Replace(chkStr, "　", "") '全角スペースとる
        End If
        
        If hasHiragana(chkStr) = True Then 'ひらがな判定
            ReDim Preserve hanteiArr(arrCnt)
            hanteiArr(arrCnt) = "ひらがなあるよ"
            arrCnt = arrCnt + 1
        Else
            ReDim Preserve hanteiArr(arrCnt)
            hanteiArr(arrCnt) = ""
            arrCnt = arrCnt + 1
        End If
    Next LoopCnt
    
    Cells(1, 3).Value = "ひらがな判定"
    Range(Cells(2, 3), Cells(UBound(hanteiArr) + 1, 3)) = WorksheetFunction.Transpose(hanteiArr) '判定結果をセルに転記
    
    Range("A:A").ClearContents
    Range("A:A").NumberFormatLocal = "0_ " 'セルの表示形式を数値にする
    
    Cells(1, 1).Value = "No."
    Cells(2, 1).Value = "1"
    Cells(2, 1).AutoFill Range(Cells(2, 1), Cells(UBound(chkArr) + 1, 1)), xlFillSeries '連番を降る
    
    '最後の手入れ
    Columns("A:C").AutoFit
    Range("A1").AutoFilter

End Sub
