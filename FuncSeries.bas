Option Explicit

Sub startmacro()
'マクロ開始前のおまじない
  Application.ScreenUpdating = False 'Stop Updating Screen
  Application.EnableEvents = False 'Stop Events
  Application.Calculation = xlCalculationManual 'Caluculation Manually
End Sub

Sub endmacro()
'マクロ終了時のおまじない
  Application.ScreenUpdating = True 'Start Updating Screen
  Application.EnableEvents = True 'Start Events
  Application.Calculation = xlCalculationAutomatic 'Caluculation Automatically
  Application.StatusBar = False 'Clear Statusbar message
End Sub

Function hasSpace(s As String) As Boolean
' =======================================================================
' 関数名   : hasSpace
' 関数概要: 引数として渡した文字列に、全角ないし半角スペースがあるかどうかを判定
' 引数     : s スペースを含んでいるか調べたい文字列
' 返り値   : 空白が含まれる→True 空白が含まれない→False
' =======================================================================

    If InStr(s, " ") > 0 Or InStr(s, "　") > 0 Then
        hasSpace = True
    Else
        hasSpace = False
    End If

End Function

Function hasHiragana(ByVal s As String) As Boolean
' =======================================================================
' 関数名   : hasHiragana
' 関数概要: 引数として渡した文字列にひらがなが含まれるかを判定
' 引数     : s ひらがなを含んでいるか調べたい文字列
' 返り値   : ひらがなが含まれる→True ひらがなが含まれない→False
' =======================================================================

    'RegExpを使えるようにオブジェクト宣言
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    
    With reg
        .Pattern = "[\u3040-\u309F]" '正規表現でひらがな判定する
        .Global = True '文字列全体を見る
        
        If .Test(s) Then 'Patternに合致したら（ひらがなだったら）
            hasHiragana = True
        Else
            hasHiragana = False
        End If
    End With

End Function
