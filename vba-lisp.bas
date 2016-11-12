Attribute VB_Name = "Module1"

Sub VbaLisp()
    Dim Workbook_path As String
    Dim WSH, wExec, sCmd As String, Result As String
    Dim LispCommand As String
    
    Thisbook_path = ThisWorkbook.Path   ''このマクロが実行されたPathを取得
    
    LispCommand = Replace(Replace(Range("B2").Value, vbLf, " "), """", "\""") ''lispコマンドを取得し、改行を空白に変換、"を\でエスケープ
    
    Set WSH = CreateObject("WScript.Shell")
'    sCmd = Thisbook_path & "\repl.exe " & """" & LispCommand & """" '' exeファイルの場合。ここにコマンドを記載
    sCmd = "ros " & Thisbook_path & "\repl.ros " & """" & LispCommand & """" ''roswellスクリプトの場合。ここにコマンドを記載
    Set wExec = WSH.Exec("%ComSpec% /c " & sCmd)        ''"%ComSpec%"はおまじない
    Do While wExec.Status = 0       ''外部のプログラムが終了するまで待ち続ける、多分・・
        DoEvents
    Loop
    Result = wExec.StdOut.ReadAll   ''Resultに標準出力の結果を格納
    Set wExec = Nothing
    Set WSH = Nothing
    
    Range("K2").Value = Result

End Sub

