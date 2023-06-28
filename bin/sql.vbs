Dim oCON
Dim oFSO
Dim oFLL
Dim oFLS
Dim strFileList
Dim strFileSQL
Dim strRec
Dim strSQL
Dim strDSN

Call MAIN()

Sub MAIN()
On Error Resume Next
	'<< 引数の取得 >>
	IF (Wscript.Arguments.Count <> 2) Then
		Wsh.Echo "引数の指定が間違っています"
		Exit SUB
	End IF
	strDSN		= Wscript.Arguments(0)
	strFileList	= Wscript.Arguments(1)
	'<< データベースＯＰＥＮ >>
	Set oCON = CreateObject("ADODB.Connection")
	oCon.Open strDSN
	IF (Err.Number <> 0) Then
		Wsh.Echo "データベースの接続に失敗"
		Wsh.Echo Err.Number & ":" & Err.Description
		Exit SUB
	End IF
	'<< ファイルシステムオブジェクト作成 >>
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	'<< メイン処理開始 >>
	Call SUB01()
	'<< データベースＣＬＯＳＥ >>
	oCon.Close
	Set oFSO = Nothing
	Set oCon = Nothing
End Sub
 
Sub SUB01()
On Error Resume Next
	Set oFLL = oFSO.OpenTextFile("prm\" & strFileList, 1, False)
	IF (Err.Number <> 0) Then
		Wsh.Echo "リストファイルの読み込みに失敗"
		Wsh.Echo Err.Number & ":" & Err.Description
		Exit Sub
	End IF
	Do While oFLL.AtEndOfLine <> True
		strFileSQL = oFLL.ReadLine()
		Set oFLS = oFSO.OpenTextFile("sql\" & strFileSQL, 1, False)
		IF (Err.Number <> 0) Then
			Wsh.Echo "ＳＱＬファイルの読み込みに失敗"
			Wsh.Echo Err.Number & ":" & Err.Description
			Exit Do
		End IF
		strSQL = oFLS.ReadAll()
		oFLS.Close
		Wsh.Echo strFileSQL & " " & now & " 処理を開始しました。"
		oCon.Execute strSQL
		IF (Err.Number <> 0) Then
			Wsh.Echo "ＳＱＬの実行に失敗しました"
			Wsh.Echo Err.Number & ":" & Err.Description
			Exit Do
		End IF
		Wsh.Echo strFileSQL & " " & now & " 処理を終了しました。"
	Loop
	oFLL.Close
End Sub
