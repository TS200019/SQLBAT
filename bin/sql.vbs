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
	'<< �����̎擾 >>
	IF (Wscript.Arguments.Count <> 2) Then
		Wsh.Echo "�����̎w�肪�Ԉ���Ă��܂�"
		Exit SUB
	End IF
	strDSN		= Wscript.Arguments(0)
	strFileList	= Wscript.Arguments(1)
	'<< �f�[�^�x�[�X�n�o�d�m >>
	Set oCON = CreateObject("ADODB.Connection")
	oCon.Open strDSN
	IF (Err.Number <> 0) Then
		Wsh.Echo "�f�[�^�x�[�X�̐ڑ��Ɏ��s"
		Wsh.Echo Err.Number & ":" & Err.Description
		Exit SUB
	End IF
	'<< �t�@�C���V�X�e���I�u�W�F�N�g�쐬 >>
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	'<< ���C�������J�n >>
	Call SUB01()
	'<< �f�[�^�x�[�X�b�k�n�r�d >>
	oCon.Close
	Set oFSO = Nothing
	Set oCon = Nothing
End Sub
 
Sub SUB01()
On Error Resume Next
	Set oFLL = oFSO.OpenTextFile("prm\" & strFileList, 1, False)
	IF (Err.Number <> 0) Then
		Wsh.Echo "���X�g�t�@�C���̓ǂݍ��݂Ɏ��s"
		Wsh.Echo Err.Number & ":" & Err.Description
		Exit Sub
	End IF
	Do While oFLL.AtEndOfLine <> True
		strFileSQL = oFLL.ReadLine()
		Set oFLS = oFSO.OpenTextFile("sql\" & strFileSQL, 1, False)
		IF (Err.Number <> 0) Then
			Wsh.Echo "�r�p�k�t�@�C���̓ǂݍ��݂Ɏ��s"
			Wsh.Echo Err.Number & ":" & Err.Description
			Exit Do
		End IF
		strSQL = oFLS.ReadAll()
		oFLS.Close
		Wsh.Echo strFileSQL & " " & now & " �������J�n���܂����B"
		oCon.Execute strSQL
		IF (Err.Number <> 0) Then
			Wsh.Echo "�r�p�k�̎��s�Ɏ��s���܂���"
			Wsh.Echo Err.Number & ":" & Err.Description
			Exit Do
		End IF
		Wsh.Echo strFileSQL & " " & now & " �������I�����܂����B"
	Loop
	oFLL.Close
End Sub
