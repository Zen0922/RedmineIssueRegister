Option Explicit
'--------------------------------------------------------------------
'    Redmine File Uploader
'      Version 1.00
'
'       Redmine�Ƀt�@�C���𑗐M��token��ԋp����\�t�g�E�F�A�ł��B
'
'   �y�g�p���@�z
'       �����ɕK�v�Ȓl���Z�b�g���ċN�����Ă��������B
'       ���̃v���O�����͎w�肷��p�X�̃t�@�C�����A�b�v���[�h���܂��B
'
'   �y�����̐����z
'       �R�}���h���C���iCScript�j�ŋN�����Ă��������B
'       ��jRedmineFileUploader.vbs [����1] [����2] [����3] [����4]
'       ����1   	Redmine��API�L�[���Z�b�g���Ă��������B
'       ����2   	Redmine��uploads.xml�̂���URI���w�肵�Ă��������B
'       ����3   	�A�b�v���[�h����t�@�C���̐�΃p�X���w�肵�Ă��������B
'               �p�X�ɃX�y�[�X���܂ޏꍇ��""�ň͂�ł��������B
'		����4	�A�X�L�[���[�h�ŃA�b�v���[�h����ꍇ�́uascii�v�𔽉f���Ă��������B
'				�o�C�i�����[�h�ŃA�b�v���[�h����ꍇ�́ubinary�v�𔽉f���Ă��������B
'
'	�y�R���\�[���ԋp�l�z
'		1�s�� [�X�e�[�^�X] [HTTP���b�Z�[�Wor�G���[���b�Z�[�W]
'		2�s�� [token]
'
'	�y�g�����z
'		�o�͂��ꂽ�R���\�[���̌��ʂ��擾���ăg�[�N�����`�P�b�g�o�^�i�X�V�j���̒l�Ƃ���
'		���f���Ă��������B
'		�����Ӂ�cscript����R�[�����Ȃ��ƃG���[�ɂȂ�܂��B
'		Excel VBA����̌Ăяo����j
'		    Set obj = CreateObject("WScript.shell")
'		    Set wExec = obj.exec("%ComSpec% /c cscript.exe " & cmdStr)
'
'	�y�ύX�����z
'		2019/01/19	1.00	����
'--------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
'���ʕϐ��錾
'--------------------------------------------------------------------------------------------------
'�N��������
Dim argAPIKey
Dim argRedmineUri
Dim argFilePath
Dim argFileType

'�I�u�W�F�N�g
Dim FSO
Dim xmlHttp
Dim xmlDoc

'���s�p�ϐ�
Dim RM_BASE_URI
Dim tmpToken
'--------------------------------------------------------------------------------------------------
'�v���O������������
'--------------------------------------------------------------------------------------------------
Call setEnvironment
Call checkArguments
Call uploadFile
'--------------------------------------------------------------------------------------------------
'�T�u���[�`���Q��������
'--------------------------------------------------------------------------------------------------

'���ݒ聡������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������
Private Sub setEnvironment()
    '�����m�F
    If Wscript.Arguments.Count <> 4 Then
        WScript.StdOut.WriteLine "999 �������s�����Ă��܂��B�v���O�������I�����܂��B"
        WScript.Quit
    End If
    argAPIKey = Wscript.Arguments(0)
    argRedmineUri =Wscript.Arguments(1)
    argFilePath = Wscript.Arguments(2)
	argFileType = Wscript.Arguments(3)

	'���ؗp�ϐ�
'	argAPIKey = "eae538eaae18bd51cbb8781e5efe88448f184b2a"
'	argRedmineUri = "http://127.0.0.1:81/redmine/uploads.xml"
'	argFilePath = "G:\Zen\Documents\RedmineFileUploader\v1.0\�A�b�v���[�h�t�@�C��\�����n�_����ԍ��������|�[�g.xlsx"
'	argFileType = "binary"

	'RM�ڑ��p��{URI����
	RM_BASE_URI = argRedmineUri & "?key=" & argAPIKey
	
	'�I�u�W�F�N�g����
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP.3.0")
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
End Sub

'�e����̓`�F�b�N������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������
Private Sub checkArguments()
	'�t�@�C�����݃`�F�b�N
	If Not FSO.FileExists(argFilePath) Then
        WScript.Echo "999 �A�b�v���[�h�Ώۂ̃t�@�C�������݂��܂���B�v���O�������I�����܂��B"
        WScript.Quit
	End If
	
	'�t�@�C���^�C�v����
	If Not (argFileType = "ascii" Or argFileType = "binary") Then
        WScript.Echo "999 �t�@�C���^�C�v�̈������s���ł��B�v���O�������I�����܂��B"
        WScript.Quit
	End If		
End Sub

'�t�@�C���A�b�v���[�h�J�n����������������������������������������������������������������������������������������������������������������������������������������������������������������������������
Private Sub uploadFile()
	Dim sendData
	Dim obj
	
	sendData = getFileBytes( argFilePath, argFileType)
	xmlHttp.open "POST", RM_BASE_URI, False
	xmlHttp.setRequestHeader "Content-Type", "application/octet-stream"
	xmlHttp.send sendData
	If xmlHttp.status <> "201" Then
        WScript.Echo "999 �A�b�v���[�h�����s���܂����BURI�����������m�F���Ă��������B�v���O�������I�����܂��B�i" & xmlHttp.status & " " & xmlHttp.statusText & "�j"
        WScript.Quit
	End If
	
	'XM���t�@�C���p�[�X
	xmlDoc.loadXML xmlHttp.responseText
	tmpToken = xmlDoc.getElementsByTagName("token")(0).nodeTypedValue

	'�R���\�[���o��
	WScript.Echo xmlHttp.status & " " & xmlHttp.statusText
	WScript.Echo tmpToken

	Set xmlHttp = Nothing
	Set xmlDoc = Nothing

End Sub

'�t�@�C���X�g���[������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������
Private Function getFileBytes(fileName, fileType)
	Dim ADOStream
    Set ADOStream = CreateObject("ADODB.Stream")

	If fileType = "binary" Then
		ADOStream.Type = 1		'Binary
	Else
		ADOStream.Type = 2		'Ascii
		ADOStream.Charset = "ascii"
	End If

	ADOStream.Open
	ADOStream.LoadFromFile fileName
	If fileType = "binary" Then
		getFileBytes = ADOStream.Read 'read binary'
	Else
		getFileBytes = ADOStream.ReadText 'read ascii'
	End If
	ADOStream.Close
	Set ADOStream = Nothing
End Function