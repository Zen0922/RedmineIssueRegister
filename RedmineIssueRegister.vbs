Option Explicit
'--------------------------------------------------------------------------------------------------
'	Redmine�`�P�b�g�����o�^�X�N���v�g
'		Version 1.00
'	
'	�w��f�B���N�g���Ɋi�[���ꂽ�ݒ�t�@�C������`�P�b�g�o�^�����W�߂�Redmine�Ƀ`�P�b�g��o�^���܂��B
'	
'	�y�N�����@�z
'		�R�}���h���C������RedmineIssueRegister.vbs�N�����Ă��������B
'		CScript�łȂ��ƃG���[�ɂȂ�ꍇ������܂��B
'		> cscript RedmineIssueRegister.vbs [�Ώۃf�B���N�g��]
'	
'	�y�ݒ�t�@�C���z
'		�ݒ�t�@�C����Templete�t�H���_��Readme.txt���������������B
'
'	�y���������z
'		�����I�����̃J�X�^���t�B�[���h�ɂ͑Ή����Ă��܂���B
'		
'	�y�X�V�����z
'		2019/01/24	v1.00	����
'
'	�y����zZen(http://relphe.s4.valueserver.jp/wp/)
'		���C�Z���X��LGPL�Ƃ��܂��B
'
'--------------------------------------------------------------------------------------------------
'�I�u�W�F�N�g
Dim FSO, F
Dim WSH
'�ݒ�t�@�C��
Dim PATH_MYSELF_DIR
Dim PATH_LIB_DIR
Dim PATH_MIMEASCII_DICFILE
'����
Dim dicMIME
Dim dicASCII
'����t�@�C�����̎w�胊�X�g�i�A�b�v���[�h�ΏۊO�j
Dim systemFilenames, sysFile
'�Ǎ��t�@�C���֌W
Dim loadFileDir
'�t�@�C���A�b�v���[�h�֌W
Dim fileToken
Dim fileName
Dim fileDescription
Dim fileContentType
Dim fileUploadType
Dim uploadXML
Dim uploadVBSPath
'���MXML����
Dim postXML
Dim customFieldXML
Dim issueFieldXML
Dim xmlHttp
Dim xmlDoc
'Redmine�̊�{���
Dim RedmineBaseURI
Dim RedmineIssueURI
Dim RedmineUploadURI
Dim RedmineAPIKey
Dim RedmineRequestURI
'�ŏI�����R�[�h�ϊ�
Dim ADOStream0
'�����g�p�ϐ�
Dim tmpCmd				'�A�b�v���[�h�p���s�R�}���h
Dim wExec				'�t�@�C���A�b�v���[�h�R�}���h���ʕۊǕϐ�
Dim fileUploadResult	'�t�@�C���A�b�v���[�h�R�}���h���ʎ擾�ϐ�
Dim uploadResultText	'�A�b�v���[�h���ʃe�L�X�g
'--------------------------------------------------------------------------------------------------
'	���C���v���O������������
'--------------------------------------------------------------------------------------------------
Call setEnvironment
'��{�t�@�C���Ǎ�
Call getIssueFile			'issue.txt�t�@�C���Ǎ�
'RedmineURL����
If Right(RedmineBaseURI,1) <> "/" Then
	RedmineBaseURI = RedmineBaseURI & "/"
End If
RedmineIssueURI = RedmineBaseURI & "issues.xml"			'�`�P�b�g���s�p
RedmineUploadURI = RedmineBaseURI & "uploads.xml"		'�t�@�C���A�b�v���[�h�p

'�t�@�C���A�b�v���[�h����
uploadXML = ""
For Each F In FSO.GetFolder(loadFileDir).Files
	'�V�X�e���t�@�C���Ɛ擪���A���_�[�o�[�Ŏn�܂�t�@�C���̓A�b�v���[�h�Ώۂ��珜�O
	sysFile = Filter(systemFilenames, F.Name)
	If Not ( UBound(sysFile) <> -1 Or Left(F.Name, 1) = "_" ) Then
		'�Ή��g���q����
		If isAsciiMode(getFileType(F.Name)) = "" Then
			WScript.Echo "999 �Ή��g���q�ɑ��݂��Ȃ��t�@�C���ł��B�A�b�v���[�h�ł��܂���B�v���O�������I�����܂��B"
			WScript.Quit
		End If
		'�t�@�C���A�b�v���[�h����
		Select Case isAsciiMode(getFileType(F.Name))
			Case 1
				fileUploadType = "ascii"
			Case 0
				fileUploadType = "binary"
		End Select
		tmpCmd = "%ComSpec% /c cscript /nologo """ & uploadVBSPath & """ " & RedmineAPIKey & " " & RedmineUploadURI & " """ & F.Path & """ " & fileUploadType
		Set wExec = WSH.Exec(tmpCmd)
		fileUploadResult = wExec.StdOut.ReadAll
		'�A�b�v���[�h���ʎ擾����
		fileUploadResult = Replace(fileUploadResult, vbCrLf, vbLf)
		fileUploadResult = Replace(fileUploadResult, vbCr, vbLf)
		If Left(fileUploadResult, 4) = "999 " Then
			WScript.Echo "999 �t�@�C���A�b�v���[�h�Ɏ��s���܂����B���s���b�Z�[�W�i" & fileUploadResult & "�j"
			WScript.Quit
		End If
		uploadResultText = Split(fileUploadResult, vbLf)
		fileToken = uploadResultText(1)
		If uploadXML = "" Then
			uploadXML = "<uploads type=""array"">" & vbCrLf
		End If
		'�o�^���ǋL
		fileName = F.Name
		fileDescription = "AutomaticUploaded"
		fileContentType = getMIMEType(getFileType(F.Name))
		'�t�@�C�����ǋL
		uploadXML = _
			uploadXML & "<upload>" & _
			"<token>" & fileToken & "</token>" & _
			"<filename>" & fileName & "</filename>" & _
			"<description>" & fileDescription & "</description>" & _
			"<content_type>" & fileContentType & "</content_type>" & _
			"</upload>" & vbCrLf
	End If
Next
'�t�@�C�����ɏ������݂�����ꍇ��uploads�v�f��CLOSE
If uploadXML <> "" Then
	uploadXML = uploadXML & "</uploads>"
End If

'XML����
postXML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
postXML = postXML & "<issue>" & vbCrLf
postXML = postXML & issueFieldXML & vbCrLf
postXML = postXML & customFieldXML & vbCrLf
postXML = postXML & uploadXML & vbCrLf
postXML = postXML & "</issue>"
'postXML�̕����R�[�h�ϊ��iSJIS->UTF-8�j
postXML = encodeUTF8(postXML)

'HTTP���N�G�X�g
RedmineRequestURI = RedmineIssueURI & "?key=" & RedmineAPIKey
Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
xmlHttp.open "POST", RedmineRequestURI, False
xmlHttp.setRequestHeader "Content-Type", "text/xml"
xmlHttp.send postXML

'�X�e�[�^�X����
If xmlHttp.status <> "201" Then
	'���N�G�X�g���s���̃G���[���b�Z�[�W
	WScript.Echo "999 ���N�G�X�g�Ɏ��s���܂����B�v���O�������I�����܂��B"
	WScript.echo "RedmineIssueURI-------------------------------------------------------------"
	WScript.Echo RedmineRequestURI
	WScript.echo "POST XML START--------------------------------------------------------------"
	WScript.Echo postXML
	WScript.echo "POST XML END----------------------------------------------------------------"
	WScript.Echo ""
	
	WScript.echo "RESPONSE CODE---------------------------------------------------------------"
	WScript.Echo xmlHttp.status
	WScript.echo "RESPONSE BODY START---------------------------------------------------------"
	WScript.Echo xmlHttp.responseText
	WScript.echo "RESPONSE BODY END-----------------------------------------------------------"
Else
	'���N�G�X�g�������̓`�P�b�gID��ԋp����
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
	xmlDoc.loadXML xmlHttp.responseText
	WScript.echo xmlDoc.getElementsByTagName("id")(0).nodeTypedValue
End If
'--------------------------------------------------------------------------------------------------
'	���C���v���O���������܂�
'--------------------------------------------------------------------------------------------------
'���ݒ聡������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������
Private Sub setEnvironment()
	'�����`�F�b�N
	If Wscript.Arguments.Count <> 1 Then
		WScript.Echo "999 �������s�����Ă��܂��B�v���O�������I�����܂��B"
		WScript.Quit
	End If
	loadFileDir = WScript.Arguments(0)
	
	Set FSO = CreateObject("Scripting.FileSystemObject")
	'�p�X�𐶐�
	PATH_MYSELF_DIR = 			FSO.GetParentFolderName(WScript.ScriptFullName)			'���s�f�B���N�g��
	PATH_LIB_DIR = 				PATH_MYSELF_DIR & "\Lib"								'���C�u�����f�B���N�g��
	PATH_MIMEASCII_DICFILE = 	PATH_LIB_DIR & "\Dictionary_MIME_ASCII.csv"				'MIME-ASCII����pCSV
	'������ݒ�
	Call setDictionary
	'�V�X�e���t�@�C�����̒�`
	systemFilenames = Array( _
						"issue.txt")
	'WSH
	Set WSH = CreateObject("WScript.shell")
	uploadVBSPath = PATH_MYSELF_DIR & "\RedmineFileUploader.vbs"
End Sub

'�g���q/MIME-TYPE/ASCII���莫��������������������������������������������������������������������������������������������������������������������������������������������������������
Private Sub setDictionary()
	Dim objFile
	Dim buf
	Dim splitBuf
	
	Set dicMIME = CreateObject("Scripting.Dictionary")
	Set dicASCII = CreateObject("Scripting.Dictionary")
	
	'�t�@�C���̑��݃`�F�b�N
	If FSO.FileExists(PATH_MIMEASCII_DICFILE) = False Then
		WScript.Echo "999 MIME-ASCII����p�t�@�C�������݂��܂���B�v���O�������I�����܂��B"
		WScript.Quit
	End If
	
	Set objFile = FSO.OpenTextFile(PATH_MIMEASCII_DICFILE, 1, False)
	If err.number > 0 Then
		WScript.Echo "999 MIME-ASCII����p�t�@�C���I�[�v���Ɏ��s���܂����B�v���O�������I�����܂��B"
		WScript.Quit
	End If	
	
	'�����o�^
	Do Until objFile.AtEndOfStream
		buf = objFile.ReadLine
		'�s�v��������
		buf = Replace(buf, """",	"")
		buf = Replace(buf, vbCrLf,	"")
		buf = Replace(buf, vbCr,	"")
		buf = Replace(buf, vbLf,	"")
		buf = Replace(buf, " ",		"")
		If InStr(buf, ",") > 0 Then
			splitBuf = Split(buf, ",")
			If splitBuf(0) <> "" Then
				dicMIME.Add LCase(splitBuf(0)), splitBuf(1)
				dicASCII.Add LCase(splitBuf(0)), splitBuf(2)
			End if
		End if
	Loop
End Sub

'�g���q�擾�֐���������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������
Private Function getFileType(argFileName)
	Dim tmpStr
	If InStrRev(argFileName, ".") = 0 Then
		getFileType = ""
		Exit Function
	End If

	'������擾
	tmpStr = Right(argFileName, Len(argfilename) - InStrRev(argFileName, "."))
	getFileType = LCase(tmpStr)
End Function

'MIME-TYPE����擾��������������������������������������������������������������������������������������������������������������������������������������������������������������������������
Private Function getMIMEType(argFileType)
	If dicMIME.Exists(argFileType) Then
		getMIMEType = dicMIME(argFileType)
		Exit Function
	Else
		getMIMEType = ""
		Exit Function
	End If
End Function

'ASCII����擾����������������������������������������������������������������������������������������������������������������������������������������������������������������������������������
Private Function isAsciiMode(argFileType)
	If dicASCII.Exists(argFileType) Then
		isAsciiMode = dicASCII(argFileType)
		Exit Function
	Else
		isAsciiMode = ""
		Exit Function
	End If
End Function

'�o�^�p�t�@�C������f�[�^�擾������������������������������������������������������������������������������������������������������������������������������������������������������������������������
Private Sub getIssueFile()
	'ADOStream1: for issue.txt		ADOStream2: for <FILE>
	Dim ADOStream1, ADOStream2
	Dim bufLine, KeyValue
	Dim tmpFilename
	Dim customField
	Dim tmpValue
	Dim setComplete
	
	'issue�t�@�C���I�[�v������
	Set ADOStream1 = CreateObject("ADODB.Stream")
	ADOStream1.Mode = 3 'adModeReadWrite
	ADOStream1.Type = 2 'Text
	ADOStream1.Charset = "Shift_JIS"
	ADOStream1.Open		'StreamOpen
	'issue.txt�I�[�v��
	ADOStream1.LoadFromFile(loadFileDir & "\issue.txt")
	ADOStream1.Position = 0		'Cursor to Header
	Do While Not ADOStream1.EOS
		setComplete = 0
	
		'1�s���Ǎ�
		bufLine = ADOStream1.ReadText(-2)
		KeyValue = Split(bufLine, "=")		'=�ŃX�v���b�g
		KeyValue(1) = Replace(KeyValue(1), vbCrLf, "")
		KeyValue(1) = Replace(KeyValue(1), vbCr, "")
		KeyValue(1) = Replace(KeyValue(1), vbLf, "")
		'�l�̃t�@�C�����͔���
		If KeyValue(1) = "<FILE>" Then
			'�t�@�C����ǂݍ���
			Set ADOStream2 = CreateObject("ADODB.Stream")
			ADOStream2.Mode = 3 'adModeReadWrite
			ADOStream2.Type = 2 'Text
			ADOStream2.Charset = "Shift_JIS"
			ADOStream2.Open		'StreamOpen
			tmpFilename = loadFileDir & "\_" & Replace(KeyValue(0),":","_") & ".txt"
			If FSO.FileExists(tmpFilename) = False Then
				WScript.Echo "999 �`�P�b�g���s�p�̎Q�ƃt�@�C����������܂���ł����B�v���O�������I�����܂��B"
				WScript.Quit
			End If
			ADOStream2.LoadFromFile(tmpFilename)
			ADOStream2.Position = 0		'Cursor to Header
			KeyValue(1) = ADOStream2.ReadText(-1)	'�S�s�Ǎ�
			tmpValue = Replace(Replace(trimReturn(KeyValue(1)),"<","&lt;"),">","&gt;")
			ADOStream2.Close
			Set ADOStream2 = Nothing
		Else
			tmpValue = trimReturn(KeyValue(1))
		End If
		
		'�V�X�e���L�[�𔻒�
		If Left(KeyValue(0),1) = "_" Then
			Select Case KeyValue(0)
				Case "_uri"
					RedminebaseURI = tmpValue
				Case "_apikey"
					RedmineAPIKey = tmpValue
				Case Else
					WScript.Echo "999 �V�X�e���L�[�ɕs���Ȓl���Z�b�g����Ă��܂��B�v���O�������I�����܂��B"
					WScript.Quit
			End Select
			setComplete = 1
		End If
		
		'custom_field�������������傫������custom_field���m���������{
		If Len(KeyValue(0)) >= Len("custom_field") And setComplete = 0 Then
			'custom_field�����o����
			If Left(KeyValue(0), Len("custom_field")) = Left(KeyValue(0), Len("custom_field")) Then
				'���ږ���custom_field��ID�ƕ���
				customField = Split(KeyValue(0), ":")
				If customFieldXML = "" Then
					customFieldXML = "<custom_fields type=""array"">" & vbCrLf
				End If
				'XML����				
				customFieldXML = customFieldXML & "<custom_field id=""" & customField(1) & """>"
				customFieldXML = customFieldXML & "<value>" & tmpValue & "</value>"
				customFieldXML = customFieldXML & "</custom_field>"
				setComplete = 1
			End If
		End If
		
		'����ȊO�̃L�[�𐶐�����
		If setComplete = 0 Then
			issueFieldXML = issueFieldXML & "<" & KeyValue(0) & ">" & tmpValue & "</" & KeyValue(0) & ">" & vbCrLf
			setComplete = 1
		End If
	Loop
	'custom_fields��CLOSE
	If customFieldXML <> "" Then
		customFieldXML = customFieldXML & "</custom_fields>"
	End If
	ADOStream1.Close
	Set ADOStream1 = Nothing
End Sub

'�擪�Ɩ����̉��s�R�[�h������������������������������������������������������������������������������������������������������������������������������������������������������������������������������
Private Function trimReturn(argString)
	Dim inputString
	inputString = argString
	'��U���s�R�[�h�����ׂ�LF�ɂ���
	inputString = Replace(inputString, vbCrLf, vbLf)
	inputString = Replace(inputString, vbCr, vbLf)
	'�S�����s�R�[�h�������ꍇ�̑Ώ�
	If Len(Replace(inputString, vbCrLf, "")) = 0 Then
		trimReturn = ""
		Exit Function
	End If
	'�����̉��s�R�[�h��S���J�b�g	
	Do While Right(inputString, 1) = vbLf
		inputString = Left(inputString, Len(inputString) -1 )
	Loop
	'�擪�̉��s�R�[�h��S���J�b�g
	Do While Left(inputString, 1) = vbLf
		inputString = Mid(inputString, 2)
	Loop
	trimReturn = Replace(inputString, vbLf, vbCrLf)
End Function

'�����R�[�h�ϊ��iS-JIS->UTF-8�j����������������������������������������������������������������������������������������������������������������������������������������������������������
Private Function encodeUTF8(argString)
	Dim resultString
	Set ADOStream0 = CreateObject("ADODB.Stream")
	ADOStream0.Open
	ADOStream0.Type = 2
	ADOStream0.Charset = "UTF-8"
	ADOStream0.WriteText argString
	ADOStream0.Position = 0
	ADOStream0.Type = 1
	ADOStream0.Position = 3
	resultString = ADOStream0.Read
	ADOStream0.Close
	Set ADOStream0 = Nothing
	encodeUTF8 = resultString
End Function