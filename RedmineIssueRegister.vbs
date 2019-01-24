Option Explicit
'--------------------------------------------------------------------------------------------------
'	Redmineチケット自動登録スクリプト
'		Version 1.00
'	
'	指定ディレクトリに格納された設定ファイルからチケット登録情報を集めてRedmineにチケットを登録します。
'	
'	【起動方法】
'		コマンドラインからRedmineIssueRegister.vbs起動してください。
'		CScriptでないとエラーになる場合があります。
'		> cscript RedmineIssueRegister.vbs [対象ディレクトリ]
'	
'	【設定ファイル】
'		設定ファイルはTempleteフォルダのReadme.txtをご覧ください。
'
'	【制限事項】
'		複数選択肢のカスタムフィールドには対応していません。
'		
'	【更新履歴】
'		2019/01/24	v1.00	初版
'
'	【著作】Zen(http://relphe.s4.valueserver.jp/wp/)
'		ライセンスはLGPLとします。
'
'--------------------------------------------------------------------------------------------------
'オブジェクト
Dim FSO, F
Dim WSH
'設定ファイル
Dim PATH_MYSELF_DIR
Dim PATH_LIB_DIR
Dim PATH_MIMEASCII_DICFILE
'辞書
Dim dicMIME
Dim dicASCII
'特定ファイル名の指定リスト（アップロード対象外）
Dim systemFilenames, sysFile
'読込ファイル関係
Dim loadFileDir
'ファイルアップロード関係
Dim fileToken
Dim fileName
Dim fileDescription
Dim fileContentType
Dim fileUploadType
Dim uploadXML
Dim uploadVBSPath
'送信XML生成
Dim postXML
Dim customFieldXML
Dim issueFieldXML
Dim xmlHttp
Dim xmlDoc
'Redmineの基本情報
Dim RedmineBaseURI
Dim RedmineIssueURI
Dim RedmineUploadURI
Dim RedmineAPIKey
Dim RedmineRequestURI
'最終文字コード変換
Dim ADOStream0
'処理使用変数
Dim tmpCmd				'アップロード用発行コマンド
Dim wExec				'ファイルアップロードコマンド結果保管変数
Dim fileUploadResult	'ファイルアップロードコマンド結果取得変数
Dim uploadResultText	'アップロード結果テキスト
'--------------------------------------------------------------------------------------------------
'	メインプログラムここから
'--------------------------------------------------------------------------------------------------
Call setEnvironment
'基本ファイル読込
Call getIssueFile			'issue.txtファイル読込
'RedmineURL生成
If Right(RedmineBaseURI,1) <> "/" Then
	RedmineBaseURI = RedmineBaseURI & "/"
End If
RedmineIssueURI = RedmineBaseURI & "issues.xml"			'チケット発行用
RedmineUploadURI = RedmineBaseURI & "uploads.xml"		'ファイルアップロード用

'ファイルアップロード処理
uploadXML = ""
For Each F In FSO.GetFolder(loadFileDir).Files
	'システムファイルと先頭がアンダーバーで始まるファイルはアップロード対象から除外
	sysFile = Filter(systemFilenames, F.Name)
	If Not ( UBound(sysFile) <> -1 Or Left(F.Name, 1) = "_" ) Then
		'対応拡張子判定
		If isAsciiMode(getFileType(F.Name)) = "" Then
			WScript.Echo "999 対応拡張子に存在しないファイルです。アップロードできません。プログラムを終了します。"
			WScript.Quit
		End If
		'ファイルアップロード処理
		Select Case isAsciiMode(getFileType(F.Name))
			Case 1
				fileUploadType = "ascii"
			Case 0
				fileUploadType = "binary"
		End Select
		tmpCmd = "%ComSpec% /c cscript /nologo """ & uploadVBSPath & """ " & RedmineAPIKey & " " & RedmineUploadURI & " """ & F.Path & """ " & fileUploadType
		Set wExec = WSH.Exec(tmpCmd)
		fileUploadResult = wExec.StdOut.ReadAll
		'アップロード結果取得処理
		fileUploadResult = Replace(fileUploadResult, vbCrLf, vbLf)
		fileUploadResult = Replace(fileUploadResult, vbCr, vbLf)
		If Left(fileUploadResult, 4) = "999 " Then
			WScript.Echo "999 ファイルアップロードに失敗しました。実行メッセージ（" & fileUploadResult & "）"
			WScript.Quit
		End If
		uploadResultText = Split(fileUploadResult, vbLf)
		fileToken = uploadResultText(1)
		If uploadXML = "" Then
			uploadXML = "<uploads type=""array"">" & vbCrLf
		End If
		'登録情報追記
		fileName = F.Name
		fileDescription = "AutomaticUploaded"
		fileContentType = getMIMEType(getFileType(F.Name))
		'ファイル情報追記
		uploadXML = _
			uploadXML & "<upload>" & _
			"<token>" & fileToken & "</token>" & _
			"<filename>" & fileName & "</filename>" & _
			"<description>" & fileDescription & "</description>" & _
			"<content_type>" & fileContentType & "</content_type>" & _
			"</upload>" & vbCrLf
	End If
Next
'ファイル情報に書き込みがある場合はuploads要素をCLOSE
If uploadXML <> "" Then
	uploadXML = uploadXML & "</uploads>"
End If

'XML生成
postXML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
postXML = postXML & "<issue>" & vbCrLf
postXML = postXML & issueFieldXML & vbCrLf
postXML = postXML & customFieldXML & vbCrLf
postXML = postXML & uploadXML & vbCrLf
postXML = postXML & "</issue>"
'postXMLの文字コード変換（SJIS->UTF-8）
postXML = encodeUTF8(postXML)

'HTTPリクエスト
RedmineRequestURI = RedmineIssueURI & "?key=" & RedmineAPIKey
Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
xmlHttp.open "POST", RedmineRequestURI, False
xmlHttp.setRequestHeader "Content-Type", "text/xml"
xmlHttp.send postXML

'ステータス処理
If xmlHttp.status <> "201" Then
	'リクエスト失敗時のエラーメッセージ
	WScript.Echo "999 リクエストに失敗しました。プログラムを終了します。"
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
	'リクエスト成功時はチケットIDを返却する
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
	xmlDoc.loadXML xmlHttp.responseText
	WScript.echo xmlDoc.getElementsByTagName("id")(0).nodeTypedValue
End If
'--------------------------------------------------------------------------------------------------
'	メインプログラムここまで
'--------------------------------------------------------------------------------------------------
'環境設定■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
Private Sub setEnvironment()
	'引数チェック
	If Wscript.Arguments.Count <> 1 Then
		WScript.Echo "999 引数が不足しています。プログラムを終了します。"
		WScript.Quit
	End If
	loadFileDir = WScript.Arguments(0)
	
	Set FSO = CreateObject("Scripting.FileSystemObject")
	'パスを生成
	PATH_MYSELF_DIR = 			FSO.GetParentFolderName(WScript.ScriptFullName)			'実行ディレクトリ
	PATH_LIB_DIR = 				PATH_MYSELF_DIR & "\Lib"								'ライブラリディレクトリ
	PATH_MIMEASCII_DICFILE = 	PATH_LIB_DIR & "\Dictionary_MIME_ASCII.csv"				'MIME-ASCII判定用CSV
	'辞書を設定
	Call setDictionary
	'システムファイル名の定義
	systemFilenames = Array( _
						"issue.txt")
	'WSH
	Set WSH = CreateObject("WScript.shell")
	uploadVBSPath = PATH_MYSELF_DIR & "\RedmineFileUploader.vbs"
End Sub

'拡張子/MIME-TYPE/ASCII判定辞書生成■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
Private Sub setDictionary()
	Dim objFile
	Dim buf
	Dim splitBuf
	
	Set dicMIME = CreateObject("Scripting.Dictionary")
	Set dicASCII = CreateObject("Scripting.Dictionary")
	
	'ファイルの存在チェック
	If FSO.FileExists(PATH_MIMEASCII_DICFILE) = False Then
		WScript.Echo "999 MIME-ASCII判定用ファイルが存在しません。プログラムを終了します。"
		WScript.Quit
	End If
	
	Set objFile = FSO.OpenTextFile(PATH_MIMEASCII_DICFILE, 1, False)
	If err.number > 0 Then
		WScript.Echo "999 MIME-ASCII判定用ファイルオープンに失敗しました。プログラムを終了します。"
		WScript.Quit
	End If	
	
	'辞書登録
	Do Until objFile.AtEndOfStream
		buf = objFile.ReadLine
		'不要文字除去
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

'拡張子取得関数■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
Private Function getFileType(argFileName)
	Dim tmpStr
	If InStrRev(argFileName, ".") = 0 Then
		getFileType = ""
		Exit Function
	End If

	'文字列取得
	tmpStr = Right(argFileName, Len(argfilename) - InStrRev(argFileName, "."))
	getFileType = LCase(tmpStr)
End Function

'MIME-TYPE判定取得■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
Private Function getMIMEType(argFileType)
	If dicMIME.Exists(argFileType) Then
		getMIMEType = dicMIME(argFileType)
		Exit Function
	Else
		getMIMEType = ""
		Exit Function
	End If
End Function

'ASCII判定取得■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
Private Function isAsciiMode(argFileType)
	If dicASCII.Exists(argFileType) Then
		isAsciiMode = dicASCII(argFileType)
		Exit Function
	Else
		isAsciiMode = ""
		Exit Function
	End If
End Function

'登録用ファイルからデータ取得■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
Private Sub getIssueFile()
	'ADOStream1: for issue.txt		ADOStream2: for <FILE>
	Dim ADOStream1, ADOStream2
	Dim bufLine, KeyValue
	Dim tmpFilename
	Dim customField
	Dim tmpValue
	Dim setComplete
	
	'issueファイルオープン処理
	Set ADOStream1 = CreateObject("ADODB.Stream")
	ADOStream1.Mode = 3 'adModeReadWrite
	ADOStream1.Type = 2 'Text
	ADOStream1.Charset = "Shift_JIS"
	ADOStream1.Open		'StreamOpen
	'issue.txtオープン
	ADOStream1.LoadFromFile(loadFileDir & "\issue.txt")
	ADOStream1.Position = 0		'Cursor to Header
	Do While Not ADOStream1.EOS
		setComplete = 0
	
		'1行ずつ読込
		bufLine = ADOStream1.ReadText(-2)
		KeyValue = Split(bufLine, "=")		'=でスプリット
		KeyValue(1) = Replace(KeyValue(1), vbCrLf, "")
		KeyValue(1) = Replace(KeyValue(1), vbCr, "")
		KeyValue(1) = Replace(KeyValue(1), vbLf, "")
		'値のファイル入力判定
		If KeyValue(1) = "<FILE>" Then
			'ファイルを読み込み
			Set ADOStream2 = CreateObject("ADODB.Stream")
			ADOStream2.Mode = 3 'adModeReadWrite
			ADOStream2.Type = 2 'Text
			ADOStream2.Charset = "Shift_JIS"
			ADOStream2.Open		'StreamOpen
			tmpFilename = loadFileDir & "\_" & Replace(KeyValue(0),":","_") & ".txt"
			If FSO.FileExists(tmpFilename) = False Then
				WScript.Echo "999 チケット発行用の参照ファイルが見つかりませんでした。プログラムを終了します。"
				WScript.Quit
			End If
			ADOStream2.LoadFromFile(tmpFilename)
			ADOStream2.Position = 0		'Cursor to Header
			KeyValue(1) = ADOStream2.ReadText(-1)	'全行読込
			tmpValue = Replace(Replace(trimReturn(KeyValue(1)),"<","&lt;"),">","&gt;")
			ADOStream2.Close
			Set ADOStream2 = Nothing
		Else
			tmpValue = trimReturn(KeyValue(1))
		End If
		
		'システムキーを判定
		If Left(KeyValue(0),1) = "_" Then
			Select Case KeyValue(0)
				Case "_uri"
					RedminebaseURI = tmpValue
				Case "_apikey"
					RedmineAPIKey = tmpValue
				Case Else
					WScript.Echo "999 システムキーに不正な値がセットされています。プログラムを終了します。"
					WScript.Quit
			End Select
			setComplete = 1
		End If
		
		'custom_fieldよりも文字数が大きい時はcustom_field検知処理を実施
		If Len(KeyValue(0)) >= Len("custom_field") And setComplete = 0 Then
			'custom_fieldを検出する
			If Left(KeyValue(0), Len("custom_field")) = Left(KeyValue(0), Len("custom_field")) Then
				'項目名とcustom_fieldのIDと分離
				customField = Split(KeyValue(0), ":")
				If customFieldXML = "" Then
					customFieldXML = "<custom_fields type=""array"">" & vbCrLf
				End If
				'XML生成				
				customFieldXML = customFieldXML & "<custom_field id=""" & customField(1) & """>"
				customFieldXML = customFieldXML & "<value>" & tmpValue & "</value>"
				customFieldXML = customFieldXML & "</custom_field>"
				setComplete = 1
			End If
		End If
		
		'それ以外のキーを生成する
		If setComplete = 0 Then
			issueFieldXML = issueFieldXML & "<" & KeyValue(0) & ">" & tmpValue & "</" & KeyValue(0) & ">" & vbCrLf
			setComplete = 1
		End If
	Loop
	'custom_fieldsをCLOSE
	If customFieldXML <> "" Then
		customFieldXML = customFieldXML & "</custom_fields>"
	End If
	ADOStream1.Close
	Set ADOStream1 = Nothing
End Sub

'先頭と末尾の改行コードを除去■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
Private Function trimReturn(argString)
	Dim inputString
	inputString = argString
	'一旦改行コードをすべてLFにする
	inputString = Replace(inputString, vbCrLf, vbLf)
	inputString = Replace(inputString, vbCr, vbLf)
	'全部改行コードだった場合の対処
	If Len(Replace(inputString, vbCrLf, "")) = 0 Then
		trimReturn = ""
		Exit Function
	End If
	'末尾の改行コードを全部カット	
	Do While Right(inputString, 1) = vbLf
		inputString = Left(inputString, Len(inputString) -1 )
	Loop
	'先頭の改行コードを全部カット
	Do While Left(inputString, 1) = vbLf
		inputString = Mid(inputString, 2)
	Loop
	trimReturn = Replace(inputString, vbLf, vbCrLf)
End Function

'文字コード変換（S-JIS->UTF-8）■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
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