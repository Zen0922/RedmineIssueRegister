Option Explicit
'--------------------------------------------------------------------
'    Redmine File Uploader
'      Version 1.00
'
'       Redmineにファイルを送信しtokenを返却するソフトウェアです。
'
'   【使用方法】
'       引数に必要な値をセットして起動してください。
'       このプログラムは指定するパスのファイルをアップロードします。
'
'   【引数の説明】
'       コマンドライン（CScript）で起動してください。
'       例）RedmineFileUploader.vbs [引数1] [引数2] [引数3] [引数4]
'       引数1   	RedmineのAPIキーをセットしてください。
'       引数2   	Redmineのuploads.xmlのあるURIを指定してください。
'       引数3   	アップロードするファイルの絶対パスを指定してください。
'               パスにスペースを含む場合は""で囲んでください。
'		引数4	アスキーモードでアップロードする場合は「ascii」を反映してください。
'				バイナリモードでアップロードする場合は「binary」を反映してください。
'
'	【コンソール返却値】
'		1行目 [ステータス] [HTTPメッセージorエラーメッセージ]
'		2行目 [token]
'
'	【使い方】
'		出力されたコンソールの結果を取得してトークンをチケット登録（更新）時の値として
'		反映してください。
'		＜注意＞cscriptからコールしないとエラーになります。
'		Excel VBAからの呼び出し例）
'		    Set obj = CreateObject("WScript.shell")
'		    Set wExec = obj.exec("%ComSpec% /c cscript.exe " & cmdStr)
'
'	【変更履歴】
'		2019/01/19	1.00	初版
'--------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
'共通変数宣言
'--------------------------------------------------------------------------------------------------
'起動時引数
Dim argAPIKey
Dim argRedmineUri
Dim argFilePath
Dim argFileType

'オブジェクト
Dim FSO
Dim xmlHttp
Dim xmlDoc

'実行用変数
Dim RM_BASE_URI
Dim tmpToken
'--------------------------------------------------------------------------------------------------
'プログラムここから
'--------------------------------------------------------------------------------------------------
Call setEnvironment
Call checkArguments
Call uploadFile
'--------------------------------------------------------------------------------------------------
'サブルーチン群ここから
'--------------------------------------------------------------------------------------------------

'環境設定■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
Private Sub setEnvironment()
    '引数確認
    If Wscript.Arguments.Count <> 4 Then
        WScript.StdOut.WriteLine "999 引数が不足しています。プログラムを終了します。"
        WScript.Quit
    End If
    argAPIKey = Wscript.Arguments(0)
    argRedmineUri =Wscript.Arguments(1)
    argFilePath = Wscript.Arguments(2)
	argFileType = Wscript.Arguments(3)

	'検証用変数
'	argAPIKey = "eae538eaae18bd51cbb8781e5efe88448f184b2a"
'	argRedmineUri = "http://127.0.0.1:81/redmine/uploads.xml"
'	argFilePath = "G:\Zen\Documents\RedmineFileUploader\v1.0\アップロードファイル\供給地点特定番号検索レポート.xlsx"
'	argFileType = "binary"

	'RM接続用基本URI生成
	RM_BASE_URI = argRedmineUri & "?key=" & argAPIKey
	
	'オブジェクト生成
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP.3.0")
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
End Sub

'各種入力チェック■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
Private Sub checkArguments()
	'ファイル存在チェック
	If Not FSO.FileExists(argFilePath) Then
        WScript.Echo "999 アップロード対象のファイルが存在しません。プログラムを終了します。"
        WScript.Quit
	End If
	
	'ファイルタイプ判定
	If Not (argFileType = "ascii" Or argFileType = "binary") Then
        WScript.Echo "999 ファイルタイプの引数が不正です。プログラムを終了します。"
        WScript.Quit
	End If		
End Sub

'ファイルアップロード開始■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
Private Sub uploadFile()
	Dim sendData
	Dim obj
	
	sendData = getFileBytes( argFilePath, argFileType)
	xmlHttp.open "POST", RM_BASE_URI, False
	xmlHttp.setRequestHeader "Content-Type", "application/octet-stream"
	xmlHttp.send sendData
	If xmlHttp.status <> "201" Then
        WScript.Echo "999 アップロードが失敗しました。URIが正しいか確認してください。プログラムを終了します。（" & xmlHttp.status & " " & xmlHttp.statusText & "）"
        WScript.Quit
	End If
	
	'XMｌファイルパース
	xmlDoc.loadXML xmlHttp.responseText
	tmpToken = xmlDoc.getElementsByTagName("token")(0).nodeTypedValue

	'コンソール出力
	WScript.Echo xmlHttp.status & " " & xmlHttp.statusText
	WScript.Echo tmpToken

	Set xmlHttp = Nothing
	Set xmlDoc = Nothing

End Sub

'ファイルストリーム生成■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
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