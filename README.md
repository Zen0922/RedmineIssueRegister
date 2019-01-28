Redmineへのチケット登録を行うVBScript
====
Redmineへのチケット登録をユーザが任意にしていたディレクトリ内にあるissue.txtに基づきRedmine REST API経由で登録を行います。  
issue.txtはユーザにて生成する必要があります。  
本スクリプトはVBScriptで記述されているため、通常はWindowsで動作することが前提となります。  

## 読込ライブラリ
* MSXML2.XMLHTTP3.0  
* MSXML2.DOMDocument  
* FileSystemObject  
* Dictionary  
* ADODB.Stream  

## スクリプトの全体構成
    -+-Lib
     |  -Dictionary_MIME_ASCII.csv [MIME-TYPEとアップロード時にASCIIを使用するかを定義したファイル]
     |  
     +-Template [issue.txt等のテンプレート]  
     |  
     +RedmineFileUploader.vbs [REST APIでファイルをアップロードしtokenを返却するスクリプト]  
     +RedmineIssueRegister.vbs [REST APIでチケットを登録するスクリプト]  
     +lastResult.txt [最後に実行した際のコンソール表示内容を出力/自動生成]

## 起動と処理の流れ
issue.txtを格納したディレクトリを引数に、RedmineIssueRegister.vbsを起動します。  
`> RedmineIssueRegister.vbs [issue.txtを格納したディレクトリ]`

issue.txtおよび、ファイルの頭文字が"_（半角アンダーバー）"で始まるファイルはチケット登録情報とみなし、それ以外のファイルは添付ファイルとしてアップロードを行います。  
複数ファイルが存在した場合でも全てアップロードを行います。  

REST APIへの登録処理はXML形式で文字コードはUTF-8で行います。  
※今のところ文字コードの指定オプションはありません。  

登録に成功した場合はチケットIDを、それが以外の場合は999から始まるエラーコードまたは電文の詳細をコンソールに出力します。  
この結果はlastResult.txtでも確認できます。  
なお、電文の詳細が出力されるのは、HTTPでPOSTした結果NGだった場合に限ります。  

## issue.txt
issue.txtの文字コードはShift-JIS、改行コードはCRLFである必要があります。  
issue.txtの記述方法は以下の通りです。  

**ファイルに直接記述する場合**  
`[key]=[value]`  
[value]に改行が存在する場合は正常に動作しません。  
（理論上は[value]の改行コードがCRまたはLFであれば問題なさそうですが試していません。）

**別ファイルの内容を[value]とする場合**  
`[key]=<FILE>`  
`<FILE>`とそのまま記載ください。  
この場合、下記の規則のファイルを検出して読み込みします。 

    "_" & [key] & ".txt"
    Ex) key が descriptionの場合
    issue.txt
     description=<FILE>
    読込ファイル名
     _description.txt
 
**`[key]`に指定する内容**  
`[key]`に指定できる文字列は、Redmine REST APIに定義されているキーを指定してください。  
なお、チケットの登録先情報として以下の[key]は必ず記述してください。  

    _uri: RedmineのURIです。ルートを指定してください。
      Ex) _uri=http://127.0.0.1/redmine/

`_apikey: Redmine REST APIのAPIキーを指定してください。`  

**カスタムフィールドの扱い**  
　カスタムフィールドに値をセットする場合は、issue.txtにcustom_field:[custom_field_id]を記述します。  
`Ex) custom_field:2=[value]`    
　`<FILE>`を指定した場合、読み込むファイルは「_custom_field_2.txt」となります。  
 （":（コロン）"がアンダーバーになっていますので注意してください。）

## Dictionary_MIME_ASCII.csvの記述
Dictionary_MIME_ASCII.csvへの記述は以下の通りです。  
この一覧に存在しないファイルはアップロードできません。
`[拡張子],[MIME-TYPE],[ASCII-MODE]`  
`[拡張子]`は半角小文字で拡張子を記載してください。  
`[MIME-TYPE]`は`[拡張子]`に対応するMIME-TYPEを記述してください。  
`[ASCII-MODE]`は、ASCIIモードでアップロードすべきものには`1`を、それ以外の場合は`0`を反映します。  

## その他
* ファイルアップロードでテキストファイルが化ける  
化けるときは対象の拡張子のDictionary_MIME_ASCII.csvの[ASCII-MODE]を`0`にしてみてください。  
私の環境ではWin10のローカル環境では`1`を指定しても問題なくアップロードでき、会社の環境では`1`だとファイルが化けます。

* 複数の選択肢をとるカスタムフィールドについて  
現在対応しておりません。必要になったら対応します。
