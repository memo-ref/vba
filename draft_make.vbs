Option Explicit

'※予め、Alt + F11 ＞ VBE画面 ＞ ツール ＞ 参照設定 ＞ Microsoft Outlook XX.0 Object Library にチェック入れをしておく必要があります。

Sub draft_make()

    'メールのテンプレートのシート名の変数
    Dim templateName As String
    
    'メールのテンプレートのシートを読み込むオブジェクト
    Dim wsTmplt As Worksheet
    
    'outlookと接続
    Dim app As Outlook.Application
    Set app = New Outlook.Application
    
    '下書きメールのオブジェクト
    Dim myMail As Outlook.MailItem
   
    
    'メールアドレス一覧を取得する
    Dim wsMailList As Worksheet
    Set wsMailList = Worksheets("メールアドレス一覧")
    
    'メールアドレス一覧の最終行番号を取得する
    Dim gmax As Long
    gmax = wsMailList.Cells(Rows.Count, 1).End(xlUp).Row
    
    Debug.Print ("gmax:" & gmax)
    
    
    'メールアドレス一覧の行番号の変数
    Dim g As Long
    
    
    'メールアドレス
    Dim meruado As String
    
    
    '下書きメールの作成件数をカウントする変数
    Dim countMail As Long
    
    For g = 2 To gmax
    
        '作成される下書きメールのオブジェクト
        Set myMail = app.CreateItem(olMailItem)
        
        
        'メールアドレス
        meruado = wsMailList.Cells(g, 1).Value
        Debug.Print (meruado)
        
        
        'テンプレートのシート名
        templateName = wsMailList.Cells(g, 2).Value
        Debug.Print (templateName)
    
    
        Set wsTmplt = WorkSheets(templateName)
    
        
		'メール宛先
		myMail.To = meruado
		
		'メールの形式
		myMail.BodyFormat = olFormatPlain
		
		'メール件名
		myMail.Subject = wsTmplt.Range("B1").Value
		
		'メール本文
		myMail.Body = wsTmplt.Range("B2").Value
		
		'下書き保存
		myMail.Save
		
		
		'表示したい場合はコメントアウトを外す
		'myMail.Display
		
        
        Set myMail = Nothing
        
        '作成件数＋１
        countMail = countMail + 1
    
    Next
    
    
    Set wsTmplt = Nothing
    
    Set app = Nothing
    
    MsgBox "下書きメールを　" & countMail & "件　作成しました。"

End Sub
