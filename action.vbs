'タスクスケジューラ実行用VBS
'Excel操作を可能にする為のオブジェクトを作成
Dim obj
Set obj = WScript.CreateObject("Excel.Application")

'Excel処理を画面表示
obj.Visible = true
obj.DisplayAlerts = False  'ポップアップメッセージを非表示にする
obj.AutomationSecurity = 1 'マクロを有効にする


'バッチファイルでこのファイルを実行する際の引数設定
'第一引数：指定したExcelマクロを開く
Dim xlWorkbook
Set xlWorkbook = obj.Workbooks.Open(WScript.Arguments(0))
'第二引数：指定したマクロを実行
obj.Application.Run WScript.Arguments(1)

xlWorkbook.Close True
obj.Quit

WScript.Quit(100)