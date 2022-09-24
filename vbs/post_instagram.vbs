Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'各シートの情報を定義
Public Const post_config_sheet = "投稿設定シート"
Public Const account_list_sheet = "アカウント情報"
Public Const post_history_sheet = "投稿履歴"
Public Const log_config_sheet = "ログ設定"
Public Const for_program_sheet = "プログラム用"

'プログラム用シートの各値のセル場所を指定
Public Const today_cell = "A3"
Public Const today_post_num_cell = "B3"
Public Const today_max_post_num_cell = "C3"
Public Const today_last_post_time_cell = "A5"

'投稿履歴シートの情報を定義
Public Const latest_post_row = 3
Public Const latest_post_num_cell = "B3"
Public Const latest_post_product_id_cell = "C3"
Public Const latest_post_product_title_cell = "D3"
Public Const delete_row = 33

'投稿設定シートの情報を定義
Public Const next_post_csv_row_cell = "J3"
Public Const next_post_product_id_cell = "K3"
Public Const next_post_product_title_cell = "L3"

Public Const selenium_sleep = 6000

Sub run_now()
  On Error GoTo ErrorHandler
    '実行時、投稿設定シートをアクティブにする
    Sheets(post_config_sheet).Select

    Dim log_file_path As String
    'ログ設定シートを読み込む
    With Worksheets(log_config_sheet)
      log_file_path = .log_file_path_field.Value
    End With

    'ログファイルを開く
    Open log_file_path For Append As #1

    Dim csv_path As String
    '投稿設定シートの入力値を読み込む
    With Worksheets(post_config_sheet)
      csv_path = .csv_path_field.Value

      flag_post_random = False
      log_post_random = "上から順"
      If (.post_order_field.Value = "ランダム") Then
        flag_post_random = True
        log_post_random = "ランダム"
      End If
      
      common_desc = .post_common_desc_field.Value

      limited_post_range = Split(.post_num_field.Value, ",")
      min_post_num = limited_post_range(0)
      max_post_num = limited_post_range(1)

      post_time_zone = Split(.post_time_zone_field.Value, ",")
      post_time_start = CDate(post_time_zone(0))
      post_time_end = CDate(post_time_zone(1))
      
      post_interval = .post_interval_field.Value
      
      wait_time_config = Split(.post_wait_time_field.Value, ",")
      min_wait = wait_time_config(0)
      max_wait = wait_time_config(1)

      next_post_csv_row = .Cells(3, "J").Value
      If (next_post_csv_row = "") Then
          .Cells(3, "J").Value = 2
          next_post_csv_row = 2
      End If
      

      Print #1, Date & " " & Time & " [INFO]下記の設定で投稿処理を開始します。" & vbCrLf _
      & "----------------------" & vbCrLf _
      & "CSVパス: " & csv_path & vbCrLf _
      & "投稿順: " & log_post_random & vbCrLf _
      & "キャプション: " & common_desc & vbCrLf _
      & "投稿回数の範囲: " & min_post_num & " - " & max_post_num & vbCrLf _
      & "投稿時間帯: " & post_time_start & " - " & post_time_end & vbCrLf _
      & "インターバル: " & post_interval & vbCrLf _
      & "実行前の待機時間: " & min_wait & " - " & max_wait & vbCrLf _
      & "投稿するCSVデータ: " & next_post_csv_row & " 行目" & vbCrLf _
      & "----------------------"

    End With

    'アカウント情報シートの設定値を読み込む
    auth_info = get_auth_info()
    account = auth_info(0)
    Password = auth_info(1)
    If account = "" Or Password = "" Then
      Print #1, Date & " " & Time & " [ERROR]アカウント情報の取得に失敗したため終了します。"
      Close #1
      Exit Sub
    End If
    Print #1, Date & " " & Time & " [INFO]アカウント情報の取得を取得しました。" & vbCrLf _
    & "----------------------" & vbCrLf _
    & "アカウント: " & Left(account, 4) & "********" & vbCrLf _
    & "パスワード: " & Left(Password, 1) & "*************" & vbCrLf _
    & "----------------------"

    'プログラム用データを読み込む
    With Worksheets(for_program_sheet)
      today = .Range(today_cell).Value
      '今日、初めての実行ならデータを更新
      With Worksheets(for_program_sheet)
        If Date <> today Then
          .Range(today_cell).Value = Date
          .Range(today_post_num_cell).Value = 0
          .Range(today_max_post_num_cell).Value = get_random_number_from_range(Int(min_post_num), Int(max_post_num))
          .Range(today_last_post_time_cell).Value = ""
          Print #1, Date & " " & Time & " [INFO]今日、初めての実行です。プログラム用データを初期化しました。"
        End If
      End With
      posted_num = .Range(today_post_num_cell).Value
      max_post_num = .Range(today_max_post_num_cell).Value
      last_post_time = .Range(today_last_post_time_cell).Value
      Print #1, Date & " " & Time & " [INFO]下記の設定で投稿処理を開始します。" & vbCrLf _
      & "----------------------" & vbCrLf _
      & "今日の投稿回数: " & posted_num & vbCrLf _
      & "今日の最大の投稿回数: " & max_post_num & vbCrLf _
      & "前回の投稿日時: " & last_post_time & vbCrLf _
      & "----------------------"

    End With


    '一日の投稿上限数を超えているなら終了
    If posted_num >= max_post_num Then
        Print #1, Date & " " & Time & " [INFO]一日の投稿上限数を超えているため終了します。"
        Close #1
        Exit Sub
    End If

    '時間外の実行なら終了
    If Time < post_time_start Or post_time_end < Time Then
        Print #1, Date & " " & Time & " [INFO]実行時間外のため終了します。現在の時間: " & Time
        Close #1
        Exit Sub
    End If

    'インターバル内の実行なら終了
    If last_post_time <> "" And DateDiff("n", last_post_time, Time) <= Int(post_interval) Then
        Print #1, Date & " " & Time & " [INFO]インターバル内の実行のため終了します。前回実行からの経過時間: " & DateDiff("n", last_post_time, Time)
        Close #1
        Exit Sub
    End If

    '指定された範囲内でランダムに待機
    sleep_time = get_random_number_from_range(Int(min_wait), Int(max_wait))
    Call Sleep(sleep_time)

    'CSVデータをメモリに読み込む
    Dim upload_data As Variant
    upload_data = read_csv(csv_path, True)
    Index = next_post_csv_row - 2
    Print #1, Date & " " & Time & " [INFO]CSVデータを読み込みました。" & vbCrLf _
    & "----------------------" & vbCrLf _
    & "商品管理番号: " & upload_data(Index, 0) & vbCrLf _
    & "商品タイトル: " & upload_data(Index, 1) & vbCrLf _
    & "価格: " & upload_data(Index, 2) & vbCrLf _
    & "商品ページURL: " & upload_data(Index, 3) & vbCrLf _
    & "ハッシュタグ: " & upload_data(Index, 4) & vbCrLf _
    & "文章1: " & upload_data(Index, 5) & vbCrLf _
    & "文章2: " & upload_data(Index, 6) & vbCrLf _
    & "文章3: " & upload_data(Index, 7) & vbCrLf _
    & "画像リンク: " & upload_data(Index, 8) & vbCrLf _
    & "CSV行: " & next_post_csv_row & vbCrLf _
    & "----------------------"

    If upload_data(Index, 0) = "" Then
      Print #1, Date & " " & Time & " [INFO]データが登録されていないためプログラムを終了します。"
      Close #1
      Exit Sub
    End If
    
    '共通文章を設定
    post_desc = Replace(common_desc, "{商品管理番号}", upload_data(Index, 0))
    post_desc = Replace(post_desc, "{商品タイトル}", upload_data(Index, 1))
    post_desc = Replace(post_desc, "{価格}", upload_data(Index, 2))
    post_desc = Replace(post_desc, "{商品ページURL}", upload_data(Index, 3))
    post_desc = Replace(post_desc, "{ハッシュタグ}", upload_data(Index, 4))
    post_desc = Replace(post_desc, "{文章1}", upload_data(Index, 5))
    post_desc = Replace(post_desc, "{文章2}", upload_data(Index, 6))
    post_desc = Replace(post_desc, "{文章3}", upload_data(Index, 7))
    post_desc = Replace(post_desc, "{画像リンク}", upload_data(Index, 8))
    Print #1, Date & " " & Time & " [INFO]キャプションを設定。" & vbCrLf _
    & "----------------------" & vbCrLf _
    & "キャプション: " & post_desc & vbCrLf _
    & "----------------------"

    'ブラウザを起動し、ログインページへ移動
    Dim driver As New Selenium.WebDriver
    driver.Start "Chrome"
    driver.Get "https://www.instagram.com/"
    driver.Window.Maximize

    Call Sleep(selenium_sleep)

    'ログインする
    field_login_account = "/html/body/div[1]/section/main/article/div[2]/div[1]/div[2]/form/div/div[1]/div/label/input"
    field_login_password = "/html/body/div[1]/section/main/article/div[2]/div[1]/div[2]/form/div/div[2]/div/label/input"
    button_login = "/html/body/div[1]/section/main/article/div[2]/div[1]/div[2]/form/div/div[3]/button"
    driver.FindElementByXPath(field_login_account).SendKeys account
    driver.FindElementByXPath(field_login_password).SendKeys Password
    Call Sleep(selenium_sleep)
    driver.FindElementByXPath(button_login).Click
    Call Sleep(10000)
    Print #1, Date & " " & Time & " [INFO]インスタグラムにログインしました。"

    'プロファイル保存に関する質問に「後で」で回答
    button_login_later = "/html/body/div[1]/div/div/div/div[1]/div/div/div/div[1]/section/main/div/div/div/div/button"
    driver.FindElementByXPath(button_login_later).Click
    Call Sleep(selenium_sleep)
    '通知に関する質問に「後で」で回答
    button_notice_later = "/html/body/div[1]/div/div/div/div[2]/div/div/div[1]/div/div[2]/div/div/div/div/div[2]/div/div/div[3]/button[2]"
    driver.FindElementByXPath(button_notice_later).Click
    Call Sleep(selenium_sleep)
    Print #1, Date & " " & Time & " [INFO]ログイン後の質問に回答しました。"

    'インスタグラムへ投稿する画像を設定する
    Dim image_urls As Variant
    image_urls = Split(upload_data(Index, 8), ",")

    '投稿ボタンをクリック
    button_post = "/html/body/div[1]/div/div/div/div[1]/div/div/div/div[1]/section/nav/div[2]/div/div/div[3]/div/div[3]/div/button"
    driver.FindElementByXPath(button_post).Click
    Call Sleep(selenium_sleep)
    '画像を選択
    button_select_image = "/html/body/div[1]/div/div/div/div[2]/div/div/div[1]/div/div[3]/div/div/div/div/div[2]/div/div/div/div[2]/div[1]/div/div/div[2]/div/button"
    driver.FindElementByXPath(button_select_image).Click
    Call Sleep(selenium_sleep)
    SendKeys image_urls(0)
    SendKeys "{ENTER}"
    Print #1, Date & " " & Time & " [INFO]画像を設定。画像URL: " & image_urls(0)
    Call Sleep(selenium_sleep)
    '右下にある「画像を追加」ボタンをクリック
    button_add_image = "/html/body/div[1]/div/div/div/div[2]/div/div/div[1]/div/div[3]/div/div/div/div/div[2]/div/div/div/div[2]/div[1]/div/div/div/div[3]/div/div[2]/div/button"
    driver.FindElementByXPath(button_add_image).Click
    Call Sleep(selenium_sleep)
    '2番目以降の画像を設定
    For i = 1 To UBound(image_urls)
      button_add_image2 = "/html/body/div[1]/div/div/div/div[2]/div/div/div[1]/div/div[3]/div/div/div/div/div[2]/div/div/div/div[2]/div[1]/div/div/div/div[3]/div/div[1]/div/div/div/div[2]/div"
      driver.FindElementByXPath(button_add_image2).Click
      Call Sleep(selenium_sleep)
      'ファイル選択ダイアログボックスにURLを入力
      SendKeys image_urls(i)
      SendKeys "{ENTER}"
      Print #1, Date & " " & Time & " [INFO]画像を設定。画像URL: " & image_urls(i)
      Call Sleep(selenium_sleep)
    Next i

    '画像を投稿
    button_next = "/html/body/div[1]/div/div/div/div[2]/div/div/div[1]/div/div[3]/div/div/div/div/div[2]/div/div/div/div[1]/div/div/div[3]/div/button"
    driver.FindElementByXPath(button_next).Click
    Call Sleep(selenium_sleep)
    button_next2 = "/html/body/div[1]/div/div/div/div[2]/div/div/div[1]/div/div[3]/div/div/div/div/div[2]/div/div/div/div[1]/div/div/div[3]/div/button"
    driver.FindElementByXPath(button_next2).Click
    Call Sleep(selenium_sleep)
    textarea_caption = "/html/body/div[1]/div/div/div/div[2]/div/div/div[1]/div/div[3]/div/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/div/div[2]/div[1]/textarea"
    driver.FindElementByXPath(textarea_caption).Click 
    SendKeys post_desc
    Call Sleep(selenium_sleep)
    button_share = "/html/body/div[1]/div/div/div/div[2]/div/div/div[1]/div/div[3]/div/div/div/div/div[2]/div/div/div/div[1]/div/div/div[3]/div/button"
    driver.FindElementByXPath(button_share).Click
    Call Sleep(20000)
    button_close = "/html/body/div[1]/div/div/div/div[2]/div/div/div[1]/div/div[2]/div"
    driver.FindElementByXPath(button_close).Click
    Call Sleep(selenium_sleep)

    Print #1, Date & " " & Time & " [INFO]画像を投稿しました。" & vbCrLf _
    & "----------------------" & vbCrLf _
    & "管理番号: " & upload_data(Index, 0) & vbCrLf _
    & "画像URL: " & upload_data(Index, 8) & vbCrLf _
    & "----------------------"

    '投稿履歴を更新
    Dim post_row As String
    Dim product_id As String
    Dim product_title As String
    post_row = upload_data(Index, 9)
    product_id = upload_data(Index, 0)
    product_title = upload_data(Index, 1)
    Call update_post_history(post_row, product_id, product_title)
    Print #1, Date & " " & Time & " [INFO]投稿履歴を更新しました。"

    '次回投稿のCSVデータを設定
    With Worksheets(post_config_sheet)
      '上から順に実行された場合
      If (Not flag_post_random) Then
        If (.Range(next_post_csv_row_cell).Value = "") Then
            .Range(next_post_csv_row_cell).Value = 2
        Else
            .Range(next_post_csv_row_cell).Value = .Range(next_post_csv_row_cell).Value + 1
        End If
      'ランダム実行された場合
      Else
          Randomize
          .Range(next_post_csv_row_cell).Value = get_random_number_from_range(2, get_csv_data_num(csv_path))
      End If

      Index = .Range(next_post_csv_row_cell).Value - 2
      .Range(next_post_product_id_cell).Value = upload_data(Index, 0)
      .Range(next_post_product_title_cell).Value = upload_data(Index, 1)
      Print #1, Date & " " & Time & " [INFO]次回、投稿するCSVデータを設定しました。" & vbCrLf _
      & "----------------------" & vbCrLf _
      & "CSV行: " & .Range(next_post_csv_row_cell).Value & vbCrLf _
      & "管理番号: " & .Range(next_post_product_id_cell).Value & vbCrLf _
      & "商品タイトル: " & .Range(next_post_product_title_cell).Value & vbCrLf _
      & "----------------------"
    End With

    With Worksheets(for_program_sheet)
      'プログラム用データを更新
      .Range(today_post_num_cell).Value = .Range(today_post_num_cell).Value + 1
      .Range(today_last_post_time_cell).Value = Time
      Print #1, Date & " " & Time & " [INFO]プログラム用データを更新しました。"
    End With

    Print #1, Date & " " & Time & " [INFO]プログラムを終了しました。"
    Close #1
    Exit Sub
  ErrorHandler:
    Print #1, Date & " " & Time & " [ERROR]エラーが発生しました。エラー番号:" & Err.Number & vbCrLf & _
    "エラーの種類:" & Err.Description
    Close #1
End Sub

'アカウント情報シートの設定値を読み込む
Private Function get_auth_info() As Variant
  'IG-User-ID, Access-token
  Dim auth_info(2) As Variant
  Dim row As Integer
  row = 3
  With Worksheets(account_list_sheet)
    auth_info(0) = ""
    auth_info(1) = ""
    Dim val As String
    For row = 3 To 13
        val = .Cells(row, "D").Value
        If (val = "〇") Then
            auth_info(0) = .Cells(row, "B").Value
            auth_info(1) = .Cells(row, "C").Value
            Exit For
        End If
        row = row + 1
    Next
    
    get_auth_info = auth_info
  End With
End Function

'指定した範囲から乱数を取得
Private Function get_random_number_from_range(min As Integer, max As Integer) As Integer
  Randomize
  get_random_number_from_range = Int((max - min + 1) * Rnd + min)
End Function

'CSVデータを読み込む
Private Function read_csv(csv_path As String, skip_header As Boolean) As Variant
  'CSVデータのサイズを定義
  Const data_size_1 = 100 '0-99
  Const data_size_2 = 10 ' 0-9
  Dim filesystem As Object
  Set filesystem = CreateObject("Scripting.FileSystemObject")
  
  Workbooks.Open csv_path

  row = Cells(Rows.Count, 1).End(xlUp).row

  'csvデータのサイズを確保
  Dim csv_data(data_size_1, data_size_2) As Variant

  Dim i As Integer
  Dim j As Integer
  For i = 0 To row - 1
      For j = 0 To data_size_2 - 1
        If (skip_header) Then
          csv_data(i, j) = Cells(i + 2, j + 1).Value
        Else
          csv_data(i, j) = Cells(i + 1, j + 1).Value
        End If
      Next

      '画像URLを取得(9列目から13列目まで)
      Dim image_url As String
      image_url = Cells(i + 2, 9).Value
      For k = 10 To 13
        If (skip_header) Then
          If Cells(i + 2, k).Value <> "" Then
            image_url = image_url + "," + Cells(i + 2, k).Value
          End If
        Else
          If Cells(i + 1, k).Value <> "" Then
            image_url = image_url + "," + Cells(i + 1, k).Value
          End If
        End If
      Next
      csv_data(i, 8) = image_url
  Next
  
  '(オプション)末尾に行数データを追加
  For i = 0 To row - 1
      If (skip_header) Then
        csv_data(i, data_size_2 - 1) = i + 2
      Else
        csv_data(i, data_size_2 - 1) = i + 1
      End If
  Next

  Workbooks(filesystem.GetFileName(csv_path)).Close SaveChanges:=False
  read_csv = csv_data
End Function

'投稿履歴を更新する
Private Function update_post_history(row As String, product_id As String, product_title As String) As String
  With Worksheets(post_history_sheet)
    .Rows(latest_post_row).Insert CopyOrigin:=xlFormatFromRightOrBelow
    .Range(latest_post_num_cell).Value = row
    .Range(latest_post_product_id_cell).Value = product_id
    .Range(latest_post_product_title_cell).Value = product_title

    '31件目を削除
    .Rows(delete_row).Delete
  End With
End Function

'CSVデータの個数をカウントする
Private Function get_csv_data_num(csv_path As String) As Integer
  Dim filesystem As Object
  Set filesystem = CreateObject("Scripting.FileSystemObject")
  Workbooks.Open Filename:=csv_path
  post_num = Cells(Rows.Count, 1).End(xlUp).row
  Workbooks(filesystem.GetFileName(csv_path)).Close SaveChanges:=False
  get_csv_data_num = post_num
End Function
