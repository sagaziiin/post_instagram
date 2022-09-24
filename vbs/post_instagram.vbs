Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'�e�V�[�g�̏����`
Public Const post_config_sheet = "���e�ݒ�V�[�g"
Public Const account_list_sheet = "�A�J�E���g���"
Public Const post_history_sheet = "���e����"
Public Const log_config_sheet = "���O�ݒ�"
Public Const for_program_sheet = "�v���O�����p"

'�v���O�����p�V�[�g�̊e�l�̃Z���ꏊ���w��
Public Const today_cell = "A3"
Public Const today_post_num_cell = "B3"
Public Const today_max_post_num_cell = "C3"
Public Const today_last_post_time_cell = "A5"

'���e�����V�[�g�̏����`
Public Const latest_post_row = 3
Public Const latest_post_num_cell = "B3"
Public Const latest_post_product_id_cell = "C3"
Public Const latest_post_product_title_cell = "D3"
Public Const delete_row = 33

'���e�ݒ�V�[�g�̏����`
Public Const next_post_csv_row_cell = "J3"
Public Const next_post_product_id_cell = "K3"
Public Const next_post_product_title_cell = "L3"

Public Const selenium_sleep = 6000

Sub run_now()
  On Error GoTo ErrorHandler
    '���s���A���e�ݒ�V�[�g���A�N�e�B�u�ɂ���
    Sheets(post_config_sheet).Select

    Dim log_file_path As String
    '���O�ݒ�V�[�g��ǂݍ���
    With Worksheets(log_config_sheet)
      log_file_path = .log_file_path_field.Value
    End With

    '���O�t�@�C�����J��
    Open log_file_path For Append As #1

    Dim csv_path As String
    '���e�ݒ�V�[�g�̓��͒l��ǂݍ���
    With Worksheets(post_config_sheet)
      csv_path = .csv_path_field.Value

      flag_post_random = False
      log_post_random = "�ォ�珇"
      If (.post_order_field.Value = "�����_��") Then
        flag_post_random = True
        log_post_random = "�����_��"
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
      

      Print #1, Date & " " & Time & " [INFO]���L�̐ݒ�œ��e�������J�n���܂��B" & vbCrLf _
      & "----------------------" & vbCrLf _
      & "CSV�p�X: " & csv_path & vbCrLf _
      & "���e��: " & log_post_random & vbCrLf _
      & "�L���v�V����: " & common_desc & vbCrLf _
      & "���e�񐔂͈̔�: " & min_post_num & " - " & max_post_num & vbCrLf _
      & "���e���ԑ�: " & post_time_start & " - " & post_time_end & vbCrLf _
      & "�C���^�[�o��: " & post_interval & vbCrLf _
      & "���s�O�̑ҋ@����: " & min_wait & " - " & max_wait & vbCrLf _
      & "���e����CSV�f�[�^: " & next_post_csv_row & " �s��" & vbCrLf _
      & "----------------------"

    End With

    '�A�J�E���g���V�[�g�̐ݒ�l��ǂݍ���
    auth_info = get_auth_info()
    account = auth_info(0)
    Password = auth_info(1)
    If account = "" Or Password = "" Then
      Print #1, Date & " " & Time & " [ERROR]�A�J�E���g���̎擾�Ɏ��s�������ߏI�����܂��B"
      Close #1
      Exit Sub
    End If
    Print #1, Date & " " & Time & " [INFO]�A�J�E���g���̎擾���擾���܂����B" & vbCrLf _
    & "----------------------" & vbCrLf _
    & "�A�J�E���g: " & Left(account, 4) & "********" & vbCrLf _
    & "�p�X���[�h: " & Left(Password, 1) & "*************" & vbCrLf _
    & "----------------------"

    '�v���O�����p�f�[�^��ǂݍ���
    With Worksheets(for_program_sheet)
      today = .Range(today_cell).Value
      '�����A���߂Ă̎��s�Ȃ�f�[�^���X�V
      With Worksheets(for_program_sheet)
        If Date <> today Then
          .Range(today_cell).Value = Date
          .Range(today_post_num_cell).Value = 0
          .Range(today_max_post_num_cell).Value = get_random_number_from_range(Int(min_post_num), Int(max_post_num))
          .Range(today_last_post_time_cell).Value = ""
          Print #1, Date & " " & Time & " [INFO]�����A���߂Ă̎��s�ł��B�v���O�����p�f�[�^�����������܂����B"
        End If
      End With
      posted_num = .Range(today_post_num_cell).Value
      max_post_num = .Range(today_max_post_num_cell).Value
      last_post_time = .Range(today_last_post_time_cell).Value
      Print #1, Date & " " & Time & " [INFO]���L�̐ݒ�œ��e�������J�n���܂��B" & vbCrLf _
      & "----------------------" & vbCrLf _
      & "�����̓��e��: " & posted_num & vbCrLf _
      & "�����̍ő�̓��e��: " & max_post_num & vbCrLf _
      & "�O��̓��e����: " & last_post_time & vbCrLf _
      & "----------------------"

    End With


    '����̓��e������𒴂��Ă���Ȃ�I��
    If posted_num >= max_post_num Then
        Print #1, Date & " " & Time & " [INFO]����̓��e������𒴂��Ă��邽�ߏI�����܂��B"
        Close #1
        Exit Sub
    End If

    '���ԊO�̎��s�Ȃ�I��
    If Time < post_time_start Or post_time_end < Time Then
        Print #1, Date & " " & Time & " [INFO]���s���ԊO�̂��ߏI�����܂��B���݂̎���: " & Time
        Close #1
        Exit Sub
    End If

    '�C���^�[�o�����̎��s�Ȃ�I��
    If last_post_time <> "" And DateDiff("n", last_post_time, Time) <= Int(post_interval) Then
        Print #1, Date & " " & Time & " [INFO]�C���^�[�o�����̎��s�̂��ߏI�����܂��B�O����s����̌o�ߎ���: " & DateDiff("n", last_post_time, Time)
        Close #1
        Exit Sub
    End If

    '�w�肳�ꂽ�͈͓��Ń����_���ɑҋ@
    sleep_time = get_random_number_from_range(Int(min_wait), Int(max_wait))
    Call Sleep(sleep_time)

    'CSV�f�[�^���������ɓǂݍ���
    Dim upload_data As Variant
    upload_data = read_csv(csv_path, True)
    Index = next_post_csv_row - 2
    Print #1, Date & " " & Time & " [INFO]CSV�f�[�^��ǂݍ��݂܂����B" & vbCrLf _
    & "----------------------" & vbCrLf _
    & "���i�Ǘ��ԍ�: " & upload_data(Index, 0) & vbCrLf _
    & "���i�^�C�g��: " & upload_data(Index, 1) & vbCrLf _
    & "���i: " & upload_data(Index, 2) & vbCrLf _
    & "���i�y�[�WURL: " & upload_data(Index, 3) & vbCrLf _
    & "�n�b�V���^�O: " & upload_data(Index, 4) & vbCrLf _
    & "����1: " & upload_data(Index, 5) & vbCrLf _
    & "����2: " & upload_data(Index, 6) & vbCrLf _
    & "����3: " & upload_data(Index, 7) & vbCrLf _
    & "�摜�����N: " & upload_data(Index, 8) & vbCrLf _
    & "CSV�s: " & next_post_csv_row & vbCrLf _
    & "----------------------"

    If upload_data(Index, 0) = "" Then
      Print #1, Date & " " & Time & " [INFO]�f�[�^���o�^����Ă��Ȃ����߃v���O�������I�����܂��B"
      Close #1
      Exit Sub
    End If
    
    '���ʕ��͂�ݒ�
    post_desc = Replace(common_desc, "{���i�Ǘ��ԍ�}", upload_data(Index, 0))
    post_desc = Replace(post_desc, "{���i�^�C�g��}", upload_data(Index, 1))
    post_desc = Replace(post_desc, "{���i}", upload_data(Index, 2))
    post_desc = Replace(post_desc, "{���i�y�[�WURL}", upload_data(Index, 3))
    post_desc = Replace(post_desc, "{�n�b�V���^�O}", upload_data(Index, 4))
    post_desc = Replace(post_desc, "{����1}", upload_data(Index, 5))
    post_desc = Replace(post_desc, "{����2}", upload_data(Index, 6))
    post_desc = Replace(post_desc, "{����3}", upload_data(Index, 7))
    post_desc = Replace(post_desc, "{�摜�����N}", upload_data(Index, 8))
    Print #1, Date & " " & Time & " [INFO]�L���v�V������ݒ�B" & vbCrLf _
    & "----------------------" & vbCrLf _
    & "�L���v�V����: " & post_desc & vbCrLf _
    & "----------------------"

    '�u���E�U���N�����A���O�C���y�[�W�ֈړ�
    Dim driver As New Selenium.WebDriver
    driver.Start "Chrome"
    driver.Get "https://www.instagram.com/"
    driver.Window.Maximize

    Call Sleep(selenium_sleep)

    '���O�C������
    field_login_account = "/html/body/div[1]/section/main/article/div[2]/div[1]/div[2]/form/div/div[1]/div/label/input"
    field_login_password = "/html/body/div[1]/section/main/article/div[2]/div[1]/div[2]/form/div/div[2]/div/label/input"
    button_login = "/html/body/div[1]/section/main/article/div[2]/div[1]/div[2]/form/div/div[3]/button"
    driver.FindElementByXPath(field_login_account).SendKeys account
    driver.FindElementByXPath(field_login_password).SendKeys Password
    Call Sleep(selenium_sleep)
    driver.FindElementByXPath(button_login).Click
    Call Sleep(10000)
    Print #1, Date & " " & Time & " [INFO]�C���X�^�O�����Ƀ��O�C�����܂����B"

    '�v���t�@�C���ۑ��Ɋւ��鎿��Ɂu��Łv�ŉ�
    button_login_later = "/html/body/div[1]/div/div/div/div[1]/div/div/div/div[1]/section/main/div/div/div/div/button"
    driver.FindElementByXPath(button_login_later).Click
    Call Sleep(selenium_sleep)
    '�ʒm�Ɋւ��鎿��Ɂu��Łv�ŉ�
    button_notice_later = "/html/body/div[1]/div/div/div/div[2]/div/div/div[1]/div/div[2]/div/div/div/div/div[2]/div/div/div[3]/button[2]"
    driver.FindElementByXPath(button_notice_later).Click
    Call Sleep(selenium_sleep)
    Print #1, Date & " " & Time & " [INFO]���O�C����̎���ɉ񓚂��܂����B"

    '�C���X�^�O�����֓��e����摜��ݒ肷��
    Dim image_urls As Variant
    image_urls = Split(upload_data(Index, 8), ",")

    '���e�{�^�����N���b�N
    button_post = "/html/body/div[1]/div/div/div/div[1]/div/div/div/div[1]/section/nav/div[2]/div/div/div[3]/div/div[3]/div/button"
    driver.FindElementByXPath(button_post).Click
    Call Sleep(selenium_sleep)
    '�摜��I��
    button_select_image = "/html/body/div[1]/div/div/div/div[2]/div/div/div[1]/div/div[3]/div/div/div/div/div[2]/div/div/div/div[2]/div[1]/div/div/div[2]/div/button"
    driver.FindElementByXPath(button_select_image).Click
    Call Sleep(selenium_sleep)
    SendKeys image_urls(0)
    SendKeys "{ENTER}"
    Print #1, Date & " " & Time & " [INFO]�摜��ݒ�B�摜URL: " & image_urls(0)
    Call Sleep(selenium_sleep)
    '�E���ɂ���u�摜��ǉ��v�{�^�����N���b�N
    button_add_image = "/html/body/div[1]/div/div/div/div[2]/div/div/div[1]/div/div[3]/div/div/div/div/div[2]/div/div/div/div[2]/div[1]/div/div/div/div[3]/div/div[2]/div/button"
    driver.FindElementByXPath(button_add_image).Click
    Call Sleep(selenium_sleep)
    '2�Ԗڈȍ~�̉摜��ݒ�
    For i = 1 To UBound(image_urls)
      button_add_image2 = "/html/body/div[1]/div/div/div/div[2]/div/div/div[1]/div/div[3]/div/div/div/div/div[2]/div/div/div/div[2]/div[1]/div/div/div/div[3]/div/div[1]/div/div/div/div[2]/div"
      driver.FindElementByXPath(button_add_image2).Click
      Call Sleep(selenium_sleep)
      '�t�@�C���I���_�C�A���O�{�b�N�X��URL�����
      SendKeys image_urls(i)
      SendKeys "{ENTER}"
      Print #1, Date & " " & Time & " [INFO]�摜��ݒ�B�摜URL: " & image_urls(i)
      Call Sleep(selenium_sleep)
    Next i

    '�摜�𓊍e
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

    Print #1, Date & " " & Time & " [INFO]�摜�𓊍e���܂����B" & vbCrLf _
    & "----------------------" & vbCrLf _
    & "�Ǘ��ԍ�: " & upload_data(Index, 0) & vbCrLf _
    & "�摜URL: " & upload_data(Index, 8) & vbCrLf _
    & "----------------------"

    '���e�������X�V
    Dim post_row As String
    Dim product_id As String
    Dim product_title As String
    post_row = upload_data(Index, 9)
    product_id = upload_data(Index, 0)
    product_title = upload_data(Index, 1)
    Call update_post_history(post_row, product_id, product_title)
    Print #1, Date & " " & Time & " [INFO]���e�������X�V���܂����B"

    '���񓊍e��CSV�f�[�^��ݒ�
    With Worksheets(post_config_sheet)
      '�ォ�珇�Ɏ��s���ꂽ�ꍇ
      If (Not flag_post_random) Then
        If (.Range(next_post_csv_row_cell).Value = "") Then
            .Range(next_post_csv_row_cell).Value = 2
        Else
            .Range(next_post_csv_row_cell).Value = .Range(next_post_csv_row_cell).Value + 1
        End If
      '�����_�����s���ꂽ�ꍇ
      Else
          Randomize
          .Range(next_post_csv_row_cell).Value = get_random_number_from_range(2, get_csv_data_num(csv_path))
      End If

      Index = .Range(next_post_csv_row_cell).Value - 2
      .Range(next_post_product_id_cell).Value = upload_data(Index, 0)
      .Range(next_post_product_title_cell).Value = upload_data(Index, 1)
      Print #1, Date & " " & Time & " [INFO]����A���e����CSV�f�[�^��ݒ肵�܂����B" & vbCrLf _
      & "----------------------" & vbCrLf _
      & "CSV�s: " & .Range(next_post_csv_row_cell).Value & vbCrLf _
      & "�Ǘ��ԍ�: " & .Range(next_post_product_id_cell).Value & vbCrLf _
      & "���i�^�C�g��: " & .Range(next_post_product_title_cell).Value & vbCrLf _
      & "----------------------"
    End With

    With Worksheets(for_program_sheet)
      '�v���O�����p�f�[�^���X�V
      .Range(today_post_num_cell).Value = .Range(today_post_num_cell).Value + 1
      .Range(today_last_post_time_cell).Value = Time
      Print #1, Date & " " & Time & " [INFO]�v���O�����p�f�[�^���X�V���܂����B"
    End With

    Print #1, Date & " " & Time & " [INFO]�v���O�������I�����܂����B"
    Close #1
    Exit Sub
  ErrorHandler:
    Print #1, Date & " " & Time & " [ERROR]�G���[���������܂����B�G���[�ԍ�:" & Err.Number & vbCrLf & _
    "�G���[�̎��:" & Err.Description
    Close #1
End Sub

'�A�J�E���g���V�[�g�̐ݒ�l��ǂݍ���
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
        If (val = "�Z") Then
            auth_info(0) = .Cells(row, "B").Value
            auth_info(1) = .Cells(row, "C").Value
            Exit For
        End If
        row = row + 1
    Next
    
    get_auth_info = auth_info
  End With
End Function

'�w�肵���͈͂��痐�����擾
Private Function get_random_number_from_range(min As Integer, max As Integer) As Integer
  Randomize
  get_random_number_from_range = Int((max - min + 1) * Rnd + min)
End Function

'CSV�f�[�^��ǂݍ���
Private Function read_csv(csv_path As String, skip_header As Boolean) As Variant
  'CSV�f�[�^�̃T�C�Y���`
  Const data_size_1 = 100 '0-99
  Const data_size_2 = 10 ' 0-9
  Dim filesystem As Object
  Set filesystem = CreateObject("Scripting.FileSystemObject")
  
  Workbooks.Open csv_path

  row = Cells(Rows.Count, 1).End(xlUp).row

  'csv�f�[�^�̃T�C�Y���m��
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

      '�摜URL���擾(9��ڂ���13��ڂ܂�)
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
  
  '(�I�v�V����)�����ɍs���f�[�^��ǉ�
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

'���e�������X�V����
Private Function update_post_history(row As String, product_id As String, product_title As String) As String
  With Worksheets(post_history_sheet)
    .Rows(latest_post_row).Insert CopyOrigin:=xlFormatFromRightOrBelow
    .Range(latest_post_num_cell).Value = row
    .Range(latest_post_product_id_cell).Value = product_id
    .Range(latest_post_product_title_cell).Value = product_title

    '31���ڂ��폜
    .Rows(delete_row).Delete
  End With
End Function

'CSV�f�[�^�̌����J�E���g����
Private Function get_csv_data_num(csv_path As String) As Integer
  Dim filesystem As Object
  Set filesystem = CreateObject("Scripting.FileSystemObject")
  Workbooks.Open Filename:=csv_path
  post_num = Cells(Rows.Count, 1).End(xlUp).row
  Workbooks(filesystem.GetFileName(csv_path)).Close SaveChanges:=False
  get_csv_data_num = post_num
End Function
