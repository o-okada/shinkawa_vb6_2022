Attribute VB_Name = "Calculation"
'******************************************************************************
'モジュール名：Calculation
'
'******************************************************************************
Option Explicit
Option Base 1

Public Const INFINITE = -1&
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwAccess As Long, ByVal flnherit As Integer, ByVal hObject As Long) As Long

Public p_name           As String                                       'ポンプ履歴表示の時の施設名
Public isRAIN           As String                                       '使用雨量  "01"=気象庁   "02"=FRICS
Public isPump           As String                                       '"00"=ポンプあり  "01"=ポンプなし
Public FRICS            As Boolean                                      'FRICS雨量で計算するときはTRUE
Public KISYO            As Boolean                                      '気象庁雨量で計算するときはTRUE
Public Pump_FULL_Data(100)   As String                                  'フルポンプデータ
Public Pump_FULL_num(100)    As Long
Public PDF_Date         As Date                                         'PDF出力ファイル名に使う
Public Rsa_Mag          As Single                                       '準線形プログラムのRSA倍率
Public Yosoku_Time_F    As String
Public Yosoku_Time_K    As String

'******************************************************************************
'サブルーチン：ADD_ERROR_Message(msg As String)
'処理概要：
'******************************************************************************
Sub ADD_ERROR_Message(msg As String)
    Error_Message_n = Error_Message_n + 1
    Error_Message(Error_Message_n) = msg
End Sub

'******************************************************************************
'サブルーチン：Awaito_Time_Read(Cat As String, dw As Date)
'処理概要：
'待機時刻をディスクから読み込む
'******************************************************************************
Sub Awaito_Time_Read(Cat As String, dw As Date)
    Dim nf As Long
    Dim F  As String
    Dim buf As String
    Select Case Cat
        Case "気象庁"
            F = App.Path & "\data\Await_Time_KISYOU.dat"
        Case "FRICS"
            F = App.Path & "\data\Await_Time_FRICS.dat"
        Case Else
            Exit Sub
    End Select
    nf = FreeFile
    Open F For Input As #nf
    Input #nf, buf
    dw = CDate(buf)
    Close #nf
End Sub

'******************************************************************************
'サブルーチン：Awaito_Time_Write(Cat As String, dw As Date)
'処理概要：
'待機時刻をディスクに書き込む
'******************************************************************************
Sub Awaito_Time_Write(Cat As String, dw As Date)
    Dim nf As Long
    Dim F  As String
    Select Case Cat
        Case "気象庁"
            F = App.Path & "\data\Await_Time_Kisyou.dat"
        Case "FRICS"
            F = App.Path & "\data\Await_Time_FRICS.dat"
        Case Else
            MsgBox "ここに来てはいけません。"
    End Select
    nf = FreeFile
    Open F For Output As #nf
    Print #nf, Format(dw, "yyyy/mm/dd hh:nn")
    Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
    Close #nf
End Sub

'******************************************************************************
'サブルーチン：Data_Time_Rewrite(da As String, FL As String)
'処理概要：
'******************************************************************************
Sub Data_Time_Rewrite(da As String, FL As String)
    Dim dw  As String
    Dim nf  As Long
    dw = Format(Now, "yyyy/mm/dd hh:nn")
    nf = FreeFile
    Open FL For Output As #nf
    Print #nf, da
    Print #nf, dw
    Close #nf
End Sub

'******************************************************************************
'サブルーチン：H_to_Pump()
'処理概要：
'実績水位からポンプのデータを作成
'******************************************************************************
Sub H_to_Pump()
End Sub

'******************************************************************************
'サブルーチン：Pre_Pump()
'処理概要：
'フルポンプデータを読み込む
'******************************************************************************
Sub Pre_Pump()
    Dim i      As Long
    Dim j      As Long
    Dim nf     As Long
    Dim buf    As String
    Dim c      As String
    LOG_Out "In  Pre_Pump"
    nf = FreeFile
    Open App.Path & "\work\Pump.org" For Input As #nf
    i = 0
    Do Until EOF(nf)
        i = i + 1
        Line Input #nf, buf
        Pump_FULL_Data(i) = buf
        c = Mid(buf, 14, 2)
        If IsNumeric(c) Then
            j = CLng(c)
            Pump_FULL_num(j) = i
        End If
    Loop
    Close #nf
    LOG_Out "Out Pre_Pump"
End Sub

'******************************************************************************
'サブルーチン：Prediction_CAL_By_KISYO_Veri(manu As Boolean)
'処理概要：
'気象庁実績降雨計算検証用
'******************************************************************************
Sub Prediction_CAL_By_KISYO_Veri(manu As Boolean)
    Dim dwj     As Date
    Dim dwy     As Date
    Dim irc     As Boolean
    Dim jrc     As Long
    Dim rc      As Boolean
    Dim i       As Integer
    Dim ns      As Long
    Dim ts      As Long
    LOG_Out "IN Prediction_CAL_By_KISYO 気象庁雨量による洪水予測開始 現時刻=" & Format(jgd, "yyyy/mm/dd/hh:nn")
    Froude = 0#
    isRAIN = "01"                                                       '"01"=気象庁  "02"=FRICS
    isPump = "00"                                                       '"00"=ノーマル "01"=ポンプ停止
    Screen.MousePointer = vbHourglass
    '久地野と五条上流端流量
    JRADAR = 0
    If MAIN.Check2 Then
        dwy = DateAdd("h", 3, jgd)
        MDB_気象庁レーダー実績2 jsd, dwy, dwj, irc
        If dwy <> dwj Then
            MsgBox "気象庁実績雨量に必要とする雨量が格納されていません。" & vbCrLf & _
                    "  jsd=" & Format(jsd, "yyyy/mm/dd hh:nn") & vbCrLf & _
                    "  dwy=" & Format(dwy, "yyyy/mm/dd hh:nn")
            End
        End If
        MDB_洗堰 jsd, jgd, jrc
        If jrc > 1 Then
            LOG_Out "気象庁雨量で計算時に洗堰がまだ洪水予測システムに取り込まれませんでした。越流量=0として計算します。"
            ORA_Message_Out "洗堰越流量データ受信", "気象庁雨量で計算時に洗堰がまだ洪水予測システムに取り込まれませんでした。越流量=0として計算します。", 1
        End If
        レーダー雨量出力_Veri
        JRADAR = 1
    End If
    ポンプ雛型データ読み込み
    ポンプ能力表読み込み
    ポンプデータ作成 jgd
    Set_Pump                                                            '水位に応じた稼動、停止ポンプを設定する。
    Flood_Data_Write_For_Calc
    Message.Label1 = "予 測 計 算 実 行 中"
    'Message.Label1 = "ＳＨＩＮＫ１０　実 行 中"
    Message.Show
    Message.ZOrder 0
    Message.Refresh
    ChDir App.Path & "\WORK"
    'Call WaitForProcessToEnd("RRSHINK10.EXE")                          '久地野フィードバック有り
    'Call WaitForProcessToEnd("RRSHINK10NF.EXE")                        '久地野フィードバック無し
    '******************************************************
    'Ver1.0.0 修正開始 2015/08/06 O.OKADA
    '******************************************************
    'Call WaitForProcessToEnd("New_RSHINK.EXE")                          'なにも無し
    Call WaitForProcessToEnd("D:\SHINKAWA\レーダー洪水予測\WORK\New_RSHINK.EXE")                          'なにも無し
    '******************************************************
    'Ver1.0.0 修正終了 2015/08/06 O.OKADA
    '******************************************************
    'Message.Label1 = "Ｎ Ｓ Ｋ　実 行 中"
    Message.Refresh
    Flood_Data_Write_For_Calc1
    'Cal_Initial_flow_profile irc                                       '全て不定流で計算するため初期水面形は固定ファイルとし不等流計算は使わない
    'If irc = False Then Exit Sub
    '******************************************************
    'Ver1.0.0 修正開始 2015/08/06 O.OKADA
    '******************************************************
    'Call WaitForProcessToEnd("NEWNSKG2.EXE")
    Call WaitForProcessToEnd("D:\SHINKAWA\レーダー洪水予測\WORK\NEWNSKG2.EXE")                          'なにも無し
    '******************************************************
    'Ver1.0.0 修正開始 2015/08/06 O.OKADA
    '******************************************************
    'Message.Hide
    Screen.MousePointer = vbDefault
    ChDir App.Path
    Input_Yosoku irc
    If Not irc Then
        Exit Sub
    End If
    Category = True                                                     'True=不定流計算時
    'If (Froude > 1.9) Or (HO(5, Now_Step) <= Un_Cal) Then
    If (HO(5, Now_Step) <= Un_Cal) Then
        Category = False
        EMG_Cal irc
        If Not manu Then AutoDrive.Timer1.Enabled = True
        If irc = False Then Exit Sub
    Else
        '縦断ＤＢ出力用
        For i = 1 To nd                                                 'nd=断面数(74)
            YHJ(0, i) = HQ(1, i, NT - 18)                               '現時刻
            YHJ(1, i) = HQ(1, i, NT - 12)                               '予測+1
            YHJ(2, i) = HQ(1, i, NT - 6)                                '予測+2
            YHJ(3, i) = HQ(1, i, NT)                                    '予測+3
        Next i
        For i = 1 To 5
            FeedBack i
        Next i
        '縦断補正用
        For i = 1 To 5
            ns = V_Sec_Num(i)                                           'V_Sec_Num(nr)は不定流上の断面位置を表す
            OHJ(0, i) = HQ(1, ns, NT - 18)
            OHJ(1, i) = HQ(1, ns, NT - 12)
            OHJ(2, i) = HQ(1, ns, NT - 6)
            OHJ(3, i) = HQ(1, ns, NT)
        Next i
    End If
    If MDBx Then MDB_履歴_Write                                         'データベースに予測値の書き込み
    Load Graph3
    Graph3.Show
    Graph3.Refresh
    If Verification2.Check3 = vbChecked Then
        Graph3.VSPDF1.ConvertDocument Graph3.VSP, App.Path & "\" & Format(PDF_Date, "yyyy_mm_dd") & "_Hydro.pdf"
    End If
    '洪水予報文案作成                                                   'テスト時はここを生かす
    If DBX_ora Then                                                     '計算結果を
        ORA_DataBase_Connection
        If OraDB_OK Then                                                '県庁ＤＢが使用可能なとき
            '洪水予報文案作成
            LOG_Out "気象庁雨量　横断予測水位書き込み開始"
            ORA_SUII_YOSOKU_KIJYUN_PUT rc
            LOG_Out "気象庁雨量　横断予測水位書き込み終了"
            LOG_Out "気象庁雨量　縦断予測水位書き込み終了"
            ORA_SUII_YOSOKU_JYUDAN_PUT rc
            LOG_Out "気象庁雨量　縦断予測水位書き込み終了"
            ORA_DataBase_Close
        End If
    Else
    End If
    LOG_Out "IN Prediction_CAL_By_KISYO 気象庁雨量による洪水予測終了 現時刻=" & Format(jgd, "yyyy/mm/dd/hh:nn")
    Message.Hide
    If Not manu Then
        Short_Break 2
        Unload Graph3
        AutoDrive.Timer1.Enabled = True
    End If
End Sub

'******************************************************************************
'サブルーチン：Prediction_CAL_By_KISYO(manu As Boolean)
'処理概要：
'******************************************************************************
Sub Prediction_CAL_By_KISYO(manu As Boolean)
'■■■修正開始2016/03/04■■■
'計算処理でエラーが発生した場合、計算中画面がポップアップしたままとなるため。
On Error GoTo ERR1:
'■■■修正終了2016/03/04■■■
    Dim dwj     As Date
    Dim dwy     As Date
    Dim irc     As Boolean
    Dim jrc     As Long
    Dim rc      As Boolean
    Dim i       As Integer
    Dim ns      As Long
    Dim ts      As Long
    LOG_Out "IN Prediction_CAL_By_KISYO 気象庁雨量による洪水予測開始 現時刻=" & Format(jgd, "yyyy/mm/dd/hh:nn")
    Froude = 0#
    isRAIN = "01"                                                       '"01"=気象庁  "02"=FRICS
    isPump = "00"                                                       '"00"=ノーマル "01"=ポンプ停止
    Screen.MousePointer = vbHourglass
    '久地野と五条上流端流量
    JRADAR = 0
    If MAIN.Check2 Then
        MDB_気象庁レーダー実績2 jsd, jgd, dwj, irc
        If dwj < jgd Then jgd = dwj
        dwj = DateAdd("h", 1, jgd)
        dwy = DateAdd("h", 3, jgd)
        MDB_気象庁レーダー予測2 dwj, dwy, irc
        If irc = False Then
            If manu Then
                MsgBox "気象庁予測雨量がまだ未受信か登録されていません。" & vbCrLf & _
                        "現時刻を再設定して下さい。"
            End If
            LOG_Out "気象庁予測雨量がまだ未受信か登録されていません、計算をスキップします。"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        MDB_洗堰 jsd, jgd, jrc
        Select Case jrc
            Case 0                                                      '正常取得
            Case 1                                                      '10分前取得
                LOG_Out "洗堰データが前時刻データを使う。"
                ORA_Message_Out "洗堰越流量データ受信", "気象庁雨量による計算において、洗堰データが取り込まれませんでした。前時刻データで計算します。", 1
            Case 2                                                      '取得できず
                LOG_Out "洗堰データが10分前も取得できません、データを0としこまま計算します。"
                ORA_Message_Out "洗堰越流量データ受信", "気象庁雨量による計算において、洗堰データが2時刻以上連続して取り込まれませんでした。越流量=0として計算します。", 1
        End Select
        レーダー雨量出力
        JRADAR = 1
    End If
    ポンプ雛型データ読み込み
    ポンプ能力表読み込み
    ポンプデータ作成 jgd
    Set_Pump                                                            '水位に応じた稼動、停止ポンプを設定する。
    Flood_Data_Write_For_Calc
    Message.Label1 = "気象庁雨量 予 測 計 算 実 行 中"
    'Message.Label1 = "ＳＨＩＮＫ１０　実 行 中"
    Message.Show
    Message.ZOrder 0
    Message.Refresh
    ChDir App.Path & "\WORK"
    '******************************************************
    'Ver1.0.0 修正開始 2015/08/06 O.OKADA
    '******************************************************
    'Call WaitForProcessToEnd("RRSHINK10.EXE")
    'Call WaitForProcessToEnd("RRSHINK10NF.EXE")                        '久地野フィードバック無し
    'Call WaitForProcessToEnd("New_RSHINK.EXE")                          'なにも無し
    Call WaitForProcessToEnd("D:\SHINKAWA\レーダー洪水予測\WORK\New_RSHINK.EXE")                          'なにも無し
    '******************************************************
    'Ver1.0.0 修正終了 2015/08/06 O.OKADA
    '******************************************************
    'Message.Label1 = "Ｎ Ｓ Ｋ　実 行 中"
    Message.Refresh
    Flood_Data_Write_For_Calc1
    'Cal_Initial_flow_profile irc                                       '全て不定流計算で行う為初期水面形は固定となった為不等流計算は使用しない
    'If irc = False Then Exit Sub
    '******************************************************
    'Ver1.0.0 修正開始 2015/08/06 O.OKADA
    '******************************************************
    'Call WaitForProcessToEnd("NEWNSKG2.EXE")
    Call WaitForProcessToEnd("D:\SHINKAWA\レーダー洪水予測\WORK\NEWNSKG2.EXE")
    '******************************************************
    'Ver1.0.0 修正終了 2015/08/06 O.OKADA
    '******************************************************
    'Message.Hide
    Screen.MousePointer = vbDefault
    ChDir App.Path
    Input_Yosoku irc
    If Not irc Then
        Message.Hide
        Exit Sub
    End If
    Category = True                                                     'True=不定流計算時
    Message.Hide
    If (HO(5, Now_Step) <= Un_Cal) Then
        'If (Froude > 0.5) Or (HO(5, Now_Step) <= Un_Cal) Then
        Category = False
        EMG_Cal irc
        If Not manu Then AutoDrive.Timer1.Enabled = True
        If irc = False Then
            Message.Hide
            Exit Sub
        End If
    Else
        '縦断ＤＢ出力用
        For i = 1 To 74
            YHJ(0, i) = HQ(1, i, NT - 18)
            YHJ(1, i) = HQ(1, i, NT - 12)
            YHJ(2, i) = HQ(1, i, NT - 6)
            YHJ(3, i) = HQ(1, i, NT)
        Next i
        For i = 1 To 5
            FeedBack i
        Next i
        '縦断補正用
        For i = 1 To 5
            ns = V_Sec_Num(i)                                           'V_Sec_Num(nr)は不定流上の断面位置を表す
            OHJ(0, i) = HQ(1, ns, NT - 18)
            OHJ(1, i) = HQ(1, ns, NT - 12)
            OHJ(2, i) = HQ(1, ns, NT - 6)
            OHJ(3, i) = HQ(1, ns, NT)
        Next i
    End If
    If MDBx Then MDB_履歴_Write                                         'データベースに予測値の書き込み
    Load Graph3
    Graph3.Show
    Graph3.Refresh
    If Verification2.Check3 = vbChecked Then
        Graph3.VSPDF1.ConvertDocument Graph3.VSP, App.Path & "\" & Format(PDF_Date, "yyyy_mm_dd") & "_Hydro.pdf"
    End If
    '洪水予報文案作成                                                   'テスト時はここを生かす
    '予報文チェック
    If DBX_ora Then                                                     '計算結果を
        ORA_DataBase_Connection
        If OraDB_OK Then                                                '県庁ＤＢが使用可能なとき
            予報文チェック
            LOG_Out "気象庁雨量　横断予測水位書き込み開始"
            ORA_SUII_YOSOKU_KIJYUN_PUT rc
            LOG_Out "気象庁雨量　横断予測水位書き込み終了"
            LOG_Out "気象庁雨量　縦断予測水位書き込み終了"
            ORA_SUII_YOSOKU_JYUDAN_PUT rc
            LOG_Out "気象庁雨量　縦断予測水位書き込み終了"
            ORA_DataBase_Close
        End If
    End If
    LOG_Out "IN Prediction_CAL_By_KISYO 気象庁雨量による洪水予測終了 現時刻=" & Format(jgd, "yyyy/mm/dd/hh:nn") & " manu=" & manu
    If Not manu Then
        Short_Break 5
        Unload Graph3
        AutoDrive.Timer1.Enabled = True
    End If
'■■■修正開始2016/03/04■■■
'計算処理でエラーが発生した場合、計算中画面がポップアップしたままとなるため。
ERR1:
    If Message.Visible = True Then
        Message.Hide
    End If
'■■■修正終了2016/03/04■■■
End Sub

'******************************************************************************
'サブルーチン：Prediction_CAL_By_FRICS(manu As Boolean)
'処理概要：
'manu True=手動計算時 False=自動計算時
'******************************************************************************
Sub Prediction_CAL_By_FRICS(manu As Boolean)
    Dim dwj     As Date
    Dim irc     As Boolean
    Dim jrc     As Long
    Dim rc      As Boolean
    Dim i       As Integer
    Dim ns      As Long
    Dim ts      As Long
    LOG_Out "IN Prediction_CAL_By_FRICS FRICS雨量による洪水予測開始 現時刻=" & Format(jgd, "yyyy/mm/dd/hh:nn")
    Froude = 0#
    isRAIN = "02"                                                       '"01"=気象庁  "02"=FRICS
    isPump = "00"                                                       '"00"=ノーマル "01"=ポンプ停止
    Screen.MousePointer = vbHourglass
    '久地野と五条上流端流量
    JRADAR = 0
    If MAIN.Check2 Then
        MDB_FRICSレーダー実績 jsd, jgd, dwj, irc
        If dwj < jgd Then jgd = dwj
        dwj = DateAdd("h", 1, jgd)
        MDB_FRICSレーダー予測 jgd, irc
        If irc = False Then
            If manu Then
                MsgBox "FRICS予測雨量がまだ未受信か登録されていません。" & vbCrLf & _
                        "現時刻を再設定して下さい。"
            End If
            LOG_Out "FRICS予測雨量がまだ未受信か登録されていません、計算をスキップします。"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        MDB_洗堰 jsd, jgd, jrc
        Select Case jrc
            Case 0                                                      '正常取得
            Case 1                                                      '10分前取得
                LOG_Out "洗堰データが前時刻データを使う。"
                ORA_Message_Out "洗堰越流量データ受信", "FRICS雨量による計算において、洗堰データが取り込まれませんでした。前時刻データで計算します。", 1
            Case 2                                                      '取得できず
                LOG_Out "洗堰データが10分前も取得できません、データを0としこまま計算します。"
                ORA_Message_Out "洗堰越流量データ受信", "FRICS雨量による計算において、洗堰データが2時刻以上連続して取り込まれませんでした。越流量=0として計算します。", 1
        End Select
        レーダー雨量出力
        JRADAR = 1
    End If
    ポンプ雛型データ読み込み
    ポンプ能力表読み込み
    ポンプデータ作成 jgd
    Set_Pump                                                            '水位に応じた稼動、停止ポンプを設定する。
    Flood_Data_Write_For_Calc
    Message.Label1 = "ＦＲＩＣＳ雨量　予 測 計 算 実 行 中"
    'Message.Label1 = "ＳＨＩＮＫ１０　実 行 中"
    Message.Show
    Message.ZOrder 0
    Message.Refresh
    ChDir App.Path & "\WORK"
    '******************************************************
    'Ver1.0.0 修正開始 2015/08/06 O.OKADA
    '******************************************************
    'Call WaitForProcessToEnd("RRSHINK10.EXE")                          '久地野フィードバック有り
    'Call WaitForProcessToEnd("RRSHINK10NF.EXE")                        '久地野フィードバック無し
    Call WaitForProcessToEnd("D:\SHINKAWA\レーダー洪水予測\WORK\New_RSHINK.EXE")                          'なにも無し
    'Message.Label1 = "Ｎ Ｓ Ｋ　実 行 中"
    Message.Refresh
    '******************************************************
    'Ver1.0.0 修正終了 2015/08/06 O.OKADA
    '******************************************************
    Flood_Data_Write_For_Calc1
    '******************************************************
    'Ver1.0.0 修正開始 2015/08/06 岡田治
    '******************************************************
    'Cal_Initial_flow_profile irc                                       '全て不定流計算で行う為初期水面形は固定となった為不等流計算は使用しない
    'If irc = False Then Exit Sub
'    Call WaitForProcessToEnd("NEWNSKG2.EXE")
    Call WaitForProcessToEnd("NEWNSKG2.EXE")
    'Message.Hide
    Screen.MousePointer = vbDefault
    ChDir App.Path
    Input_Yosoku irc
    If Not irc Then
        Message.Hide
        Exit Sub
    End If
    Category = True                                                     'True=不定流計算時
    Message.Hide
    If (HO(5, Now_Step) <= Un_Cal) Then
        'If (Froude > 0.5) Or (HO(5, Now_Step) <= Un_Cal) Then
        Category = False
        EMG_Cal irc
        If Not manu Then AutoDrive.Timer1.Enabled = True
        If irc = False Then
            Message.Hide
            Exit Sub
        End If
    Else
        '縦断ＤＢ出力用
        For i = 1 To 74
            YHJ(0, i) = HQ(1, i, NT - 18)
            YHJ(1, i) = HQ(1, i, NT - 12)
            YHJ(2, i) = HQ(1, i, NT - 6)
            YHJ(3, i) = HQ(1, i, NT)
        Next i
        For i = 1 To 5
            FeedBack i
        Next i
        '縦断補正用
        For i = 1 To 5
            ns = V_Sec_Num(i)                                           'V_Sec_Num(nr)は不定流上の断面位置を表す
            OHJ(0, i) = HQ(1, ns, NT - 18)
            OHJ(1, i) = HQ(1, ns, NT - 12)
            OHJ(2, i) = HQ(1, ns, NT - 6)
            OHJ(3, i) = HQ(1, ns, NT)
        Next i
    End If
    If MDBx Then MDB_履歴_Write                                         'データベースに予測値の書き込み
    Load Graph3
    Graph3.Show
    Graph3.Refresh
    '洪水予報文案作成                                                   'テスト時はここを生かす
    '予報文チェック
    If DBX_ora Then
        ORA_DataBase_Connection
        If OraDB_OK Then
            予報文チェック
            LOG_Out "FRICS雨量　予測雨量書き込み開始"
            ORA_FRICS_RAIN                                              'ＦＲＩＣＳ予測雨量書き込み
            LOG_Out "FRICS雨量　予測雨量書き込み終了"
            LOG_Out "FRICS雨量　横断予測水位書き込み開始"
            ORA_SUII_YOSOKU_KIJYUN_PUT rc
            LOG_Out "FRICS雨量　横断予測水位書き込み終了"
            LOG_Out "FRICS雨量　縦断予測水位書き込み開始"
            ORA_SUII_YOSOKU_JYUDAN_PUT rc
            LOG_Out "FRICS雨量　縦断予測水位書き込み終了"
            ORA_DataBase_Close
        End If
    Else
    End If
    LOG_Out "Out Prediction_CAL_By_FRICS FRICS雨量による洪水予測終了 現時刻=" & Format(jgd, "yyyy/mm/dd/hh:nn") & _
            " manu=" & manu
    If Not manu Then
        Short_Break 5
        Unload Graph3
        AutoDrive.Timer1.Enabled = True
    End If
End Sub

'******************************************************************************
'サブルーチン：Set_Pump()
'処理概要：
'水位に応じた稼動、停止ポンプを設定する。
'******************************************************************************
Sub Set_Pump()
    Dim ConR        As New ADODB.Recordset
    Dim SQL         As String
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long
    Dim nf          As Integer
    LOG_Out "In  Set_Pump"
    Pre_Pump                                                            'フルポンプデータを読む
    SQL = "select * from ポンプ履歴 where Time = '" & Format(jgd, "yyyy/mm/dd hh:nn") & "'"
    ConR.Open SQL, Con_水文, adOpenKeyset, adLockReadOnly
    'j = 0
    '******************************************************
    '下之一色チェック
    '******************************************************
    j = ConR.Fields("下之一色").Value
    If j = 1 Then
        For i = 1 To 8
            k = Pump_FULL_num(i)
            Mid(Pump_FULL_Data(k), 51, 5) = "    0"
        Next i
    End If
    '******************************************************
    '水場川外水位チェック
    '******************************************************
    j = ConR.Fields("水場川外水位").Value
    If j = 1 Then
        For i = 9 To 20
            k = Pump_FULL_num(i)
            Mid(Pump_FULL_Data(k), 51, 5) = "    0"
        Next i
    End If
    '******************************************************
    '春日チェック
    '******************************************************
    j = ConR.Fields("春日").Value
    If j = 1 Then
        For i = 21 To 27
            k = Pump_FULL_num(i)
            Mid(Pump_FULL_Data(k), 51, 5) = "    0"
        Next i
    End If
    ConR.Close
    '******************************************************
    'ポンプデータ出力
    '******************************************************
    nf = FreeFile
    Open App.Path & "\work\Pump.dat" For Output As #nf
    For i = 1 To 79
        Print #nf, Pump_FULL_Data(i)
    Next i
    Close #nf
    LOG_Out "Out Set_Pump"
End Sub

'******************************************************************************
'サブルーチン：WaitForProcessToEnd(cmdLine As String)
'処理概要：
'******************************************************************************
Sub WaitForProcessToEnd(cmdLine As String)
    'INFINITEをミリ秒単位の時間に置き換える事が出来る
    Dim retVal As Long, pID As Long, pHandle As Long
    pID = Shell(cmdLine, vbMinimizedFocus)                              ' vbMinimizedFocus   vbNormalFocus
    pHandle = OpenProcess(&H100000, True, pID)
    retVal = WaitForSingleObject(pHandle, INFINITE)
End Sub
