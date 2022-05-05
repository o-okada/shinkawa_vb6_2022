VERSION 5.00
Begin VB.Form OracleDB 
   BackColor       =   &H00404040&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "OraDB"
   ClientHeight    =   4590
   ClientLeft      =   3075
   ClientTop       =   2220
   ClientWidth     =   9435
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   9.75
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "OlacleDB.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   9435
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   195
      TabIndex        =   8
      Top             =   2085
      Width           =   6735
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   750
      Top             =   3855
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   165
      Top             =   3840
   End
   Begin VB.CommandButton Command3 
      Caption         =   "終　了"
      Height          =   495
      Left            =   7380
      TabIndex        =   6
      Top             =   3480
      Width           =   1795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DBﾒｲﾝﾃﾅﾝｽ"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7380
      TabIndex        =   5
      Top             =   2797
      Visible         =   0   'False
      Width           =   1795
   End
   Begin VB.CommandButton Command2 
      Caption         =   "予報文書き込み"
      Height          =   495
      Left            =   7380
      TabIndex        =   4
      Top             =   2115
      Visible         =   0   'False
      Width           =   1795
   End
   Begin VB.Label Label4 
      Alignment       =   2  '中央揃え
      Caption         =   "Label4"
      Height          =   255
      Left            =   1230
      TabIndex        =   7
      Top             =   870
      Width           =   6810
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  '実線
      Caption         =   "取り込み待ち"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1725
      TabIndex        =   3
      Top             =   1305
      Width           =   6150
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "現在の状態"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   495
      TabIndex        =   2
      Top             =   1290
      Width           =   1125
   End
   Begin VB.Label Time_Disp 
      BorderStyle     =   1  '実線
      Caption         =   "Label4"
      Height          =   255
      Left            =   -15
      TabIndex        =   1
      Top             =   4320
      Width           =   9435
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      BorderStyle     =   1  '実線
      Caption         =   " 新川洪水予測水文データ取得システム　"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   870
      TabIndex        =   0
      Top             =   315
      Width           =   7590
   End
End
Attribute VB_Name = "OracleDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'モジュール名：OraDB
'
'OracleDB.Check_Araizeki_Time()を修正した。【01-01】
'※OracleDB.Time1_Timer()の処理を修正すること。【01-01-01】
'
'OracleDB.Check_F_MESSYU_10MIN_1_Time()を、、、【03-01】
'※Timer1_Timer()の処理を修正すること。【03-01-01】
'
'OracleDB.Check_F_MESSYU_10MIN_2_Time()を、、、【04-01】
'※OracleDB.Timer1_Timer()の処理を修正すること。【04-01-01】
'
'OracleDB.Check_F_RADAR_TIME()を、、、【06-01】
'※OracleDB.Timer1_Timer()の処理を修正すること。【06-01-01】
'
'OracleDB.Check_P_MESSYU_10MIN_Time()を、、、【07-01】
'※OracleDB.Check_P_MESSYU_10MIN_Time()を修正すること。【07-01-01】
'
'OracleDB.Check_P_MESSYU_1Hour_Time()を、、、【08-01】
'
'OracleDB.Check_P_RADAR_Time()を、、、【09-01】
'※OracleDB.Timer1_Timer()の処理を修正すること。【09-01-01】
'
'******************************************************************************
Option Explicit
Option Base 1
Dim jobg As Boolean

'******************************************************************************
'サブルーチン：Check_Araizeki_Time()
'処理概要：
'洗堰越流量データをチェックする。
'******************************************************************************
'Sub Check_Araizeki_Time(ic As Boolean)
'    Dim nf    As Integer
'    Dim n     As Long
'    Dim d1    As Date
'    Dim d2    As Date
'    Dim d3    As Date
'    Dim ans   As Long
'    Dim buf   As String
'    Dim irc   As Boolean
'    Dim d1st  As String
'    Dim d2st  As String
'    nf = FreeFile
'    'Ver0.0.0 修正開始 1900/01/01 00:00
'    'Debug.Print " Freefile="; nf
'    'Ver0.0.0 修正終了 1900/01/01 00:00
'    Open App.Path & "\data\Araizeki.DAT" For Input As #nf
'    Line Input #nf, buf
'    d1 = CDate(buf)
'    Close #nf
'    d3 = Format(Now, "yyyy/mm/dd hh") & ":00"
'    n = DateDiff("h", d1, d3)
'    If n > 25 Then
'        d1 = DateAdd("h", -25, d3)          '前回終了が２５時間より前だったので取り込み開始時刻を変更
'    End If
'    '******************************************************
'    'Ver1.0.0 修正開始 2015/08/04 O.OKADA【01-01】
'    '※OracleDB.Time1_Timer()の処理を修正すること。【01-01-01】
'    '******************************************************
'    'ORA_KANSOKU_JIKOKU_GET "ARAIZEKI", d2, irc
'    Exit Sub
'    '******************************************************
'    'Ver1.0.0 修正終了 2015/08/04 O.OKADA【01-01】
'    '******************************************************
'    If irc = False Then
'        ic = irc
'        GoTo JUMP
'    End If
'    d1st = Format(d1, "yyyy/mm/dd hh:nn")
'    d2st = Format(d2, "yyyy/mm/dd hh:nn")
'    If d2st > d1st Then
'        n = DateDiff("h", d1, d2) + 1
'        If n > 100 Then                     '100は適当に決めた値、要するにif文に引っかからないようにした。2002/08/07 in YOKOHAMA
'            ans = MsgBox("追加で取得しようとしている洗堰データステップが２４ｈｒの" & vbCrLf & _
'                         "間隔があります。作業を継続しますか？" & vbCrLf & _
'                         "新規の洪水計算ではじめることをお進めします。" & vbCrLf & _
'                         "[はい]でこのジョブは終了します、[いいえ]で継続します。", vbInformation + vbYesNo)
'            If ans = vbYes Then
'                Close
'                End
'            End If
'        End If
'        d1 = DateAdd("n", 10, d1)
'        ORA_LOG "洗堰データ取り込み開始 " & d1 & " から " & d2 & "まで"
'        ORA_Araizeki d1, d2, ic
'        If Not ic Then
'            ORA_LOG "オラクルデータベースより洗堰データを取得しようとした時に" & vbCrLf & _
'                    "エラーが発生しています。"
'            GoTo JUMP
'        Else
'            ORA_LOG "洗堰データ取り込み正常終了"
'            ORA_LOG "洗堰データ時刻書き込み開始 " & d2
'            nf = FreeFile
'            Open App.Path & "\data\Araizeki.DAT" For Output As #nf
'            Print #nf, Format(d2, "yyyy/mm/dd hh:nn")
'            Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
'            Close #nf
'            ORA_LOG "洗堰データ時刻書き込み終了"
'        End If
'    End If
'JUMP:
'End Sub

'******************************************************************************
'サブルーチン：Check_P_RADAR_Time()
'処理概要：
'FRICS実績雨量データをチェックする。
'******************************************************************************
'Sub Check_P_RADAR_Time(ic As Boolean)
'    Dim nf     As Integer
'    Dim n      As Long
'    Dim d1     As Date
'    Dim d2     As Date
'    Dim d3     As Date
'    Dim ans    As Long
'    Dim buf    As String
'    Dim irc    As Boolean
'    Dim d1st   As String
'    Dim d2st   As String
'    nf = FreeFile
'    Open App.Path & "\data\P_RADAR.DAT" For Input As #nf
'    Line Input #nf, buf
'    d1 = CDate(buf)
'    Close #nf
'    d3 = Format(Now, "yyyy/mm/dd hh") & ":00"
'    n = DateDiff("h", d1, d3)
'    If n > 25 Then
'        d1 = DateAdd("h", -25, d3)          '前回終了が２５時間より前だったので取り込み開始時刻を変更
'    End If
'    ORA_KANSOKU_JIKOKU_GET "P_RADAR", d2, irc
'    If irc = False Then
'        GoTo JUMP
'    End If
'    d1st = Format(d1, "yyyy/mm/dd hh:nn")
'    d2st = Format(d2, "yyyy/mm/dd hh:nn")
'    If d2st > d1st Then
'        n = DateDiff("h", d1, d2) + 1
'        If n > 100 Then                     '100は適当に決めた値、要するにif文に引っかからないようにした。2002/08/07 in YOKOHAMA
'            ans = MsgBox("追加で取得しようとしているFRICSレーダデータステップが２４ｈｒの" & vbCrLf & _
'                         "間隔があります。作業を継続しますか？" & vbCrLf & _
'                         "新規の洪水計算ではじめることをお進めします。" & vbCrLf & _
'                         "[はい]でこのジョブは終了します、[いいえ]で継続します。", vbInformation + vbYesNo)
'            If ans = vbYes Then
'                Close
'                End
'            End If
'        End If
'        d1 = DateAdd("n", 10, d1)
'        ORA_LOG "FRICS実績レーダデータ取り込み開始 " & d1 & " から " & d2 & "まで"
'        '******************************************************
'        'Ver1.0.0 修正開始 2015/08/05 O.OKADA【09-01】
'        '※OracleDB.Timer1_Timer()の処理を修正すること。【09-01-01】
'        '******************************************************
'        'ORA_P_RADAR d1, d2, ic
'        Exit Sub
'        '******************************************************
'        'Ver1.0.0 修正終了 2015/08/05 O.OKADA【09-01】
'        '******************************************************
'        If Not ic Then
'            ORA_LOG "オラクルデータベースよりFRICS実績レーダデータを取得しようとした時に" & vbCrLf & _
'                    "エラーが発生しています。"
'            GoTo JUMP
'        Else
'            ORA_LOG "FRICS実績レーダデータ取り込み正常終了"
'        End If
'    End If
'JUMP:
'End Sub

'******************************************************************************
'サブルーチン：Check_Suii_Time()
'処理概要：
'水位データをチェックする。
'******************************************************************************
Sub Check_Suii_Time(ic As Boolean)
    Dim nf  As Integer
    Dim n   As Long
    Dim d1  As Date
    Dim d2  As Date
    Dim d3  As Date
    Dim ans As Long
    Dim buf As String
    Dim irc As Boolean
    nf = FreeFile
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'Debug.Print " Freefile="; nf
    'Ver0.0.0 修正終了 1900/01/01 00:00
    Open App.Path & "\data\P_WATER.DAT" For Input As #nf
    Line Input #nf, buf
    d1 = CDate(buf)
    Close #nf
'    d3 = Format(Now, "yyyy/mm/dd hh") & ":00"
'    n = DateDiff("h", d1, d3)
'    If n > 25 Then
'        d1 = DateAdd("h", -25, d3)          '前回終了が２５時間より前だったので取り込み開始時刻を変更
'    End If
    Call WaterDataNewTime(d2, irc)
    If irc = False Then
        ic = irc
      GoTo JUMP
    End If
    If d2 > d1 Then
'        n = DateDiff("h", d1, d2) + 1
'        If n > 100 Then                     '100は適当に決めた値、要するにif文に引っかからないようにした。2002/08/07 in YOKOHAMA
'            ans = MsgBox("追加で取得しようとしている水位データステップが２４ｈｒの" & vbCrLf & _
'                         "間隔があります。作業を継続しますか？" & vbCrLf & _
'                         "新規の洪水計算ではじめることをお進めします。" & vbCrLf & _
'                         "[はい]でこのジョブは終了します、[いいえ]で継続します。", vbInformation + vbYesNo)
'            If ans = vbYes Then
'                Close
'                End
'            End If
'        End If
        d1 = DateAdd("n", 10, d1)
        ORA_LOG "水位データ取り込み開始 " & d1 & " から " & d2 & "まで"
        ORA_P_WATER d1, d2, ic
        If Not ic Then
            ORA_LOG "愛知県河川情報システムデータベース装置より水位データを取得しようとした時に" & vbCrLf & _
                    "エラーが発生しています。"
            GoTo JUMP
        Else
            ORA_LOG "水位データ取り込み正常終了"
            ORA_LOG "水位データ時刻書き込み開始 " & d2
            nf = FreeFile
            Open App.Path & "\data\P_WATER.DAT" For Output As #nf
            Print #nf, Format(d2, "yyyy/mm/dd hh:nn")
'            Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
            Print #nf, Format(d1, "yyyy/mm/dd hh:nn")
'            d1 = CDate(buf)
            Close #nf
            ORA_LOG "水位データ時刻書き込み終了"
        End If
    End If
JUMP:
End Sub

'******************************************************************************
'サブルーチン：Check_ORA_OWARI_WATER()
'処理概要：
'光ケーブル水位データをチェックする。
'******************************************************************************
Sub Check_ORA_OWARI_WATER(ic As Boolean)
    Dim nf  As Integer
    Dim n   As Long
    Dim d1  As Date
    Dim d2  As Date
    Dim d3  As Date
    Dim ans As Long
    Dim buf As String
    Dim irc As Boolean
    
    Exit Sub
    
    nf = FreeFile
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'Debug.Print " Freefile="; nf
    'Ver0.0.0 修正終了 1900/01/01 00:00
    Open App.Path & "\data\OWARI_WATER.DAT" For Input As #nf
    Line Input #nf, buf
    d1 = CDate(buf)
    Close #nf
'    d3 = Format(Now, "yyyy/mm/dd hh") & ":00"
'    n = DateDiff("h", d1, d3)
'    If n > 25 Then
'        d1 = DateAdd("h", -25, d3)          '前回終了が２５時間より前だったので取り込み開始時刻を変更
'    End If
    ORA_KANSOKU_JIKOKU_GET "OWARI_WATER", d2, irc
    If irc = False Then
        ic = irc
      GoTo JUMP
    End If
    If d2 > d1 Then
'        n = DateDiff("h", d1, d2) + 1
'        If n > 100 Then                     '100は適当に決めた値、要するにif文に引っかからないようにした。2002/08/07 in YOKOHAMA
'            ans = MsgBox("追加で取得しようとしている水位データステップが２４ｈｒの" & vbCrLf & _
'                         "間隔があります。作業を継続しますか？" & vbCrLf & _
'                         "新規の洪水計算ではじめることをお進めします。" & vbCrLf & _
'                         "[はい]でこのジョブは終了します、[いいえ]で継続します。", vbInformation + vbYesNo)
'            If ans = vbYes Then
'                Close
'                End
'            End If
'        End If
        d1 = DateAdd("n", 10, d1)
        ORA_LOG "光水位データ取り込み開始 " & d1 & " から " & d2 & "まで"
        ORA_OWARI_WATER d1, d2, ic
        If Not ic Then
            ORA_LOG "オラクルデータベースより光水位データを取得しようとした時に" & vbCrLf & _
                    "エラーが発生しています。"
            GoTo JUMP
        Else
            ORA_LOG "光水位データ取り込み正常終了"
            ORA_LOG "光水位データ時刻書き込み開始 " & d2
            nf = FreeFile
            Open App.Path & "\data\OWARI_WATER.DAT" For Output As #nf
            Print #nf, Format(d2, "yyyy/mm/dd hh:nn")
'            Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
            Print #nf, Format(d1, "yyyy/mm/dd hh:nn")
'            d1 = CDate(buf)
            Close #nf
            ORA_LOG "水位データ時刻書き込み終了"
        End If
    End If
JUMP:
End Sub

'******************************************************************************
'サブルーチン：Check_F_RADAR_Time()
'処理概要：
' FRICS予測雨量データをチェックする。
'******************************************************************************
'Sub Check_F_RADAR_Time(ic As Boolean)
'    Dim nf     As Integer
'    Dim n      As Long
'    Dim d1     As Date
'    Dim d2     As Date
'    Dim d3     As Date
'    Dim ans    As Long
'    Dim buf    As String
'    Dim irc    As Boolean
'    Dim d1st   As String
'    Dim d2st   As String
'    nf = FreeFile
'    Open App.Path & "\data\F_RADAR.DAT" For Input As #nf
'    Line Input #nf, buf
'    d1 = CDate(buf)
'    Close #nf
'    d3 = Format(Now, "yyyy/mm/dd hh") & ":00"
'    n = DateDiff("h", d1, d3)
'    If n > 25 Then
'        d1 = DateAdd("h", -25, d3)          '前回終了が２５時間より前だったので取り込み開始時刻を変更
'    End If
'    ORA_KANSOKU_JIKOKU_GET "F_RADAR", d2, irc
'    If irc = False Then
'        ic = irc
'        GoTo JUMP
'    End If
'    d1st = Format(d1, "yyyy/mm/dd hh:nn")
'    d2st = Format(d2, "yyyy/mm/dd hh:nn")
'    If d2st > d1st Then
'        ORA_LOG "FRICS予測レーダデータ取り込み開始 " & d2
'        '******************************************************
'        'Ver1.0.0 修正開始 2015/08/05 O.OKADA【06-01】
'        '※OracleDB.Timer1_Timer()の処理を修正すること。【06-01-01】
'        '******************************************************
'        'ORA_F_RADAR d2, ic
'        Exit Sub
'        '******************************************************
'        'Ver1.0.0 修正終了 2015/08/05 O.OKADA【06-01】
'        '******************************************************
'        If Not ic Then
'            ORA_LOG "オラクルデータベースよりFRICS予測レーダデータを取得しようとした時に" & vbCrLf & _
'                    "エラーが発生しています。"
'            GoTo JUMP
'        Else
'            ORA_LOG "FRICS予測レーダデータ取り込み正常終了"
'        End If
'    End If
'JUMP:
'End Sub

'******************************************************************************
'サブルーチン：Check_F_MESSYU_10MIN_1_Time()
'処理概要：
'気象庁予測定時雨量データをチェックする。
'毎１０分予測（１時間分）１０分を６個
'******************************************************************************
'Sub Check_F_MESSYU_10MIN_1_Time(ic As Boolean)
'    Dim nf     As Integer
'    Dim n      As Long
'    Dim d1     As Date
'    Dim d2     As Date
'    Dim d3     As Date
'    Dim dw     As Date
'    Dim ans    As Long
'    Dim buf    As String
'    Dim irc    As Boolean
'    Dim da     As String
'    Dim db     As String
'    Dim d1st   As String
'    Dim d2st   As String
'    nf = FreeFile
'    Open App.Path & "\data\F_MESSYU_10MIN_1.DAT" For Input As #nf
'    Line Input #nf, buf
'    d1 = CDate(buf)
'    Close #nf
'    d3 = Format(Now, "yyyy/mm/dd hh") & ":00"
'    n = DateDiff("h", d1, d3)
'    If n > 25 Then
'        d1 = DateAdd("h", -25, d3)          '前回終了が２５時間より前だったので取り込み開始時刻を変更
'    End If
'    ORA_KANSOKU_JIKOKU_GET "F_MESSYU_10MIN_1", d2, irc
'    If irc = False Then
'        ic = irc
'        GoTo JUMP
'    End If
'    d1st = Format(d1, "yyyy/mm/dd hh:nn")
'    d2st = Format(d2, "yyyy/mm/dd hh:nn")
'    If d2st > d1st Then
'        n = DateDiff("h", d1, d2) + 1
'        If n > 100 Then                     '100は適当に決めた値、要するにif文に引っかからないようにした。2002/08/07 in YOKOHAMA
'            ans = MsgBox("追加で取得しようとしている気象庁レーダデータステップが２４ｈｒの" & vbCrLf & _
'                         "間隔があります。作業を継続しますか？" & vbCrLf & _
'                         "新規の洪水計算ではじめることをお進めします。" & vbCrLf & _
'                         "[はい]でこのジョブは終了します、[いいえ]で継続します。", vbInformation + vbYesNo)
'            If ans = vbYes Then
'                Close
'                End
'            End If
'        End If
'        dw = DateAdd("n", 10, d1)
'        ORA_LOG "気象庁１０分予測雨量データ取り込み開始 " & dw & " から " & d2 & "まで"
'        da = Format(dw, "yyyy/mm/dd hh:nn")
'        db = Format(d2, "yyyy/mm/dd hh:nn")
'        Do Until da > db
'            '******************************************************
'            'Ver1.0.0 修正開始 2015/08/05【03-01】
'            'Timer1_Timer()の処理を修正すること。【03-01-01】
'            '******************************************************
'            'ORA_F_MESSYU_10MIN_1 dw, ic
'            Exit Sub
'            '******************************************************
'            'Ver1.0.0 修正終了 2015/08/05 O.OKADA【03-01】
'            '******************************************************
'            If Not ic Then
'                ORA_LOG "オラクルデータベースより気象庁１０分予測雨量データを取得しようとした時に" & vbCrLf & _
'                        "エラーが発生しています。dw=" & Format(dw, "yyyy/mm/dd hh:nn")
'            End If
'            dw = DateAdd("n", 10, dw)
'            da = Format(dw, "yyyy/mm/dd hh:nn")
'        Loop
'    End If
'JUMP:
'End Sub

'******************************************************************************
'サブルーチン；Check_F_MESSYU_10MIN_2_Time()
'処理概要：
'気象庁予測正時雨量データをチェックする。
'毎正時予測（６時間分）１０分を１８個
'******************************************************************************
'Sub Check_F_MESSYU_10MIN_2_Time(ic As Boolean)
'    Dim i   As Integer
'    Dim nf  As Integer
'    Dim n   As Long
'    Dim d1  As Date
'    Dim d2  As Date
'    Dim d3  As Date
'    Dim dw  As Date
'    Dim ans As Long
'    Dim buf As String
'    Dim irc As Boolean
'    nf = FreeFile
'    Open App.Path & "\data\F_MESSYU_10MIN_2.DAT" For Input As #nf
'    Line Input #nf, buf
'    d1 = CDate(buf)
'    Close #nf
'    d3 = Format(Now, "yyyy/mm/dd hh") & ":00"
'    n = DateDiff("h", d1, d3)
'    If n > 25 Then
'        d1 = DateAdd("h", -25, d3)          '前回終了が２５時間より前だったので取り込み開始時刻を変更
'    End If
'    ORA_KANSOKU_JIKOKU_GET "F_MESSYU_10MIN_2", d2, irc
'    If irc = False Then
'        ic = irc
'        GoTo JUMP
'    End If
'    dw = DateAdd("h", 1, d1)
'    If d2 > d1 Then
'        n = DateDiff("h", dw, d2) + 1
'        If n > 100 Then                     '100は適当に決めた値、要するにif文に引っかからないようにした。2002/08/07 in YOKOHAMA
'            ans = MsgBox("追加で取得しようとしている気象庁レーダデータステップが２４ｈｒの" & vbCrLf & _
'                         "間隔があります。作業を継続しますか？" & vbCrLf & _
'                         "新規の洪水計算ではじめることをお進めします。" & vbCrLf & _
'                         "[はい]でこのジョブは終了します、[いいえ]で継続します。", vbInformation + vbYesNo)
'            If ans = vbYes Then
'                Close
'                End
'            End If
'        End If
'        ORA_LOG "気象庁正時予測雨量データ取り込み開始 " & dw & " から " & d2 & "まで"
'        For i = 1 To n
'            '******************************************************
'            'Ver1.0.0 修正開始 2015/08/05 O.OKADA【04-01】
'            'OracleDB.Timer1_Timer()の処理を修正すること。【04-01-01】
'            '******************************************************
'            'ORA_F_MESSYU_10MIN_2 dw, ic
'            Exit Sub
'            '******************************************************
'            'Ver1.0.0 修正終了 2015/08/05 O.OKADA【04-01】
'            '******************************************************
'            If Not ic Then
'                ORA_LOG "オラクルデータベースより気象庁正時予測雨量データを取得しようとした時に" & vbCrLf & _
'                        "エラーが発生しています。"
'                GoTo JUMP
'            Else
'                ORA_LOG "気象庁正時予測雨量データ取り込み正常終了"
'                ORA_LOG "気象庁正時予測雨量データ時刻書き込み開始 " & d2
'                nf = FreeFile
'                Open App.Path & "\data\F_MESSYU_10MIN_2.DAT" For Output As #nf
'                Print #nf, Format(d2, "yyyy/mm/dd hh:nn")
'                Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
'                Close #nf
'                ORA_LOG "気象庁正時予測雨量データ時刻書き込み終了"
'            End If
'            dw = DateAdd("h", 1, dw)
'        Next i
'    End If
'JUMP:
'End Sub

'******************************************************************************
'サブルーチン：Check_P_MESSYU_1HOUR_Time()
'処理概要：
'気象庁実績正時雨量データをチェックする。
'※使用していない。
'******************************************************************************
'Sub Check_P_MESSYU_1HOUR_Time(ic As Boolean)
'    Dim nf  As Integer
'    Dim n   As Long
'    Dim d1  As Date
'    Dim d2  As Date
'    Dim d3  As Date
'    Dim ans As Long
'    Dim buf As String
'    Dim irc As Boolean
'    nf = FreeFile
'    Open App.Path & "\data\P_MESSYU_1HOUR.DAT" For Input As #nf
'    Line Input #nf, buf
'    d1 = CDate(buf)
'    Close #nf
'    d3 = Format(Now, "yyyy/mm/dd hh") & ":00"
'    n = DateDiff("h", d1, d3)
'    If n > 25 Then
'        d1 = DateAdd("h", -25, d3)          '前回終了が２５時間より前だったので取り込み開始時刻を変更
'    End If
'    ORA_KANSOKU_JIKOKU_GET "P_MESSYU_1HOUR", d2, irc
'    If irc = False Then
'        ic = irc
'        GoTo JUMP
'    End If
'    If d2 > d1 Then
'        n = DateDiff("h", d1, d2) + 1
'        If n > 100 Then                     '100は適当に決めた値、要するにif文に引っかからないようにした。2002/08/07 in YOKOHAMA
'            ans = MsgBox("追加で取得しようとしている気象庁レーダデータステップが２４ｈｒの" & vbCrLf & _
'                         "間隔があります。作業を継続しますか？" & vbCrLf & _
'                         "新規の洪水計算ではじめることをお進めします。" & vbCrLf & _
'                         "[はい]でこのジョブは終了します、[いいえ]で継続します。", vbInformation + vbYesNo)
'            If ans = vbYes Then
'                Close
'                End
'            End If
'        End If
'        d1 = DateAdd("n", 10, d1)
'        ORA_LOG "気象庁実績雨量データ取り込み開始 " & d1 & " から " & d2 & "まで"
'        '******************************************************
'        'Ver1.0.0 修正開始 2015/08/05 O.OKADA【08-01】
'        '※使用していなく、このサブルーチンも呼び出されておらず、影響範囲なしと判断する。
'        '******************************************************
'        'ORA_P_MESSYU_1Hour d1, d2, ic
'        Exit Sub
'        '******************************************************
'        'Ver1.0.0 修正終了 2015/08/05 O.OKADA【08-01】
'        '******************************************************
'        If Not ic Then
'            ORA_LOG "オラクルデータベースより気象庁レーダ実績データを取得しようとした時に" & vbCrLf & _
'                    "エラーが発生しています。"
'            GoTo JUMP
'        Else
'            ORA_LOG "気象庁レーダ実績データ取り込み正常終了"
'            ORA_LOG "気象庁レーダ実績データ時刻書き込み開始 " & d2
'            nf = FreeFile
'            Open App.Path & "\data\P_MESSYU_1HOUR.DAT" For Output As #nf
'            Print #nf, Format(d2, "yyyy/mm/dd hh:nn")
'            Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
'            d1 = CDate(buf)
'            Close #nf
'            ORA_LOG "気象庁レーダ実績データ時刻書き込み終了"
'        End If
'    End If
'JUMP:
'End Sub

'******************************************************************************
'サブルーチン：Check_P_MESSYU_10MIN_Time()
'処理概要：
'気象庁実績定時雨量データチェック
'******************************************************************************
'Sub Check_P_MESSYU_10MIN_Time(ic As Boolean)
'    Dim nf  As Integer
'    Dim n   As Long
'    Dim d1  As Date
'    Dim d2  As Date
'    Dim d3  As Date
'    Dim ans As Long
'    Dim buf As String
'    Dim irc As Boolean
'    nf = FreeFile
'    Open App.Path & "\data\P_MESSYU_10MIN.DAT" For Input As #nf
'    Line Input #nf, buf
'    d1 = CDate(buf)
'    Close #nf
'    d3 = Format(Now, "yyyy/mm/dd hh") & ":00"
'    n = DateDiff("h", d1, d3)
'    If n > 25 Then
'        d1 = DateAdd("h", -25, d3)          '前回終了が２５時間より前だったので取り込み開始時刻を変更
'    End If
'    ORA_KANSOKU_JIKOKU_GET "P_MESSYU_10MIN", d2, irc
'    If irc = False Then
'        ic = irc
'        GoTo JUMP
'    End If
'    If d2 > d1 Then
'        n = DateDiff("h", d1, d2) + 1
'        If n > 100 Then                     '100は適当に決めた値、要するにif文に引っかからないようにした。2002/08/07 in YOKOHAMA
'            ans = MsgBox("追加で取得しようとしている気象庁レーダデータステップが２４ｈｒの" & vbCrLf & _
'                         "間隔があります。作業を継続しますか？" & vbCrLf & _
'                         "新規の洪水計算ではじめることをお進めします。" & vbCrLf & _
'                         "[はい]でこのジョブは終了します、[いいえ]で継続します。", vbInformation + vbYesNo)
'            If ans = vbYes Then
'                Close
'                End
'            End If
'        End If
'        d1 = DateAdd("n", 10, d1)
'        ORA_LOG "気象庁実績雨量データ取り込み開始 " & d1 & " から " & d2 & "まで"
'        '******************************************************
'        'Ver1.0.0 修正開始 2015/08/05 O.OKADA 【07-01】
'        '※OracleDB.Check_P_MESSYU_10MIN_Time()を修正すること。【07-01-01】
'        '******************************************************
'        ORA_P_MESSYU_10MIN d1, d2, ic
'        '******************************************************
'        'Ver1.0.0 修正終了 2015/08/05 O.OKADA【07-01】
'        '******************************************************
'        If Not ic Then
'            ORA_LOG "オラクルデータベースより気象庁レーダ実績データを取得しようとした時に" & vbCrLf & _
'                    "エラーが発生しています。"
'            ORA_LOG "気象庁レーダー実績データエラー時刻書き込み開始 " & Format(d2, "yyyy/mm/dd h:nn")
'            nf = FreeFile
'            Open App.Path & "\data\P_MESSYU_10MIN.DAT" For Output As #nf
'            Print #nf, Format(d2, "yyyy/mm/dd h:nn")
'            Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
'            Close #nf
'            ORA_LOG "気象庁レーダー実績データエラー時刻書き込み終了"
'            GoTo JUMP
'        Else
'            ORA_LOG "気象庁レーダ実績データ取り込み正常終了"
'        End If
'    End If
'JUMP:
'End Sub

'******************************************************************************
'サブルーチン：Command1_Click()
'処理概要：
'データベースをメインテナンスする。
'******************************************************************************
Private Sub Command1_Click()
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'Me.Timer1.Enabled = False
    'CompactMDB
    'Ver0.0.0 修正終了 1900/01/01 00:00
End Sub

'******************************************************************************
'サブルーチン：Command2_Click()
'処理概要：
'使用していない。
'******************************************************************************
'Private Sub Command2_Click()
'    Timer1.Enabled = False
'    Load 予報文テスト送信
'    予報文テスト送信.Show
'End Sub

'******************************************************************************
'サブルーチン：Command3_Click()
'処理概要：
'******************************************************************************
Private Sub Command3_Click()
    Timer1.Enabled = False
    ORA_DataBase_Close
    MsgBox "終わりました"
    Close
    End
End Sub

'******************************************************************************
'サブルーチン：Form_Click()
'処理概要：
'******************************************************************************
Private Sub Form_Click()
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'Dim ic As Boolean
    'ORA_DataBase_Connection ic
    'Dim d1 As Date
    'Dim d2 As Date
    'Dim ic As Boolean
    'd1 = "2006/04/02 01:20"
    'd2 = "2006/04/02 01:20"
    'ORA_DataBase_Connection ic
    'ORA_P_WATER d1, d2, ic
    'ORA_DataBase_Close
    'Check_OWARI_PUMP ic                    '尾張土木ポンプデータ
    'Check_ORA_OWARI_WATER ic               '光ケーブル水位データ
    'Dim Name    As String                  'ポンプ所名
    'Dim Code    As Long                    'ポンプ所コード
    'Dim sv      As Long                    'sv番号
    'Dim N_P     As Long                    'ポンプ数
    'Dim np      As Long                    'ポンプ場の通し番号
    'Dim d1      As Date
    'Dim d2      As Date
    'Dim n       As Long
    'd1 = "2005/05/21 01:00"
    'd2 = "2005/05/21 02:00"
    'n = DateDiff("n", d1, d2) / 10 + 1     '10分データの個数
    'ReDim Pump(17, n)  '17=ポンプ場数  n=時刻ステップ数
    'Name = "土器野ポンプ場"
    'Code = 2605
    'sv = 1
    'N_P = 4
    'np = 12
    'Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic
    'Ver0.0.0 修正終了 1900/01/01 00:00
End Sub

'******************************************************************************
'サブルーチン：Form_Load()
'処理概要：
'******************************************************************************
Private Sub Form_Load()
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'Me.Timer1.Enabled = True
    'Ver0.0.0 修正終了 1900/01/01 00:00
    If App.PrevInstance Then
        MsgBox "このプログラムはすでに起動されています。"
        End
    End If
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'Me.Left = (Screen.Width - Me.Width) * 0.5
    'Me.Top = (Screen.Height - Me.Height) * 0.3
    'Ver0.0.0 修正終了 1900/01/01 00:00
    Me.Left = Screen.Width - Me.Width
    Me.Top = Screen.Height - Me.Height
    Dim i    As Integer
    Dim ic   As Boolean
    Dim nf   As Integer
    Dim buf  As String
    Dim dw   As Date
    Dim d1   As Date
    Dim d2   As Date
    Dim Rtry As Long
    'Ver0.0.0 修正開始 1900/01/01 00:00
    '2002/05/20 11:28 Frics Data center にて修正
    'nf = FreeFile
    'Open App.Path & "\DBpath.dat" For Input As #nf
    'Input #nf, MDB_Path
    'Close #nf
    'Ver0.0.0 修正終了 1900/01/01 00:00
    LOG_File = "\data\LOGFILE" & Format(Now, "yyyy-mm-dd-hh-nn") & ".DAT"
    LOG_N = FreeFile
    Open App.Path & LOG_File For Output As #LOG_N
    Mesh_2km_to_1km_data
    レーダーティーセン読み込み
    jobg = False                            '取り込み可能状態
    FRICS_CVT_DATA                          '2次メッシュデータを315流域に振り分けるデータを読む
    Bit_Intial
    M_Link_Read
    ic = True
    Rtry = 0
ret_MDB:
    MDB_Connection ic
    Pump_Inital
    MDB_最新時刻
    If Not ic Then
        ORA_LOG " MDB Connetion リトライ中"
        Rtry = Rtry + 1
        If Rtry > 10 Then
            MsgBox "ローカルＤＢに接続できませんでした。" & vbCrLf & _
                   "ジョブを終了します。"
            End
        End If
        Short_Break 10
        GoTo ret_MDB
    End If
    ic = True
    Rtry = 0
ret_Ora:
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'ORA_DataBase_Connection ic
    'Ver0.0.0 修正終了 1900/01/01 00:00
    If Not ic Then
       ORA_LOG " OraDB Connetion リトライ中"
        Rtry = Rtry + 1
        If Rtry > 20 Then
            MsgBox "愛知県オラクルＤＢに接続できませんでした。" & vbCrLf & _
                   "ジョブを終了します。"
            End
        End If
        Short_Break 20
        GoTo ret_Ora
    End If
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'ORA_KANSOKU_JIKOKU buf, dw
    'd1 = "2002/05/23 12:10"
    'd2 = "2002/05/25 12:10"
    'ORA_P_EATER d1, d2, ic
    'Ver0.0.0 修正終了 1900/01/01 00:00
End Sub

'******************************************************************************
'サブルーチン：Timer1_Timer()
'業務概要：
'******************************************************************************
Private Sub Timer1_Timer()
    Dim sec   As Integer
    Dim ic    As Boolean
    Dim rc    As Boolean
    Dim nf    As Integer
    Dim n     As Long
    Dim buf   As String
    Dim d1    As Date
    Dim d2    As Date
    Dim ans   As Long
    Dim ret   As Long
    Dim Rrun  As Boolean
    DoEvents
    '******************************************************
    'ローカルDBを圧縮する。
    '******************************************************
    If (Day(Now) Mod 10) = 0 And Hour(Now) = 0 And Minute(Now) = 0 And Second(Now) < 2 Then
        Me.Timer1.Enabled = False
        Pre_Compact rc, Rrun
        If rc Then
            CompactMDB
        End If
        If Rrun Then ret = Shell("D:\SHINKAWA\レーダー洪水予測\RSHINKAWA.EXE", 1)
        Me.Timer1.Enabled = True
    End If
    '******************************************************
    'その他の処理
    '******************************************************
    sec = Second(Now)
    If (sec = 0 Or sec = 30) And Not jobg Then
        Short_Break 1
        jobg = True
        Me.Command1.Enabled = False
        Me.Command2.Enabled = False
        Me.Command3.Enabled = False
        ORA_LOG "愛知県河川情報システムデータベース装置と接続開始"
        ORA_DataBase_Connection ic          '愛知県庁オラクルサーバーとセッションを開始
        If ic Then
            ORA_LOG "愛知県河川情報システムデータベース装置と接続完了"
        Else
            ORA_LOG "愛知県河川情報システムデータベース装置と接続できませんでした。"
            jobg = False
            Me.Command1.Enabled = True
            Me.Command2.Enabled = True
            Me.Command3.Enabled = True
            Exit Sub
        End If
        Me.Timer1.Enabled = False
        '**************************************************
        'サブルーチンをコールする。
        '**************************************************
        '**************************************************
        'Ver1.0.0 修正開始 2015/08/05 O.OKADA【01-01-01】【03-01-01】【04-01-01】【06-01-01】【07-01-01】【09-01-01】
        '**************************************************
        'Check_Araizeki_Time ic                  '洗堰データ
        'Check_ORA_OWARI_WATER ic                '光ケーブル水位データ
        'Check_Suii_Time ic                      '水位データ
        'Check_P_MESSYU_10MIN_Time ic            '気象庁雨量実績データ
        'Check_P_RADAR_Time ic                   'FRICSレーダー実績
        'Check_F_RADAR_Time ic                   'FRICSレーダー予測
        
        'Check_Araizeki_Time ic                  '洗堰データ
'        Check_ORA_OWARI_WATER ic                '光ケーブル水位データ
        Check_Suii_Time ic                      '水位データ
        'Check_P_MESSYU_10MIN_Time ic            '気象庁雨量実績データ
        'Check_P_RADAR_Time ic                   'FRICSレーダー実績
        'Check_F_RADAR_Time ic                   'FRICSレーダー予測
        '**************************************************
        'Ver1.0.0 修正終了 2015/08/05 O.OKADA【01-01-01】【03-01-01】【04-01-01】【06-01-01】【07-01-01】【09-01-01】
        '**************************************************
        
        'Ver0.0.0 修正開始 1900/01/01 00:00
        'Check_F_MESSYU_10MIN_2_Time ic         '気象庁毎正時予測雨量データ
        'Check_F_MESSYU_10MIN_1_Time ic         '気象庁毎10分予測雨量データ
        'Ver0.0.0 修正終了 1900/01/01 00:00
        
'        Check_OWARI_PUMP ic                     '尾張土木ポンプデータ
'        ORA_NEW_DATA_TIME
        
        'Ver0.0.0 修正開始 1900/01/01 00:00
        '気象庁1kmメッシュ雨量追加 2007/05/02
        Check_1kmMesh_Time "VDXA70", ic         '気象庁実績レーダ雨量データ
        Check_1kmMesh_Time "VCXB70", ic         '気象庁降雨短時間レーダ雨量データ(1-3)
        Check_1kmMesh_Time "VCXB71", ic         '気象庁降雨短時間レーダ雨量データ(4-5)
        Check_1kmMesh_Time "VCXB75", ic         '気象庁降雨短時間レーダ雨量データ30(1-3)
        Check_1kmMesh_Time "VCXB76", ic         '気象庁降雨短時間レーダ雨量データ30(4-5)
        Check_1kmMesh_Time "VDXB70", ic         '気象庁ナウキャストデータ
        'Ver0.0.0 修正終了 1900/01/01 00:00
        Me.Timer1.Enabled = True
JUMP:
        OracleDB.Label3 = "取り込み待機中"
        OracleDB.Label3.Refresh
        jobg = False
        ORA_DataBase_Close
        ORA_LOG "愛知県河川情報システムデータベース装置と接続解除"
        Me.Command1.Enabled = True
        Me.Command2.Enabled = True
        Me.Command3.Enabled = True
    End If
End Sub

'******************************************************************************
'サブルーチン：Timer2_Timer()
'処理概要：
'******************************************************************************
Private Sub Timer2_Timer()
    Label4 = Format(Now, "yyyy年mm月dd日 hh時nn分ss秒")
End Sub
