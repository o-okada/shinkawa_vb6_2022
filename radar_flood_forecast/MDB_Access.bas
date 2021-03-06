Attribute VB_Name = "MDB_Access"
Option Explicit
Option Base 1

Public MDBx              As Boolean              '予測データベース  接続可=True  接続不可=False
Public 水文MDB           As String               '水文MDBのフルパス
Public 履歴MDB           As String               '履歴MDBのフルパス
Public Con_水文          As New ADODB.Connection
Public Con_履歴          As New ADODB.Connection
Public Rec_水文          As New ADODB.Recordset
Public Rec_履歴          As New ADODB.Recordset

Public Con_予報文        As New ADODB.Connection
Public Rst_予報文        As New ADODB.Recordset
Public DB_予報文         As Boolean

Public H_Pred(500, 5, 4) As Single               '水位予測履歴
Public R_Pred(500, 5, 4) As Single               '雨量予測履歴
Public T_Pred(500)       As Date                 '予測計算現時刻

Public History           As Boolean              '履歴表示コントロール 表示有り=True  無し=False
Sub MDB_水文_Close()

    Con_水文.Close
    Set Rec_水文.ActiveConnection = Nothing

End Sub
Sub MDB_水文_Connection()

    Dim Con_str As String
    Dim a

    On Error GoTo ER1

    Con_str = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & 水文MDB
    Con_水文.ConnectionString = Con_str
    Con_水文.Open

    Set Rec_水文.ActiveConnection = Con_水文
    MDBx = True
    On Error GoTo 0
    Exit Sub

ER1:
    a = MsgBox("MDB_水文のDBにアクセスできません、DBの有無、ODBC等の設定を確認してください。" & vbCrLf & _
           "計算を続行しますか(続行の場合は予測値の履歴は保存されません)？", vbYesNo + vbInformation)
    If a = vbYes Then
        MDBx = False
        Exit Sub
    Else
        End
    End If


End Sub
Sub MDB_履歴_Close()

    Con_履歴.Close
    Set Rec_履歴.ActiveConnection = Nothing

End Sub


Sub MDB_履歴_Connection()

    Dim Con_str As String
    Dim a

    On Error GoTo ER1

    Con_str = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & 履歴MDB
    Con_履歴.ConnectionString = Con_str
    Con_履歴.Open

    Set Rec_履歴.ActiveConnection = Con_履歴
    MDBx = True
    On Error GoTo 0
    Exit Sub

ER1:
    a = MsgBox("予測履歴のDBにアクセスできません、DBの有無、ODBC等の設定を確認してください。" & vbCrLf & _
           "計算を続行しますか(続行の場合は予測値の履歴は保存されません)？", vbYesNo + vbInformation)
    On Error GoTo 0
    If a = vbYes Then
        MDBx = False
        Exit Sub
    Else
        End
    End If

End Sub
Sub MDB_履歴_Read()


    Dim i       As Integer
    Dim j       As Integer
    Dim k       As Integer
    Dim n       As Integer
    Dim d       As Date
    Dim sr      As Single
    Dim SQL     As String
    Dim rw      As Variant
    Dim buf     As String
    Dim fn      As String

    fn = Format(Minute(jgd), "00")
    If isRAIN = "02" Then
        SQL = "Select * From FRICS履歴 Where Time Between  '" & Format(jsd, "yyyy/mm/dd hh:nn") & _
              "' And '" & Format(jgd, "yyyy/mm/dd hh:nn") & "' AND Minute=" & fn
    Else
        SQL = "Select * From 気象庁履歴 Where Time Between  '" & Format(jsd, "yyyy/mm/dd hh:nn") & _
              "' And '" & Format(jgd, "yyyy/mm/dd hh:nn") & "' AND Minute=" & fn
    End If


    Rec_履歴.Open SQL, Con_履歴, adOpenDynamic, adLockReadOnly

    For i = 1 To 500
        For j = 1 To 5
            For k = 1 To 4
                H_Pred(i, j, k) = -99#
                R_Pred(i, j, k) = -99#
            Next k
        Next j
    Next i

    If Rec_履歴.BOF Or Rec_履歴.EOF Then
        Rec_履歴.Close
        Exit Sub
    End If

    Do Until Rec_履歴.EOF

        d = CDate(Rec_履歴.Fields("Time").Value)
        n = DateDiff("h", jsd, d) + 1
        T_Pred(n) = d

'下之一色
        buf = Rec_履歴.Fields("下之一色").Value
        rw = Split(buf, ",")
        k = UBound(rw) - 1
        For j = 0 To k
            sr = rw(j)
            H_Pred(n, 1, j + 1) = sr
        Next j
'大治
        buf = Rec_履歴.Fields("大治").Value
        rw = Split(buf, ",")
        k = UBound(rw) - 1
        For j = 0 To k
            sr = rw(j)
            H_Pred(n, 2, j + 1) = sr
        Next j
'水場外水位
        buf = Rec_履歴.Fields("水場川外水位").Value
        rw = Split(buf, ",")
        k = UBound(rw) - 1
        For j = 0 To k
            sr = rw(j)
            H_Pred(n, 3, j + 1) = sr
        Next j
'久地野
        buf = Rec_履歴.Fields("久地野").Value
        rw = Split(buf, ",")
        k = UBound(rw) - 1
        For j = 0 To k
            sr = rw(j)
            H_Pred(n, 4, j + 1) = sr
        Next j
'春日
        buf = Rec_履歴.Fields("春日").Value
        rw = Split(buf, ",")
        k = UBound(rw) - 1
        For j = 0 To k
            sr = rw(j)
            H_Pred(n, 5, j + 1) = sr
        Next j
'予測雨量
        buf = Rec_履歴.Fields("予測降雨").Value
        rw = Split(buf, ",")
        k = UBound(rw) - 1
        For j = 0 To k
            sr = rw(j)
            R_Pred(n, 1, j + 1) = sr
        Next j

        Rec_履歴.MoveNext
    Loop

    Rec_履歴.Close

End Sub
'
'予測値をMDBに保存する。
'
'
Sub MDB_履歴_Write()

    Dim i   As Integer
    Dim j   As Integer
    Dim ns  As Integer
    Dim buf As String
    Dim SQL As String

    Const f2 = "##0.00"
    Const f1 = "###0.0"

    If isRAIN = "02" Then
        SQL = "Select * From FRICS履歴 Where Time = '" & Format(jgd, "yyyy/mm/dd hh:nn") & "'"
    Else
        SQL = "Select * From 気象庁履歴 Where Time = '" & Format(jgd, "yyyy/mm/dd hh:nn") & "'"
    End If

    Rec_履歴.Open SQL, Con_履歴, adOpenDynamic, adLockOptimistic

    If Rec_履歴.BOF Or Rec_履歴.EOF Then
        Rec_履歴.AddNew
        Rec_履歴.Fields("Time").Value = Format(jgd, "yyyy/mm/dd hh:nn")
    End If
    Rec_履歴.Fields("Minute").Value = Format(Minute(jgd), "00")

'日光川外水位
    buf = Format(DH_Tide, f2) & ","                   '現時刻天文潮位と実績水位との差
    buf = buf & Format(HO(1, Now_Step), f2) & ","     '現時刻水位(実績)
    buf = buf & Format(HO(1, Now_Step + 1), f2) & "," '1時間後
    buf = buf & Format(HO(1, Now_Step + 2), f2) & "," '2時間後
    buf = buf & Format(HO(1, Now_Step + 3), f2) & "," '3時間後
    Rec_履歴.Fields("日光川外水位").Value = buf
'下之一色
    ns = V_Sec_Num(1)
    buf = Format(HO(3, Now_Step), f2) & ","           '現時刻水位(実績)
    buf = buf & Format(HQ(1, ns, NT - 12), f2) & ","  '1時間後
    buf = buf & Format(HQ(1, ns, NT - 6), f2) & ","   '2時間後
    buf = buf & Format(HQ(1, ns, NT), f2) & ","       '3時間後
    Rec_履歴.Fields("下之一色").Value = buf
'大治
    ns = V_Sec_Num(2)
    buf = Format(HO(4, Now_Step), f2) & ","           '現時刻水位(実績)
    buf = buf & Format(HQ(1, ns, NT - 12), f2) & ","  '1時間後
    buf = buf & Format(HQ(1, ns, NT - 6), f2) & ","   '2時間後
    buf = buf & Format(HQ(1, ns, NT), f2) & ","       '3時間後
    Rec_履歴.Fields("大治").Value = buf
'水場外水位
    ns = V_Sec_Num(3)
    buf = Format(HO(5, Now_Step), f2) & ","           '現時刻水位(実績)
    buf = buf & Format(HQ(1, ns, NT - 12), f2) & ","  '1時間後
    buf = buf & Format(HQ(1, ns, NT - 6), f2) & ","   '2時間後
    buf = buf & Format(HQ(1, ns, NT), f2) & ","       '3時間後
    Rec_履歴.Fields("水場川外水位").Value = buf
'久地野
    ns = V_Sec_Num(4)
    buf = Format(HO(6, Now_Step), f2) & ","           '現時刻水位(実績)
    buf = buf & Format(HQ(1, ns, NT - 12), f2) & ","  '1時間後
    buf = buf & Format(HQ(1, ns, NT - 6), f2) & ","   '2時間後
    buf = buf & Format(HQ(1, ns, NT), f2) & ","       '3時間後
    Rec_履歴.Fields("久地野").Value = buf
'春日
    ns = V_Sec_Num(5)
    buf = Format(HO(7, Now_Step), f2) & ","           '現時刻水位(実績)
    buf = buf & Format(HQ(1, ns, NT - 12), f2) & ","  '1時間後
    buf = buf & Format(HQ(1, ns, NT - 6), f2) & ","   '2時間後
    buf = buf & Format(HQ(1, ns, NT), f2) & ","       '3時間後
    Rec_履歴.Fields("春日").Value = buf
'予測雨量
    buf = Format(RO(1, Now_Step), f1) & ","           '現時刻流域平均雨量
    buf = buf & Format(RO(1, Now_Step + 1), f1) & "," '1時間後流域平均雨量
    buf = buf & Format(RO(1, Now_Step + 2), f1) & "," '2時間後流域平均雨量
    buf = buf & Format(RO(1, Now_Step + 3), f1) & "," '3時間後流域平均雨量
    Rec_履歴.Fields("予測降雨").Value = buf

'DB書き込み
    Rec_履歴.Update
    Rec_履歴.Close

End Sub
'
'現在計算に使われている予測雨量の状態をセーブする。
'
'
'
Sub RAIN_SELECT_READ()

    Dim Con    As String
    Dim R_Con  As New ADODB.Connection
    Dim R_Rst  As New ADODB.Recordset

    LOG_Out "IN    RAIN_SELECT_READ"

    Con = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & 履歴MDB

    R_Con.ConnectionString = Con
    R_Con.Open

    Set R_Rst.ActiveConnection = R_Con
    R_Rst.Open "SELECT * FROM RAIN_SELECT", R_Con, adOpenDynamic, adLockOptimistic
    
    If R_Rst.Fields("気象庁").Value Then
        AutoDrive.Check1 = vbChecked
        KISYO = True
    Else
        AutoDrive.Check1 = vbUnchecked
        KISYO = False
    End If

    If R_Rst.Fields("FRICS").Value Then
        AutoDrive.Check2 = vbChecked
        FRICS = True
    Else
        AutoDrive.Check2 = vbUnchecked
        FRICS = False
    End If

    R_Rst.Update
    R_Rst.Close
    R_Con.Close

    Set R_Rst = Nothing
    Set R_Con = Nothing

    LOG_Out "OUT   RAIN_SELECT_READ"

End Sub
'
'現在計算に使われている予測雨量の状態をセーブする。
'
'
'
Sub RAIN_SELECT_SAVE()

    Dim Con    As String
    Dim R_Con  As New ADODB.Connection
    Dim R_Rst  As New ADODB.Recordset

    LOG_Out "IN    RAIN_SELECT_SAVE"

    Con = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & 履歴MDB

    R_Con.ConnectionString = Con
    R_Con.Open

    Set R_Rst.ActiveConnection = R_Con
    R_Rst.Open "SELECT * FROM RAIN_SELECT", R_Con, adOpenDynamic, adLockOptimistic

    R_Rst.Fields("気象庁").Value = KISYO
    R_Rst.Fields("FRICS").Value = FRICS

    R_Rst.Update
    R_Rst.Close
    R_Con.Close

    Set R_Rst = Nothing
    Set R_Con = Nothing

    LOG_Out "OUT   RAIN_SELECT_SAVE"

End Sub
Sub 予報文履歴DB_Close()

    On Error Resume Next

    If Rst_予報文.State = 1 Then
        Rst_予報文.Close
    End If
    Set Rst_予報文 = Nothing
    Set Con_予報文 = Nothing

End Sub
Sub 予報文履歴DB_Connection()

    Dim Con  As String

    LOG_Out "IN    予報文履歴DB_Connection"

    On Error GoTo ERH1
    
    Con = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & 履歴MDB

    Con_予報文.ConnectionString = Con
    Con_予報文.Open

    Set Rst_予報文.ActiveConnection = Con_予報文
    DB_予報文 = True

    LOG_Out "OUT   予測履歴DB_Connection Normal Return"

    On Error GoTo 0
    Exit Sub

ERH1:

    DB_予報文 = False
    MsgBox "予報文履歴データベースに接続できませんでした、履歴は残りません。"
    LOG_Out "予報文履歴データベースに接続できませんでした、履歴は残りません。"

    LOG_Out "OUT   予報文履歴DB_Connection ABNormal Return"

    On Error GoTo 0

End Sub
Sub 予報文履歴DB_Read()

    Dim SQL     As String
    Dim dw      As String
    Dim T_Last  As Date
    Dim n       As Long

    LOG_Out "IN    予報文履歴DB_Read"

    予報文履歴DB_Connection

    If DB_予報文 = False Then
        LOG_Out "OT   予測履歴DB_Read DB_予報文 = False"
        Exit Sub
    End If

    SQL = "Select MAX(TIME) From 予報文履歴 Where RAIN_KIND = '" & isRAIN & "'"

    Rst_予報文.Open SQL, Con_予報文, adOpenDynamic, adLockOptimistic

    If Rst_予報文.BOF Or Rst_予報文.EOF Then
       'ここにはこないはずだがもしきたら
        BP = 0
        Rst_予報文.Close
        予報文履歴DB_Close
        LOG_Out "OUT   予測履歴DB_Read ここにはこないはずだがもしきたら"
        Exit Sub
    End If

    dw = Rst_予報文.Fields(0).Value
    T_Last = CDate(dw)

    Rst_予報文.Close

    n = DateDiff("h", T_Last, jgd) + 1
    If n > 25 Then
        BP = 0
        LOG_Out "                   jgd=" & TIMEC(jgd)
        LOG_Out "                T_Last=" & TIMEC(T_Last)
        予報文履歴DB_Close
        LOG_Out "OUT   予報文履歴DB_Read n=" & str(n)
        Exit Sub
    End If

    SQL = "Select * From 予報文履歴 Where TIME = '" & dw & "' AND  RAIN_KIND = '" & isRAIN & "'"
    Rst_予報文.Open SQL, Con_予報文, adOpenDynamic, adLockOptimistic
    If Rst_予報文.EOF Then
        BP = 0
        Wng_Last_Time = 0
    Else
        BP = Rst_予報文.Fields("予報フラグ").Value
        Wng_Last_Time = Rst_予報文.Fields("Course").Value
    End If

    Rst_予報文.Close

    予報文履歴DB_Close

    LOG_Out "OUT   予報文履歴DB_Read SQL=" & SQL

End Sub
Sub 予報文履歴DB_Write()

    Dim SQL    As String

    LOG_Out "IN    予報文履歴DB_Write"

    予報文履歴DB_Connection

    If DB_予報文 = False Then
        Exit Sub
    End If

    SQL = "Select * From 予報文履歴 Where TIME = '" & Format(jgd, "yyyy/mm/dd hh:nn") & "' AND RAIN_KIND = '" & isRAIN & "'"

    Rst_予報文.Open SQL, Con_予報文, adOpenDynamic, adLockOptimistic

    If Rst_予報文.BOF Or Rst_予報文.EOF Then
        Rst_予報文.AddNew
        Rst_予報文.Fields("Time").Value = Format(jgd, "yyyy/mm/dd hh:nn")
        Rst_予報文.Fields("RAIN_KIND").Value = isRAIN
    End If

    If Pattan_Now <> 4 Then
        Rst_予報文.Fields("予報フラグ").Value = Pattan_Now  '予報文パターン
        Rst_予報文.Fields("予報種別コード").Value = Messag(Pattan_Now).Patn(16) 'Kind_N
        Rst_予報文.Fields("予報種別").Value = Messag(Pattan_Now).Patn(2) 'Kind_S
    Else
        '予報文発生終了
        Rst_予報文.Fields("予報フラグ").Value = 0                    '予報文パターン
        Rst_予報文.Fields("予報種別コード").Value = "0"            'Kind_N
        Rst_予報文.Fields("予報種別").Value = "洪水注意情報解除"   'Kind_S
    End If
    Rst_予報文.Fields("Course").Value = Wng_Last_Time

    If isRAIN = "01" Then
        Rst_予報文.Fields("RAIN_NAME").Value = "気象庁"
    Else
        Rst_予報文.Fields("RAIN_NAME").Value = "FRICS"
    End If

    If PRACTICE_FLG_CODE = "40" Then
        Rst_予報文.Fields("PRACTICE").Value = "予報"
    Else
        Rst_予報文.Fields("PRACTICE").Value = "演習"
    End If

    Rst_予報文.Update

    Rst_予報文.Close
    
    予報文履歴DB_Close

    LOG_Out "OUT   予報文履歴DB_Write"

End Sub
