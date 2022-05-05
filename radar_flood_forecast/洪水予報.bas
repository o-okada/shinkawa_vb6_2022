Attribute VB_Name = "洪水予報"
Option Explicit
Option Base 1

Public B1                 As String
Public B2                 As String
Public Y_FLAG             As Integer ' 0=計算開始時 1=洪水注意報時 2=洪水警報時 3=洪水注意解除時
Public Kind_S             As String  '主文種別
Public Kind_N             As String  '主文種別コード
Public hx                 As Single  '(警戒水位＋危険水位)*0.5
Public SYUBN              As String
Public Course             As String
Public Wng_Last_Time      As Integer '前ステップの注意文番号

Public PRACTICE_FLG_CODE  As String  '"40"=予報  "99"=演習

Public 危険水位           As Single  '= 5.2
Public 警戒水位           As Single  '= 3#
Public 指定水位           As Single  '= 2#

Public Con_予報文         As New ADODB.Connection
Public Rst_予報文         As New ADODB.Recordset
Public DB_予報文          As Boolean

Public Const 主文A = "　　新川水場川外水位水位観測所では、" & vbLf & _
                     "　　警戒水位を大幅に超える出水となる見込みですので" & vbLf & _
                     "　　各地とも厳重な警戒をして下さい。"

Public Const 主文B = "　　新川水場川外水位水位観測所では、" & vbLf & _
                     "　　警戒水位を超える出水となる見込みですので各地" & vbLf & _
                     "　　とも十分な注意をして下さい。"

Public Const 主文C = "　　新川水場川外水位水位観測所では、" & vbLf & _
                     "　　当分の間警戒水位以上の水位が続く見込みですので各地とも" & vbLf & _
                     "　　十分な注意をして下さい。"

Public Const 主文D = "　　新川洪水注意報を洪水警報に切換えます。" & vbCrLf & _
                     "　　新川水場川外水位水位観測所では、" & vbLf & _
                     "　　危険水位を超える出水となる見込みですので各地とも厳重な" & vbLf & _
                     "　　警戒をして下さい。"

Public Const 主文E = "　　新川水場川外水位水位観測所では、" & vbLf & _
                     "　　危険水位を超える出水となる見込みですので各地とも厳重な" & vbLf & _
                     "　　警戒をして下さい。"

Public Const 主文F = "　　新川水場川外水位水位観測所では、" & vbLf & _
                     "　　危険水位を大幅に超える出水となる見込みですので各地とも" & vbLf & _
                     "　　厳重な警戒をして下さい。"

Public Const 主文G = "　　新川水場川外水位水位観測所では、" & vbLf & _
                     "　　当分の間危険水位以上の出水が続く見込みですので各地とも" & vbLf & _
                     "　　厳重な警戒をして下さい。"

Public Const 主文H = "　　新川洪水警報を洪水注意報に切換えます。" & vbLf & _
                     "　　新川水場川外水位水位観測所では、" & vbLf & _
                     "　　当分の間警戒水位以上の水位が続く見込ですので各地とも" & vbLf & _
                     "　　十分な注意をして下さい。"

Public Const 主文I = "　　新川水場川外水位水位観測所では、" & vbLf & _
                     "　　警戒水位を下回り危険はなくなったものと思われます。"

Public Const 主文J = "　　新川水場川外水位水位観測所では、" & vbLf & _
                     "　　当分の間警戒水位以上の水位が続く見込ですので各地とも" & vbLf & _
                     "　　十分な注意をして下さい。"

Public Const CYUBN_1 = "　　今回の出水は、平成3年9月の台風17・18号に匹敵" & vbLf & _
                       "　　する規模と見込まれます。"

Public Const CYUBN_2 = "　　今回の出水は、平成3年9月の台風17・18号を上回" & vbLf & _
                       "　　る規模と見込まれます。"

Public Const CYUBN_3 = "　　今回の出水は、平成12年9月の東海豪雨に匹敵する" & vbLf & _
                       "　　規模と見込まれます。"



Public Log_Repo           As Integer   'レポファイルに書き出すファイル番号
Sub H2Z(strH As String, strZ As String)

    Dim ZN(0 To 9)    As String
    Dim i        As Long
    Dim j        As Long
    Dim L        As Long
    Dim w
    Dim ww

    ZN(0) = "０": ZN(1) = "１": ZN(2) = "２": ZN(3) = "３": ZN(4) = "４"
    ZN(5) = "５": ZN(6) = "６": ZN(7) = "７": ZN(8) = "８": ZN(9) = "９"

    strZ = ""
    L = Len(strH)
    For i = 1 To L
        w = Mid(strH, i, 1)
        Select Case w
         Case "0" To "9" 'IsNumeric(w)
            j = CInt(w)
            ww = ZN(j)
        Case "."
            ww = "．"
        Case " "
            ww = "　"
        Case Else
            ww = w
        End Select
        strZ = strZ & ww
    Next i

End Sub
'******************************************************************
'
'
'
'愛知県サーバーに予報文を書き込む
'
'
'
'
'
'
'
'
'
'******************************************************************
Sub ORA_YOHOUBUNAN(Return_Code As Boolean)

    Dim sql_SELECT   As String
    Dim sql_WHERE    As String
    Dim SQL          As String
    Dim N_rec        As Long
    Dim n            As Integer
    Dim i            As Long
    Dim SDATE        As String
    Dim Edate        As String
    Dim jssd         As Date
    Dim NTim         As Date
    Dim dw           As Date
    Dim Timew        As String
    Dim c1           As String
    Dim c2           As String
    Dim c3           As String
    Dim c4           As String
    Dim c5           As String

    LOG_Out "IN    ORA_YOHOUBUNAN"

    NTim = Now
    dw = DateAdd("n", 30, jgd)
    c1 = Format(NTim, "yyyy/mm/dd hh:nn") 'DB書き込み時刻
    c2 = ""
    c3 = ""
    c4 = Format(jgd, "yyyy/mm/dd hh:nn")  '水文データの現時刻 ESTIMATE_TIME
    c5 = Format(dw, "yyyy/mm/dd hh:nn")   '発表時刻 データの現時刻+30分 ANNOUNCE_TIME

    jssd = jgd

    SDATE = "'" & Format(jssd, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"

'SELECT
    sql_SELECT = "SELECT * FROM oracle.YOHOUBUNAN"

'WHERE
    sql_WHERE = " WHERE  ESTIMATE_TIME = TO_DATE(" & SDATE & ") AND" & _
                " DATA_KIND_CODE = 'フケンコウズイアン01' AND" & _
                " SENDING_STATION_CODE ='23001' AND" & _
                " RAIN_KIND = '" & isRAIN & "'"

    SQL = sql_SELECT & sql_WHERE

'    SQL = sql_SELECT
'
'------------ フィールド名を取得する -----------------
'    Dim Tw
'    n = RST_YB.Fields.Count
'    For i = 0 To n - 1
'        Tw = RST_YB.Fields(i).Name
'        Debug.Print " Number=" & Format(Str(i), "@@@") & " フィールド名="; Tw
'    Next i
'---------------------------------------------------

    ' SQLステートメントを指定してダイナセットを取得する
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)

    Dim nf As Integer
    Dim buf As String

    nf = FreeFile
    Open App.Path & "\Data\DB_YB.DAT" For Output As #nf
    If dynOra.EOF Then
        dynOra.AddNew
    Else
        dynOra.Edit
    End If
    dynOra.Fields("WRITE_TIME").Value = c1                      '書き込み時刻
    dynOra.Fields("DATA_KIND_CODE").Value = "フケンコウズイアン01"
    dynOra.Fields("DATA_KIND").Value = "予報文案（水位部分）"
    dynOra.Fields("SENDING_STATION_CODE").Value = "23001"
    dynOra.Fields("SENDING_STATION").Value = "愛知県尾張建設事務所"
    dynOra.Fields("APPOINTED_CODE").Value = ""
    dynOra.Fields("ESTIMATE_TIME").Value = c4
    dynOra.Fields("PRACTICE_FLG_CODE").Value = PRACTICE_FLG_CODE  '"40"=予報  "99"=演習
    If PRACTICE_FLG_CODE = "40" Then
        dynOra.Fields("PRACTICE_FLG").Value = "予報"
    Else
        dynOra.Fields("PRACTICE_FLG").Value = "演習"
    End If
    dynOra.Fields("SEQ_NO").Value = ""
    dynOra.Fields("ANNOUNCE_TIME").Value = c5
    dynOra.Fields("RIVER_NAME").Value = "愛知県庄内川水系　新川"
    dynOra.Fields("RIVER_NO_CODE").Value = "85053002"
    dynOra.Fields("RIVER_NO").Value = "新川"
    dynOra.Fields("RIVER_DIV_CODE").Value = "00"
    dynOra.Fields("RIVER_DIV").Value = ""
    dynOra.Fields("ANNOUNCE_NO").Value = ""
    dynOra.Fields("FORECAST_KIND").Value = Kind_S
    dynOra.Fields("FORECAST_KIND_CODE").Value = Kind_N
    dynOra.Fields("BUNSHO1").Value = B1
    dynOra.Fields("BUNSHO2").Value = B2
    dynOra.Fields("BUNSHO3").Value = ""
    dynOra.Fields("RAIN_KIND").Value = isRAIN '01=気象庁  02=FRICS

    dynOra.Update
    dynOra.Close

'予報文対象河川

'SELECT
    sql_SELECT = "SELECT * FROM oracle.YOHOU_TARGET_RIVER"
'WHERE

    sql_WHERE = " WHERE  ESTIMATE_TIME = TO_DATE(" & SDATE & ") AND" & _
                " BUNAN_CODE = '01' AND" & _
                " DATA_KIND_CODE = 'フケンコウズイアン01' AND" & _
                " SENDING_STATION_CODE ='23001' AND" & _
                " TRIVER_NO_CODE = '85053002' AND" & _
                " RAIN_KIND = '" & isRAIN & "'"

    SQL = sql_SELECT & sql_WHERE

    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)

    If dynOra.EOF Then
        dynOra.AddNew
    Else
        dynOra.Edit
    End If

    dynOra.Fields("WRITE_TIME").Value = c1                      '書き込み時刻
    dynOra.Fields("BUNAN_CODE").Value = "01"
    dynOra.Fields("DATA_KIND_CODE").Value = "フケンコウズイアン01"
    dynOra.Fields("SENDING_STATION_CODE").Value = "23001"
    dynOra.Fields("ESTIMATE_TIME").Value = c4
    dynOra.Fields("TRIVER_NAME").Value = "新川"
    dynOra.Fields("TRIVER_NO_CODE").Value = "85053002"
    dynOra.Fields("TRIVER_NO").Value = "新川"
    dynOra.Fields("TRIVER_DIV_CODE").Value = "00"
    dynOra.Fields("FORECAST_KIND").Value = Kind_S              'c2
    dynOra.Fields("FORECAST_KIND_CODE").Value = Kind_N         'c3
    dynOra.Fields("RAIN_KIND").Value = isRAIN         '02=FRICS   01=気象庁
    dynOra.Fields("OUT_NO").Value = 1

    dynOra.Update
    dynOra.Close

    DoEvents
    Close #nf
    Set dynOra = Nothing

    LOG_Out "OUT   ORA_YOHOUBUNAN"

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
    Dim a


    LOG_Out "IN    RAIN_SELECT_READ"

    a = Dir(履歴MDB)

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
Function Raise(S As Single) As Single

    Dim c As Long
    Dim d As Double

    d = S
    c = Fix(d * 10.00001)
    Raise = c + 1#
    Raise = Raise / 10#

End Function
Sub ST1(H2 As Single)

    If H2 >= hx Then
        SYUBN = 主文A
        Course = "1"
    Else
        SYUBN = 主文B
        Course = "2"
    End If
    Y_FLAG = 1
    Kind_S = "洪水注意報発表"
    Kind_N = "10"

End Sub
Sub ST2()

    SYUBN = 主文I
    Course = Course & "O"
    Kind_S = "洪水注意報解除"
    Kind_N = "30"
    Y_FLAG = 0

End Sub
Sub ST3(H2 As Single, Hm1 As Single, Hm2 As Single, c1 As Integer)

    If 警戒水位 <= Hm1 And 警戒水位 <= Hm2 Then
        SYUBN = 主文C
        Y_FLAG = 2
        Course = Course & "5"
        Kind_S = "洪水情報発表"
        Kind_N = "30"
    Else               '③
        If c1 = 1 Then
            Course = Course & "ｲ"
            ST1 H2
        End If
    End If

End Sub
Sub ST4(H2 As Single, H3 As Single, Hm1 As Single, Hm2 As Single, c1 As Integer)

    If 危険水位 <= H3 Then
        SYUBN = 主文D
        Y_FLAG = 3
        Course = Course & "6"
        Kind_S = "洪水警報発表"
        Kind_N = "20"
    Else
        ST3 H2, Hm1, Hm2, c1
    End If

End Sub
Sub ST5(H0 As Single, H1 As Single, H2 As Single, H3 As Single)

    Kind_S = "洪水情報発表"
    Kind_N = "30"
    If 危険水位 <= H1 Or 危険水位 <= H2 Or 危険水位 <= H3 Then '⑤
        Course = "9"
        If H0 <= 危険水位 And H3 > 危険水位 Then
            SYUBN = 主文E   '⑥
            Y_FLAG = 4
            Course = Course & "A"
            Exit Sub
        End If
        If 危険水位 < H0 And 危険水位 < H3 Then
            SYUBN = 主文F   '⑦
            Y_FLAG = 4
            Course = Course & "B"
            Exit Sub
        End If
        If 危険水位 < H0 And H3 < 危険水位 Then
            SYUBN = 主文G   '⑧
            Y_FLAG = 4
            Course = Course & "C"
            Exit Sub
        End If
        If H0 < 危険水位 And H3 < 危険水位 Then
            SYUBN = 主文G   '⑧
            Y_FLAG = 4
            Course = Course & "Ca"
            Exit Sub
        End If
    End If

    If 危険水位 <= H0 Then
        SYUBN = 主文G   '⑧
        Y_FLAG = 4
        Course = Course & "Cb"
        Exit Sub
    Else
        ST6
    End If

End Sub
Sub ST6()

    SYUBN = 主文H
    Y_FLAG = 5
    Course = Course & "G"
    Kind_S = "洪水注意報発表"
    Kind_N = "10"
    
End Sub
Sub ST7()

    SYUBN = 主文J
    Y_FLAG = 6
    Course = Course & "L"
    Kind_S = "洪水注意報発表"
    Kind_N = "10"

End Sub
Sub ST8()

    SYUBN = 主文I
    Y_FLAG = 7
    Course = Course & "La"
    Kind_S = "洪水注意報解除発表"
    Kind_N = "01"

End Sub
Sub 洪水予報文初期化()

    Dim nf   As Integer
    Dim j    As Integer
    Dim buf  As String
    Dim a

    LOG_Out "IN  洪水予報文初期化"

    nf = FreeFile
    Open App.Path & "\data\予報文出力.txt" For Input As #nf
    Input #nf, buf
    j = CInt(Mid(buf, 1, 5))
    If j = 1 Then
        DBX_ora = True
        AutoDrive.Option1(0).Value = True
    Else
        DBX_ora = False
        AutoDrive.Option1(1).Value = True
    End If

    Input #nf, buf '水位タイトル
    Input #nf, buf
    a = Mid(buf, 1, 10)
    If IsNumeric(a) Then
        危険水位 = CSng(a)
    Else
        MsgBox "入力した危険水位は数値ではありません" & vbLf & _
               "オラクルＤＢには出力しないモードで計算ます。" & vbLf & _
               "計算を中止します。"
        End
    End If
    a = Mid(buf, 11, 10)
    If IsNumeric(a) Then
        警戒水位 = CSng(a)
    Else
        MsgBox "入力した警戒水位は数値ではありません" & vbLf & _
               "オラクルＤＢには出力しないモードで計算ます。" & vbLf & _
               "計算を中止します。"
        End
    End If
    a = Mid(buf, 20, 10)
    If IsNumeric(a) Then
        指定水位 = CSng(a)
    Else
        MsgBox "入力した指定水位は数値ではありません" & vbLf & _
               "オラクルＤＢには出力しないモードで計算ます。" & vbLf & _
               "計算を中止します。"
        End
    End If

    Close #nf

    PRACTICE_FLG_CODE = "40" '予報文本ちゃんモードを初期値とする
    AutoDrive.Option2(0).Value = True

    LOG_Out "OUT 洪水予報文初期化"


End Sub
'**************************************************
'水場川外水位を判定し洪水予報文を作成する
'
'
'
'
'
'
'
'**************************************************
Sub 洪水予報文案作成()

    Dim i           As Long
    Dim j           As Long
    Dim Hm2         As Single   '実績水位
    Dim Hm1         As Single   '実績水位
    Dim H0          As Single   '実績水位
    Dim H1          As Single   '1時間後予測水位
    Dim H2          As Single   '2時間後予測水位
    Dim H3          As Single   '3時間後予測水位
    Dim HM          As Single
    Dim H2r         As Single   '2時間後予測切り上げ水位
    Dim H3r         As Single   '3時間後予測切り上げ水位
    Dim HV          As Single
    Dim c1          As Integer
    Dim CYUBN       As String
    Dim CYUBN1      As String
    Dim Wng         As Integer
    Dim nf          As Integer
    Dim buf         As String
    Dim Kind(6, 2)  As String  '種別コードと文言
    Dim Bun1        As String
    Dim Bun2        As String
    Dim 水位状況     As String
    Dim jsx         As Date
    Dim Bunw        As String
    Dim w           As Single
    Dim M1          As String
    Dim Mw          As String
    Dim irc         As Boolean
    Dim Kind_M      As String

    LOG_Out "IN    洪水予報文案作成"

    Const LF = vbLf

'    Kind(1, 1) = "10": Kind(1, 2) = "洪水注意報発表"
'    Kind(2, 1) = "11": Kind(2, 2) = "洪水注意情報発表（切換）"
'    Kind(3, 1) = "20": Kind(3, 2) = "洪水警報発表"
'    Kind(4, 1) = "21": Kind(4, 2) = "洪水警報発表（切換）"
'    Kind(5, 1) = "30": Kind(5, 2) = "洪水情報発表"
'    Kind(6, 1) = "01": Kind(6, 2) = "洪水注意報解除"

    SYUBN = ""
    Kind_M = ""
    Kind_S = ""
    Kind_N = ""
    Course = ""
    CYUBN = ""
    CYUBN1 = ""

    予測履歴DB_Read

    hx = (危険水位 + 警戒水位) * 0.5
    Hm2 = HO(5, Now_Step - 2)
    Hm1 = HO(5, Now_Step - 1)
    H0 = HO(5, Now_Step)
    H1 = HQ(1, 41, NT - 12)
    H2 = HQ(1, 41, NT - 6)
    H3 = HQ(1, 41, NT)
    HM = H1
    If H2 > HM Then HM = H2
    If H3 > HM Then HM = H3
    H2r = Raise(HQ(1, 41, NT - 6))
    H3r = Raise(HQ(1, 41, NT))

    c1 = Y_FLAG

    Select Case Y_FLAG

        Case 0
            If H2 < 警戒水位 Then
                Exit Sub
            Else
                ST1 H2
            End If

        Case 1, 2                             '②
            If H1 < 警戒水位 And H2 < 警戒水位 And H3 < 警戒水位 Then
                If H0 < 警戒水位 Then
                    ST8
                    Course = "ｱ"
                Else
                    Course = "3"
                End If
            Else
                If H1 < 危険水位 And H2 < 危険水位 And H3 < 危険水位 Then
                    ST3 H2, Hm1, Hm2, c1
                    Course = Course & "4"
                Else
                    ST4 H2, H3, Hm1, Hm2, c1
                End If
            End If

        Case 3, 4
            ST5 H0, H1, H2, H3

        Case 5, 6
            If H0 >= 危険水位 Then '⑩
                Course = Course & "H"
                ST4 H2, H3, Hm1, Hm2, c1
                GoTo J1
            End If
            If H1 >= 危険水位 Or H2 >= 危険水位 Or H3 >= 危険水位 Then '⑪
                Course = Course & "I"
                ST4 H2, H3, Hm1, Hm2, c1
                GoTo J1
            End If
            If H1 < 警戒水位 And H2 < 警戒水位 Or H3 < 警戒水位 Then '⑫
                If H0 < 警戒水位 Then
                    ST8
                    GoTo J1
                Else
                    Course = Course & "K"
                    GoTo J1
                End If
            Else
                If 警戒水位 <= H0 Then
                    If 警戒水位 <= Hm1 And 警戒水位 <= Hm2 Then
                        ST7
                        GoTo J1
                    End If
                Else
                    If c1 = 5 Then  '⑬
                        ST6
                        GoTo J1
                    Else
                        Course = Course & "N"
                        GoTo J1
                    End If
                End If
            End If

        Case 7
            ST8
            GoTo J1

    End Select

J1:

'注意文
    If Y_FLAG = 3 Or Y_FLAG = 4 Then
        HV = H3 - H2
        If Y_FLAG >= 2 Then
            Select Case HV
                Case Is < 0.5
                    Wng = 1
                Case Is < 1#
                    Wng = 2
                Case Is >= 1#
                    Wng = 3
            End Select
'           注意事項文面番号保存
            nf = FreeFile
            Open App.Path & "\Data\注意事項.dat" For Output As #nf
            Print #nf, Format(jgd, "yyyy/mm/dd hh;nn")
            Print #nf, Wng
            Close #nf
'        Else
'            nf = FreeFile
'            Open App.Path & "\Data\注意事項.dat" For Input As #nf
'            Line Input #nf, buf
'            If IsDate(buf) Then
'                j = DateDiff("h", CDate(buf), jgd) + 1
'                Input #nf, Wng
'                Select Case Wng
'                    Case 3
'                        If j > 2 Then
'                            CYUBN = "　　今回の出水は、平成3年9月の台風17・18号を上回る規模と見込まれます。"
'                        Else
'                            CYUBN = "　　今回の出水は、平成3年9月の台風17・18号に匹敵する規模と見込まれます。"
'                        End If
'                    Case 2
'                        If j > 6 Then
'                            CYUBN = "　　今回の出水は、平成12年9月の東海豪雨に匹敵する規模と見込まれます。"
'                        Else
'                            CYUBN = "　　今回の出水は、平成3年9月の台風17・18号を上回る規模と見込まれます。"
'                        End If
'                    Case 1
'                        CYUBN = "　　今回の出水は、平成12年9月の東海豪雨に匹敵する規模と見込まれます。"
'                End Select
'            End If
'            Close #nf
        End If
        If Wng_Last_Time > Wng Then Wng = Wng_Last_Time
        Wng_Last_Time = Wng
        Select Case Wng
            Case 1
                CYUBN = CYUBN_1
            Case 2
                CYUBN = CYUBN_2
            Case 3
                CYUBN = CYUBN_3
        End Select
        If H0 >= 6.2 Or H3 >= 6.2 Then '計画堤防高(T.P 6.2m)を超える
            CYUBN1 = "　　また、越水の恐れがありますので厳重な警戒が必要です。"
        End If
    End If

'洪水状況発表状況
    Select Case Y_FLAG
        Case 1, 2, 5, 6
           Kind_M = "洪水注意報発表中"
        Case 3, 4
           Kind_M = "洪水警報発表中"
        Case 7
           Kind_M = " "
   End Select



'水位状況
    w = H0 - Hm2
    If w <= -0.1 Then 水位状況 = "下降中"
    If -0.1 < w And w <= 0.1 Then 水位状況 = "横ばい"
    If 0.1 < w And w <= 0.3 Then 水位状況 = "上昇中"
    If 0.3 < w Then 水位状況 = "急上昇中"

    Print #Log_Repo, ""
    Print #Log_Repo, Format(jgd, "yyyy/mm/dd hh:nn") & "  " & Kind_S
    Print #Log_Repo, SYUBN
    Print #Log_Repo, "現時刻前２時間水位 " & Format(Format(Hm2, "##0.00"), "@@@@@@@") & " " & IIf((Hm2 - 警戒水位) < 0#, "<", ">=") & " 警戒水位  " & IIf((Hm2 - 危険水位) < 0#, "<", ">=") & " 危険水位"
    Print #Log_Repo, "現時刻前１時間水位 " & Format(Format(Hm1, "##0.00"), "@@@@@@@") & " " & IIf((Hm1 - 警戒水位) < 0#, "<", ">=") & " 警戒水位  " & IIf((Hm1 - 危険水位) < 0#, "<", ">=") & " 危険水位"
    Print #Log_Repo, "現時刻水位 　　　　" & Format(Format(H0, "##0.00"), "@@@@@@@") & " " & IIf((H0 - 警戒水位) < 0#, "<", ">=") & " 警戒水位  " & IIf((H0 - 危険水位) < 0#, "<", ">=") & " 危険水位"
    Print #Log_Repo, "現時刻＋１時間水位 " & Format(Format(H1, "##0.00"), "@@@@@@@") & " " & IIf((H1 - 警戒水位) < 0#, "<", ">=") & " 警戒水位  " & IIf((H1 - 危険水位) < 0#, "<", ">=") & " 危険水位"
    Print #Log_Repo, "現時刻＋２時間水位 " & Format(Format(H2, "##0.00"), "@@@@@@@") & " " & IIf((H2 - 警戒水位) < 0#, "<", ">=") & " 警戒水位  " & IIf((H2 - 危険水位) < 0#, "<", ">=") & " 危険水位"
    Print #Log_Repo, "現時刻＋３時間水位 " & Format(Format(H3, "##0.00"), "@@@@@@@") & " " & IIf((H3 - 警戒水位) < 0#, "<", ">=") & " 警戒水位  " & IIf((H3 - 危険水位) < 0#, "<", ">=") & " 危険水位"
    Print #Log_Repo, "予測最大水位       " & Format(Format(HM, "##0.00"), "@@@@@@@") & " " & IIf((HM - 警戒水位) < 0#, "<", ">=") & " 警戒水位  " & IIf((HM - 危険水位) < 0#, "<", ">=") & " 危険水位"
    Print #Log_Repo, "洪水現況=" & Y_FLAG
    Print #Log_Repo, "Course=" & Course

    If SYUBN = "" Then
        Exit Sub
    End If

    Bun1 = "主文" & LF & SYUBN & LF
    If CYUBN <> "" Then
        Bun1 = Bun1 & "注意・警戒事項" & LF & CYUBN & LF
        If CYUBN1 <> "" Then
            Bun1 = Bun1 & CYUBN1 & LF
        End If
    End If
    Bun1 = Bun1 & " " & LF
    Bun1 = Bun1 & "現況・予想" & LF
    
    Bun2 = ""
    If Y_FLAG <> 1 Then
        jsx = DateAdd("h", 3, jgd)
    Else
        jsx = DateAdd("h", 2, jgd)
    End If
    buf = "　　　　"
    M1 = Format(Day(jgd), "##") & "日" & _
         Format(Hour(jgd), "#0") & "時" & _
         Format(Minute(jgd), "#0") & "分"
'    H2Z M1, Mw
    Mw = M1
    buf = "　　新川の水位は" & Mw & "現在、次のとおりとなっています。" & LF
    buf = buf & "　　水場川外水位水位観測所［新川町大字阿原地内］で" & LF
    M1 = Format(Format(H0, "##0.00"), "@@@@@@")
'    H2Z M1, Mw
    Mw = M1
    buf = buf & "　　　　　　" & Mw & "メートル（" & 水位状況 & "）" & LF
    If Y_FLAG <> 7 Then
        M1 = Format(Day(jsx), "##") & "日" & _
             Format(Hour(jsx), "#0") & "時" & _
             Format(Minute(jsx), "#0") & "分"
'        H2Z M1, Mw
        Mw = M1
        buf = buf & "　　新川の水位は" & Mw & "頃には、次のように見込まれます。" & LF
        buf = buf & "　　水場川外水位水位観測所［新川町大字阿原地内］で" & LF
        If Y_FLAG <> 1 Then
            M1 = Format(Format(H3r, "###0.00"), "@@@@@@")
        Else
            M1 = Format(Format(H2r, "###0.00"), "@@@@@@")
        End If
'        H2Z M1, Mw
        Mw = M1
        buf = buf & "　　　　　　" & Mw & "メートル程度" & LF & " " & LF
    Else
        buf = buf & "　　　　　　" & LF
        buf = buf & "　　　　　　" & LF
        buf = buf & "　　　　　　" & LF
        buf = buf & "　　　　　　" & LF

    End If
'    H2Z buf, Bunw
    Bunw = buf
    Bun2 = Bun2 & Bunw

    Bunw = "　　【参考】" & LF & _
           "　　水場川外水位水位観測所［新川町大字阿原地内］" & LF & _
           "　　堤防高 6.20m  危険水位 5.20m  警戒水位 3.00m  指定水位 2.00m" & LF
    Bun2 = Bun2 & Bunw & " " & LF

    Bunw = "　　【新川の洪水予報発表状況】" & LF
    Bunw = Bunw & "　　　　　" & Kind_M & LF

    Bun2 = Bun2 & Bunw & " " & LF

    Bunw = "　　問い合わせ先" & LF & _
           "　　水位関係　愛知県尾張建設事務所　　維持管理課　ＴＥＬ052(961)4421" & LF & _
           "　　気象関係　気象庁名古屋地方気象台　観測予報課　ＴＥＬ052(763)2449" & LF & " "

    Bun2 = Bun2 & Bunw

    Print #Log_Repo, Bun1
    Print #Log_Repo, Bun2
    B1 = Bun1
    B2 = Bun2

    If DBX_ora Then   '予報文出力が指示されていたら
       ORA_YOHOUBUNAN irc
    End If

    If Y_FLAG = 7 And c1 = 7 Then
        Y_FLAG = 0
    End If

    予測履歴DB_Write

    LOG_Out "IN    洪水予報文案作成"

End Sub
Sub 予測履歴DB_Close()

    If Rst_予報文.State = 1 Then
        Rst_予報文.Close
    End If
    Set Rst_予報文 = Nothing
    Set Con_予報文 = Nothing

End Sub
Sub 予測履歴DB_Connection()

    Dim Con  As String

    LOG_Out "IN    予測履歴DB_Connection"

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

    LOG_Out "OUT   予測履歴DB_Connection ABNormal Return"

    On Error GoTo 0

End Sub
Sub 予測履歴DB_Read()

    Dim SQL     As String
    Dim dw      As String
    Dim T_Last  As Date
    Dim n       As Long

    LOG_Out "IN    予測履歴DB_Read"

    予測履歴DB_Connection

    If DB_予報文 = False Then
        LOG_Out "OT   予測履歴DB_Read DB_予報文 = False"
        Exit Sub
    End If

    SQL = "Select MAX(TIME) From 予報文履歴 Where RAIN_KIND = '" & isRAIN & "'"

    Rst_予報文.Open SQL, Con_予報文, adOpenDynamic, adLockOptimistic

    If Rst_予報文.BOF Or Rst_予報文.EOF Then
       'ここにはこないはずだがもしきたら
        Y_FLAG = 0
        Rst_予報文.Close
        予測履歴DB_Close
        LOG_Out "OUT   予測履歴DB_Read ここにはこないはずだがもしきたら"
        Exit Sub
    End If

    dw = Rst_予報文.Fields(0).Value
    T_Last = CDate(dw)

    Rst_予報文.Close

    n = DateDiff("h", T_Last, jgd) + 1
    If n > 6 Then
        Y_FLAG = 0
        予測履歴DB_Close
        LOG_Out "OUT   予測履歴DB_Read n=" & str(n)
        Exit Sub
    End If

    SQL = "Select * From 予報文履歴 Where TIME = '" & dw & "' AND  RAIN_KIND = '" & isRAIN & "'"
    Rst_予報文.Open SQL, Con_予報文, adOpenDynamic, adLockOptimistic

    Y_FLAG = Rst_予報文.Fields("予報フラグ").Value

    Rst_予報文.Close

    予測履歴DB_Close

    LOG_Out "OUT   予測履歴DB_Read SQL=" & SQL

End Sub
Sub 予測履歴DB_Write()

    Dim SQL    As String

    LOG_Out "IN    予測履歴DB_Write"

    予測履歴DB_Connection

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
    Rst_予報文.Fields("予報フラグ").Value = Y_FLAG
    Rst_予報文.Fields("予報種別コード").Value = Kind_N
    Rst_予報文.Fields("予報種別").Value = Kind_S
    Rst_予報文.Fields("Course").Value = Course

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
    
    予測履歴DB_Close

    LOG_Out "OUT   予測履歴DB_Write"

End Sub
