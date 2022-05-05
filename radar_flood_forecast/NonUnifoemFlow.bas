Attribute VB_Name = "NonUnifoemFlow"
Option Explicit
Option Base 1

'Public DLIB As New DLIB1.Class1            'ライブラリー

Public Const NSEC = 500                    '最大計算断面数
Public Const NSPC = 15                     '断面特性の数
Public Const NSEP = 1                      '断面分割数
Public Const Froude_Number_Limit = 1#      'フルード数チェック値
'--------  断面諸元  ---------------------------------------------
Public Num_Of_Sec              As Integer  '計算断面数
Public Sec_Name(NSEC)          As String   '断面名
Public DeltaX(NSEC)            As Single   '区間距離
Public NBLBR(NSEC)             As Integer  '各断面毎のＢＬＢＲの数
Public BLBR(NSEP + 1, NSEC)    As Single   '各断面のＢＬＢＲ
Public n(NSEP, NSEC)           As Single   '租度係数
Public CS(NSEC)                As Single   '断面角度補正

'--------  断面特性  ---------------------------------------------
Public H(NSPC, NSEC)           As Single   '水位
Public ZS(NSEC)                As Single   '最深河床座標
Public AG(NSPC, NSEC)          As Single   '合成断面河積
Public RG(NSPC, NSEC)          As Single   '合成断面径深
Public BG(NSPC, NSEC)          As Single   '合成断面水面幅
Public PG(NSPC, NSEC)          As Single   '合成断面潤辺
Public NG(NSPC, NSEC)          As Single   '合成断面租度

'--------  不等流計算結果  ----------------------------------------
Public CQ(0 To NSEP, NSEC)     As Single   '流量
Public CV(NSEC)                As Single   '流速
Public ch(NSEC)                As Single   '水位
Public CR(NSEC)                As Single   '径深
Public FR(NSEC)                As Single   'フルード数
Public CA(0 To NSEP, NSEC)     As Single   '分割断面毎の河積
Public CD(NSEC)                As Single   'エネルギー補正係数
Public CFLAG(NSEC)             As String   '計算フラグ

'--------  不等流計算パラメータ  ----------------------------------
Public Alpha                   As Single   'エネルギー補正係数
'Public Froude_Number_Limit     As Single   'フルード数限界値
Public Start_Sec               As String   '不等流計算開始断面記号
Public End_Sec                 As String   '不等流計算終了断面記号
Public Start_Num               As Integer  '不等流計算開始断面順番号
Public End_Num                 As Integer  '不等流計算終了断面順番号
'--------  不等流計算境界条件  ------------------------------------
Public QU                      As Single   '流量
Public H_Start                 As Single   '下流端水位
'--------  ファイル関係  ------------------------------------------
Public open_data               As String
Public Log_CALC_ERROR          As String   '計算中止ログファイル番号
Public Log_CALC_N              As Long     '計算ログ出力ライン数
'概要      :不等流計算用ベースデータ読み込み。
'説明      :データ読み込み。
Sub Base_Data_Read()

    Dim i As Integer, j As Integer, nf As Integer, buf As String
    Dim ii As Integer, k As Integer
    Dim SFdx As Single, SFn As Single, t As String, c As String
    Dim msg    As String
    Dim SF     As Single

    On Error GoTo 0


    Const NSPCx = 14

    nf = FreeFile
    Open App.Path & "\WORK\nsk.dat" For Input As #nf
'断面数
    Do
        Line Input #nf, buf
        If Mid(buf, 1, 2) = "AR" Then
            Num_Of_Sec = CInt(Mid(buf, 6, 5))
            Exit Do
        End If
    Loop
'断面名,租度係数,区間距離読み込み
    Line Input #nf, buf 'スケールファクター
    If Mid(buf, 1, 1) <> "S" Then
        MsgBox "不定流計算のデータ構成が違う" & vbCrLf & _
               "租度計数、区間距離データのスケールファクターを読み込むところに違うデータがある" & vbCrLf & _
               "データ=(" & buf & ")" & vbCrLf & _
               "計算を中止します。", vbExclamation
        End
    End If
    SFn = CSng(Mid(buf, 11, 5))    '租度係数用
    SFdx = CSng(Mid(buf, 16, 5))   '区間距離用
    For i = 1 To Num_Of_Sec
        Line Input #nf, buf
            Sec_Name(i) = Mid(buf, 5, 6)              '断面名
            NG(1, i) = CSng(Mid(buf, 11, 5)) * SFn    '租度計数
            DeltaX(i) = CSng(Mid(buf, 16, 5)) * SFdx  '区間距離
            ZS(i) = CSng(Mid(buf, 36, 5))             '最深河床
    Next i
'断面特性 水位読み込み
    Line Input #nf, buf 'スケールファクター
    If Mid(buf, 1, 1) <> "H" Then
        MsgBox "不定流計算のデータ構成が違う" & vbCrLf & _
               "水位データのスケールファクターを読み込むところに違うデータがある" & vbCrLf & _
               "データ=(" & buf & ")" & vbCrLf & _
               "計算を中止します。", vbExclamation
        End
    End If
    SF = CSng(Mid(buf, 11, 5)) '水位用
    For i = 1 To Num_Of_Sec
        Line Input #nf, buf
        If Mid(buf, 5, 6) <> Sec_Name(i) Then
            MsgBox "水位断面特性を読み込み中にエラーが発生した。" & vbCrLf & _
                   "エラー＝断面記号が違う(" & Sec_Name(i) & ")が(" & buf & ")になっている。" & vbCrLf & _
                   "計算を中止します。", vbExclamation
            End
        End If
        H(1, i) = ZS(i)
        For j = 1 To NSPCx '断面特性の数
            H(j + 1, i) = CSng(Mid(buf, 11 + (j - 1) * 5, 5)) * SF
        Next j
    Next i
'断面特性 水面幅読み込み
    Line Input #nf, buf 'スケールファクター
    If Mid(buf, 1, 1) <> "B" Then
        MsgBox "不定流計算のデータ構成が違う" & vbCrLf & _
               "水面幅データのスケールファクターを読み込むところに違うデータがある" & vbCrLf & _
               "データ=(" & buf & ")" & vbCrLf & _
               "計算を中止します。", vbExclamation
        End
    End If
    SF = CSng(Mid(buf, 11, 5)) '水面幅用
    For i = 1 To Num_Of_Sec
        Line Input #nf, buf
        If Mid(buf, 5, 6) <> Sec_Name(i) Then
            MsgBox "水位断面特性を読み込み中にエラーが発生した。" & vbCrLf & _
                   "エラー＝断面記号が違う(" & Sec_Name(i) & ")が(" & buf & ")になっている。" & vbCrLf & _
                   "計算を中止します。", vbExclamation
            End
        End If
        BG(1, i) = 0#
        For j = 1 To NSPCx '断面特性の数
            BG(j + 1, i) = CSng(Mid(buf, 11 + (j - 1) * 5, 5)) * SF
        Next j
    Next i
'断面特性 河積読み込み
    Line Input #nf, buf 'スケールファクター
    If Mid(buf, 1, 1) <> "A" Then
        MsgBox "不定流計算のデータ構成が違う" & vbCrLf & _
               "河積データのスケールファクターを読み込むところに違うデータがある" & vbCrLf & _
               "データ=(" & buf & ")" & vbCrLf & _
               "計算を中止します。", vbExclamation
        End
    End If
    SF = CSng(Mid(buf, 11, 5)) '河積用
    For i = 1 To Num_Of_Sec
        Line Input #nf, buf
        If Mid(buf, 5, 6) <> Sec_Name(i) Then
            MsgBox "河積断面特性を読み込み中にエラーが発生した。" & vbCrLf & _
                   "エラー＝断面記号が違う(" & Sec_Name(i) & ")が(" & buf & ")になっている。" & vbCrLf & _
                   "計算を中止します。", vbExclamation
            End
        End If
        AG(1, i) = 0#
        For j = 1 To NSPCx '断面特性の数
            AG(j, i) = CSng(Mid(buf, 11 + (j - 1) * 5, 5)) * SF
        Next j
    Next i
'断面特性 径深読み込み
    Line Input #nf, buf 'スケールファクター
    If Mid(buf, 1, 1) <> "R" Then
        MsgBox "不定流計算のデータ構成が違う" & vbCrLf & _
               "径深データのスケールファクターを読み込むところに違うデータがある" & vbCrLf & _
               "データ=(" & buf & ")" & vbCrLf & _
               "計算を中止します。", vbExclamation
        End
    End If
    SF = CSng(Mid(buf, 11, 5)) '径深用
    For i = 1 To Num_Of_Sec
        Line Input #nf, buf
        If Mid(buf, 5, 6) <> Sec_Name(i) Then
            MsgBox "径深断面特性を読み込み中にエラーが発生した。" & vbCrLf & _
                   "エラー＝断面記号が違う(" & Sec_Name(i) & ")が(" & buf & ")になっている。" & vbCrLf & _
                   "計算を中止します。", vbExclamation
            End
        End If
        RG(1, i) = 0#
        For j = 1 To NSPCx '断面特性の数
            RG(j, i) = CSng(Mid(buf, 11 + (j - 1) * 5, 5)) * SF
        Next j
    Next i
'ＢＬＢＲ設定
    For i = 1 To Num_Of_Sec
        NBLBR(i) = 2
    Next i

    Close #nf

End Sub
Sub Cal_Nonuniform_Flow_Parameter(m As Integer, H1 As Single, AW As Single, BW As Single, Return_Code As Boolean)

    Dim j As Integer, buf As String
    Dim QA As Single, QR As Single, QB As Single, QP As Single
    Dim sar As Single, sn1 As Single, sn2 As Single, na As Single
    Dim aa As Single, da As Single, db As Single
    Dim nn As Single, nn3 As Single, ar As Single
    Dim R1 As Single, n1 As Single, d1 As Single
    
    Const g2 = 9.8 * 2#
    Const P1 = 5# / 3#
    Const P2 = 2# / 3#
    
    sar = 0#
    sn1 = 0#
    sn2 = 0#
    na = 0#
    aa = 0#
    da = 0#
    db = 0#

        nn = NG(1, m)
        nn3 = 1# / nn ^ 3
        nn = 1# / nn
        Call Inner_point_G(m, H1, QA, QR, QB, QP, Return_Code)
        If Not Return_Code Then
'            MsgBox "水位が低すぎて計算できません、もう少し水位が上昇してから計算してください。"
            Exit Sub
        End If
        ar = QA * QR ^ P2
        sar = sar + ar
        sn1 = sn1 + ar * nn
        aa = aa + QA
        na = na + ar * nn
        If QB > 0# Then
            QR = QA / QB
        Else
            QR = 0#
        End If
        da = da + QR ^ 3 * nn3 * QB
        db = db + QR ^ P1 * nn * QB
            
        CA(1, m) = QA
'        If Sec_Name(m) = "1.40 " Then
'            buf = " j=" & Format(Str(j), "@@@") & "  QR=" & Format(Str(QR), "@@@@@@@") & _
'                "  nn=" & Format(Str(nn), "@@@@@@@") & "  da=" & Format(Str(da), "@@@@@@@@") & _
'                "  db=" & Format(Str(db), "@@@@@@@@")
'            Print #7, buf
'        End If

    R1 = (sar / aa) ^ 1.5
    n1 = sar / sn1
    d1 = Alpha * aa * aa * da / db ^ 3
    If d1 < 1# Then d1 = 1#
    AW = H1 + d1 / g2 * (QU / aa) ^ 2
    BW = (n1 ^ 2 * QU ^ 2) / (aa ^ 2 * R1 ^ 1.33333)
    
    CD(m) = d1
    CR(m) = R1

End Sub
Sub Log_Calc(msg As String)

    If Log_CALC_N > 3000 Then
        Close #Log_CALC_ERROR
        Log_CALC_ERROR = FreeFile
        Open App.Path & "\Log_Calc.dat" For Output As #Log_CALC_ERROR
        Log_CALC_N = 0
    End If

    Print #Log_CALC_ERROR, Format(Now, "yyyy/mm/dd hh:nn:ss") & "  jgd=" & _
                           Format(jgd, "yyyy/mm/dd hh:nn") & "   " & msg
    Log_CALC_N = Log_CALC_N + 1

End Sub
Sub Test_CalCulation_GO()

    Dim i   As Integer
    Dim i1  As Integer
    Dim nf  As Integer
    Dim buf As String
    Dim irc As Boolean


    nf = FreeFile
    Open Wpath & "\Non_Flow.log" For Output As #nf

    QU = 112.8
    H_Start = 0.88
    Start_Sec = "S0.000"
    End_Sec = "SP1   "
    Nonuniform_Flow irc
    i1 = Start_Num

    H_Start = ch(End_Num)
    QU = 109.8
    Start_Sec = "SP1   "
    End_Sec = "S3.200"
    Nonuniform_Flow irc
    H_Start = ch(End_Num)
    QU = 104.7
    Start_Sec = "S3.200"
    End_Sec = "S4.600"
    Nonuniform_Flow irc

    H_Start = ch(End_Num)
    QU = 99.9
    Start_Sec = "S4.600"
    End_Sec = "S6.600"
    Nonuniform_Flow irc

    H_Start = ch(End_Num)
    QU = 94.6
    Start_Sec = "S6.600"
    End_Sec = "S7.000"
    Nonuniform_Flow irc
    H_Start = ch(End_Num)
    QU = 88.2
    Start_Sec = "S7.000"
    End_Sec = "S8.000"
    Nonuniform_Flow irc


    Print #nf, "    N   断面       H         A         Q       V       FR"
    For i = i1 To End_Num
        buf = Format(Format(i, "####0"), "@@@@@  ") & Sec_Name(i)
        buf = buf & Format(Format(ch(i), "###0.000"), "@@@@@@@@@@")
        buf = buf & Format(Format(CA(1, i), "###0.000"), "@@@@@@@@@@")
        buf = buf & Format(Format(CQ(1, i), "###0.000"), "@@@@@@@@@@")
        buf = buf & Format(Format(CV(i), "###0.000"), "@@@@@@@@")
        buf = buf & Format(Format(FR(i), "###0.000"), "@@@@@@@@")
        Print #nf, buf
    Next i

    Close #nf

End Sub
'概要      :フルード数を計算する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :Msec          ,I  ,Integer   ,断面順番号
'          :HH            ,I  ,Single    ,水位
'          :FrX           ,O  ,Single    ,フルード数
'説明      :
Sub Check_Froude_Number(Msec As Integer, Hh As Single, qq As Single, FrX As Single)

    Dim i As Integer, j As Integer, Fr1 As Single, Fr2 As Single
    Dim fv As Single, fx As Single, hmin As Single, HMAX As Single, hw As Single
    Dim fz As Single, Return_Code As Boolean
    Dim XA As Single, XR As Single, XB As Single, XP As Single
    
    Const eps = 0.0001
        
    Call Inner_point_G(Msec, Hh, XA, XR, XB, XP, Return_Code)
    
    fv = qq / XA
    fx = fv / Sqr(9.8 * XR)

    CA(0, Msec) = XA

    If fx <= Froude_Number_Limit Then
        FrX = fx
        Exit Sub
    End If

'限界水深
    hmin = H(1, Msec)
    HMAX = H(NSPC, Msec)
    hw = (hmin + HMAX) * 0.5
    For i = 1 To 50
        hw = (hmin + HMAX) * 0.5
        Call Inner_point_G(Msec, hw, XA, XR, XB, XP, Return_Code)
        fv = qq / XA
        fx = fv / Sqr(9.8 * XR)
        fz = fx - Froude_Number_Limit
        If Abs(fz) < eps Then
           Hh = hw
           CA(0, Msec) = XA
           Exit Sub
        End If
        If fz > 0# Then
            hmin = hw
        Else
            HMAX = hw
        End If
        If Abs(HMAX - hmin) < eps Then
            Hh = hw
            FrX = fx
            CA(0, Msec) = XA
            Exit Sub
        End If
    Next i
    
    FrX = 9999#

End Sub
Sub Inner_point_G(Msec As Integer, Hh As Single, XA As Single, _
                  XR As Single, XB As Single, XP As Single, _
                  Return_Code As Boolean)

    Dim i As Integer, j As Integer, msg As String
    Dim H1 As Single, H2 As Single
    Dim A1 As Single, A2 As Single
    Dim R1 As Single, R2 As Single
    Dim B1 As Single, B2 As Single
    Dim P1 As Single, P2 As Single
    Dim x As Single

    Return_Code = False   'とりあえず

    If Hh < H(1, Msec) Then
        msg = "Error In Inner_point_G " & _
              "断面特性を内挿計算しようとした時にエラー 日光川外水位の異常が考えられる " & _
              "断面名＝（" & Sec_Name(Msec) & ")" & _
              "入力値水位（" & str(Hh) & ")が断面特性表の最小値より小さい" & _
              "断面特性表最小値＝（" & str(H(1, Msec)) & ")"
'        MsgBox MSG
        Log_Calc msg
        Exit Sub
    End If

    x = Hh
    For j = 2 To NSPC   '断面特性の数
        If x < H(j, Msec) Then
            i = j - 1
            H1 = H(i, Msec)
            H2 = H(j, Msec)
             
            A1 = AG(i, Msec)
            A2 = AG(j, Msec)

            R1 = RG(i, Msec)
            R2 = RG(j, Msec)

            B1 = BG(i, Msec)
            B2 = BG(j, Msec)

            P1 = PG(i, Msec)
            P2 = PG(j, Msec)

            XA = (A2 - A1) / (H2 - H1) * (x - H1) + A1
            XR = (R2 - R1) / (H2 - H1) * (x - H1) + R1
            XB = (B2 - B1) / (H2 - H1) * (x - H1) + B1
            XP = (P2 - P1) / (H2 - H1) * (x - H1) + P1

            Return_Code = True
            Exit Sub
        End If
    Next j

    msg = "Error In Inner_point_G" & vbCrLf & _
          "断面特性表を内挿計算しようとした時にエラー  日光川外水位の異常が考えられる " & _
          "断面名＝（" & Sec_Name(Msec) & ")" & _
          "入力値水位（" & str(Hh) & ")が断面特性表の最大値より大きい" & _
          "断面特性表最大値＝（" & str(H(NSPC, Msec)) & ")"
'    MsgBox MSG
    Log_Calc msg


End Sub
Sub Nonuniform_Flow(irc As Boolean)

    Dim i As Integer, j As Integer, m As Integer
    Dim H1 As Single, H2 As Single, hx As Single
    Dim Return_Code As Boolean
    Dim AW1 As Single, BW1 As Single
    Dim AW2 As Single, BW2 As Single
    Dim LX  As Single, RX As Single
    Dim qq  As Single, FrX As Single
    Dim er As Single, msg As String, ans As Integer

    Start_Num = 0
    For i = 1 To Num_Of_Sec
        If Start_Sec = Sec_Name(i) Then
            Start_Num = i
            Exit For
        End If
    Next i
    End_Num = 0
    For i = 1 To Num_Of_Sec
        If End_Sec = Sec_Name(i) Then
            End_Num = i
            Exit For
        End If
    Next i
    If Start_Num = 0 Then
        MsgBox "計算開始の断面が見つからない、計算中止" & vbCrLf & _
               "計算開始断面=(" & Start_Sec & ")"
        End
    End If
    If End_Num = 0 Then
        MsgBox "計算終了の断面が見つからない、計算中止" & vbCrLf & _
               "計算開始断面=(" & End_Sec & ")"
        End
    End If

    Const eps = 0.00001

    qq = QU
    For m = Start_Num To End_Num
        
        CFLAG(m) = " "
        If m = Start_Num Then
            Call Cal_Nonuniform_Flow_Parameter(m, H_Start, AW2, BW2, irc)
            If irc = False Then
                Log_Calc "不等流計算が出来ませんでした、今時刻の予測計算は中止します。"
                Exit Sub
            End If
            hx = H_Start
        Else
            H1 = H(NSPC, m)
            H2 = H(1, m)
            Do
                hx = (H1 + H2) / 2
                Call Cal_Nonuniform_Flow_Parameter(m, hx, AW2, BW2, irc)
                If irc = False Then
                    Log_Calc "不等流計算が出来ませんでした、今時刻の予測計算は中止します。"
                    Exit Sub
                End If
                er = (AW2 - AW1) - (BW1 + BW2) * DeltaX(m) * 0.5
                If Abs(er) < eps Then GoTo CALOK
                If er > 0# Then
                    H1 = hx
                Else
                    H2 = hx
                End If
'                If Abs(h1 - h2) < eps Then GoTo CALBAD
                If Abs(H1 - H2) < eps Then
                    CFLAG(m) = "+"
                    GoTo CALOK
                End If
            Loop
CALBAD:
            msg = "不等流計算収束しませんでした、でもこのまま計算を続ける。"
            Log_Calc msg
'            ans = MsgBox(MSG, vbYesNo)
'            If ans = vbNo Then
'                End
'            End If
CALOK:
        End If
        Call Check_Froude_Number(m, hx, qq, FrX)
        CQ(1, m) = qq
        ch(m) = hx
        CV(m) = qq / CA(1, m)
        FR(m) = Abs(FrX)
        If FrX < 0# Then
            CFLAG(m) = "*"
        End If

        AW1 = AW2
        BW1 = BW2

    Next m

End Sub
