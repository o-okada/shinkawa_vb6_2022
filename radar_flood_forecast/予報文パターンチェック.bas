Attribute VB_Name = "予報文"
'******************************************************************************
'モジュール名：予報文パターンチェック
'
'******************************************************************************
Option Explicit
Option Base 1
Public BP              As Long                                          '前回予報文コード
Public Patan           As Long                                          '今回パターン
Public Xnum            As String                                        'チェックするパターン番号
Public rch             As Boolean
Public PRACTICE_FLG_CODE  As String                                     '"40"=予報  "99"=演習
Public Const はん濫注意水位 = 3#
Public Const 避難判断水位 = 4.4
Public Const はん濫危険水位 = 5.2

'******************************************************************************
'サブルーチン：Pattern_Check()
'処理概要：
'******************************************************************************
Sub Pattern_Check()
    rch = False
    Select Case BP
        Case 0
            Xnum = "1,5,10"
        Case 1
            Xnum = "2,4,5,6,7,10"
        Case 2
            Xnum = "4,5,6,7,10"
        Case 3
            Xnum = "4,5,6,7,10,11"
        Case 4
            Patan = 0
            BP = 0
            Wng_Last_Time = 0                                           '注意文のランクを初期化
            rch = True
            Xnum = "0"
        Case 5
            Xnum = "3,4,8,10"
        Case 6
            Xnum = "3,4,8,10"
        Case 7
            Xnum = "3,4,8,10"
        Case 8
            Xnum = "3,4,10,12"
        Case 9
            Xnum = "3,4,10,12"
        Case 10
            Xnum = "3,4,9,13"
        Case 11
            Xnum = "4,5,6,7,10"
        Case 12
            Xnum = "3,4,10"
        Case 13
            Xnum = "3,4,9"
    End Select
    水位_Check
    If rch Then                                                         '以下の文は本番では関係ないテスト時のみ有効 2008/08/30 check
        BP = Patan
    End If
End Sub

'******************************************************************************
'サブルーチン：洪水予報文初期化()
'処理概要：
'******************************************************************************
Sub 洪水予報文初期化()
    Dim nf   As Integer
    Dim j    As Integer
    Dim buf  As String
    Dim a
    LOG_Out "IN  洪水予報文初期化"
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    '******************************************************
    'nf = FreeFile
    'Open App.Path & "\data\予報文出力.txt" For Input As #nf
    'Input #nf, buf
    'j = CInt(Mid(buf, 1, 5))
    'If j = 1 Then
    '    DBX_ora = True
    '    AutoDrive.Option1(0).Value = True
    'Else
    '    DBX_ora = False
    '    AutoDrive.Option1(1).Value = True
    'End If
    'Input #nf, buf '水位タイトル
    'Input #nf, buf
    'a = Mid(buf, 1, 10)
    'If IsNumeric(a) Then
    '    危険水位 = CSng(a)
    'Else
    '    MsgBox "入力した危険水位は数値ではありません" & vbLf & _
    '           "オラクルＤＢには出力しないモードで計算ます。" & vbLf & _
    '           "計算を中止します。"
    '    End
    'End If
    'a = Mid(buf, 11, 10)
    'If IsNumeric(a) Then
    '    警戒水位 = CSng(a)
    'Else
    '    MsgBox "入力した警戒水位は数値ではありません" & vbLf & _
    '           "オラクルＤＢには出力しないモードで計算ます。" & vbLf & _
    '           "計算を中止します。"
    '    End
    'End If
    'a = Mid(buf, 20, 10)
    'If IsNumeric(a) Then
    '    指定水位 = CSng(a)
    'Else
    '    MsgBox "入力した指定水位は数値ではありません" & vbLf & _
    '           "オラクルＤＢには出力しないモードで計算ます。" & vbLf & _
    '           "計算を中止します。"
    '    End
    'End If
    'Close #nf
    '******************************************************
    'Ver0.0.0 修正終了 1900/01/01 00:00
    '******************************************************
    PRACTICE_FLG_CODE = "40"                                            '予報文本ちゃんモードを初期値とする
    AutoDrive.Option2(0).Value = True
    LOG_Out "OUT 洪水予報文初期化"
End Sub

'******************************************************************************
'サブルーチン：水位_Check()
'処理概要：
'******************************************************************************
Sub 水位_Check()
    Dim i      As Long
    Dim n      As Long
    Dim m      As Long
    Dim w
    Dim HOM2   As Single                                                '実績2時間前水位
    Dim HOM1   As Single                                                '実績1時間前水位
    Dim HON    As Single                                                '実績 現時刻水位
    Dim HC1    As Single                                                '予測1時間後水位
    Dim HC2    As Single                                                '予測2時間後水位
    Dim HC3    As Single                                                '予測3時間後水位
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    '******************************************************
    'HOM2 = 水位.hm2
    'HOM1 = 水位.hm1
    'HON = 水位.h
    'HC1 = 水位.hy1
    'HC2 = 水位.hy2
    'HC3 = 水位.hy3
    '******************************************************
    'Ver0.0.0 修正終了 1900/01/01 00:00
    '******************************************************
    水位.hm2 = HO(5, Now_Step - 2)
    水位.hm1 = HO(5, Now_Step - 1)
    水位.H = HO(5, Now_Step)
    水位.hy1 = HQ(1, 41, NT - 12)
    水位.hy2 = HQ(1, 41, NT - 6)
    水位.hy3 = HQ(1, 41, NT)
    HOM2 = 水位.hm2
    HOM1 = 水位.hm1
    HON = 水位.H
    HC1 = 水位.hy1
    HC2 = 水位.hy2
    HC3 = 水位.hy3
    w = Split(Xnum, ",")
    n = UBound(w)
    For i = 0 To n
        m = w(i)                                                        'パターン番号
        Select Case m
            Case 1
                If (HON < 避難判断水位) And (HON >= はん濫注意水位) Then
                    If (HC3 < はん濫危険水位) Then
                        If (HC1 < はん濫危険水位) And (HC1 >= はん濫注意水位) Then
                            If (HC2 < はん濫危険水位) And (HC2 >= はん濫注意水位) Then
                                Patan = m
                                rch = True
                            End If
                        End If
                    End If
                End If
            Case 2
                If (はん濫危険水位 > HON) And (HON >= 避難判断水位) Then
                    If (避難判断水位 > HC3) And (HC3 >= はん濫注意水位) Then
                        Patan = m
                        rch = True
                    End If
                End If
            Case 3
                If (避難判断水位 > HON) And (HON >= はん濫注意水位) Then
                    If (避難判断水位 > HC3) Then
                        Patan = m
                        rch = True
                    End If
                End If
            Case 4
                If (HON < はん濫注意水位) Then
                    If (HC1 < はん濫注意水位) Then
                        If (HC2 < はん濫注意水位) Then
                            If (HC3 < はん濫注意水位) Then
                                Patan = m
                                rch = True
                            End If
                        End If
                    End If
                End If
            Case 5
                If (避難判断水位 > HON) Then
                    If (はん濫危険水位 <= HC3) Then
                        Patan = m
                        rch = True
                    End If
                End If
            Case 6
                If (はん濫危険水位 > HON) And (HON >= 避難判断水位) Then
                    If (はん濫危険水位 <= HC3) Then
                        Patan = m
                        rch = True
                    End If
                End If
            Case 7
                If (はん濫危険水位 > HON) And (HON >= 避難判断水位) Then
                    If (はん濫危険水位 > HC3) And (HC3 >= 避難判断水位) Then
                        Patan = m
                        rch = True
                    End If
                End If
            Case 8
                If (はん濫危険水位 > HON) And (HON >= 避難判断水位) Then
                    If (はん濫危険水位 <= HC3) Then
                        Patan = m
                        rch = True
                    End If
                End If
            Case 9
                If (はん濫危険水位 > HON) And (HON >= 避難判断水位) Then
                    If (はん濫危険水位 > HC3) Then
                        Patan = m
                        rch = True
                    End If
                End If
            Case 10
                If (はん濫危険水位 <= HON) And (はん濫危険水位 > HOM1) And (はん濫危険水位 > HOM2) Then
                    Patan = m
                    rch = True
                End If
            Case 11
                If (避難判断水位 > HON) And (はん濫注意水位 <= HON) Then
                    If (避難判断水位 > HOM1) And (はん濫注意水位 <= HOM1) Then
                        If (避難判断水位 > HOM2) And (はん濫注意水位 <= HOM2) Then
                            If (避難判断水位 > HC1) And (はん濫注意水位 <= HC1) Then
                                If (避難判断水位 > HC2) And (はん濫注意水位 <= HC2) Then
                                    If (避難判断水位 > HC3) And (はん濫注意水位 <= HC3) Then
                                        Patan = m
                                        rch = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Case 12
                If (はん濫危険水位 > HON) And (HON >= 避難判断水位) Then
                    If (はん濫危険水位 > HOM1) And (HOM1 >= 避難判断水位) Then
                        If (はん濫危険水位 > HOM2) And (HOM2 >= 避難判断水位) Then
                            If (はん濫危険水位 > HC1) And (HC1 >= 避難判断水位) Then
                                If (はん濫危険水位 > HC2) And (HC2 >= 避難判断水位) Then
                                    If (はん濫危険水位 > HC3) And (HC3 >= 避難判断水位) Then
                                        Patan = m
                                        rch = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Case 13
                If (はん濫危険水位 <= HON) And (はん濫危険水位 <= HOM1) And (はん濫危険水位 <= HOM2) Then
                    If (はん濫危険水位 <= HC1) And (はん濫危険水位 <= HC2) And (はん濫危険水位 <= HC3) Then
                        Patan = m
                        rch = True
                    End If
                End If
        End Select
    Next i
    Exit Sub
jump:
    Patan = m
End Sub

'******************************************************************************
'サブルーチン：予報文チェック()
'処理概要：
'******************************************************************************
Sub 予報文チェック()
    '******************************************************
    'Ver1.0.0 修正開始 2015/08/06 O.OKADA【01-01】
    '※オラクルデータベースのテーブル「」の削除に対応し、下記のとおり修正する。
    '******************************************************
    Exit Sub
    '******************************************************
    'Ver1.0.0 修正終了 2015/08/06 O.OKADA【01-01】
    '******************************************************
    予報文履歴DB_Read
    Pattern_Check
    If Patan > 0 Then
        '******************************************************
        'Ver0.0.0 修正開始 1900/01/01 00:00
        '******************************************************
        'パターン文集_Read                                              'AutoDriveのForm Loadで読むようにした
        '******************************************************
        'Ver0.0.0 修正終了 1900/01/01 00:00
        '******************************************************
        Pattan_Now = Patan
        主文作成2
        If DBX_ora Then
            ORA_YOHOUBUNAN
        End If
        予報文履歴DB_Write
    End If
    Patan = 0
End Sub
