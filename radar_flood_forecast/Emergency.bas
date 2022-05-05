Attribute VB_Name = "Emergency"
Option Explicit
Option Base 1
Sub EMG_Cal(irc As Boolean)

    Dim nf             As Long
    Dim buf            As String
    Dim i              As Long
    Dim j              As Long
    Dim J1             As Long
    Dim j2             As Long
    Dim msg            As String
    Dim Haruhi_Q(120)  As Single
    Dim Kujino_Q(120)  As Single
    Dim Nikkou_H(120)  As Single
    Dim Q_kuji         As Single
    Dim Q_Haru         As Single
    Dim H_Sea          As Single
    Dim i1             As Integer
    Dim nr             As Integer
    Dim ns             As Integer
    Dim hd             As Single
    Dim qq             As Single
'DBóp
    Dim GH(5)          As Single

    NT = 18

    ReDim HQ(3, nd, NT)

'    On Error GoTo ERH1


'ãvínñÏó¨ì¸ó ì«Ç›çûÇ›
    msg = "ãvínñÏó¨ì¸ó ì«Ç›çûÇ›íÜÉGÉâÅ[Ç™î≠ê∂ÇµÇΩÅB" & vbCrLf & _
          "ERROR IN EMG_Cal"
    nf = FreeFile
    Open App.Path & "\Work\NSK_KUJINO.DAT" For Input As #nf
    Line Input #nf, buf
    i = 0
    Do Until EOF(nf)
        i = i + 1
        Line Input #nf, buf
        If IsNumeric(Mid(buf, 6, 10)) Then
            Kujino_Q(i) = CSng(Mid(buf, 6, 10))
        Else
            Kujino_Q(i) = 11
        End If
    Loop
    Close #nf

'ètì˙ó¨ì¸ó ì«Ç›çûÇ›
    msg = "ètì˙ó¨ì¸ó ì«Ç›çûÇ›íÜÉGÉâÅ[Ç™î≠ê∂ÇµÇΩÅB" & vbCrLf & _
          "ERROR IN EMG_Cal"
    nf = FreeFile
    Open App.Path & "\Work\NSK_HARUHI.DAT" For Input As #nf
    Line Input #nf, buf
    i = 0
    Do Until EOF(nf)
        i = i + 1
        Line Input #nf, buf
        If IsNumeric(Mid(buf, 6, 10)) Then
            Haruhi_Q(i) = CSng(Mid(buf, 6, 10))
        Else
            Haruhi_Q(i) = 11
        End If
    Loop
    Close #nf

'â∫ó¨í[êÖà ì«Ç›çûÇ›
    msg = "â∫ó¨í[êÖà ì«Ç›çûÇ›íÜÉGÉâÅ[Ç™î≠ê∂ÇµÇΩÅB" & vbCrLf & _
          "ERROR IN EMG_Cal"
    nf = FreeFile
    Open App.Path & "\Work\NSK_â∫ó¨í[êÖà .DAT" For Input As #nf
    Line Input #nf, buf
    i = 0
    Do Until EOF(nf)
        i = i + 1
        Line Input #nf, buf
        Nikkou_H(i) = CSng(Mid(buf, 17, 10))
    Loop
    Close #nf
    J1 = i - 3
    j2 = i

    For i = J1 To j2

        Q_kuji = Kujino_Q(i)
        Q_Haru = Haruhi_Q(i)
        H_Sea = Nikkou_H(i)

        Print #Log_Num, " ã´äEèåèÅ@ãvínñÏó¨ì¸= " & Format(Q_kuji, "###0.00") & _
                        "   ètì˙ó¨ì¸= " & Format(Q_Haru, "###0.00") & _
                        "   â∫ó¨í[êÖà = " & Format(H_Sea, "##0.000")


        QU = Q_kuji + Q_Haru
        H_Start = H_Sea
        Start_Sec = "S0.000"
        End_Sec = "S12.40"
        Nonuniform_Flow irc
        If irc = False Then GoTo ERH2
        i1 = Start_Num

        QU = Q_kuji
        H_Start = ch(End_Num)
        Start_Sec = "S12.40"
        End_Sec = "S20.00"
        Nonuniform_Flow irc
        If irc = False Then GoTo ERH2

        H_Start = ch(Start_Num)
        QU = Q_Haru
        Start_Sec = "G0.000"
        End_Sec = "G8.200"
        Nonuniform_Flow irc
        If irc = False Then GoTo ERH2

'        Print #Log_Num, "    N   ífñ        H         A         Q       V       FR     FLAG"
        For j = i1 To End_Num
'            buf = Format(Format(j, "####0"), "@@@@@  ") & Sec_Name(j)
'            buf = buf & Format(Format(ch(j), "###0.000"), "@@@@@@@@@@")
'            buf = buf & Format(Format(CA(1, j), "###0.000"), "@@@@@@@@@@")
'            buf = buf & Format(Format(CQ(1, j), "###0.000"), "@@@@@@@@@@")
'            buf = buf & Format(Format(CV(j), "###0.000"), "@@@@@@@@")
'            buf = buf & Format(Format(FR(j), "###0.000"), "@@@@@@@@")
'            buf = buf & Space(5) & CFLAG(j)
'            Print #Log_Num, buf
            YHJ(i - J1, j) = ch(j)  'ècífê}ÇcÇaóp
        Next j

        For nr = 1 To 5
            ns = V_Sec_Num(nr)
            Nonuni_H(nr, i - J1) = ch(ns)
        Next nr

    Next i

    'åªéûçèçáÇÌÇπ
    For nr = 1 To 5
        CO(nr, 1) = Nonuni_H(nr, 0)
        CO(nr, 2) = Nonuni_H(nr, 1)
        CO(nr, 3) = Nonuni_H(nr, 2)
        CO(nr, 4) = Nonuni_H(nr, 3)
'---
        Debug.Print "  nr="; nr; "  0="; Nonuni_H(nr, 0);
        Debug.Print "  1="; Nonuni_H(nr, 1);
        Debug.Print "  2="; Nonuni_H(nr, 2);
        Debug.Print "  3="; Nonuni_H(nr, 3)

        hd = HO(nr + 2, Now_Step) - Nonuni_H(nr, 0)
        YHK(nr, 0) = HO(nr + 2, Now_Step)
        Slide1(nr) = hd
        Slide2(nr) = 0#
        Delta_H(nr) = 0#
        For i = 0 To 3
            Nonuni_H(nr, i) = Nonuni_H(nr, i) + hd
            If i > 0 Then
                CF(nr, i) = Nonuni_H(nr, i)
            End If
        Next i
    Next nr

    For i = 1 To 5
        ns = V_Sec_Num(i)
        HQ(1, ns, NT - 12) = Nonuni_H(i, 1)
        HQ(1, ns, NT - 6) = Nonuni_H(i, 2)
        HQ(1, ns, NT) = Nonuni_H(i, 3)
        YHK(i, NT - 12) = Nonuni_H(i, 1)  'åßÇcÇaèoóÕóp
        YHK(i, NT - 6) = Nonuni_H(i, 2)   'åßÇcÇaèoóÕóp
        YHK(i, NT) = Nonuni_H(i, 3)       'åßÇcÇaèoóÕóp
        'ècífï‚ê≥óp
        OHJ(0, i) = Nonuni_H(i, 1)
        OHJ(1, i) = HQ(1, ns, NT - 12)
        OHJ(2, i) = HQ(1, ns, NT - 6)
        OHJ(3, i) = HQ(1, ns, NT)
    Next i

    For i = 1 To 5
        Naisou i
    Next i

    On Error GoTo 0
    Exit Sub

ERH1:
'    MsgBox MSG
    Log_Calc msg
    On Error GoTo 0
    Exit Sub
ERH2:
    Log_Calc "ïsìôó¨åvéZäJénífñ ÅÅ" & Start_Sec & " èIóπífñ ÅÅ" & End_Sec
    On Error GoTo 0
    irc = False

End Sub
Sub Naisou(j As Long)

    Dim i    As Long
    Dim k    As Long
    Dim i1   As Long
    Dim i2   As Long
    Dim w    As Single
    Dim m    As Long


    For m = 1 To 3
        Select Case m
            Case 1
                i1 = 0
                i2 = 6
            Case 2
                i1 = 6
                i2 = 12
            Case 3
                i1 = 12
                i2 = 18
        End Select

        w = (YHK(j, i2) - YHK(j, i1)) / 6#
        k = 1
        For i = i1 + 1 To i2 - 1
            YHK(j, i) = YHK(j, i1) + w * k
            k = k + 1
        Next i
    Next m

End Sub
