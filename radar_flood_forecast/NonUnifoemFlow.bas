Attribute VB_Name = "NonUnifoemFlow"
Option Explicit
Option Base 1

'Public DLIB As New DLIB1.Class1            '���C�u�����[

Public Const NSEC = 500                    '�ő�v�Z�f�ʐ�
Public Const NSPC = 15                     '�f�ʓ����̐�
Public Const NSEP = 1                      '�f�ʕ�����
Public Const Froude_Number_Limit = 1#      '�t���[�h���`�F�b�N�l
'--------  �f�ʏ���  ---------------------------------------------
Public Num_Of_Sec              As Integer  '�v�Z�f�ʐ�
Public Sec_Name(NSEC)          As String   '�f�ʖ�
Public DeltaX(NSEC)            As Single   '��ԋ���
Public NBLBR(NSEC)             As Integer  '�e�f�ʖ��̂a�k�a�q�̐�
Public BLBR(NSEP + 1, NSEC)    As Single   '�e�f�ʂ̂a�k�a�q
Public n(NSEP, NSEC)           As Single   '�d�x�W��
Public CS(NSEC)                As Single   '�f�ʊp�x�␳

'--------  �f�ʓ���  ---------------------------------------------
Public H(NSPC, NSEC)           As Single   '����
Public ZS(NSEC)                As Single   '�Ő[�͏����W
Public AG(NSPC, NSEC)          As Single   '�����f�ʉ͐�
Public RG(NSPC, NSEC)          As Single   '�����f�ʌa�[
Public BG(NSPC, NSEC)          As Single   '�����f�ʐ��ʕ�
Public PG(NSPC, NSEC)          As Single   '�����f�ʏ���
Public NG(NSPC, NSEC)          As Single   '�����f�ʑd�x

'--------  �s�����v�Z����  ----------------------------------------
Public CQ(0 To NSEP, NSEC)     As Single   '����
Public CV(NSEC)                As Single   '����
Public ch(NSEC)                As Single   '����
Public CR(NSEC)                As Single   '�a�[
Public FR(NSEC)                As Single   '�t���[�h��
Public CA(0 To NSEP, NSEC)     As Single   '�����f�ʖ��̉͐�
Public CD(NSEC)                As Single   '�G�l���M�[�␳�W��
Public CFLAG(NSEC)             As String   '�v�Z�t���O

'--------  �s�����v�Z�p�����[�^  ----------------------------------
Public Alpha                   As Single   '�G�l���M�[�␳�W��
'Public Froude_Number_Limit     As Single   '�t���[�h�����E�l
Public Start_Sec               As String   '�s�����v�Z�J�n�f�ʋL��
Public End_Sec                 As String   '�s�����v�Z�I���f�ʋL��
Public Start_Num               As Integer  '�s�����v�Z�J�n�f�ʏ��ԍ�
Public End_Num                 As Integer  '�s�����v�Z�I���f�ʏ��ԍ�
'--------  �s�����v�Z���E����  ------------------------------------
Public QU                      As Single   '����
Public H_Start                 As Single   '�����[����
'--------  �t�@�C���֌W  ------------------------------------------
Public open_data               As String
Public Log_CALC_ERROR          As String   '�v�Z���~���O�t�@�C���ԍ�
Public Log_CALC_N              As Long     '�v�Z���O�o�̓��C����
'�T�v      :�s�����v�Z�p�x�[�X�f�[�^�ǂݍ��݁B
'����      :�f�[�^�ǂݍ��݁B
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
'�f�ʐ�
    Do
        Line Input #nf, buf
        If Mid(buf, 1, 2) = "AR" Then
            Num_Of_Sec = CInt(Mid(buf, 6, 5))
            Exit Do
        End If
    Loop
'�f�ʖ�,�d�x�W��,��ԋ����ǂݍ���
    Line Input #nf, buf '�X�P�[���t�@�N�^�[
    If Mid(buf, 1, 1) <> "S" Then
        MsgBox "�s�藬�v�Z�̃f�[�^�\�����Ⴄ" & vbCrLf & _
               "�d�x�v���A��ԋ����f�[�^�̃X�P�[���t�@�N�^�[��ǂݍ��ނƂ���ɈႤ�f�[�^������" & vbCrLf & _
               "�f�[�^=(" & buf & ")" & vbCrLf & _
               "�v�Z�𒆎~���܂��B", vbExclamation
        End
    End If
    SFn = CSng(Mid(buf, 11, 5))    '�d�x�W���p
    SFdx = CSng(Mid(buf, 16, 5))   '��ԋ����p
    For i = 1 To Num_Of_Sec
        Line Input #nf, buf
            Sec_Name(i) = Mid(buf, 5, 6)              '�f�ʖ�
            NG(1, i) = CSng(Mid(buf, 11, 5)) * SFn    '�d�x�v��
            DeltaX(i) = CSng(Mid(buf, 16, 5)) * SFdx  '��ԋ���
            ZS(i) = CSng(Mid(buf, 36, 5))             '�Ő[�͏�
    Next i
'�f�ʓ��� ���ʓǂݍ���
    Line Input #nf, buf '�X�P�[���t�@�N�^�[
    If Mid(buf, 1, 1) <> "H" Then
        MsgBox "�s�藬�v�Z�̃f�[�^�\�����Ⴄ" & vbCrLf & _
               "���ʃf�[�^�̃X�P�[���t�@�N�^�[��ǂݍ��ނƂ���ɈႤ�f�[�^������" & vbCrLf & _
               "�f�[�^=(" & buf & ")" & vbCrLf & _
               "�v�Z�𒆎~���܂��B", vbExclamation
        End
    End If
    SF = CSng(Mid(buf, 11, 5)) '���ʗp
    For i = 1 To Num_Of_Sec
        Line Input #nf, buf
        If Mid(buf, 5, 6) <> Sec_Name(i) Then
            MsgBox "���ʒf�ʓ�����ǂݍ��ݒ��ɃG���[�����������B" & vbCrLf & _
                   "�G���[���f�ʋL�����Ⴄ(" & Sec_Name(i) & ")��(" & buf & ")�ɂȂ��Ă���B" & vbCrLf & _
                   "�v�Z�𒆎~���܂��B", vbExclamation
            End
        End If
        H(1, i) = ZS(i)
        For j = 1 To NSPCx '�f�ʓ����̐�
            H(j + 1, i) = CSng(Mid(buf, 11 + (j - 1) * 5, 5)) * SF
        Next j
    Next i
'�f�ʓ��� ���ʕ��ǂݍ���
    Line Input #nf, buf '�X�P�[���t�@�N�^�[
    If Mid(buf, 1, 1) <> "B" Then
        MsgBox "�s�藬�v�Z�̃f�[�^�\�����Ⴄ" & vbCrLf & _
               "���ʕ��f�[�^�̃X�P�[���t�@�N�^�[��ǂݍ��ނƂ���ɈႤ�f�[�^������" & vbCrLf & _
               "�f�[�^=(" & buf & ")" & vbCrLf & _
               "�v�Z�𒆎~���܂��B", vbExclamation
        End
    End If
    SF = CSng(Mid(buf, 11, 5)) '���ʕ��p
    For i = 1 To Num_Of_Sec
        Line Input #nf, buf
        If Mid(buf, 5, 6) <> Sec_Name(i) Then
            MsgBox "���ʒf�ʓ�����ǂݍ��ݒ��ɃG���[�����������B" & vbCrLf & _
                   "�G���[���f�ʋL�����Ⴄ(" & Sec_Name(i) & ")��(" & buf & ")�ɂȂ��Ă���B" & vbCrLf & _
                   "�v�Z�𒆎~���܂��B", vbExclamation
            End
        End If
        BG(1, i) = 0#
        For j = 1 To NSPCx '�f�ʓ����̐�
            BG(j + 1, i) = CSng(Mid(buf, 11 + (j - 1) * 5, 5)) * SF
        Next j
    Next i
'�f�ʓ��� �͐ϓǂݍ���
    Line Input #nf, buf '�X�P�[���t�@�N�^�[
    If Mid(buf, 1, 1) <> "A" Then
        MsgBox "�s�藬�v�Z�̃f�[�^�\�����Ⴄ" & vbCrLf & _
               "�͐σf�[�^�̃X�P�[���t�@�N�^�[��ǂݍ��ނƂ���ɈႤ�f�[�^������" & vbCrLf & _
               "�f�[�^=(" & buf & ")" & vbCrLf & _
               "�v�Z�𒆎~���܂��B", vbExclamation
        End
    End If
    SF = CSng(Mid(buf, 11, 5)) '�͐ϗp
    For i = 1 To Num_Of_Sec
        Line Input #nf, buf
        If Mid(buf, 5, 6) <> Sec_Name(i) Then
            MsgBox "�͐ϒf�ʓ�����ǂݍ��ݒ��ɃG���[�����������B" & vbCrLf & _
                   "�G���[���f�ʋL�����Ⴄ(" & Sec_Name(i) & ")��(" & buf & ")�ɂȂ��Ă���B" & vbCrLf & _
                   "�v�Z�𒆎~���܂��B", vbExclamation
            End
        End If
        AG(1, i) = 0#
        For j = 1 To NSPCx '�f�ʓ����̐�
            AG(j, i) = CSng(Mid(buf, 11 + (j - 1) * 5, 5)) * SF
        Next j
    Next i
'�f�ʓ��� �a�[�ǂݍ���
    Line Input #nf, buf '�X�P�[���t�@�N�^�[
    If Mid(buf, 1, 1) <> "R" Then
        MsgBox "�s�藬�v�Z�̃f�[�^�\�����Ⴄ" & vbCrLf & _
               "�a�[�f�[�^�̃X�P�[���t�@�N�^�[��ǂݍ��ނƂ���ɈႤ�f�[�^������" & vbCrLf & _
               "�f�[�^=(" & buf & ")" & vbCrLf & _
               "�v�Z�𒆎~���܂��B", vbExclamation
        End
    End If
    SF = CSng(Mid(buf, 11, 5)) '�a�[�p
    For i = 1 To Num_Of_Sec
        Line Input #nf, buf
        If Mid(buf, 5, 6) <> Sec_Name(i) Then
            MsgBox "�a�[�f�ʓ�����ǂݍ��ݒ��ɃG���[�����������B" & vbCrLf & _
                   "�G���[���f�ʋL�����Ⴄ(" & Sec_Name(i) & ")��(" & buf & ")�ɂȂ��Ă���B" & vbCrLf & _
                   "�v�Z�𒆎~���܂��B", vbExclamation
            End
        End If
        RG(1, i) = 0#
        For j = 1 To NSPCx '�f�ʓ����̐�
            RG(j, i) = CSng(Mid(buf, 11 + (j - 1) * 5, 5)) * SF
        Next j
    Next i
'�a�k�a�q�ݒ�
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
'            MsgBox "���ʂ��Ⴗ���Čv�Z�ł��܂���A�����������ʂ��㏸���Ă���v�Z���Ă��������B"
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


    Print #nf, "    N   �f��       H         A         Q       V       FR"
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
'�T�v      :�t���[�h�����v�Z����B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :Msec          ,I  ,Integer   ,�f�ʏ��ԍ�
'          :HH            ,I  ,Single    ,����
'          :FrX           ,O  ,Single    ,�t���[�h��
'����      :
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

'���E���[
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

    Return_Code = False   '�Ƃ肠����

    If Hh < H(1, Msec) Then
        msg = "Error In Inner_point_G " & _
              "�f�ʓ�������}�v�Z���悤�Ƃ������ɃG���[ ������O���ʂُ̈킪�l������ " & _
              "�f�ʖ����i" & Sec_Name(Msec) & ")" & _
              "���͒l���ʁi" & str(Hh) & ")���f�ʓ����\�̍ŏ��l��菬����" & _
              "�f�ʓ����\�ŏ��l���i" & str(H(1, Msec)) & ")"
'        MsgBox MSG
        Log_Calc msg
        Exit Sub
    End If

    x = Hh
    For j = 2 To NSPC   '�f�ʓ����̐�
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
          "�f�ʓ����\����}�v�Z���悤�Ƃ������ɃG���[  ������O���ʂُ̈킪�l������ " & _
          "�f�ʖ����i" & Sec_Name(Msec) & ")" & _
          "���͒l���ʁi" & str(Hh) & ")���f�ʓ����\�̍ő�l���傫��" & _
          "�f�ʓ����\�ő�l���i" & str(H(NSPC, Msec)) & ")"
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
        MsgBox "�v�Z�J�n�̒f�ʂ�������Ȃ��A�v�Z���~" & vbCrLf & _
               "�v�Z�J�n�f��=(" & Start_Sec & ")"
        End
    End If
    If End_Num = 0 Then
        MsgBox "�v�Z�I���̒f�ʂ�������Ȃ��A�v�Z���~" & vbCrLf & _
               "�v�Z�J�n�f��=(" & End_Sec & ")"
        End
    End If

    Const eps = 0.00001

    qq = QU
    For m = Start_Num To End_Num
        
        CFLAG(m) = " "
        If m = Start_Num Then
            Call Cal_Nonuniform_Flow_Parameter(m, H_Start, AW2, BW2, irc)
            If irc = False Then
                Log_Calc "�s�����v�Z���o���܂���ł����A�������̗\���v�Z�͒��~���܂��B"
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
                    Log_Calc "�s�����v�Z���o���܂���ł����A�������̗\���v�Z�͒��~���܂��B"
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
            msg = "�s�����v�Z�������܂���ł����A�ł����̂܂܌v�Z�𑱂���B"
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
