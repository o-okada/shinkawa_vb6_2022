Attribute VB_Name = "Module1"
'******************************************************************************
'���W���[�����FModule1
'
'******************************************************************************
Option Explicit
Option Base 1

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'******************************************************************************
'
'******************************************************************************
Public Run_Mode             As Boolean                                  'False=�f�o�b�O���[�h��  True=�ʏ탂�[�h��
Public Current_Path         As String                                   '�J�����g�t�H���_�̃p�X(D:\Shinkaw\ �̂悤��)
Public DBX_ora              As Boolean                                  'true=���ʂ��I���N���ɏ��� false=���ʂ��I���N���ɏ����Ȃ�
Public Log_Repo             As Long
'******************************************************************************
'
'******************************************************************************
Public Flood_Name           As String                                   '�^���f�[�^�^�C�g��
Public Now_Step             As Integer                                  '�v�Z�J�n�������猻�����܂ł̃X�e�b�v��
Public Const Yosoku_Step = 3
Public All_Step             As Integer                                  '�v�Z�J�n��������\���I���܂ł̃X�e�b�v��
Public Data_Steps           As Integer                                  '���̓f�[�^�X�e�b�v��
Public js(6)                As Integer                                  '�f�[�^�J�n����
Public jg(6)                As Integer                                  '�f�[�^�I������
Public jx(6)                As Integer                                  '�������[�N
Public jsd                  As Date                                     'js()�̓��t�^
Public jgd                  As Date                                     'jg()�̓��t�^
Public jxd                  As Date                                     'jx()�̓��t�^
Public NSK_jsd              As Date                                     '�s�藬�v�Z�J�n����
Public Data_Pich(5)         As Single
'******************************************************************************
'
'******************************************************************************
Public Return_Code_fm       As Boolean                                  '
'******************************************************************************
'
'******************************************************************************
Public Log_Num              As Integer                                  '���O�t�@�C���ɏ����o���t�@�C���ԍ�
Public Log_Time             As Date                                     '���O�t�@�C�����N���A���邽�߂̎����Ǘ��p
'******************************************************************************
'
'******************************************************************************
Public Const nd = 74 '215                                               '�f�ʐ�
Public Const aksk = -99#                                                '�����萔
'******************************************************************************
'
'******************************************************************************
Public Y_DS                 As Date                                     '�f�[�^�J�n����
Public Y_DE                 As Date                                     '�f�[�^�I������
Public NDN()                As String                                   '�f�ʋL��
Public NT                   As Integer                                  '�v�Z���Ԑ�
Public HQ()                 As Single                                   '�\���v�Z����
Public DX()                 As Single                                   '��ԋ���
Public sdx()                As Single                                   '�݉���ԋ���
Public MDX                  As Single                                   '�͓���
Public MAX_H()              As Single                                   '�ő吅��
'******************************************************************************
'
'******************************************************************************
Public Input_file           As String                                   '�����^���f�[�^�̃t�@�C����
'******************************************************************************
'
'******************************************************************************
Public Initial_HQ(2, nd)    As Single                                   '�v�Z�J�n���ʌ`
Public Tide(500)            As Single                                   '�����[���E����
Public Const Hnum = 7                                                   '���ʊϑ�����
Public Name_H(10)           As String                                   '���ʊϑ�����
Public HO(10, 500)          As Single                                   '�ϑ�������
Public HO_Title             As String                                   '���ʃ^�C�g��
Public Name_R(10)           As String                                   '�J�ʊϑ�����
Public RO(10, 500)          As Single                                   '��n�_�㗬����J��
Public RO_Title             As String                                   '�J�ʃ^�C�g��
Public Const Rnum = 10                                                  '�J�ʊϑ�����
Public Wpath                As String                                   '�f�[�^�o�̓f�B���N�g���p�X
Public HQA(2)               As Single                                   'H-Q����A
Public HQB(2)               As Single                                   'H-Q����B
'******************************************************************************
'
'******************************************************************************
Public V_Sec_Name(10, 2)    As String                                   '���ؒn�_��  1=�v�Z�f�ʖ�  2=�ϑ�����
Public V_Sec_Num(10)        As Integer                                  '���ؒn�_�f�ʏ��ԍ�
Public V_Sec_Cnt            As Integer                                  '���ؒn�_��
Public Froude               As Single                                   '�s�藬�v�Z���ʂ̊�f�ʂ̕��σt���[�h��
'******************************************************************************
'
'******************************************************************************
Public H_Scale(5, 3)        As Single                                   '���ʖڐ���  1=���ڐ��� 2=��ڐ��� 3=�s�b�`
Public Q_Scale(5, 3)        As Single                                   '���ʖڐ���  1=���ڐ��� 2=��ڐ��� 3=�s�b�`
Public H_Stand1(5, 5)       As Single                                   '��n�_�A���ʁ�( 1=�g�v�k 2=��O�  3=��� 4=��� 5=�[���_�� )
Public H_Stand1t(5, 5)      As String                                   '��n�_���ʖ���
Public H_Stand2(5, 3)       As Single                                   '��n�_�A�|���v���ʁ�( 1=��~���� 2=�ĊJ���� 3=�������� )
Public H_Stand2t(5, 3)      As String                                   '�|���v���ʖ���
Public H_Standi(5, 2)       As Integer                                  '��n�_�A����ʐ���( 1=� 2=�|���v )
'******************************************************************************
'
'******************************************************************************
Public OBS1                 As Boolean                                  '���уf�[�^�S���v���b�g����Ƃ�=True
Public CAL1                 As Integer                                  '1=�v�Z�l�v���b�g�L��  0=�v�Z�l�v���b�g����
'******************************************************************************
'�s�����v�Z���p�����[�^
'******************************************************************************
Public Q_kuji               As Single                                   '�s���v�Z�p�v�n�쏉������
Public Q_Haru               As Single                                   '�s�����v�Z�p�t����������
Public H_Sea                As Single                                   '�s�����v�Z�p�����[����
'******************************************************************************
'
'******************************************************************************
Public Nonuni_H(5, 0 To 3)  As Single                                   '�s�����v�Z�ɂ��\������
Public CO(5, 4)             As Single                                   '�s�藬�v�Z���l
Public CF(5, 0 To 3)        As Single                                   '�s�藬�v�Z�t�B�[�h�o�b�N��
'******************************************************************************
'�t�B�[�h�o�b�N�␳�l
'******************************************************************************
Public Slide1(5)            As Single                                   '�������O�Q���Ԃ̂P��ڃX���C�h��
Public Slide2(5)            As Single                                   '�������@�@�@�@�̂Q��ڃX���C�h��
Public Delta_H(5)           As Single                                   '�P���ԓ���̐��ʕ␳�l�i�Q���Ԍ�͂Q�{�R���Ԍ�͂R�{�j
Public OBS_Pump             As Boolean                                  '���у|���v���v�Z�Ɏg����True
Public Beer                 As Boolean                                  '�s�藬�v�Z���l�n�C�h���v���b�g��TRUE
Public Un_Cal               As Single                                   '����̐��ʂ� [ UN_Cal ] �ȉ��̎��͕s�藬�ł͂Ȃ��s�����Ƃ���B
Public Category             As Boolean                                  'True=�s�藬�Ōv�Z���ꂽ�Ƃ�  False=�s�����Ōv�Z���ꂽ�Ƃ�
Public Error_Message()      As String                                   '�I���N���ɏo�͂���G���[���b�Z�[�W
Public Error_Message_n      As Long                                     '�I���N���ɏo�͂���G���[���b�Z�[�W�̐�
Public Error_Cal_Type       As String                                   'Error���o���̉J�ʃ^�C�v

'******************************************************************************
'�T�u���[�`���FCal_Initial_flow_profile(irc As Boolean)
'�����T�v�F
'�s�����v�Z�ɂ��s�藬�v�Z�p�̏������ʌ`�̌v�Z���s���B
'******************************************************************************
Sub Cal_Initial_flow_profile(irc As Boolean)
    Dim i   As Integer
    Dim i1  As Integer
    Dim nf  As Integer
    Dim buf As String
    nf = FreeFile
    Open Wpath & "\Non_Flow.log" For Output As #nf
    Print #nf, " ���E�����@�v�n�여��= " & Format(Q_kuji, "###0.00") & _
               "   �t������= " & Format(Q_Haru, "###0.00") & _
               "   �����[����= " & Format(H_Sea, "##0.000")
    '******************************************************
    '
    '******************************************************
    QU = Q_kuji + Q_Haru
    H_Start = H_Sea
    Start_Sec = "S0.000"
    End_Sec = "S12.40"
    Nonuniform_Flow irc
    If irc = False Then GoTo jump1
    i1 = Start_Num
    '******************************************************
    '
    '******************************************************
    QU = Q_kuji
    H_Start = ch(End_Num)
    Start_Sec = "S12.40"
    End_Sec = "S20.00"
    Nonuniform_Flow irc
    If irc = False Then GoTo jump1
    '******************************************************
    '
    '******************************************************
    H_Start = ch(Start_Num)
    QU = Q_Haru
    Start_Sec = "G0.000"
    End_Sec = "G8.200"
    Nonuniform_Flow irc
    If irc = False Then GoTo jump1
    '******************************************************
    '
    '******************************************************
    Print #nf, "    N   �f��       H         A         Q       V       FR     FLAG"
    For i = i1 To End_Num
        buf = Format(Format(i, "####0"), "@@@@@  ") & Sec_Name(i)
        buf = buf & Format(Format(ch(i), "###0.000"), "@@@@@@@@@@")
        buf = buf & Format(Format(CA(1, i), "###0.000"), "@@@@@@@@@@")
        buf = buf & Format(Format(CQ(1, i), "###0.000"), "@@@@@@@@@@")
        buf = buf & Format(Format(CV(i), "###0.000"), "@@@@@@@@")
        buf = buf & Format(Format(FR(i), "###0.000"), "@@@@@@@@")
        buf = buf & Space(5) & CFLAG(i)
        Print #nf, buf
    Next i
    Close #nf
    '******************************************************
    '
    '******************************************************
    nf = FreeFile
    Open Wpath & "\NSK_�������ʌ`.Temp" For Output As #nf
    Print #nf, " INITIAL        H          Q"
    For i = i1 To End_Num
        buf = Space(4) & Sec_Name(i)
        buf = buf & Format(Format(ch(i), "#####0.000"), "@@@@@@@@@@")
        buf = buf & Format(Format(CQ(1, i), "######0.00"), "@@@@@@@@@@")
        Print #nf, buf
    Next i
    Close #nf
    Exit Sub
jump1:
    Log_Calc "�s�藬�v�Z�̏������ʌ`�̌v�Z�Ɏ��s�����B"
    Close #nf
    Exit Sub
End Sub

'******************************************************************************
'�T�u���[�`���FConstant_Read()
'�����T�v�F
'******************************************************************************
Sub Constant_Read()
    Dim i     As Integer
    Dim j     As Integer
    Dim nf    As Integer
    Dim buf   As String
    Dim msg   As String
    LOG_Out "IN  Constant_Read"
    On Error GoTo ERH
    For i = 1 To 5
        For j = 1 To 5
            H_Stand1(i, j) = 99#
        Next j
        For j = 1 To 3
            H_Stand2(i, j) = 99#
        Next j
    Next i
    msg = "DATA�t�H���_�Ɋ���ʂ̃t�@�C�����Ȃ�"
    nf = FreeFile
    Open App.Path & "\DATA\�����.DAT" For Input As #nf
    Line Input #nf, buf                                                 '�^�C�g���ǂݔ�΂�
    '******************************************************
    '����ʓǂݍ���
    '******************************************************
    For i = 1 To 5
        Select Case i
            Case 1
            msg = "���V��F�F"
            Case 2
            msg = "��@���F"
            Case 3
            msg = "�����O���ʁF"
            Case 4
            msg = "�v�n��F"
            Case 5
            msg = "�t�@���F"
        End Select
        Line Input #nf, buf
            H_Standi(i, 1) = CInt(Mid(buf, 1, 5))                       '����ʂ̐�
            H_Standi(i, 2) = CInt(Mid(buf, 6, 5))                       '�|���v���ʂ̐�
        For j = 1 To H_Standi(i, 1)
            Line Input #nf, buf
            H_Stand1(i, j) = CSng(Mid(buf, 1, 5))
            H_Stand1t(i, j) = Mid(buf, 11, 4)
        Next j
        For j = 1 To H_Standi(i, 2)
            Line Input #nf, buf
            H_Stand2(i, j) = CSng(Mid(buf, 1, 5))
            H_Stand2t(i, j) = Mid(buf, 11, 13)
        Next j
    Next i
    Close #nf
    LOG_Out "OUT Constant_Read Normal Exit"
    On Error GoTo 0
    Exit Sub
ERH:
    If InStr(msg, "�F") > 0 Then
        MsgBox msg & "�n�_�̊���ʓǂݍ��ݒ��ɃG���[���������܂����A����ʂ𖳌��Ƃ��܂��B" & vbCrLf & _
                     "DATA�t�H���_�̊����.DAT���C�����ĉ������B"
        Resume ERH1
    Else
        MsgBox "����ʓǂݍ��ݒ��ɃG���[���������܂����ADATA�t�H���_�̊����.DAT�������\��������܂��B" & vbCrLf & _
               "����ʂ𖳌��Ƃ��܂��B"
        Resume ERH1
    End If
ERH1:
    For i = 1 To 5
        For j = 1 To 5
            H_Stand1(i, j) = 99#
        Next j
        For j = 1 To 3
            H_Stand2(i, j) = 99#
        Next j
    Next i
    Close #nf
    LOG_Out "OUT Constant_Read ABNormal Exit"
    On Error GoTo 0
    Exit Sub
End Sub

'******************************************************************************
'�T�u���[�`���FDate_dim(d As Date, x() As Integer)
'�����T�v�F
'******************************************************************************
Sub Date_dim(d As Date, x() As Integer)
    x(1) = Year(d)
    x(2) = Month(d)
    x(3) = Day(d)
    x(4) = Hour(d)
    x(5) = Minute(d)
    x(6) = Second(d)
End Sub

'******************************************************************************
'�T�u���[�`���FC_Date(jx() As Integer) As Date
'�����T�v�F
'******************************************************************************
Function C_Date(jx() As Integer) As Date
    Dim d As String
    d = Format(jx(1), "####") & "/" & Format(jx(2), "00") & "/" & _
        Format(jx(3), "00") & " " & Format(jx(4), "00") & ":" & _
        Format(jx(5), "00")
    If IsDate(d) Then
        C_Date = CDate(d)
    Else
        MsgBox d & " �͓��t�ł͂Ȃ�"
        '******************************************************
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        '******************************************************
        'End
        '******************************************************
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        '******************************************************
    End If
End Function

'******************************************************************************
'�T�u���[�`���FFeedBack(m As Integer)
'�����T�v�F
'******************************************************************************
Sub FeedBack(m As Integer)
    Dim i      As Integer
    Dim j      As Integer
    Dim nr     As Integer
    Dim ns     As Integer
    Dim hxj(6) As Single
    Dim hxc(6) As Single
    Dim frd(6) As Single
    Dim H1     As Single
    Dim H2     As Single
    Dim hx     As Single
    Dim hav    As Single
    Dim fd     As Single
    nr = m
    ns = V_Sec_Num(nr)                                                  'V_Sec_Num(nr)�͕s�藬��̒f�ʈʒu��\��
    hxj(1) = HO(nr + 2, Now_Step - 2)                                   '+2�͂P�ɓ�����O����(�����[����)��2�ɐ􉁗����ʂ������Ă���
    hxj(2) = HO(nr + 2, Now_Step - 1)
    hxj(3) = HO(nr + 2, Now_Step - 0)
    hxj(4) = HO(nr + 2, Now_Step + 1)
    hxj(5) = HO(nr + 2, Now_Step + 2)
    hxj(6) = HO(nr + 2, Now_Step + 3)
    hxc(1) = HQ(1, ns, NT - 30)
    hxc(2) = HQ(1, ns, NT - 24)
    hxc(3) = HQ(1, ns, NT - 18)
    hxc(4) = HQ(1, ns, NT - 12)
    hxc(5) = HQ(1, ns, NT - 6)
    hxc(6) = HQ(1, ns, NT - 0)
    frd(1) = HQ(1, ns, NT - 30)
    frd(2) = HQ(1, ns, NT - 24)
    frd(3) = HQ(1, ns, NT - 18)
    frd(4) = HQ(1, ns, NT - 12)
    frd(5) = HQ(1, ns, NT - 6)
    frd(6) = HQ(1, ns, NT - 0)
    CO(nr, 1) = hxc(3)
    CO(nr, 2) = hxc(4)
    CO(nr, 3) = hxc(5)
    CO(nr, 4) = hxc(6)
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    'Print #Log_Num, V_Sec_Name(nr, 2)
    'Print #Log_Num, nr; " hxj(1)="; fmt(hxj(1)); " hxc(1)="; fmt(hxc(1)); "  FROUD="; fmt(HQ(3, ns, NT - 30))
    'Print #Log_Num, nr; " hxj(2)="; fmt(hxj(2)); " hxc(2)="; fmt(hxc(2)); "  FROUD="; fmt(HQ(3, ns, NT - 24))
    'Print #Log_Num, nr; " hxj(3)="; fmt(hxj(3)); " hxc(3)="; fmt(hxc(3)); "  FROUD="; fmt(HQ(3, ns, NT - 18))
    'Print #Log_Num, nr; " hxj(4)="; fmt(hxj(4)); " hxc(4)="; fmt(hxc(4)); "  FROUD="; fmt(HQ(3, ns, NT - 12))
    'Print #Log_Num, nr; " hxj(5)="; fmt(hxj(5)); " hxc(5)="; fmt(hxc(5)); "  FROUD="; fmt(HQ(3, ns, NT - 6))
    'Print #Log_Num, nr; " hxj(6)="; fmt(hxj(6)); " hxc(6)="; fmt(hxc(6)); "  FROUD="; fmt(HQ(3, ns, NT - 0))
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    For i = 1 To 3
        If hxj(i) < -20# Then hxj(i) = hxc(i)
    Next i
    hx = hxj(1) - hxc(1)
    Slide1(nr) = hx
    For i = 1 To 6
        hxc(i) = hxc(i) + hx
    Next i
    H1 = (hxj(2) - hxc(2)) / 6#
    H2 = (hxj(3) - hxc(3)) / 12#
    If H1 > H2 Then
        '******************************************************
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        '******************************************************
        hav = (H1 + H2) * 0.5
        'hav = h1
        '******************************************************
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        '******************************************************
    Else
        hav = H2
    End If
    '******************************************************
    'Ver0.0.0 �C���J�n 2003/07/14 00:00
    '******************************************************
    '2003/07/14 �ǉ� �����O���ʂ�3.0m(�x������)��菬����������X���C�h���킹�Ƃ���B
    'If HO(5, Now_Step) < 3# Then
    '   hav = 0#
    'End If
    '******************************************************
    'Ver0.0.0 �C���I�� 2003/07/14 00:00
    '******************************************************
    '******************************************************
    'Ver0.0.0 �C���J�n 2003/08/12 00:00
    '******************************************************
    If hav < 0# And Froude < 0.1 Then hav = 0#
    'If hav > 0.03 Then hav = 0.03 '2003/0719 �C�� ����ł������ԈႤ(2003/08/12)
    '******************************************************
    'Ver0.0.0 �C���I�� 2003/08/12
    '******************************************************
    If hav > 0.06 Then hav = 0.06
    hav = 0#
    Print #Log_Num, nr; " hxj(1)="; fmt(hxj(1)); " hxc(1)="; fmt(hxc(1)); " hxj(1)-hxc(1)="; fmt(hxj(1) - hxc(1))
    Print #Log_Num, nr; " hxj(2)="; fmt(hxj(2)); " hxc(2)="; fmt(hxc(2)); " hxj(2)-hxc(2)="; fmt(hxj(2) - hxc(2))
    Print #Log_Num, nr; " hxj(3)="; fmt(hxj(3)); " hxc(3)="; fmt(hxc(3)); " hxj(3)-hxc(3)="; fmt(hxj(3) - hxc(3))
    Print #Log_Num, nr; " hxj(4)="; fmt(hxj(4)); " hxc(4)="; fmt(hxc(4)); fmt(hxc(4) + hav * 6)
    Print #Log_Num, nr; " hxj(5)="; fmt(hxj(5)); " hxc(5)="; fmt(hxc(5)); fmt(hxc(5) + hav * 12)
    Print #Log_Num, nr; " hxj(6)="; fmt(hxj(6)); " hxc(6)="; fmt(hxc(6)); fmt(hxc(6) + hav * 18)
    Print #Log_Num, nr; "     h1="; fmt(H1); "     h2="; fmt(H2); "           hav="; fmt(hav)
    j = 0
    For i = 18 To 0 Step -1
        HQ(1, ns, NT - i) = HQ(1, ns, NT - i) + hav * j
        j = j + 1
    Next i
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    'hav = hav + 1#
    'For i = 0 To 18
    '    HQ(1, ns, NT - i) = HQ(1, ns, NT - i) + hav * (18 - i)
    'Next i
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    If hxj(3) < -10# Then                                               '�������͌v�Z�l�Ƃ��� 2004/04/26
        hx = 0#
    Else
        hx = hxj(3) - HQ(1, ns, NT - 18)
    End If
    Slide2(nr) = hx
    j = 0
    For i = NT - 18 To NT
        HQ(1, ns, i) = HQ(1, ns, i) + hx
        YHK(nr, j) = HQ(1, ns, i)                                       '���c�a�o�͗p
        Debug.Print "  nr="; nr; "  j="; j; " YHK="; YHK(nr, j)
        j = j + 1
    Next i
    CF(nr, 0) = HQ(1, ns, NT - 18)
    CF(nr, 1) = HQ(1, ns, NT - 12)
    CF(nr, 2) = HQ(1, ns, NT - 6)
    CF(nr, 3) = HQ(1, ns, NT - 0)
    Delta_H(nr) = hav * 6
End Sub

'******************************************************************************
'�T�u���[�`���FFeedBack_Slide_Only(m As Integer)
'�����T�v�F
'******************************************************************************
Sub FeedBack_Slide_Only(m As Integer)
    Dim i      As Integer
    Dim j      As Integer
    Dim nr     As Integer
    Dim ns     As Integer
    Dim hxj    As Single
    Dim hx     As Single
    nr = m
    ns = V_Sec_Num(nr)
    hxj = HO(nr + 2, Now_Step)
    hx = hxj - HQ(1, ns, NT - 18)
    If hxj < -90# Then Exit Sub
    For i = NT - 18 To NT
        HQ(1, ns, i) = HQ(1, ns, i) + hx
    Next i
End Sub

'******************************************************************************
'�T�u���[�`���FFlood_Data_Write_For_Calc()
'�����T�v�F
'******************************************************************************
Sub Flood_Data_Write_For_Calc()
    Dim i       As Integer
    Dim j       As Integer
    Dim nf      As Integer
    Dim htw     As Single
    Dim buf     As String
    Dim d       As Date
    Dim ht(500) As Single
    LOG_Out " IN Flood_Data_Write_For_Calc"
    nf = FreeFile
    Open Wpath & "����.DAT" For Output As #nf
    '******************************************************
    '�f�[�^�J�n����
    '******************************************************
    buf = ""
    For i = 1 To 6
        buf = buf & Format(str(js(i)), "@@@@@")
    Next i
    buf = buf & "      �f�[�^�J�n����"
    Print #nf, buf
    '******************************************************
    '�f�[�^�I������
    '******************************************************
    jxd = DateAdd("h", 3, jgd)
    Date_dim jxd, jx()
    buf = ""
    For i = 1 To 6
        buf = buf & Format(str(jx(i)), "@@@@@")
    Next i
    buf = buf & "      �f�[�^�I������"
    Print #nf, buf
    '******************************************************
    '�f�[�^�s�b�`
    '******************************************************
    buf = ""
    For i = 1 To 5
        Data_Pich(i) = 3600
        buf = buf & Format(str(Data_Pich(i)), "@@@@@")
    Next i
    buf = buf & "           �f�[�^�̎��ԃs�b�`(�b)"
    Print #nf, buf
    Print #nf, Space(4) & Format(JRADAR, "0") & Format(Format(Rsa_Mag, "#0.00"), "@@@@@") '0=�e�����[�^�J�� 1=���[�_�[�J��
    Close #nf
    '******************************************************
    '���щJ��+�\���~�J
    '******************************************************
    d = jsd
    nf = FreeFile
    Open Wpath & "�J.DAT" For Output As #nf
    Print #nf, RO_Title
    For i = 1 To Now_Step + Yosoku_Step
        buf = Format(d, "yyyy/mm/dd hh:nn")
        For j = 1 To Rnum
            buf = buf & Format(Format(RO(j, i), "#######0.0"), "@@@@@@@@@@")
        Next j
        Print #nf, buf
        d = DateAdd("h", 1, d)
    Next i
    Close #nf
    '******************************************************
    '�����[����(����)
    '******************************************************
    d = jsd
    nf = FreeFile
    Open Wpath & "�����[����.DAT" For Output As #nf
    Print #nf, "  DATE     TIME      Cal_H      Tide   Suiba_H"
    '******************************************************
    '�v�Z��ʂ����߂̋���̍�
    '******************************************************
    For i = 1 To All_Step
        ht(i) = HO(1, i)
    Next i
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    'If HO(3, Now_Step) <= 2.5 Then                                     '����������O���ʂ�2.5���ȉ��̂Ƃ�
    'htw = -99#
    'For j = Now_Step - 3 To 4
    '    If HO(1, j) > htw Then htw = HO(1, j)
    'Next j
    'For j = Now_Step - 3 To 4
    '    ht(j) = htw
    'Next j
    'End If
    '�� �I���
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    For i = 1 To All_Step
        buf = Format(d, "yyyy/mm/dd hh:nn")
        buf = buf & Format(Format(ht(i), "######0.00"), "@@@@@@@@@@")
        buf = buf & Format(Format(HO(1, i), "######0.00"), "@@@@@@@@@@")
        buf = buf & Format(Format(HO(3, i), "######0.00"), "@@@@@@@@@@")
        Print #nf, buf
        d = DateAdd("h", 1, d)
    Next i
    Close #nf
    '******************************************************
    '�􉁗�����
    '******************************************************
    d = jsd
    nf = FreeFile
    Open Wpath & "��.DAT" For Output As #nf
    Print #nf, "  DATE     TIME       ��"
    For i = 1 To All_Step
        buf = Format(d, "yyyy/mm/dd hh:nn")
        buf = buf & Format(Format(HO(2, i), "######0.00"), "@@@@@@@@@@")
        Print #nf, buf
        d = DateAdd("h", 1, d)
    Next i
    Close #nf
    LOG_Out " OUT Flood_Data_Write_For_Calc"
End Sub

'******************************************************************************
'�T�u���[�`���FFlood_Data_Write_For_Calc1()
'�����T�v�F
'******************************************************************************
Sub Flood_Data_Write_For_Calc1()
    Dim i       As Long
    Dim j       As Long
    Dim m       As Long
    Dim fi      As Long
    Dim fo      As Long
    Dim buf     As String
    Dim Titl    As String
    Dim d       As Date
    Dim dd(6)   As Integer
    Dim x(500)  As String
    Dim sx(500) As String
    Dim nf
    Dim ht(500) As Single
    Dim htw     As Single
    Dim Steps   As Long
    Dim ht_max  As Single
    Dim iht_max As Long
    Dim BeforeTime As Long
    Dim Start_h As Single
    Dim d_nsk   As Date                                                 '�m�r�j�p�̊J�n����
    Dim s_nsk   As Integer                                              '�m�r�j�p�̃f�[�^�J�n�X�e�b�v
    LOG_Out " IN Flood_Data_Write_For_Calc1"
    Start_h = MAIN.Text2                                                 '�s�藬�v�Z���X���[�Y�ɍs���׏o�����ʂ������Ƃ��납��v�Z���������ׂ̐ݒ�l
    For i = 1 To All_Step
        ht(i) = HO(1, i)
    Next i
    '******************************************************
    '�v�Z�J�n�X�e�b�v��T���A���ʂ��w�肳�ꂽ����(Main.Text2)�ȏ��T���B
    '******************************************************
    ht_max = -999
    For j = 1 To All_Step - 5
        If ht(j) >= Start_h Then
            m = j
            GoTo jump1
        End If
        If ht_max > ht(j) Then
            ht_max = ht(j)
            iht_max = j
        End If
    Next j
    LOG_Out "����Ȃ��Ƃ͂����Ă͂Ȃ�Ȃ�"
    m = iht_max
jump1:
    Steps = m
    d_nsk = DateAdd("h", Steps - 1, jsd)                                '�s�藬�v�Z�J�n����
    '******************************************************
    '�����[����(����)
    '******************************************************
    LOG_Out "NSK_�����[����.DAT  For Output ---OPEN"
    fo = FreeFile
    Open Wpath & "NSK_�����[����.DAT" For Output As #fo
    Print #fo, "  DATE     TIME      Cal_H      Tide   Suiba_H"
    d = d_nsk
    For i = Steps To All_Step
        buf = Format(d, "yyyy/mm/dd hh:nn")
        buf = buf & Format(Format(ht(i), "######0.00"), "@@@@@@@@@@")
        buf = buf & Format(Format(HO(1, i), "######0.00"), "@@@@@@@@@@")
        buf = buf & Format(Format(HO(3, i), "######0.00"), "@@@@@@@@@@")
        Print #fo, buf
        d = DateAdd("h", 1, d)
    Next i
    Close #fo
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    'H_Sea = HO(1, Now_Step - 3)                                        '�s�����v�Z�p�����[����+++++++++++++++++++++++++++�������ʌ`���E����
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    H_Sea = HO(1, Steps)                                                '�s�����v�Z�p�����[����+++++++++++++++++++++++++++�������ʌ`���E����
    LOG_Out "NSK_�����[����.DAT  For Output ---CLOSE"
    '******************************************************
    '�v�Z�����ݒ�
    '******************************************************
    Date_dim d_nsk, dd()
    LOG_Out "NSK_����.DAT For Output ---OPEN"
    fo = FreeFile
    Open Wpath & "NSK_����.DAT" For Output As #fo
    '******************************************************
    '�f�[�^�J�n����
    '******************************************************
    buf = ""
    For i = 1 To 6
        buf = buf & Format(str(dd(i)), "@@@@@")
    Next i
    buf = buf & "      �f�[�^�J�n����"
    NSK_jsd = d_nsk
    Print #fo, buf
    '******************************************************
    '�f�[�^�I������
    '******************************************************
    jxd = DateAdd("h", 3, jgd)
    Date_dim jxd, jx()
    buf = ""
    For i = 1 To 6
        buf = buf & Format(str(jx(i)), "@@@@@")
    Next i
    buf = buf & "      �f�[�^�I������"
    Print #fo, buf
    '******************************************************
    '�f�[�^�s�b�`
    '******************************************************
    buf = ""
    For i = 1 To 5
        buf = buf & " 3600"
    Next i
    buf = buf & "           �f�[�^�̎��ԃs�b�`(�b)"
    Print #fo, buf
    Close #fo
    LOG_Out "NSK_����.DAT For Output ---CLOSE"
    '******************************************************
    'SHINK10.U07
    '******************************************************
    LOG_Out "SHINK10.U07  For Input ---OPEN"
    fi = FreeFile
    Open Wpath & "SHINK10.U07" For Input As #fi
    Line Input #fi, Titl
    i = 0
    Do Until EOF(fi)
        i = i + 1
        Line Input #fi, x(i)
    Loop
    Close #fi
    LOG_Out "SHINK10.U07  For Input ---CLOSE"

    LOG_Out "NSK_SHINK10.U07  For Output ---OPEN"
    m = DateDiff("h", d_nsk, jxd) + 1
    Mid(Titl, 1, 10) = " 3600" & Format(str(m), "@@@@@")
    fo = FreeFile
    Open Wpath & "NSK_SHINK10.U07" For Output As #fo
    Print #fo, Titl
    For j = Steps To i
        Print #fo, x(j)
    Next j
    Close #fo
    LOG_Out "NSK_SHINK10.U07  For Output ---CLOSE"
    '******************************************************
    '�􉁗�����
    '******************************************************
    LOG_Out "NSK_��.DAT  For Output ---OPEN"
    d = d_nsk
    fo = FreeFile
    Open Wpath & "NSK_��.DAT" For Output As #fo
    Print #fo, "  DATE     TIME       ��"
    For i = Steps To All_Step
        buf = Format(d, "yyyy/mm/dd hh:nn")
        buf = buf & Format(Format(HO(2, i), "######0.00"), "@@@@@@@@@@")
        Print #fo, buf
        d = DateAdd("h", 1, d)
    Next i
    Close #fo
    LOG_Out "NSK_��.DAT  For Output ---CLOSE"
    '******************************************************
    'KUJINO.DAT
    '******************************************************
    LOG_Out "KUJINO.DAT For Input ---OPEN"
    fi = FreeFile
    Open Wpath & "KUJINO.DAT" For Input As #fi
    Line Input #fi, Titl
    i = 0
    Do Until EOF(fi)
        i = i + 1
        Line Input #fi, x(i)
    Loop
    Close #fi
    LOG_Out "KUJINO.DAT For Input ---CLOSE"
    LOG_Out "NSK_KUJINO.DAT  For Output ---OPEN"
    fo = FreeFile
    Open Wpath & "NSK_KUJINO.DAT" For Output As #fo
    Print #fo, Titl
    For j = Steps To i
        Print #fo, x(j)
    Next j
    Close #fo
    Q_kuji = CSng(Mid(x(Steps), 7, 10))                                 '�s�����v�Z�p �v�n�여����++++++++++�������ʌ`���E����
    LOG_Out "NSK_KUJINO.DAT  For Output ---CLOSE"
    '******************************************************
    'HARUHI.DAT
    '******************************************************
    LOG_Out "HARUHI.DAT  For Input ---OPEN"
    fi = FreeFile
    Open Wpath & "HARUHI.DAT" For Input As #fi
    Line Input #fi, Titl
    i = 0
    Do Until EOF(fi)
        i = i + 1
        Line Input #fi, x(i)
    Loop
    Close #fi
    LOG_Out "HARUHI.DAT  For Input ---CLOSE"

    LOG_Out "NSK_HARUHI.DAT  For Output ---OPEN"
    fo = FreeFile
    Open Wpath & "NSK_HARUHI.DAT" For Output As #fo
    Print #fo, Titl
    For j = Steps To i
        Print #fo, x(j)
    Next j
    Close #fo
    LOG_Out "NSK_HARUHI.DAT  For Output ---CLOSE"
    Q_Haru = CSng(Mid(x(Steps), 7, 10))                                 '�s�����v�Z�p �t��������+++++++++++�������ʌ`���E����
    LOG_Out " OUT Flood_Data_Write_For_Calc1"
End Sub

'******************************************************************************
'�T�u���[�`���Ffmt(c As Variant) As String
'�����T�v�F
'******************************************************************************
Function fmt(c As Variant) As String
   fmt = Format(Format(c, "###0.0000"), "@@@@@@@@@")
End Function

'******************************************************************************
'�T�u���[�`���FFroude_Check(irc As Boolean)
'�����T�v�F
'******************************************************************************
Sub Froude_Check(irc As Boolean)
    Dim i      As Integer
    Dim j      As Integer
    Dim nr     As Integer
    Dim ns     As Integer
    Dim fds    As Single
    Dim fd     As Single
    ReDim frd(6, V_Sec_Cnt) As Single
    irc = True
    fds = 0#
    For nr = 1 To V_Sec_Cnt
        ns = V_Sec_Num(nr)                                              'V_Sec_Num(nr)�͕s�藬��̒f�ʈʒu��\��
        frd(1, nr) = HQ(3, ns, NT - 30)
        frd(2, nr) = HQ(3, ns, NT - 24)
        frd(3, nr) = HQ(3, ns, NT - 18)
        frd(4, nr) = HQ(3, ns, NT - 12)
        frd(5, nr) = HQ(3, ns, NT - 6)
        frd(6, nr) = HQ(3, ns, NT - 0)
        CO(nr, 1) = HQ(1, ns, NT - 12)
        CO(nr, 2) = HQ(1, ns, NT - 6)
        CO(nr, 3) = HQ(1, ns, NT - 0)
        Print #Log_Num, V_Sec_Name(nr, 2)
        Print #Log_Num, nr; "  FROUD="; fmt(frd(1, nr))
        Print #Log_Num, nr; "  FROUD="; fmt(frd(2, nr))
        Print #Log_Num, nr; "  FROUD="; fmt(frd(3, nr))
        Print #Log_Num, nr; "  FROUD="; fmt(frd(4, nr))
        Print #Log_Num, nr; "  FROUD="; fmt(frd(5, nr))
        Print #Log_Num, nr; "  FROUD="; fmt(frd(6, nr))
        fd = 0#
        For i = 1 To 6
            fd = fd + frd(i, nr)
        Next i
        fds = fds + fd / 6
        Print #Log_Num, nr; " fd="; fmt(fd / 6)
    Next nr
    Froude = fds / V_Sec_Cnt
    Print #Log_Num, "  Average Froude="; fmt(Froude)
End Sub

'******************************************************************************
'�T�u���[�`���FHydro_Graph(OBS_Point As Integer)
'�����T�v�F
'******************************************************************************
Sub Hydro_Graph(OBS_Point As Integer)
    '******************************************************
    '�n�_�v���b�g
    '******************************************************
    Dim dl     As Single
    Dim dp     As Single
    Dim xs     As Single
    Dim xe     As Single
    Dim ys     As Single
    Dim ye     As Single
    Dim xl     As Single
    Dim yl     As Single
    Dim Hp     As Single
    Dim xw     As Single
    Dim yw     As Single
    Dim fj     As Single
    Dim amu    As Single
    Dim amd    As Single
    Dim amp    As Single
    Dim sc     As Single
    Dim ysm    As Single
    Dim psm    As Single
    Dim msize  As Single
    Dim hqyl   As Single
    Dim ps     As Single
    Dim x      As Single
    Dim y      As Single
    Dim i      As Integer
    Dim j      As Integer
    Dim nday   As Integer
    Dim nbun   As Integer
    Dim niti   As String
    Dim J1     As Integer
    Dim j2     As Integer
    Dim j3     As Integer
    Dim n      As Integer
    Dim mn     As Integer
    Dim xt     As Single
    Dim nr     As Integer
    Dim dw     As Date
    Dim KanG   As Integer
    Dim moji   As String
    Dim w      As Single
    Dim ns     As Integer
    Dim Mp     As Single
    Dim na     As Integer
    Dim rs     As Single
    Dim t1     As String
    Dim t2     As String
    Dim t3     As String
    Dim T4     As String
    Dim Kijun_Name As String
    xl = 215: yl = 155: hqyl = 115
    xs = 35: ys = 30: xe = xs + xl: ye = ys + yl
    KanG = OBS_Point
    Select Case OBS_Point
        Case 1                                                          '���V��F
            Kijun_Name = "���V��F"
            na = 1
        Case 2                                                          '�厡
            Kijun_Name = "��@���@"
            na = 2
        Case 3                                                          '����O����
            Kijun_Name = "����O����"
            na = 3
        Case 4                                                          '�v�n��
            Kijun_Name = "�v�n��"
            na = 4
        Case 5                                                          '�t��
            Kijun_Name = "�t�@���@"
            na = 5
    End Select
    nr = KanG
    nday = DateDiff("d", jsd, jgd) + 1
    If nday < 3 Then nday = 3
    VS_Box xs, ys, xe, ye, 0, 0.4, 15, 1
    If isRAIN = "02" Then
        VS_symbol xs, ys - 1#, 12#, "�e�q�h�b�r�J�ʎg�p", 3
    Else
        VS_symbol xs, ys - 1#, 12#, "�C�ے��J�ʎg�p", 3
    End If
    dw = jsd
    dl = xl / nday
    '******************************************************
    '�����ڐ���
    '******************************************************
    Hp = dl / 24
    Mp = Hp / 60
    For i = 1 To nday
        x = xs + dl * (i - 1)
        If i <> nday Then
            J1 = 0
        Else
            J1 = 1
        End If
        For j = 0 To 23 + J1
            xw = x + Hp * j
            If (j Mod 6) = 0 Then
                VS_Line xw, ye + 1.5, xw, ys, 0, 0
                If j = 0 Then
                    VS_Line xw, ys, xw, ye, 0, 0
                Else
                    VS_Line xw, ys, xw, ye, 8, 0
                End If
                fj = j Mod 24
                VS_symbol xw, ye + 2#, 8.5, Cvt_2byte(Trim(str(j Mod 24))), 4
            Else
                VS_Line xw, ye + 1#, xw, ye, 0, 0
            End If
            If j = 12 Then
                If i > 1 Then
                    dw = DateAdd("d", 1, dw)
                    j2 = Month(dw)
                    j3 = Day(dw)
                Else
                    j2 = Month(dw)
                    j3 = Day(dw)
                End If
                niti = Cvt_2byte(Format(j2, "##") + "/" + Format(j3, "##"))
                VS_symbol xw, ye + 8#, 12.5, niti, 4
            End If
        Next j
    Next i
    niti = "������ " + Cvt_2byte(Format(jg(2), "##")) & "��" & _
                       Cvt_2byte(Format(jg(3), "##")) & "��" & _
                       Cvt_2byte(Format(jg(4), "#0")) & "��" & _
                       Cvt_2byte(Format(jg(5), "#0")) & "��"
    VS_symbol xe, ys - 1#, 12#, niti, 9
    VS_symbol xs + xl * 0.5, ys - 2.5, 14, Kijun_Name & "�ϑ���", 6
    xt = 20#
    '******************************************************
    '�J�ʖڐ���
    '******************************************************
    rs = 0.5
    For y = 20 To 60 Step 20
        yw = ys + y * 0.5
        VS_Line xs - 1.5, yw, xe, yw, 8, 0
        VS_symbol xs - 2#, yw, 9#, Cvt_2byte(str(y)), 8
    Next y
    VS_symbol xs - xt, ys + 12#, 11.5, "�J", 5
    VS_symbol xs - xt, ys + 20#, 11.5, "��", 5
    moji = Cvt_2byte("(mm)")
    VS_symbol xs - xt, ys + 25#, 9#, moji, 5
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    'If Graph.Option2(0) Then
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
        mn = 1
        VS_symbol xs - xt, ye - 80#, 11.5, "��", 5
        VS_symbol xs - xt, ye - 50#, 11.5, "��", 5
        moji = Cvt_2byte("( m )")
        VS_symbol xs - xt, ye - 40#, 10#, moji, 5
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    'End If
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    'If Graph.Option2(1) Then
    '    mn = 2
    '    Call tv_symbol(xs - xt, ys + 80#, 11.5, "��", 5)
    '    Call tv_symbol(xs - xt, ys + 50#, 11.5, "��", 5)
    '    Call tv_symbol(xs - xt, ys + 40#, 10#, "(m3/s)", 5)
    'End If
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    '******************************************************
    '�ڐ���
    '******************************************************
     If mn = 1 Then
        amd = H_Scale(KanG, 1)
        amu = H_Scale(KanG, 2)
        amp = H_Scale(KanG, 3)
    End If
    sc = hqyl / (amu - amd)
    For fj = amd To amu + 0.01 Step amp
        y = ye - (fj - amd) * sc
        VS_Line xs - 1.5, y, xe, y, 8, 0
        If mn = 1 Then
            VS_number xs - 2#, y, 9#, fj, 1, 8                          '����
        Else
            VS_number xs - 2#, y, 9#, fj, -1, 8                         '����
        End If
    Next fj
    If mn = 1 Then
        '******************************************************
        '�����
        '******************************************************
        If H_Stand1(nr, 4) < amu And H_Stand1(nr, 4) > amd Then
            y = ye - (H_Stand1(nr, 4) - amd) * sc
            VS_Line xs, y, xe, y, 2, 0.5
            VS_symbol xe + 1#, y, 9#, H_Stand1t(nr, 4), 3
            VS_number xe + 1#, y, 8#, H_Stand1(nr, 4), 2, 1
        End If
        If H_Stand1(nr, 3) < amu And H_Stand1(nr, 3) > amd Then
            y = ye - (H_Stand1(nr, 3) - amd) * sc
            VS_Line xs, y, xe, y, 5, 0.5
            VS_symbol xe + 1#, y, 9#, H_Stand1t(nr, 3), 3
            VS_number xe + 1#, y, 8#, H_Stand1(nr, 3), 2, 1
        End If
        If H_Stand1(nr, 2) < amu And H_Stand1(nr, 2) > amd Then
            y = ye - (H_Stand1(nr, 2) - amd) * sc
            VS_Line xs, y, xe, y, 5, 0.5
            VS_symbol xe + 1#, y, 9#, H_Stand1t(nr, 2), 3
            VS_number xe + 1#, y, 8#, H_Stand1(nr, 2), 2, 1
        End If
        If H_Stand1(nr, 1) < amu And H_Stand1(nr, 1) > amd Then
            y = ye - (H_Stand1(nr, 1) - amd) * sc
            VS_Line xs, y, xe, y, 12, 0.5
            VS_symbol xe + 1#, y, 9#, H_Stand1t(nr, 1), 3
            VS_number xe + 1#, y, 8#, H_Stand1(nr, 1), 2, 1
        End If
        '******************************************************
        '�|���v����
        '******************************************************
        If H_Stand2(nr, 1) < amu And H_Stand2(nr, 1) > amd Then
            y = ye - (H_Stand2(nr, 1) - amd) * sc
            VS_Line xs, y, xe, y, 12, 0.5
            VS_symbol xs + 1#, y, 9#, H_Stand2t(nr, 1), 3
            VS_number xs + 1#, y, 8#, H_Stand2(nr, 1), 2, 1
        End If
        If H_Stand2(nr, 2) < amu And H_Stand2(nr, 2) > amd Then
            y = ye - (H_Stand2(nr, 2) - amd) * sc
            VS_Line xs, y, xe, y, 12, 0.5
            VS_symbol xs + 17#, y, 9#, H_Stand2t(nr, 2), 3
            VS_number xs + 17#, y, 8#, H_Stand2(nr, 2), 2, 1
        End If
        If H_Stand2(nr, 3) < amu And H_Stand2(nr, 3) > amd Then
            y = ye - (H_Stand2(nr, 3) - amd) * sc
            VS_Line xs, y, xe, y, 12, 0.5
            VS_symbol xs + 1#, y, 9#, H_Stand2t(nr, 3), 3
            VS_number xs + 1#, y, 8#, H_Stand2(nr, 3), 2, 1
        End If
    End If
    '******************************************************
    '�J�ʃv���b�g
    '******************************************************
    xw = xs + js(4) * Hp
    For i = 1 To Now_Step                                               '�v�Z�J�n���猻�����܂�
        If RO(na, i) > 1# Then
            w = RO(na, i) * rs                                          '���搔�{�P����e��n�_�㗬�敽�ωJ��
            '******************************************************
            'Ver0.0.0 �C���J�n 1900/01/01 00:00
            '******************************************************
            'Debug.Print "  RO="; RO(na, i)
            '******************************************************
            'Ver0.0.0 �C���I�� 1900/01/01 00:00
            '******************************************************
            ps = xw + (i - 1) * Hp
            VS_Box ps - Hp, ys, ps, ys + w, QBColor(9), 0, QBColor(9), 0
        End If
    Next i
    For i = Now_Step + 1 To All_Step                                    '�������{�P����\�����ԃX�e�b�v�܂�
        If RO(na, i) > 1# Then
            w = RO(na, i) * rs                                          '���搔�{�P����e��n�_�㗬�敽�ωJ��
            ps = xw + (i - 1) * Hp
            '******************************************************
            'Ver0.0.0 �C���J�n 1900/01/01 00:00
            '******************************************************
            'Debug.Print "  RO="; RO(na, i)
            '******************************************************
            'Ver0.0.0 �C���I�� 1900/01/01 00:00
            '******************************************************
            VS_Box ps - Hp, ys, ps, ys + w, QBColor(14), 0, QBColor(14), 0
        End If
    Next i
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    'iam.DrawWidth = 2
    'If mn = 1 Then
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
        '******************************************************
        '���ѐ��ʃv���b�g
        '******************************************************
        xw = xs + js(4) * Hp + js(5) * Mp
        If OBS1 Then
            n = All_Step
        Else
            n = Now_Step
        End If
        For i = 1 To n
            If HO(nr + 2, i) <> aksk Then
                w = ye - (HO(nr + 2, i) - amd) * sc
                ps = xw + (i - 1) * Hp
                VS_Circle ps, w, 0.7, 0, 0, 0, 0
            End If
        Next i
        '******************************************************
        '�������܂Ōv�Z���ʃv���b�g
        '******************************************************
        If CAL1 = 1 Then
            ns = V_Sec_Num(nr)
            xw = xs + (js(4) + Now_Step - 4) * Hp + js(5) * Mp          '�s�藬�v�Z�������f�[�^�J�n����p
            ysm = ye - (HQ(1, ns, 1) - amd) * sc
            psm = xw
            For i = 1 To NT - 17
                amu = ye - (HQ(1, ns, i) - amd) * sc
                ps = xw + i * Mp
                VS_Line psm, ysm, ps, amu, 4, 0.5
                psm = ps
                ysm = amu
            Next i
        End If
        '******************************************************
        '�\���v�Z���ʃv���b�g
        '******************************************************
        If Category Then                                                '�s�藬��
            ns = V_Sec_Num(nr)
            xw = xs + (DateDiff("h", jsd, jgd) + js(4)) * Hp + js(5) * Mp '�s�藬�v�Z�������|�S���ԗp
            If HO(nr + 2, Now_Step) > -90# Then
                ysm = ye - (HO(nr + 2, Now_Step) - amd) * sc
            Else
                ysm = ye - (HQ(1, ns, NT - 18) - amd) * sc
            End If
            psm = xw
            j = 0
            For i = NT - 17 To NT
                j = j + 1
                amu = ye - (HQ(1, ns, i) - amd) * sc
                ps = xw + j * Mp * 10
                VS_Line psm, ysm, ps, amu, 2, 0.5
                psm = ps
                ysm = amu
            Next i
        '******************************************************
        '�s������
        '******************************************************
        Else
            ysm = ye - (Nonuni_H(nr, 0) - amd) * sc
            xw = xs + (DateDiff("h", jsd, jgd) + js(4)) * Hp + js(5) * Mp
            psm = xw
            For i = 1 To 3
                amu = ye - (Nonuni_H(nr, i) - amd) * sc
                ps = xw + i * Hp
                VS_Line psm, ysm, ps, amu, 2, 0.5
                psm = ps
                ysm = amu
            Next i
        End If
        '******************************************************
        '��
        '******************************************************
        If Beer Then
            xw = xs + js(4) * Hp + js(5) * Mp                           '�s�藬�v�Z�������|�S���ԗp
            ysm = ye - (CO(nr, 1) - amd) * sc
            psm = xw
            For i = 2 To 4
                amu = ye - (CO(nr, i) - amd) * sc
                ps = xw + Hp * (i - 1)
                VS_Line psm, ysm, ps, amu, 4, 0.4
                psm = ps
                ysm = amu
            Next i
        End If
        VS_ShowPage 3
        '******************************************************
        '�\�������v���b�g
        '******************************************************
        If History Then
            Dim yp0 As Single, YP1 As Single, YP2 As Single, YP3 As Single
            Dim tp As Date
            Dim XB As Single
            MDB_����_Read
            XB = xs + js(4) * Hp
            For i = 1 To Now_Step
                If H_Pred(i, OBS_Point, 1) > -80# Then
                    tp = T_Pred(i)
                    xw = XB + DateDiff("h", jsd, tp) * Hp + Minute(tp) * Mp
                    yp0 = ye - (H_Pred(i, OBS_Point, 1) - amd) * sc
                    YP1 = ye - (H_Pred(i, OBS_Point, 2) - amd) * sc
                    YP2 = ye - (H_Pred(i, OBS_Point, 3) - amd) * sc
                    YP3 = ye - (H_Pred(i, OBS_Point, 4) - amd) * sc
                    VS_Line xw, yp0, xw + Hp, YP1, 1, 0.2
                    VS_Line xw + Hp, YP1, xw + Hp * 2, YP2, 1, 0.2
                    VS_Line xw + Hp * 2, YP2, xw + Hp * 3, YP3, 1, 0.2
                End If
            Next i
        End If
    'End If
End Sub

'******************************************************************************
'�T�u���[�`���FInitial_Constant()
'�����T�v�F
'******************************************************************************
Sub Initial_Constant()
    '******************************************************
    '���V��F
    '******************************************************
    H_Scale(1, 1) = -1
    H_Scale(1, 2) = 5
    H_Scale(1, 3) = 0.5
    '******************************************************
    '�厡
    '******************************************************
    H_Scale(2, 1) = 0
    H_Scale(2, 2) = 5
    H_Scale(2, 3) = 0.5
    '******************************************************
    '����O
    '******************************************************
    H_Scale(3, 1) = 0
    H_Scale(3, 2) = 7
    H_Scale(3, 3) = 0.5
    '******************************************************
    '�v�n��
    '******************************************************
    H_Scale(4, 1) = 0
    H_Scale(4, 2) = 8
    H_Scale(4, 3) = 0.5
    '******************************************************
    '�t��
    '******************************************************
    H_Scale(5, 1) = 1
    H_Scale(5, 2) = 6
    H_Scale(5, 3) = 0.5
End Sub

'******************************************************************************
'�T�u���[�`���FInitial_Data()
'�����T�v�F
'******************************************************************************
Sub Initial_Data()
    '******************************************************
    '�J�ʊϑ�����
    '******************************************************
    Name_R(1) = "���@�R"
    Name_R(2) = "��m�{��"
    Name_R(3) = "��m�{�C"
    Name_R(4) = "���@�q"
    Name_R(5) = "���É���"
    Name_R(6) = "�t����"
    Name_R(7) = "��@��"
    Name_R(8) = "���É���"
    Name_R(9) = "�I�@�]"
    Name_R(10) = "���@�{"
    '******************************************************
    '���ʊϑ�����
    '******************************************************
    Name_H(1) = "������O"
    Name_H(2) = "��@��"
    Name_H(3) = "���V��F"
    Name_H(4) = "��@��"
    Name_H(5) = "�����O"
    Name_H(6) = "�v�n��"
    Name_H(7) = "�t�@��"
End Sub

'******************************************************************************
'�T�u���[�`���FPDF_Check()
'�����T�v�F
'******************************************************************************
Sub PDF_Check()
    Dim i        As Long
    Dim ns       As Long
    Dim F        As String
    Dim PDF_Out  As Boolean
    Dim T        As String
    PDF_Out = False
    '******************************************************
    '�����O���ʂ̎��т��`�F�b�N
    '******************************************************
    For i = 1 To Now_Step
        If HO(5, i) >= 2# Then
            PDF_Out = True
            Exit For
        End If
    Next i
    If PDF_Out Then
        GoTo PDF_Put
    End If
    '******************************************************
    '�����O���ʂ̗\�����`�F�b�N
    '******************************************************
    For i = NT - 17 To NT
        If HQ(1, 3, i) >= 2# Then                                       '���h�c�ҋ@���ʂ𒴂��鐅�ʂ���������PDF���o��
            PDF_Out = True
            Exit For
        End If
    Next i
    If PDF_Out = False Then
        Exit Sub
    End If
PDF_Put:
    T = Format(jgd, "yyyymmddhhnn")
    If isRAIN = 1 Then
        T = "�C�ے�" & T
    Else
        T = "FRICS" & T
    End If
    Graph3.VSPDF1.ConvertDocument Graph3.VSP, App.Path & "\DATA\PDF\" & T & ".pdf"
End Sub

'******************************************************************************
'�T�u���[�`���FPump_Full()
'�����T�v�F
'******************************************************************************
Sub Pump_Full()
    Dim i    As Integer
    Dim n1   As Integer
    Dim n2   As Integer
    Dim buf  As String
    n1 = FreeFile
    Open App.Path & "\DATA\PFULL.DAT" For Input As #n1
    n2 = FreeFile
    Open Wpath & "Pump.dat" For Output As #n2
    Do Until EOF(n1)
        Line Input #n1, buf
        Print #n2, buf
    Loop
    Close #n1
    Close #n2
End Sub

'******************************************************************************
'�T�u���[�`���FSection_Read()
'�����T�v�F
'******************************************************************************
Sub Section_Read()
    Dim i    As Integer
    Dim j    As Integer
    Dim i1   As Integer
    Dim i2   As Integer
    Dim nf   As Integer
    Dim NDx  As Integer
    Dim DXS  As Single
    Dim buf  As String
    nf = FreeFile
    Open App.Path & "\data\Section.dat" For Input As #nf
    Input #nf, buf
    NDx = CInt(Mid(buf, 1, 5))
    CAL1 = CInt(Mid(buf, 6, 5))                                             'NDx=�f�ʐ�  CAL1=�v�Z�l�v���b�g�̗L=1��=0
    ReDim DX(NDx), NDN(NDx)
    V_Sec_Cnt = 0
    For i = 1 To NDx
        Input #nf, buf
        NDN(i) = Mid(buf, 1, 6)
        DX(i) = CSng(Mid(buf, 10, 6))
        ZS(i) = CSng(Mid(buf, 21, 10))
        MDX = MDX + DX(i)
        If Mid(buf, 32, 5) <> "" Then
            V_Sec_Cnt = V_Sec_Cnt + 1
            V_Sec_Num(V_Sec_Cnt) = i
            V_Sec_Name(V_Sec_Cnt, 1) = NDN(i)
            V_Sec_Name(V_Sec_Cnt, 2) = Mid(buf, 32, 5)
        End If
    Next i
    MDX = 0
    For i = 1 To 53
        MDX = MDX + DX(i)
    Next i
    ReDim sdx(V_Sec_Cnt)
    i1 = 1
    DXS = 0#
    For j = 1 To V_Sec_Cnt - 1
        i2 = V_Sec_Num(j)
        For i = i1 To i2
            DXS = DXS + DX(i)
            sdx(j) = DXS
        Next i
        i1 = i2 + 1
    Next j
    Close #nf
End Sub

'******************************************************************************
'�T�u���[�`���FShort_Break(S As Long)
'�����T�v�F
'******************************************************************************
Public Sub Short_Break(S As Long)
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    'Sleep s * 1000   '�V�X�e�����~�߂�̂Ŕ������Ȃ��Ȃ邩�炾��
    '                  �������A�b�o�t���g��Ȃ��Ȃ�B
    '                  ���̕��@�͔������邪�b�o�t���P�O�O���g������
    '                  �Ȃ�B
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    Dim i   As Date
    Dim j   As Date
    Dim k   As Long
    i = Now
    Do
        j = Now
        k = DateDiff("s", i, j)
        If k >= S Then
            Exit Do
        End If
        DoEvents
    Loop
    Exit Sub
End Sub

'******************************************************************************
'�T�u���[�`���FFlood_Data_Read()
'�����T�v�F
'******************************************************************************
Sub Flood_Data_Read()
    Dim a
    Dim b
    Dim i      As Integer
    Dim j      As Integer
    Dim nf     As Integer
    Dim NG     As Integer
    Dim buf    As String
    LOG_Out " IN Flood_Data_Read"
    nf = FreeFile
    Open Input_file For Input As #nf
    '******************************************************
    '�^���f�[�^�^�C�g��
    '******************************************************
    Line Input #nf, Flood_Name
    '******************************************************
    '�f�[�^�J�n����
    '******************************************************
    Line Input #nf, buf
    js(1) = CInt(Mid(buf, 1, 5))
    js(2) = CInt(Mid(buf, 6, 5))
    js(3) = CInt(Mid(buf, 11, 5))
    js(4) = CInt(Mid(buf, 16, 5))
    js(5) = CInt(Mid(buf, 21, 5))
    js(6) = 0
    jsd = C_Date(js())
    '******************************************************
    '�f�[�^�I������
    '******************************************************
    Line Input #nf, buf
    jg(1) = CInt(Mid(buf, 1, 5))
    jg(2) = CInt(Mid(buf, 6, 5))
    jg(3) = CInt(Mid(buf, 11, 5))
    jg(4) = CInt(Mid(buf, 16, 5))
    jg(5) = CInt(Mid(buf, 21, 5))
    jg(6) = 0
    jgd = C_Date(jg())
    '******************************************************
    '�f�[�^�s�b�`
    '******************************************************
    Line Input #nf, buf
    Data_Pich(1) = CSng(Mid(buf, 1, 5))
    Data_Pich(2) = CSng(Mid(buf, 6, 5))
    Data_Pich(3) = CSng(Mid(buf, 11, 5))
    Data_Pich(4) = CSng(Mid(buf, 16, 5))
    Data_Pich(5) = CSng(Mid(buf, 21, 5))
    Data_Steps = DateDiff("h", jsd, jgd) + 1
    Now_Step = Data_Steps
    All_Step = Now_Step + Yosoku_Step
    '******************************************************
    '���[�_�[�J�ʂ̗L��
    '******************************************************
    Line Input #nf, buf
    IRADAR = CInt(Mid(buf, 1, 5))
    If IRADAR = 1 Then
        Radar_File = Trim(Mid(buf, 6, 30))                              '���[�_�[�J�ʃt�@�C����
        MAIN.Check2.Enabled = True
    Else
        Radar_File = ""
        MAIN.Check2.Enabled = False
    End If
    '******************************************************
    '���щJ��
    '******************************************************
    Line Input #nf, RO_Title
    For i = 1 To Data_Steps
        Line Input #nf, buf
        For j = 1 To Rnum
            RO(j, i) = CSng(Mid(buf, (j - 1) * 10 + 17, 10))
        Next j
    Next i
    '******************************************************
    '���ѐ��ʁA���ʁA��
    '******************************************************
    Line Input #nf, HO_Title
    For i = 1 To Data_Steps
        Line Input #nf, buf
        For j = 1 To Hnum
            HO(j, i) = CSng(Mid(buf, (j - 1) * 10 + 17, 10))
        Next j
    Next i
    '******************************************************
    '�|���v�f�[�^
    '******************************************************
    If EOF(nf) Then
        Pump_Full
    Else
        Input #nf, buf
        a = InStr(buf, "�|���v")
        If a > 0 Then
            If OBS_Pump Then
                '******************************************************
                'Ver0.0.0 �C���J�n 1900/01/01 00:00
                '******************************************************
                'B = MsgBox("���̃f�[�^�ɂ̓|���v���т�����܂��A�g���܂����H", vbYesNo + vbInformation)
                '******************************************************
                'Ver0.0.0 �C���I�� 1900/01/01 00:00
                '******************************************************
                b = vbYes
            Else
                b = vbNo
            End If
            If b = vbYes Then
                NG = FreeFile
                Open Wpath & "Pump.dat" For Output As #NG
                Do
                    Line Input #nf, buf
                    If InStr(buf, "INIT") > 0 Then
                        Close #NG
                        GoTo EXT
                    End If
                    Print #NG, buf
                Loop
            Else
                Pump_Full
                Do
                    Line Input #nf, buf
                    If InStr(buf, "INIT") > 0 Then
                        GoTo EXT
                    End If
                Loop
            End If
        End If
    End If
    GoTo EXT1
EXT:
    '******************************************************
    '�������ʌ`�f�[�^
    '******************************************************
    NG = FreeFile
    Open Wpath & "�������ʌ`.dat" For Output As #NG
    Print #NG, buf
    Do Until EOF(nf)
        Line Input #nf, buf
        Print #NG, buf
    Loop
EXT1:
    Close #NG
    Close #nf
    �v�n��ƌ܏��㗬�[����
    LOG_Out " OUT Flood_Data_Read"
End Sub

'******************************************************************************
'�T�u���[�`���FInput_Yosoku(irc As Boolean)
'�����T�v�F
'�s�藬�v�Z���ʂ�ǂݍ���
'******************************************************************************
Sub Input_Yosoku(irc As Boolean)
    Dim i          As Integer
    Dim j          As Integer
    Dim k          As Integer
    Dim ns         As Integer
    Dim buf        As String
    Dim nf         As Integer
    Dim w          As Single
    Dim Froude_Max As Single
    On Error GoTo ERH1
    irc = True
    NT = 18
    ReDim HQ(3, nd, NT)                                                 '3=(1=H 2=Q 3=flood)
    ReDim MAX_H(nd)
    nf = FreeFile
    Open App.Path & "\WORK\newnskg2.u08" For Input As #nf

    Froude_Max = 0#
    For i = 1 To nd
        Line Input #nf, buf
        NDN(i) = Mid(buf, 25, 6)
        w = CSng(Mid(buf, 31, 10))                                      '����
        HQ(1, i, 1) = w
        MAX_H(i) = w
        HQ(2, i, 1) = CSng(Mid(buf, 41, 10))                            '����
        HQ(3, i, 1) = CSng(Mid(buf, 51, 10))                            '�t���[�h��
        If HQ(3, i, 1) > Froude_Max Then Froude_Max = HQ(3, i, 1)
    Next i
    Froude = Froude_Max
    j = 1
    Do Until EOF(nf)
        j = j + 1
        If j > NT Then
            NT = NT + 1
            ReDim Preserve HQ(3, nd, NT)
        End If
        For i = 1 To nd
            Line Input #nf, buf
            NDN(i) = Mid(buf, 25, 6)
            w = CSng(Mid(buf, 31, 10))                                  '����
            HQ(1, i, j) = w
            HQ(2, i, j) = CSng(Mid(buf, 41, 10))                        '����
            HQ(3, i, j) = CSng(Mid(buf, 51, 10))                        '�t���[�h��
            If w > MAX_H(i) Then MAX_H(i) = w
        Next i
    Loop
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    '��������c�a�p�̗\���l����낤�Ƃ�����
    'For i = 1 To 5                         '���̌v�Z�l�Ȃ̂ł�߂��������s�������ɂ�
    '    ns = V_Sec_Num(i)                  '�ǂ��Ȃ��Ă��邩�킩��Ȃ�
    '    k = 0
    '    For j = NT - 18 To NT
    '        YHK(i, k) = HQ(1, ns, j)
    '        k = k + 1
    '    Next j
    'Next i
    'Froude_Check irc
    'If Froude_Max > 1# Then
    '    Froude = 0.4
    'Else
    '    Froude = 0.03
    'End If
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    Close #nf
    On Error GoTo 0
    Exit Sub
ERH1:
    MsgBox "�s�藬�v�Z���ʂ�ǂݍ��ݒ��G���[�����������B" & vbCrLf & _
           "������x�v�Z�����肢���܂��A�Q��ڈȍ~���G���[����������Ƃ���" & vbCrLf & _
           "�s�藬�v�Z�̈ȏ�I�����l�����܂��A���̓f�[�^�����m���߉������B", vbInformation
    On Error GoTo 0
    irc = False
    Close #nf
End Sub

'******************************************************************************
'�T�u���[�`���FLOG_Out(msg As String)
'�����T�v�F
'******************************************************************************
Sub LOG_Out(msg As String)
    If LOF(Log_Num) > 3000000 Then
        Close #Log_Num
        Open_Log_File
    End If
    Print #Log_Num, Format(Now, "yyyy/mm/dd hh:nn:ss") & Space(2) & msg
End Sub

'******************************************************************************
'�T�u���[�`���FOpen_Log_File()
'�����T�v�F
'******************************************************************************
Sub Open_Log_File()
    Dim File    As String
    Dim L       As Long
    Log_Num = FreeFile
    File = App.Path & "\data\Log_file.dat"
    If Len(Dir(File)) > 0 Then
        L = FileLen(File)
        If L < 3000000 Then
            Open File For Append As #Log_Num
        Else
            Open File For Output As #Log_Num
        End If
    Else
        Open File For Output As #Log_Num
    End If
End Sub

'******************************************************************************
'�֐��G
'�����T�v�FTIMEC(dw As Date) As String
'******************************************************************************
Public Function TIMEC(dw As Date) As String
    TIMEC = Format(dw, "yyyy/mm/dd hh:nn")
End Function

'******************************************************************************
'�T�u���[�`���G�v�n��ƌ܏��㗬�[����()
'�����T�v�FTIMEC(dw As Date) As String
'******************************************************************************
Sub �v�n��ƌ܏��㗬�[����()
    Dim i     As Integer
    Dim j     As Integer
    Dim buf   As String
    Dim a(2)  As Single
    Dim b(2)  As Single
    Dim qk    As Single
    Dim qh    As Single
    Dim nf    As Integer
    Dim d     As Date
    LOG_Out " �v�n��ƌ܏��㗬�[���� In"
    '******************************************************
    '
    '******************************************************
    nf = FreeFile
    Open App.Path & "\data\HQ.DAT" For Input As #nf
    Line Input #nf, buf
    Line Input #nf, buf                                                 '�v�n��g�|�p��
    HQA(1) = CSng(Mid(buf, 1, 10))
    HQB(1) = CSng(Mid(buf, 11, 10))
    Line Input #nf, buf
    Line Input #nf, buf                                                 '�t���g�|�p��
    HQA(2) = CSng(Mid(buf, 1, 10))
    HQB(2) = CSng(Mid(buf, 11, 10))
    Close #nf
    '******************************************************
    '
    '******************************************************
    nf = FreeFile
    Open Wpath & "OBSQ.DAT" For Output As #nf
    Print #nf, "    DATE    TIME    �v�n��      �t��"
    '******************************************************
    '2001/11/14 12:10*******.**-------.--
    '******************************************************
    d = jsd
    For i = 1 To Now_Step
        If HO(6, i) > -50# Then
            qk = HQA(1) * (HO(6, i) + HQB(1)) ^ 2
        Else
            qk = -99#
        End If
        If HO(7, i) > -50# Then
            qh = HQA(2) * (HO(7, i) + HQB(2)) ^ 2
        Else
            qh = -99#
        End If
        buf = Format(d, "yyyy/mm/dd hh:nn") & Format(Format(qk, "#####0.000"), "@@@@@@@@@@") & _
                                              Format(Format(qh, "#####0.000"), "@@@@@@@@@@")
        Print #nf, buf
        d = DateAdd("h", 1, d)
    Next i
    Close #nf
    LOG_Out " �v�n��ƌ܏��㗬�[���� Out"
End Sub

'******************************************************************************
'�T�u���[�`���F�O���M�����`�F�b�N(Cat As String, irc As Long)
'�d�l
'�O��󂯎������������15���߂��Ă��f�[�^����M�ł��Ȃ�������
'�v�Z���X�L�b�v������悤�ɂ���
'
'����
'Cat......"KISYO"=�C�ے��v�Z�� "FRICS"=FRICS�v�Z��
'
'******************************************************************************
Sub �O���M�����`�F�b�N(Cat As String, irc As Long)
    Dim da   As String
    Dim Dm   As String
    Dim dw   As Date
    Dim FLw  As String
    Dim nf   As Long
    Dim nn   As Long
    irc = True
    Select Case Cat
        Case "KISYO"
            '******************************************************
            '�C�ے��i�E�L���X�g�f�[�^�`�F�b�N
            '******************************************************
            FLw = Current_Path & "Oracletest\oraora\Data\F_MESSYU_10MIN_1.DAT"
            nf = FreeFile
            Open FLw For Input As #nf
            Line Input #nf, da
            Line Input #nf, Dm
            Close #nf
            dw = CDate(Dm)
            nn = DateDiff("n", dw, Now)
            If nn > 15 Then
                ADD_ERROR_Message "�C�ے��i�E�L���X�g�f�[�^�����͂���܂���ł��� " & Yosoku_Time_K & " �̌v�Z���X�L�b�v���܂��B"
                Data_Time_Rewrite da, FLw
                irc = False
            End If
            '******************************************************
            '�C�ے����щJ�ʎ���
            '******************************************************
            nf = FreeFile
            FLw = Current_Path & "Oracletest\oraora\Data\P_MESSYU_10MIN.dat"
            Open FLw For Input As #nf
            Line Input #nf, da
            Line Input #nf, Dm
            Close #nf
            dw = CDate(Dm)
            nn = DateDiff("n", dw, Now)
            If nn > 15 Then
                ADD_ERROR_Message "�C�ے����щJ�ʃf�[�^�����͂���܂���ł��� " & Yosoku_Time_K & " �̌v�Z���X�L�b�v���܂��B"
                Data_Time_Rewrite da, FLw
                irc = False
            End If
        Case "FRICS"
            '******************************************************
            'FRICS ���щJ�ʎ���
            '******************************************************
            nf = FreeFile
            FLw = Current_Path & "Oracletest\oraora\Data\P_RADAR.dat.dat"
            Open FLw For Input As #nf
            Line Input #nf, da
            Line Input #nf, Dm
            Close #nf
            dw = CDate(Dm)
            nn = DateDiff("n", dw, Now)
            If nn > 15 Then
                ADD_ERROR_Message "FRICS���щJ�ʃf�[�^�����͂���܂���ł��� " & Yosoku_Time_F & " �̌v�Z���X�L�b�v���܂��B"
                Data_Time_Rewrite da, FLw
                irc = False
            End If
            '******************************************************
            'FRICS �\���J�ʎ���
            '******************************************************
            nf = FreeFile
            FLw = Current_Path & "Oracletest\oraora\Data\F_RADAR.dat.dat"
            Open FLw For Input As #nf
            Line Input #nf, da
            Line Input #nf, Dm
            Close #nf
            dw = CDate(Dm)
            nn = DateDiff("n", dw, Now)
            If nn > 15 Then
                ADD_ERROR_Message "FRICS�\���J�ʃf�[�^�����͂���܂���ł��� " & Yosoku_Time_F & " �̌v�Z���X�L�b�v���܂��B"
                Data_Time_Rewrite da, FLw
                irc = False
            End If
    End Select
End Sub
