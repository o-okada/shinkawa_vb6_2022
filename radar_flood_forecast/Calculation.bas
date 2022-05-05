Attribute VB_Name = "Calculation"
'******************************************************************************
'���W���[�����FCalculation
'
'******************************************************************************
Option Explicit
Option Base 1

Public Const INFINITE = -1&
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwAccess As Long, ByVal flnherit As Integer, ByVal hObject As Long) As Long

Public p_name           As String                                       '�|���v����\���̎��̎{�ݖ�
Public isRAIN           As String                                       '�g�p�J��  "01"=�C�ے�   "02"=FRICS
Public isPump           As String                                       '"00"=�|���v����  "01"=�|���v�Ȃ�
Public FRICS            As Boolean                                      'FRICS�J�ʂŌv�Z����Ƃ���TRUE
Public KISYO            As Boolean                                      '�C�ے��J�ʂŌv�Z����Ƃ���TRUE
Public Pump_FULL_Data(100)   As String                                  '�t���|���v�f�[�^
Public Pump_FULL_num(100)    As Long
Public PDF_Date         As Date                                         'PDF�o�̓t�@�C�����Ɏg��
Public Rsa_Mag          As Single                                       '�����`�v���O������RSA�{��
Public Yosoku_Time_F    As String
Public Yosoku_Time_K    As String

'******************************************************************************
'�T�u���[�`���FADD_ERROR_Message(msg As String)
'�����T�v�F
'******************************************************************************
Sub ADD_ERROR_Message(msg As String)
    Error_Message_n = Error_Message_n + 1
    Error_Message(Error_Message_n) = msg
End Sub

'******************************************************************************
'�T�u���[�`���FAwaito_Time_Read(Cat As String, dw As Date)
'�����T�v�F
'�ҋ@�������f�B�X�N����ǂݍ���
'******************************************************************************
Sub Awaito_Time_Read(Cat As String, dw As Date)
    Dim nf As Long
    Dim F  As String
    Dim buf As String
    Select Case Cat
        Case "�C�ے�"
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
'�T�u���[�`���FAwaito_Time_Write(Cat As String, dw As Date)
'�����T�v�F
'�ҋ@�������f�B�X�N�ɏ�������
'******************************************************************************
Sub Awaito_Time_Write(Cat As String, dw As Date)
    Dim nf As Long
    Dim F  As String
    Select Case Cat
        Case "�C�ے�"
            F = App.Path & "\data\Await_Time_Kisyou.dat"
        Case "FRICS"
            F = App.Path & "\data\Await_Time_FRICS.dat"
        Case Else
            MsgBox "�����ɗ��Ă͂����܂���B"
    End Select
    nf = FreeFile
    Open F For Output As #nf
    Print #nf, Format(dw, "yyyy/mm/dd hh:nn")
    Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
    Close #nf
End Sub

'******************************************************************************
'�T�u���[�`���FData_Time_Rewrite(da As String, FL As String)
'�����T�v�F
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
'�T�u���[�`���FH_to_Pump()
'�����T�v�F
'���ѐ��ʂ���|���v�̃f�[�^���쐬
'******************************************************************************
Sub H_to_Pump()
End Sub

'******************************************************************************
'�T�u���[�`���FPre_Pump()
'�����T�v�F
'�t���|���v�f�[�^��ǂݍ���
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
'�T�u���[�`���FPrediction_CAL_By_KISYO_Veri(manu As Boolean)
'�����T�v�F
'�C�ے����э~�J�v�Z���ؗp
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
    LOG_Out "IN Prediction_CAL_By_KISYO �C�ے��J�ʂɂ��^���\���J�n ������=" & Format(jgd, "yyyy/mm/dd/hh:nn")
    Froude = 0#
    isRAIN = "01"                                                       '"01"=�C�ے�  "02"=FRICS
    isPump = "00"                                                       '"00"=�m�[�}�� "01"=�|���v��~
    Screen.MousePointer = vbHourglass
    '�v�n��ƌ܏��㗬�[����
    JRADAR = 0
    If MAIN.Check2 Then
        dwy = DateAdd("h", 3, jgd)
        MDB_�C�ے����[�_�[����2 jsd, dwy, dwj, irc
        If dwy <> dwj Then
            MsgBox "�C�ے����щJ�ʂɕK�v�Ƃ���J�ʂ��i�[����Ă��܂���B" & vbCrLf & _
                    "  jsd=" & Format(jsd, "yyyy/mm/dd hh:nn") & vbCrLf & _
                    "  dwy=" & Format(dwy, "yyyy/mm/dd hh:nn")
            End
        End If
        MDB_�� jsd, jgd, jrc
        If jrc > 1 Then
            LOG_Out "�C�ے��J�ʂŌv�Z���ɐ􉁂��܂��^���\���V�X�e���Ɏ�荞�܂�܂���ł����B�z����=0�Ƃ��Čv�Z���܂��B"
            ORA_Message_Out "�􉁉z���ʃf�[�^��M", "�C�ے��J�ʂŌv�Z���ɐ􉁂��܂��^���\���V�X�e���Ɏ�荞�܂�܂���ł����B�z����=0�Ƃ��Čv�Z���܂��B", 1
        End If
        ���[�_�[�J�ʏo��_Veri
        JRADAR = 1
    End If
    �|���v���^�f�[�^�ǂݍ���
    �|���v�\�͕\�ǂݍ���
    �|���v�f�[�^�쐬 jgd
    Set_Pump                                                            '���ʂɉ������ғ��A��~�|���v��ݒ肷��B
    Flood_Data_Write_For_Calc
    Message.Label1 = "�\ �� �v �Z �� �s ��"
    'Message.Label1 = "�r�g�h�m�j�P�O�@�� �s ��"
    Message.Show
    Message.ZOrder 0
    Message.Refresh
    ChDir App.Path & "\WORK"
    'Call WaitForProcessToEnd("RRSHINK10.EXE")                          '�v�n��t�B�[�h�o�b�N�L��
    'Call WaitForProcessToEnd("RRSHINK10NF.EXE")                        '�v�n��t�B�[�h�o�b�N����
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/06 O.OKADA
    '******************************************************
    'Call WaitForProcessToEnd("New_RSHINK.EXE")                          '�Ȃɂ�����
    Call WaitForProcessToEnd("D:\SHINKAWA\���[�_�[�^���\��\WORK\New_RSHINK.EXE")                          '�Ȃɂ�����
    '******************************************************
    'Ver1.0.0 �C���I�� 2015/08/06 O.OKADA
    '******************************************************
    'Message.Label1 = "�m �r �j�@�� �s ��"
    Message.Refresh
    Flood_Data_Write_For_Calc1
    'Cal_Initial_flow_profile irc                                       '�S�ĕs�藬�Ōv�Z���邽�ߏ������ʌ`�͌Œ�t�@�C���Ƃ��s�����v�Z�͎g��Ȃ�
    'If irc = False Then Exit Sub
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/06 O.OKADA
    '******************************************************
    'Call WaitForProcessToEnd("NEWNSKG2.EXE")
    Call WaitForProcessToEnd("D:\SHINKAWA\���[�_�[�^���\��\WORK\NEWNSKG2.EXE")                          '�Ȃɂ�����
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/06 O.OKADA
    '******************************************************
    'Message.Hide
    Screen.MousePointer = vbDefault
    ChDir App.Path
    Input_Yosoku irc
    If Not irc Then
        Exit Sub
    End If
    Category = True                                                     'True=�s�藬�v�Z��
    'If (Froude > 1.9) Or (HO(5, Now_Step) <= Un_Cal) Then
    If (HO(5, Now_Step) <= Un_Cal) Then
        Category = False
        EMG_Cal irc
        If Not manu Then AutoDrive.Timer1.Enabled = True
        If irc = False Then Exit Sub
    Else
        '�c�f�c�a�o�͗p
        For i = 1 To nd                                                 'nd=�f�ʐ�(74)
            YHJ(0, i) = HQ(1, i, NT - 18)                               '������
            YHJ(1, i) = HQ(1, i, NT - 12)                               '�\��+1
            YHJ(2, i) = HQ(1, i, NT - 6)                                '�\��+2
            YHJ(3, i) = HQ(1, i, NT)                                    '�\��+3
        Next i
        For i = 1 To 5
            FeedBack i
        Next i
        '�c�f�␳�p
        For i = 1 To 5
            ns = V_Sec_Num(i)                                           'V_Sec_Num(nr)�͕s�藬��̒f�ʈʒu��\��
            OHJ(0, i) = HQ(1, ns, NT - 18)
            OHJ(1, i) = HQ(1, ns, NT - 12)
            OHJ(2, i) = HQ(1, ns, NT - 6)
            OHJ(3, i) = HQ(1, ns, NT)
        Next i
    End If
    If MDBx Then MDB_����_Write                                         '�f�[�^�x�[�X�ɗ\���l�̏�������
    Load Graph3
    Graph3.Show
    Graph3.Refresh
    If Verification2.Check3 = vbChecked Then
        Graph3.VSPDF1.ConvertDocument Graph3.VSP, App.Path & "\" & Format(PDF_Date, "yyyy_mm_dd") & "_Hydro.pdf"
    End If
    '�^���\�񕶈č쐬                                                   '�e�X�g���͂����𐶂���
    If DBX_ora Then                                                     '�v�Z���ʂ�
        ORA_DataBase_Connection
        If OraDB_OK Then                                                '�����c�a���g�p�\�ȂƂ�
            '�^���\�񕶈č쐬
            LOG_Out "�C�ے��J�ʁ@���f�\�����ʏ������݊J�n"
            ORA_SUII_YOSOKU_KIJYUN_PUT rc
            LOG_Out "�C�ے��J�ʁ@���f�\�����ʏ������ݏI��"
            LOG_Out "�C�ے��J�ʁ@�c�f�\�����ʏ������ݏI��"
            ORA_SUII_YOSOKU_JYUDAN_PUT rc
            LOG_Out "�C�ے��J�ʁ@�c�f�\�����ʏ������ݏI��"
            ORA_DataBase_Close
        End If
    Else
    End If
    LOG_Out "IN Prediction_CAL_By_KISYO �C�ے��J�ʂɂ��^���\���I�� ������=" & Format(jgd, "yyyy/mm/dd/hh:nn")
    Message.Hide
    If Not manu Then
        Short_Break 2
        Unload Graph3
        AutoDrive.Timer1.Enabled = True
    End If
End Sub

'******************************************************************************
'�T�u���[�`���FPrediction_CAL_By_KISYO(manu As Boolean)
'�����T�v�F
'******************************************************************************
Sub Prediction_CAL_By_KISYO(manu As Boolean)
'�������C���J�n2016/03/04������
'�v�Z�����ŃG���[�����������ꍇ�A�v�Z����ʂ��|�b�v�A�b�v�����܂܂ƂȂ邽�߁B
On Error GoTo ERR1:
'�������C���I��2016/03/04������
    Dim dwj     As Date
    Dim dwy     As Date
    Dim irc     As Boolean
    Dim jrc     As Long
    Dim rc      As Boolean
    Dim i       As Integer
    Dim ns      As Long
    Dim ts      As Long
    LOG_Out "IN Prediction_CAL_By_KISYO �C�ے��J�ʂɂ��^���\���J�n ������=" & Format(jgd, "yyyy/mm/dd/hh:nn")
    Froude = 0#
    isRAIN = "01"                                                       '"01"=�C�ے�  "02"=FRICS
    isPump = "00"                                                       '"00"=�m�[�}�� "01"=�|���v��~
    Screen.MousePointer = vbHourglass
    '�v�n��ƌ܏��㗬�[����
    JRADAR = 0
    If MAIN.Check2 Then
        MDB_�C�ے����[�_�[����2 jsd, jgd, dwj, irc
        If dwj < jgd Then jgd = dwj
        dwj = DateAdd("h", 1, jgd)
        dwy = DateAdd("h", 3, jgd)
        MDB_�C�ے����[�_�[�\��2 dwj, dwy, irc
        If irc = False Then
            If manu Then
                MsgBox "�C�ے��\���J�ʂ��܂�����M���o�^����Ă��܂���B" & vbCrLf & _
                        "���������Đݒ肵�ĉ������B"
            End If
            LOG_Out "�C�ے��\���J�ʂ��܂�����M���o�^����Ă��܂���A�v�Z���X�L�b�v���܂��B"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        MDB_�� jsd, jgd, jrc
        Select Case jrc
            Case 0                                                      '����擾
            Case 1                                                      '10���O�擾
                LOG_Out "�􉁃f�[�^���O�����f�[�^���g���B"
                ORA_Message_Out "�􉁉z���ʃf�[�^��M", "�C�ے��J�ʂɂ��v�Z�ɂ����āA�􉁃f�[�^����荞�܂�܂���ł����B�O�����f�[�^�Ōv�Z���܂��B", 1
            Case 2                                                      '�擾�ł���
                LOG_Out "�􉁃f�[�^��10���O���擾�ł��܂���A�f�[�^��0�Ƃ����܂܌v�Z���܂��B"
                ORA_Message_Out "�􉁉z���ʃf�[�^��M", "�C�ے��J�ʂɂ��v�Z�ɂ����āA�􉁃f�[�^��2�����ȏ�A�����Ď�荞�܂�܂���ł����B�z����=0�Ƃ��Čv�Z���܂��B", 1
        End Select
        ���[�_�[�J�ʏo��
        JRADAR = 1
    End If
    �|���v���^�f�[�^�ǂݍ���
    �|���v�\�͕\�ǂݍ���
    �|���v�f�[�^�쐬 jgd
    Set_Pump                                                            '���ʂɉ������ғ��A��~�|���v��ݒ肷��B
    Flood_Data_Write_For_Calc
    Message.Label1 = "�C�ے��J�� �\ �� �v �Z �� �s ��"
    'Message.Label1 = "�r�g�h�m�j�P�O�@�� �s ��"
    Message.Show
    Message.ZOrder 0
    Message.Refresh
    ChDir App.Path & "\WORK"
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/06 O.OKADA
    '******************************************************
    'Call WaitForProcessToEnd("RRSHINK10.EXE")
    'Call WaitForProcessToEnd("RRSHINK10NF.EXE")                        '�v�n��t�B�[�h�o�b�N����
    'Call WaitForProcessToEnd("New_RSHINK.EXE")                          '�Ȃɂ�����
    Call WaitForProcessToEnd("D:\SHINKAWA\���[�_�[�^���\��\WORK\New_RSHINK.EXE")                          '�Ȃɂ�����
    '******************************************************
    'Ver1.0.0 �C���I�� 2015/08/06 O.OKADA
    '******************************************************
    'Message.Label1 = "�m �r �j�@�� �s ��"
    Message.Refresh
    Flood_Data_Write_For_Calc1
    'Cal_Initial_flow_profile irc                                       '�S�ĕs�藬�v�Z�ōs���׏������ʌ`�͌Œ�ƂȂ����וs�����v�Z�͎g�p���Ȃ�
    'If irc = False Then Exit Sub
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/06 O.OKADA
    '******************************************************
    'Call WaitForProcessToEnd("NEWNSKG2.EXE")
    Call WaitForProcessToEnd("D:\SHINKAWA\���[�_�[�^���\��\WORK\NEWNSKG2.EXE")
    '******************************************************
    'Ver1.0.0 �C���I�� 2015/08/06 O.OKADA
    '******************************************************
    'Message.Hide
    Screen.MousePointer = vbDefault
    ChDir App.Path
    Input_Yosoku irc
    If Not irc Then
        Message.Hide
        Exit Sub
    End If
    Category = True                                                     'True=�s�藬�v�Z��
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
        '�c�f�c�a�o�͗p
        For i = 1 To 74
            YHJ(0, i) = HQ(1, i, NT - 18)
            YHJ(1, i) = HQ(1, i, NT - 12)
            YHJ(2, i) = HQ(1, i, NT - 6)
            YHJ(3, i) = HQ(1, i, NT)
        Next i
        For i = 1 To 5
            FeedBack i
        Next i
        '�c�f�␳�p
        For i = 1 To 5
            ns = V_Sec_Num(i)                                           'V_Sec_Num(nr)�͕s�藬��̒f�ʈʒu��\��
            OHJ(0, i) = HQ(1, ns, NT - 18)
            OHJ(1, i) = HQ(1, ns, NT - 12)
            OHJ(2, i) = HQ(1, ns, NT - 6)
            OHJ(3, i) = HQ(1, ns, NT)
        Next i
    End If
    If MDBx Then MDB_����_Write                                         '�f�[�^�x�[�X�ɗ\���l�̏�������
    Load Graph3
    Graph3.Show
    Graph3.Refresh
    If Verification2.Check3 = vbChecked Then
        Graph3.VSPDF1.ConvertDocument Graph3.VSP, App.Path & "\" & Format(PDF_Date, "yyyy_mm_dd") & "_Hydro.pdf"
    End If
    '�^���\�񕶈č쐬                                                   '�e�X�g���͂����𐶂���
    '�\�񕶃`�F�b�N
    If DBX_ora Then                                                     '�v�Z���ʂ�
        ORA_DataBase_Connection
        If OraDB_OK Then                                                '�����c�a���g�p�\�ȂƂ�
            �\�񕶃`�F�b�N
            LOG_Out "�C�ے��J�ʁ@���f�\�����ʏ������݊J�n"
            ORA_SUII_YOSOKU_KIJYUN_PUT rc
            LOG_Out "�C�ے��J�ʁ@���f�\�����ʏ������ݏI��"
            LOG_Out "�C�ے��J�ʁ@�c�f�\�����ʏ������ݏI��"
            ORA_SUII_YOSOKU_JYUDAN_PUT rc
            LOG_Out "�C�ے��J�ʁ@�c�f�\�����ʏ������ݏI��"
            ORA_DataBase_Close
        End If
    End If
    LOG_Out "IN Prediction_CAL_By_KISYO �C�ے��J�ʂɂ��^���\���I�� ������=" & Format(jgd, "yyyy/mm/dd/hh:nn") & " manu=" & manu
    If Not manu Then
        Short_Break 5
        Unload Graph3
        AutoDrive.Timer1.Enabled = True
    End If
'�������C���J�n2016/03/04������
'�v�Z�����ŃG���[�����������ꍇ�A�v�Z����ʂ��|�b�v�A�b�v�����܂܂ƂȂ邽�߁B
ERR1:
    If Message.Visible = True Then
        Message.Hide
    End If
'�������C���I��2016/03/04������
End Sub

'******************************************************************************
'�T�u���[�`���FPrediction_CAL_By_FRICS(manu As Boolean)
'�����T�v�F
'manu True=�蓮�v�Z�� False=�����v�Z��
'******************************************************************************
Sub Prediction_CAL_By_FRICS(manu As Boolean)
    Dim dwj     As Date
    Dim irc     As Boolean
    Dim jrc     As Long
    Dim rc      As Boolean
    Dim i       As Integer
    Dim ns      As Long
    Dim ts      As Long
    LOG_Out "IN Prediction_CAL_By_FRICS FRICS�J�ʂɂ��^���\���J�n ������=" & Format(jgd, "yyyy/mm/dd/hh:nn")
    Froude = 0#
    isRAIN = "02"                                                       '"01"=�C�ے�  "02"=FRICS
    isPump = "00"                                                       '"00"=�m�[�}�� "01"=�|���v��~
    Screen.MousePointer = vbHourglass
    '�v�n��ƌ܏��㗬�[����
    JRADAR = 0
    If MAIN.Check2 Then
        MDB_FRICS���[�_�[���� jsd, jgd, dwj, irc
        If dwj < jgd Then jgd = dwj
        dwj = DateAdd("h", 1, jgd)
        MDB_FRICS���[�_�[�\�� jgd, irc
        If irc = False Then
            If manu Then
                MsgBox "FRICS�\���J�ʂ��܂�����M���o�^����Ă��܂���B" & vbCrLf & _
                        "���������Đݒ肵�ĉ������B"
            End If
            LOG_Out "FRICS�\���J�ʂ��܂�����M���o�^����Ă��܂���A�v�Z���X�L�b�v���܂��B"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        MDB_�� jsd, jgd, jrc
        Select Case jrc
            Case 0                                                      '����擾
            Case 1                                                      '10���O�擾
                LOG_Out "�􉁃f�[�^���O�����f�[�^���g���B"
                ORA_Message_Out "�􉁉z���ʃf�[�^��M", "FRICS�J�ʂɂ��v�Z�ɂ����āA�􉁃f�[�^����荞�܂�܂���ł����B�O�����f�[�^�Ōv�Z���܂��B", 1
            Case 2                                                      '�擾�ł���
                LOG_Out "�􉁃f�[�^��10���O���擾�ł��܂���A�f�[�^��0�Ƃ����܂܌v�Z���܂��B"
                ORA_Message_Out "�􉁉z���ʃf�[�^��M", "FRICS�J�ʂɂ��v�Z�ɂ����āA�􉁃f�[�^��2�����ȏ�A�����Ď�荞�܂�܂���ł����B�z����=0�Ƃ��Čv�Z���܂��B", 1
        End Select
        ���[�_�[�J�ʏo��
        JRADAR = 1
    End If
    �|���v���^�f�[�^�ǂݍ���
    �|���v�\�͕\�ǂݍ���
    �|���v�f�[�^�쐬 jgd
    Set_Pump                                                            '���ʂɉ������ғ��A��~�|���v��ݒ肷��B
    Flood_Data_Write_For_Calc
    Message.Label1 = "�e�q�h�b�r�J�ʁ@�\ �� �v �Z �� �s ��"
    'Message.Label1 = "�r�g�h�m�j�P�O�@�� �s ��"
    Message.Show
    Message.ZOrder 0
    Message.Refresh
    ChDir App.Path & "\WORK"
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/06 O.OKADA
    '******************************************************
    'Call WaitForProcessToEnd("RRSHINK10.EXE")                          '�v�n��t�B�[�h�o�b�N�L��
    'Call WaitForProcessToEnd("RRSHINK10NF.EXE")                        '�v�n��t�B�[�h�o�b�N����
    Call WaitForProcessToEnd("D:\SHINKAWA\���[�_�[�^���\��\WORK\New_RSHINK.EXE")                          '�Ȃɂ�����
    'Message.Label1 = "�m �r �j�@�� �s ��"
    Message.Refresh
    '******************************************************
    'Ver1.0.0 �C���I�� 2015/08/06 O.OKADA
    '******************************************************
    Flood_Data_Write_For_Calc1
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/06 ���c��
    '******************************************************
    'Cal_Initial_flow_profile irc                                       '�S�ĕs�藬�v�Z�ōs���׏������ʌ`�͌Œ�ƂȂ����וs�����v�Z�͎g�p���Ȃ�
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
    Category = True                                                     'True=�s�藬�v�Z��
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
        '�c�f�c�a�o�͗p
        For i = 1 To 74
            YHJ(0, i) = HQ(1, i, NT - 18)
            YHJ(1, i) = HQ(1, i, NT - 12)
            YHJ(2, i) = HQ(1, i, NT - 6)
            YHJ(3, i) = HQ(1, i, NT)
        Next i
        For i = 1 To 5
            FeedBack i
        Next i
        '�c�f�␳�p
        For i = 1 To 5
            ns = V_Sec_Num(i)                                           'V_Sec_Num(nr)�͕s�藬��̒f�ʈʒu��\��
            OHJ(0, i) = HQ(1, ns, NT - 18)
            OHJ(1, i) = HQ(1, ns, NT - 12)
            OHJ(2, i) = HQ(1, ns, NT - 6)
            OHJ(3, i) = HQ(1, ns, NT)
        Next i
    End If
    If MDBx Then MDB_����_Write                                         '�f�[�^�x�[�X�ɗ\���l�̏�������
    Load Graph3
    Graph3.Show
    Graph3.Refresh
    '�^���\�񕶈č쐬                                                   '�e�X�g���͂����𐶂���
    '�\�񕶃`�F�b�N
    If DBX_ora Then
        ORA_DataBase_Connection
        If OraDB_OK Then
            �\�񕶃`�F�b�N
            LOG_Out "FRICS�J�ʁ@�\���J�ʏ������݊J�n"
            ORA_FRICS_RAIN                                              '�e�q�h�b�r�\���J�ʏ�������
            LOG_Out "FRICS�J�ʁ@�\���J�ʏ������ݏI��"
            LOG_Out "FRICS�J�ʁ@���f�\�����ʏ������݊J�n"
            ORA_SUII_YOSOKU_KIJYUN_PUT rc
            LOG_Out "FRICS�J�ʁ@���f�\�����ʏ������ݏI��"
            LOG_Out "FRICS�J�ʁ@�c�f�\�����ʏ������݊J�n"
            ORA_SUII_YOSOKU_JYUDAN_PUT rc
            LOG_Out "FRICS�J�ʁ@�c�f�\�����ʏ������ݏI��"
            ORA_DataBase_Close
        End If
    Else
    End If
    LOG_Out "Out Prediction_CAL_By_FRICS FRICS�J�ʂɂ��^���\���I�� ������=" & Format(jgd, "yyyy/mm/dd/hh:nn") & _
            " manu=" & manu
    If Not manu Then
        Short_Break 5
        Unload Graph3
        AutoDrive.Timer1.Enabled = True
    End If
End Sub

'******************************************************************************
'�T�u���[�`���FSet_Pump()
'�����T�v�F
'���ʂɉ������ғ��A��~�|���v��ݒ肷��B
'******************************************************************************
Sub Set_Pump()
    Dim ConR        As New ADODB.Recordset
    Dim SQL         As String
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long
    Dim nf          As Integer
    LOG_Out "In  Set_Pump"
    Pre_Pump                                                            '�t���|���v�f�[�^��ǂ�
    SQL = "select * from �|���v���� where Time = '" & Format(jgd, "yyyy/mm/dd hh:nn") & "'"
    ConR.Open SQL, Con_����, adOpenKeyset, adLockReadOnly
    'j = 0
    '******************************************************
    '���V��F�`�F�b�N
    '******************************************************
    j = ConR.Fields("���V��F").Value
    If j = 1 Then
        For i = 1 To 8
            k = Pump_FULL_num(i)
            Mid(Pump_FULL_Data(k), 51, 5) = "    0"
        Next i
    End If
    '******************************************************
    '�����O���ʃ`�F�b�N
    '******************************************************
    j = ConR.Fields("�����O����").Value
    If j = 1 Then
        For i = 9 To 20
            k = Pump_FULL_num(i)
            Mid(Pump_FULL_Data(k), 51, 5) = "    0"
        Next i
    End If
    '******************************************************
    '�t���`�F�b�N
    '******************************************************
    j = ConR.Fields("�t��").Value
    If j = 1 Then
        For i = 21 To 27
            k = Pump_FULL_num(i)
            Mid(Pump_FULL_Data(k), 51, 5) = "    0"
        Next i
    End If
    ConR.Close
    '******************************************************
    '�|���v�f�[�^�o��
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
'�T�u���[�`���FWaitForProcessToEnd(cmdLine As String)
'�����T�v�F
'******************************************************************************
Sub WaitForProcessToEnd(cmdLine As String)
    'INFINITE���~���b�P�ʂ̎��Ԃɒu�������鎖���o����
    Dim retVal As Long, pID As Long, pHandle As Long
    pID = Shell(cmdLine, vbMinimizedFocus)                              ' vbMinimizedFocus   vbNormalFocus
    pHandle = OpenProcess(&H100000, True, pID)
    retVal = WaitForSingleObject(pHandle, INFINITE)
End Sub
