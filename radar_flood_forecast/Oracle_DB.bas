Attribute VB_Name = "Oracle_DB"
'******************************************************************************
'���W���[�����FOracle_DB
'
'�C�������FVer0.0.0
'��2015�Ɩ��ŋ��\�[�X�R�[�h��̎��_�Ŋ��ɃR�����g�A�E�g����Ă������̂ɂ��ẮA
'Ver0.0.0�ɂ��C���Ƃ����B
'
'�C�������FVer1.0.0
'2015/08/06�F�I���N���f�[�^�x�[�X�̃e�[�u���폜�ɑΉ����A���L�e�[�u���ւ�
'�A�N�Z�X�����������R�����g�A�E�g����B
'C_FUKENKOZUIAN
'P_MESSYU_1HOUR
'P_MESSYU_10MIN
'RYUIKI_1HOUR_FRICS
'RYUIKI_10MIN_FRICS
'
'C_FUKENKOZUIAN
'Oracle_DB.ORA_YOHOUBUNAN()���C�������B                     �y01�z�ς�
'���\�񕶃p�^�[���`�F�b�N.�\�񕶃`�F�b�N()���C�����邱�ƁB  �y01-01�z�ς�

'��Calculation.Prediction_CAL_By_FRICS()���C�����邱�ƁB    �y01-01-01�z
'��Calculation.Prediction_CAL_By_KISYO()���C�����邱�ƁB    �y01-01-02�z

'��AutoDrive.Data_Check_FRICS_z()���C�����邱�ƁB           �y01-01-01-01�z
'��AutoDrive.Data_Check_FRICS_u()���C�����邱�ƁB           �y01-01-01-02�z
'��SHINKAWA.Command1_Click()���C�����邱�ƁB                �y01-01-01-03�z
'��Verification2.Command1_Click()���C�����邱�ƁB           �y01-01-01-04�z
'��AutoDrive.Timer1_Timer()���C�����邱�ƁB                 �y01-01-01-01-01�z
'
'��AutoDrive.Data_Check_KISYO_z()���C�����邱�ƁB           �y01-01-02-01�z
'��AutoDrive.Data_Check_KISYO_u()���C�����邱�ƁB           �y01-01-02-02�z
'��SHINKAWA.Command1_Click()���C�����邱�ƁB                �y01-01-02-03�z
'��Verification2.Command1_Click()���C�����邱�ƁB           �y01-01-02-04�z
'��AutoDrive.Timer1_Timer()���C�����邱�ƁB                 �y01-01-02-01-01�z
'
'P_MESSYU_1HOUR
'Oracle_DB.Data_Time_Check()���C�������B                    �y02�z
'�����Y�T�u���[�`�����Ăяo���Ă��鏈�������݂��Ȃ����߁A�e���Ȃ��B
'
'P_MESSYU_10MIN
'Oracle_DB.Data_Time_Check()���C�������B                    �y03�z
'�����Y�T�u���[�`�����Ăяo���Ă��鏈�������݂��Ȃ����߁A�e���Ȃ��B
'
'RYUIKI_1HOUR_FRICS
'Oracle_DB.ORA_FRICS_RAIN()���C�������B                     �y04�z
'��Calculation.Prediction_CAL_By_FRICS()���C�����邱�ƁB    �y04-01�z
'��AutoDrive.Data_Check_FRICS_z()���C�����邱�ƁB           �y04-01-01�z
'��AutoDrive.Data_Check_FRICS_u()���C�����邱�ƁB           �y04-01-02�z
'��SHINKAWA.Command1_Click()���C�����邱�ƁB                �y04-01-03�z
'��Verification2.Command1_Click()���C�����邱�ƁB           �y04-01-04�z
'��AutoDrive.Timer1_Timer()���C�����邱�ƁB                 �y04-01-01-01�z
'
'RYUIKI_1HOUR_FRICS
'Oracle_DB.ORA_FRICS_RAIN()���C�������B                     �y05�z
'��Calculation.Prediction_CAL_By_FRICS()���C�����邱�ƁB    �y05-01�z
'��AutoDrive.Data_Check_FRICS_z()���C�����邱�ƁB           �y05-01-01�z
'��AutoDrive.Data_Check_FRICS_u()���C�����邱�ƁB           �y05-01-02�z
'��SHINKAWA.Command1_Click()���C�����邱�ƁB                �y05-01-03�z
'��Verification2.Command1_Click()���C�����邱�ƁB           �y05-01-04�z
'��AutoDrive.Timer1_Timer()���C�����邱�ƁB                 �y05-01-01-01�z
'
'******************************************************************************
Option Explicit
Option Base 1

'******************************************************************************
'OO4O�̃I�u�W�F�N�g�ϐ���錾����
'******************************************************************************
'******************************************************************************
'Ver1.0.0 �C���J�n 2015/08/07 O.OKADA
'��OraDatabase�AOraDynaset�ł́AFields("")�A�N�Z�X���ɁA
'�u�����̐�����v���Ă��܂���B�܂��͕s���ȃv���p�e�B���w�肵�Ă��܂��B�v��
'�G���[���\�������B
'��OraDBT.exe�����l�̂��߁A�C������K�v������B
'******************************************************************************
'Public ssOra              As Object
'Public dbOra              As OraDatabase
'Public dynOra             As OraDynaset
'Public OraDB_OK           As Boolean
'Public ssOra              As Object
Public dbOra              As Object
Public dynOra             As Object
Public gAdoCon As ADODB.Connection
Public gAdoRst As ADODB.Recordset
Public OraDB_OK           As Boolean
'******************************************************************************
'Ver1.0.0 �C���I�� 2015/08/07 O.OKADA
'******************************************************************************
'******************************************************************************
'���f�}�p�\���l
'******************************************************************************
Public YHK(6, 0 To 18)    As Single
'******************************************************************************
'�c�f�}�f�[�^�p
'******************************************************************************
Public Ed1(4, 2, 50)      As Single                                     '�c�f�f�[�^�␳�W��
Public Ied(4)             As Long                                       '��n�_�Ԓf�ʐ�
Public YHJ(0 To 3, 75)    As Single                                     '1���ԃs�b�`��3���Ԍ�܂Ł@�V��52�f�� �܏��22�f��
Public OHJ(0 To 3, 5)     As Single                                     '�t�B�[�h�o�b�N�O�̌v�Z�l�i�c�f�␳�p�j
Public Base_Time          As Date                                       '�\���v�Z�J�n����(��)
Public Suii_Time          As Date                                       '�e�����[�^���ʂ̍ŐV����
Public Obs_Time           As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function GetConfigData(ByVal aSection As String, ByVal aKey As String, ByVal aFilename As String) As String

    Const intMaxSize As Integer = 255
    Dim strBuffer As String

    strBuffer = Space(intMaxSize)
    If GetPrivateProfileString(aSection, aKey, vbNullString, strBuffer, intMaxSize, aFilename) Then
        GetConfigData = SetNullCharCut(strBuffer)
    Else
        GetConfigData = vbNullString
    End If

End Function

Public Function SetNullCharCut(ByVal aChar As String) As String

    Dim intNullPos As Integer

    intNullPos = InStr(aChar, vbNullChar)
    If intNullPos > 0 Then
        SetNullCharCut = Left(aChar, intNullPos - 1)
    Else
        SetNullCharCut = aChar
    End If

End Function

'******************************************************************************
'�T�u���[�`���FData_Time_Check(Kind As Integer, Data_Time As Date, irc As Boolean)
'�����T�v�F
'�����̃T�u���[�`���͎g�p����Ă��܂���B
'******************************************************************************
Sub Data_Time_Check(Kind As Integer, Data_Time As Date, irc As Boolean)





'*************
'DB�o�͗}�~�ŁA�I���N�����b�Z�[�W�o�͗}�~�ŁiEXIT SUB����j�y�~�z
'    Exit Sub
'*************

'*************
'DB�o�͎��{�A�I���N�����b�Z�[�W�o�͗}�~�ŁiEXIT SUB���Ȃ��j�y���z
    
'*************






    Dim SQL      As String
    Dim i        As Integer
    Dim w
    SQL = "SELECT * FROM KANSOKU_JIKOKU WHERE "
    Select Case Kind
        Case 1                                                          '�e�����[�^����
            SQL = SQL & "table_name=P_WATER"
        Case 2                                                          '�C�ے����[�_�[�J�ʎ��ђ莞
            SQL = SQL & "table_name=P_MESSYU_10MIN"
        Case 3                                                          '�C�ے����[�_�[�J�ʎ��ѐ���
            SQL = SQL & "table_name=P_MESSYU_1HOUR"
        Case 4                                                          '�C�ے����[�_�[�J�ʗ\���莞
            SQL = SQL & "table_name=P_MESSYU_10MIN_1"
        Case 5                                                          '�C�ے����[�_�[�J�ʗ\������
            SQL = SQL & "table_name=P_MESSYU_10MIN_2"
        Case 6                                                          'FRICS���[�_�[�J�ʎ���
            SQL = SQL & "table_name=P_REDAR"
        Case 7                                                          'FRICS���[�_�[�J�ʗ\��
            SQL = SQL & "table_name=F_REDAR"
    End Select
    '******************************************************
    '�y�R�����g�z
    '2015/08/06 O.OKADA
    '�e�[�u���uP_MESSYU_10MIN_1�v�͑��݂��Ȃ��B�܂��A
    '�e�[�u���uP_MESSYU_10MIN_2�v�͑��݂��Ȃ��B���ꂼ��A
    '�uF_MESSYU_10MIN_1�v�uF_MESSYU_10MIN_2�v���������̂ł́H
    '******************************************************
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    'Rst_oraDB.Open SQL, Con_oraDB, adOpenDynamic, adLockReadOnly
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    '******************************************************
    '�t�B�[���h�����擾����
    '******************************************************
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    'n = Rst_oraDB.Fields.Count
    'For i = 0 To n - 1
    '    Tw = Rst_oraDB.Fields(i).Name
    '    Debug.Print " �t�B�[���h��="; Tw
    'Next i
    'If Rst_oraDB.EOF Then
    '    MsgBox "�ŐV�f�[�^�����̃e�[�u���Ƀf�[�^������܂���B"
    '    Rst_oraDB.Close
    '    irc = False
    '    Exit Sub
    'End If
    'w = Rst_oraDB.Fields("last_data_time").Value
    'If IsDate(w) Then
    '    Data_Time = CDate(w)
    '    irc = True
    'Else
    '    irc = False
    'End If
    'Rst_oraDB.Close
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
End Sub

'******************************************************************************
'�T�u���[�`���FF_2_I(F As Single)
'�����T�v�F
'******************************************************************************
Function F_2_I(F As Single)
    If F >= 0# Then
        F_2_I = Format(F * 100#, "00000")
    Else
        F_2_I = Format(F * 100#, "0000")
    End If
End Function

'******************************************************************************
'�T�u���[�`���FJyudan_Hosei()
'�����T�v�F
'******************************************************************************
Sub Jyudan_Hosei()






'*************
'DB�o�͗}�~�ŁA�I���N�����b�Z�[�W�o�͗}�~�ŁiEXIT SUB���Ȃ��j�y�~�z

'*************

'*************
'DB�o�͎��{�A�I���N�����b�Z�[�W�o�͗}�~�ŁiEXIT SUB���Ȃ��j�y���z
    
'*************






    Dim i     As Long
    Dim j     As Long
    Dim k     As Long
    Dim m     As Long
    Dim i1    As Long
    Dim i2    As Long
    Dim sdx   As Single
    Dim H1    As Single
    Dim H2    As Single
    Dim buf   As String
    Debug.Print V_Sec_Num(1); V_Sec_Num(2); V_Sec_Num(3); V_Sec_Num(4)
    For i = 1 To 53
        buf = Format(i, " 00") & "   "
        buf = buf & F_2_I(YHJ(1, i)) & "  " & F_2_I(YHJ(2, i)) & "  " & F_2_I(YHJ(3, i))
        Debug.Print buf
    Next i
    For i = 0 To 3
        i1 = 1
        H1 = 0#
        For m = 1 To 4                                                  '�V�� ��n�_�ԃu���b�N��
            Select Case m
                Case 1                                                  '������O���ʂ��牺�V��F�܂�
                    i1 = 1
                    i2 = V_Sec_Num(1)
                Case 2                                                  '���V��F����厡�܂�
                    i1 = V_Sec_Num(1) + 1
                    i2 = V_Sec_Num(2)
                Case 3                                                  '�厡���琅���O���ʂ܂�
                    i1 = V_Sec_Num(2) + 1
                    i2 = V_Sec_Num(3)
                Case 4                                                  '�����O���ʂ��牺�V��F�܂�
                    i1 = V_Sec_Num(3) + 1
                    i2 = V_Sec_Num(4)
            End Select
            j = 1
            H2 = OHJ(i, m) - YHJ(i, i2)
            For k = i1 To i2
                YHJ(i, k) = YHJ(i, k) + H1 * Ed1(m, 1, j) + H2 * Ed1(m, 2, j)
                j = j + 1
            Next k
            H1 = H2
            YHJ(i, 52) = YHJ(i, 51) + (YHJ(i, 51) - YHJ(i, 50))
            YHJ(i, 53) = YHJ(i, 52) + (YHJ(i, 51) - YHJ(i, 50)) / 2
        Next m
    Next i
    Debug.Print V_Sec_Num(1); V_Sec_Num(2); V_Sec_Num(3); V_Sec_Num(4)
    For i = 1 To 53
        buf = Format(i, " 00") & "   "
        buf = buf & F_2_I(YHJ(1, i)) & "  " & F_2_I(YHJ(2, i)) & "  " & F_2_I(YHJ(3, i))
        Debug.Print buf
    Next i
End Sub

'******************************************************************************
'�T�u���[�`���FJyudan_Hosei_Pre()
'�����T�v�F
'******************************************************************************
Sub Jyudan_Hosei_Pre()






'*************
'DB�o�͗}�~�ŁA�I���N�����b�Z�[�W�o�͗}�~�ŁiEXIT SUB���Ȃ��j�y�~�z

'*************

'*************
'DB�o�͎��{�A�I���N�����b�Z�[�W�o�͗}�~�ŁiEXIT SUB���Ȃ��j�y���z
    
'*************






    Dim i       As Long
    Dim j       As Long
    Dim k       As Long
    Dim m       As Long
    Dim i1      As Long
    Dim i2      As Long
    Dim sdx     As Single
    Dim f1      As Single
    Dim f2      As Single
    LOG_Out "  Jyudan_Hosei_Pre In"
    '******************************************************
    '������O���ʂ��牺�V��F�܂�
    '******************************************************
    For m = 1 To 4
        Select Case m
            Case 1                                                      '������O���ʂ��牺�V��F�܂�
                i1 = 1
                i2 = V_Sec_Num(1)
            Case 2                                                      '���V��F����厡�܂�
                i1 = V_Sec_Num(1)
                i2 = V_Sec_Num(2)
            Case 3                                                      '�厡���琅���O���ʂ܂�
                i1 = V_Sec_Num(2)
                i2 = V_Sec_Num(3)
            Case 4                                                      '�����O���ʂ��牺�V��F�܂�
                i1 = V_Sec_Num(3)
                i2 = V_Sec_Num(4)
        End Select
        sdx = 0#
        j = 1
        For i = i1 To i2
            If i = i1 Then
                Ed1(m, 2, j) = 0#
            Else
                sdx = sdx + DX(i)
                Ed1(m, 2, j) = sdx
            End If
            j = j + 1
        Next i
        Ied(m) = i2 - i1 + 1
        For i = 1 To Ied(m)
            Ed1(m, 2, i) = Ed1(m, 2, i) / sdx
            Ed1(m, 1, i) = 1# - Ed1(m, 2, i)
            '******************************************************
            'Ver0.0.0 �C���J�n 1900/01/01 00:00
            '******************************************************
            'LOG_Out " i=" & Format(i, "000") & " dx=" & fmt(DX(i1 + i - 1)) & "  ed2=" & fmt(Ed1(m, 2, i)) & "  ed1=" & fmt(Ed1(m, 1, i))
            '******************************************************
            'Ver0.0.0 �C���I�� 1900/01/01 00:00
            '******************************************************
        Next i
        '******************************************************
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        '******************************************************
        'Debug.Print "   "
        '******************************************************
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        '******************************************************
    Next m
    LOG_Out "  Jyudan_Hosei_Pre Out"
End Sub

'******************************************************************************
'�T�u���[�`���FORA_KANSOKU_JIKOKU_PUT(TBL As String, dw As Date)
'�����T�v�F
'�e�[�u�� KANSOKU_JIKOU �ɍŐV������������
' Number=  0 �t�B�[���h��=WRITE_TIME
' Number=  1 �t�B�[���h��=TABLE_NAME
' Number=  2 �t�B�[���h��=LAST_DATA_TIME
' Number=  3 �t�B�[���h��=DETAIL
' Number=  1 �t�B�[���h��=TABLE_NAME
' Number=  2 �t�B�[���h��=RIVER_NO
' Number=  3 �t�B�[���h��=RIVER_DIV
' Number=  4 �t�B�[���h��=LAST_DATA_TIME
' Number=  5 �t�B�[���h��=RAIN_KIND
' Number=  6 �t�B�[���h��=PONPU
'******************************************************************************
Sub ORA_KANSOKU_JIKOKU_PUT(TBL As String, dw As Date)
    Dim cDwn  As String
    Dim cDw   As String
    Dim SQL   As String
    Dim buf   As String
    On Error GoTo ErrOracle:
    SQL = "SELECT * FROM oracle.YOSOKU_SUII_JIKOKU WHERE TABLE_NAME='" & TBL & "'" & _
          " AND RAIN_KIND='" & isRAIN & "' AND PONPU='" & isPump & "'" & _
          " AND RIVER_NO ='85053002'"
    LOG_Out "IN  ORA_KANSOKU_JIKOKU_PUT  SQL=" & SQL
    
    
    
    
    
    
    
    
'*************
'DB�o�͗}�~�ŁA�I���N�����b�Z�[�W�o�͗}�~�ŁiEXIT SUB����j�y�~�z
'    Exit Sub
'*************

'*************
'DB�o�͎��{�A�I���N�����b�Z�[�W�o�͗}�~�ŁiEXIT SUB���Ȃ��j�y���z
    
'*************
    
    
    
    
    
    
    
    
    '******************************************************
    'SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����B
    '******************************************************
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    '******************************************************
    '�t�B�[���h�����擾����B
    '******************************************************
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    'Dim tw, i, n
    'n = dynOra.Fields.Count
    'For i = 0 To n - 1
    '    tw = dynOra.Fields(i).Name
    '    Debug.Print " Number=" & Format(str(i), "@@@") & " �t�B�[���h��="; tw
    'Next i
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    If dynOra.EOF Then
        dynOra.AddNew
        LOG_Out "�_�C�i�Z�b�g�擾 OK Addnew mode"
    Else
        dynOra.Edit
        LOG_Out "�_�C�i�Z�b�g�擾 OK Edit mode"
    End If
    cDwn = Format(Now, "yyyy/mm/dd hh:nn")
    dynOra.Fields("write_time").Value = cDwn
    dynOra.Fields("table_name").Value = TBL
    cDw = Format(jgd, "yyyy/mm/dd hh:nn")
    dynOra.Fields("last_data_time").Value = cDw
    dynOra.Fields("rain_kind").Value = isRAIN '"02"
    dynOra.Fields("ponpu").Value = isPump
    dynOra.Fields("river_no").Value = "85053002"
    dynOra.Fields("river_div").Value = "00"
    dynOra.Update
    dynOra.Close
    DoEvents
    Set dynOra = Nothing
    LOG_Out "OUT ORA_KANSOKU_JIKOKU_PUT �I��"
    On Error GoTo 0
    Exit Sub
ErrOracle:
    '******************************************************
    '��������G���[��������
    '******************************************************
    Dim strMessage As String
    If dbOra.LastServerErr <> 0 Then
        strMessage = dbOra.LastServerErrText                            'DB�����ɂ�����G���[
    Else
        strMessage = Err.Description                                    '�ʏ�̃G���[
    End If
    LOG_Out "###### ERROR ERROR ERROR ERROR ERROR ERROR ERROR ######"
    LOG_Out "IN ORA_KANSOKU_JIKOKU_PU " & strMessage
    LOG_Out "     SQL=" & SQL
    LOG_Out "###### ERROR ERROR ERROR ERROR ERROR ERROR ERROR ######"
    On Error GoTo 0
End Sub

'******************************************************************************
'�T�u���[�`���FORA_FRICS_RAIN()
'�����T�v
'FRICS�J�ʂ̏�������
'******************************************************************************
Sub ORA_FRICS_RAIN()
    Dim sql_SELECT   As String
    Dim sql_WHERE    As String
    Dim SQL          As String
    Dim N_rec        As Long
    Dim n            As Integer
    Dim i            As Long
    Dim j            As Long
    Dim k            As Long
    Dim Ndate        As String
    Dim Mdate        As String
    Dim Gdate        As String
    Dim TMAX         As String
    Dim HMAX         As Single
    Dim dw           As Date
    Dim nf           As Integer
    Dim buf          As String
    Dim K_CODE       As String
    Dim IYHK(18)     As String
    Dim IPYHK        As String
    Dim F_RAIN       As String
    Dim P_RAIN       As String
    Dim Point_Name   As String
    Dim Ryouiki_Name As String
    Dim r            As Long
    Dim rc           As Boolean
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/06 O.OKADA�y04�z�y05�z
    '��Oracle_DB�̃w�b�_�����Q�Ƃ��A�K�v�ӏ����C�����邱�ƁB
    '******************************************************
    Exit Sub
    '******************************************************
    'Ver1.0.0 �C���I�� 2015/08/06 O.OKADA�y04�z�y05�z
    '******************************************************
    Message.Label1 = "FRICS�\������J�ʂ��c�a�ɓo�^��"
    Message.ZOrder 0
    Message.Label1.Refresh
    LOG_Out "IN  ORA_FRICS_RAIN"
    MDB_FRICS���[�_�[�\��_For_HANS jgd, rc
    Ndate = Format(Now, "yyyy/mm/dd hh:nn")
    Mdate = Format(jgd, "yyyy/mm/dd hh:nn")
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/07 O.OKADA
    '���ŏI�I��hh24:mi:ss�ɖ߂����B
    '******************************************************
    Gdate = "'" & Format(jgd, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'Gdate = "'" & Format(jgd, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi" & "'"
    '******************************************************
    'Ver1.0.0 �C���I�� 2015/08/07 O.OKADA
    '******************************************************
    On Error GoTo ErrOracle:
    '******************************************************
    'SELECT
    '******************************************************
    sql_SELECT = "SELECT * FROM oracle.RYUIKI_10MIN_FRICS"
    '1002=������O����
    '1015=�V�쉺�V��F
    '1016=�厡
    '1017=�����O����
    '1019=�v�n��
    '1020=�t��
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    'For j = 1 To 5
    For j = 3 To 3
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
        Select Case j
            Case 1
                K_CODE = "231015"                                       '1015=�V�쉺�V��F
                Point_Name = "�V�쉺�V��F"
                Ryouiki_Name = "�V�쉺�V��F�㗬����"
            Case 2
                K_CODE = "231016"                                       '1016=�厡
                Point_Name = "�厡"
                Ryouiki_Name = "�厡�㗬����"
            Case 3
                K_CODE = "231017"                                       '1017=�����O����
                Point_Name = "�V�쐅���O����"
                Ryouiki_Name = "�V�쐅���O���ʑS"
            Case 4
                K_CODE = "231019"                                       '1019=�v�n��
                Point_Name = "�v�n��"
                Ryouiki_Name = "�v�n��㗬����"
            Case 5
                K_CODE = "231020"                                       '1020=�t��
                Point_Name = "�t��"
                Ryouiki_Name = "�t���㗬����"
        End Select
        '******************************************************
        'WHERE
        '******************************************************
        sql_WHERE = " WHERE  JIKOKU = TO_DATE(" & Gdate & ") " & _
                    "and RIVER_NUMBER='85053002' " & _
                    "and RIVER_KIND='00' " & _
                    "and GUN_NUMBER='01' " & _
                    "and POINT_NUMBER=" & K_CODE & " " & _
                    "and RYOUIKI_NUMBER='000'"
        SQL = sql_SELECT & sql_WHERE
        
        
        
        
        
        
'*************
'DB�o�͗}�~�ŁA�I���N�����b�Z�[�W�o�͗}�~�ŁiEXIT SUB����j�y�~�z
    Exit Sub
'*************

'*************
'DB�o�͎��{�A�I���N�����b�Z�[�W�o�͗}�~�ŁiEXIT SUB����j�y���z
    Exit Sub
'*************
        
        
        
        
        
        
        '******************************************************
        'SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����
        '******************************************************
        Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
        '******************************************************
        '�t�B�[���h�����擾����
        '******************************************************
        '******************************************************
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        '******************************************************
        'Dim Tw
        'n = dynOra.Fields.Count
        'For i = 0 To n - 1
        '    Tw = dynOra.Fields(i).Name
        '        Debug.Print " Number=" & Format(str(i), "@@@") & " �t�B�[���h��="; Tw
        'Next i
        '******************************************************
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        '******************************************************
        If dynOra.EOF Then
            dynOra.AddNew
        Else
            dynOra.Edit
        End If
        P_RAIN = Format(Round(RO(j, Now_Step) / 6 * 10), "000")
        For k = 1 To 18
            r = Round(RhY(j, k) * 10)
            F_RAIN = F_RAIN & Format(r, "000 ")
        Next k
        dynOra.Fields("WRITE_TIME").Value = Ndate                       '�������ݎ����|�|�|�|�|�|�|(DATE)
        dynOra.Fields("JIKOKU").Value = Mdate                           '���������|�|�|�|�|�|�|�|�|(DATE)
        dynOra.Fields("RIVER_NUMBER").Value = "85053002"                '�\��W��͐�ԍ��|�|�|�|�|(CHAR)
        dynOra.Fields("RIVER_KIND").Value = "00"                        '�\��W��͐�敪�ԍ��|�|�|(CHAR)
        dynOra.Fields("GUN_NUMBER").Value = "01"                        '���ʕW�|�|�|�|�|�|�|�|�|�|(VARCHAR2)
        dynOra.Fields("NOW_SEKISAN").Value = "010"                      '�����J�ʐώZ���ԁ|�|�|�|�|(CHAR)
        dynOra.Fields("YOSOKU_SEKISAN").Value = "010"                   '�\���J�ʐώZ���ԁ|�|�|�|�|(CHAR)
        dynOra.Fields("YOSOKU_SPAN").Value = "010"                      '�\�����ԊԊu�|�|�|�|�|�|�|(CHAR)
        dynOra.Fields("YOSOKU_COUNT").Value = "18"                      '�\�����Ԑ��|�|�|�|�|�|�|�|(CHAR)
        dynOra.Fields("TANI").Value = "01"                              '�J�ʒP�ʁ|�|�|�|�|�|�|�|�|(CHAR)
        dynOra.Fields("POINT_NUMBER").Value = K_CODE                    '���ʗ\���Ώۊϑ��n�_�ԍ��|(CHAR)
        dynOra.Fields("POINT_NAME").Value = Point_Name                  '�n�_���|�|�|�|�|�|�|�|�|�|(VARCHAR2)
        dynOra.Fields("RYOUIKI_NUMBER").Value = "000"                   '�̈搮���ԍ��|�|�|�|�|�|�|(CHAR)
        dynOra.Fields("RYOUIKI_NAME").Value = Ryouiki_Name              '�̈於�|�|�|�|�|�|�|�|�|�|(VARCHAR2)
        dynOra.Fields("P_RAIN").Value = P_RAIN                          '��������J�ʁ|�|�|�|�|�|�|(CHAR)
        dynOra.Fields("F_RAIN").Value = F_RAIN                          '�\������J�ʁ|�|�|�|�|�|�|(VARCHAR2)
        dynOra.Update
        LOG_Out "IN  ORA_FRICS_RAIN  SQL=" & SQL & vbCrLf & _
                "    P_RAIN=" & P_RAIN & vbCrLf & _
                "    F_RAIN=" & F_RAIN
        DoEvents
        dynOra.Close
        '******************************************************
        '�����J�ʂ̏�������
        '******************************************************
        If Minute(jgd) = 0 Then
            '******************************************************
            'SELECT
            '******************************************************
            sql_SELECT = "SELECT * FROM oracle.RYUIKI_1HOUR_FRICS"
            '******************************************************
            'WHERE
            '******************************************************
            sql_WHERE = " WHERE  JIKOKU = TO_DATE(" & Gdate & ") " & _
                        "and RIVER_NUMBER='85053002' " & _
                        "and RIVER_KIND='00' " & _
                        "and GUN_NUMBER='01' " & _
                        "and POINT_NUMBER=" & K_CODE & " " & _
                        "and RYOUIKI_NUMBER='000'"
            SQL = sql_SELECT & sql_WHERE
            '**************************************************
            'SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����
            '**************************************************
            Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
            If dynOra.EOF Then
                dynOra.AddNew
            Else
                dynOra.Edit
            End If
            P_RAIN = Format(Round(RO(j, Now_Step)), "000")
            F_RAIN = ""
            For k = 1 To 3
                r = Round(RO(j, Now_Step + k))
                F_RAIN = F_RAIN & Format(r, "000 ")
            Next k
            dynOra.Fields("WRITE_TIME").Value = Ndate                   '�������ݎ����|�|�|�|�|�|�|(DATE)
            dynOra.Fields("JIKOKU").Value = Mdate                       '���������|�|�|�|�|�|�|�|�|(DATE)
            dynOra.Fields("RIVER_NUMBER").Value = "85053002"            '�\��W��͐�ԍ��|�|�|�|�|(CHAR)
            dynOra.Fields("RIVER_KIND").Value = "00"                    '�\��W��͐�敪�ԍ��|�|�|(CHAR)
            dynOra.Fields("GUN_NUMBER").Value = "01"                    '���ʕW�|�|�|�|�|�|�|�|�|�|(VARCHAR2)
            dynOra.Fields("NOW_SEKISAN").Value = "060"                  '�����J�ʐώZ���ԁ|�|�|�|�|(CHAR)
            dynOra.Fields("YOSOKU_SEKISAN").Value = "060"               '�\���J�ʐώZ���ԁ|�|�|�|�|(CHAR)
            dynOra.Fields("YOSOKU_SPAN").Value = "060"                  '�\�����ԊԊu�|�|�|�|�|�|�|(CHAR)
            dynOra.Fields("YOSOKU_COUNT").Value = "03"                   '�\�����Ԑ��|�|�|�|�|�|�|�|(CHAR)
            dynOra.Fields("TANI").Value = "10"                          '�J�ʒP�ʁ|�|�|�|�|�|�|�|�|(CHAR)
            dynOra.Fields("POINT_NUMBER").Value = K_CODE                '���ʗ\���Ώۊϑ��n�_�ԍ��|(CHAR)
            dynOra.Fields("POINT_NAME").Value = Point_Name              '�n�_���|�|�|�|�|�|�|�|�|�|(VARCHAR2)
            dynOra.Fields("RYOUIKI_NUMBER").Value = "000"               '�̈搮���ԍ��|�|�|�|�|�|�|(CHAR)
            dynOra.Fields("RYOUIKI_NAME").Value = Ryouiki_Name          '�̈於�|�|�|�|�|�|�|�|�|�|(VARCHAR2)
            dynOra.Fields("P_RAIN").Value = P_RAIN                      '��������J�ʁ|�|�|�|�|�|�|(CHAR)
            dynOra.Fields("F_RAIN").Value = F_RAIN                      '�\������J�ʁ|�|�|�|�|�|�|(VARCHAR2)
            dynOra.Update
            LOG_Out "IN  ORA_FRICS_RAIN  SQL=" & SQL & vbCrLf & _
                    "    P_RAIN=" & P_RAIN & vbCrLf & _
                    "    F_RAIN=" & F_RAIN
            dynOra.Close
        End If
    Next j
    LOG_Out "OUT ORA_FRICS_RAIN"
    On Error GoTo 0
    Exit Sub
ErrOracle:
    '******************************************************
    '��������G���[��������
    '******************************************************
    Dim strMessage As String
    If dbOra.LastServerErr <> 0 Then
        strMessage = dbOra.LastServerErrText                            'DB�����ɂ�����G���[
    Else
        strMessage = Err.Description                                    '�ʏ�̃G���[
    End If
    LOG_Out "###### ERROR ERROR ERROR ERROR ERROR ERROR ERROR ######"
    LOG_Out "IN ORA_FRICS_RAIN " & strMessage
    LOG_Out "     SQL=" & SQL
    LOG_Out "###### ERROR ERROR ERROR ERROR ERROR ERROR ERROR ######"
    On Error GoTo 0
End Sub

'******************************************************************************
'�T�u���[�`���FORA_DataBase_Connection()
'�����T�v
'******************************************************************************
Sub ORA_DataBase_Connection()
    Dim msg   As String
    
    
    
    
    
'*************
'DB�o�͗}�~�ŁA�I���N�����b�Z�[�W�o�͗}�~�ŁiEXIT SUB����j�y�~�z
'    Exit Sub
'*************

'*************
'DB�o�͎��{�A�I���N�����b�Z�[�W�o�͗}�~�ŁiEXIT SUB���Ȃ��j�y���z
    
'*************
    
    
    
    
    
    '******************************************************
    'OO4O �� Oracle �ɐڑ�����
    '******************************************************
    On Error Resume Next
    '******************************************************
    '�Z�b�V�����̍쐬
    '******************************************************
'    Set ssOra = CreateObject("OracleInProcServer.XOraSession")
'    If Err <> 0 Then
'        msg = "���m���I���N���f�[�^�x�[�X�ɐڑ��o���܂���B" & Chr(10) & _
'               "CreateObject - Oracle oo4o �G���["
'        GoTo ERRHAND
'    End If
    '******************************************************
    '�T�[�r�X���i�T�[�o���j�� ���[�U��/�p�X���[�h ���w�肷��
    '******************************************************
    
    Dim strConfigFile As String
    Dim strProvider As String
    Dim strServer As String
    Dim strDBS As String
    Dim strUID As String
    Dim strPWD As String
    Dim strConn As String
    strConfigFile = App.Path
    If Right(strConfigFile, 1) <> "\" Then strConfigFile = strConfigFile & "\"
    strConfigFile = strConfigFile & "dbsinfo.cfg"
    If Len(Dir(strConfigFile, vbNormal)) < 1 Then
        msg = "���m���͐���V�X�e���f�[�^�x�[�X���u�̐ڑ����t�@�C��������܂���B"
        GoTo ERRHAND
    End If
    
    strProvider = GetConfigData("databases", "provider", strConfigFile)
    strServer = GetConfigData("databases", "server", strConfigFile)
    strDBS = GetConfigData("databases", "dbs", strConfigFile)
    strUID = GetConfigData("databases", "uid", strConfigFile)
    strPWD = GetConfigData("databases", "pwd", strConfigFile)
    If Len(strServer) < 1 Or Len(strDBS) < 1 Then
        msg = "���m���͐���V�X�e���f�[�^�x�[�X���u�̐ڑ���񂪂���܂���B"
        GoTo ERRHAND
    End If
    
    strConn = vbNullString
    If Len(strProvider) > 0 Then
        strConn = strConn & "Provider="
        strConn = strConn & strProvider
        strConn = strConn & ";"
    End If
    strConn = strConn & "Data Source="
    strConn = strConn & strServer
    strConn = strConn & ";"
    strConn = strConn & "Initial Catalog="
    strConn = strConn & strDBS
    strConn = strConn & ";"
    strConn = strConn & "User ID="
    strConn = strConn & strUID
    strConn = strConn & ";"
    strConn = strConn & "Password="
    strConn = strConn & strPWD
    strConn = strConn & ";"
    
'    Set dbOra = ssOra.OpenDatabase("ORACLE", "oracle/oracle", 0&)
    Set gAdoCon = New ADODB.Connection
    gAdoCon.ConnectionTimeout = 60
    gAdoCon.CommandTimeout = 60
    gAdoCon.Open strConn
    If Err.Number <> 0 Then
        msg = "���m���͐���V�X�e���f�[�^�x�[�X���u�ɐڑ��o���܂���B" & vbCrLf & _
              Err & ": " & Err.Description
        GoTo ERRHAND
    End If
    On Error GoTo 0
    OraDB_OK = True
    Exit Sub
ERRHAND:
'    Dim strMessage As String
'    If dbOra.LastServerErr <> 0 Then
'        strMessage = dbOra.LastServerErrText                            'DB�����ɂ�����G���[
'    Else
'        strMessage = Err.Description                                    '�ʏ�̃G���[
'    End If
    LOG_Out "###### ERROR ERROR ERROR ERROR ERROR ERROR ERROR ######"
    LOG_Out "IN ORA_DataBase_Connection " & msg
    LOG_Out "###### ERROR ERROR ERROR ERROR ERROR ERROR ERROR ######"
    OraDB_OK = False
    On Error GoTo 0
    Call ORA_DataBase_Close
End Sub

Public Sub SQLdbsDeleteRecordset(aobjAdorst As ADODB.Recordset)
    On Error Resume Next
    If Not (aobjAdorst Is Nothing) Then
        If aobjAdorst.State = adStateOpen Then aobjAdorst.Close
    End If
    Set aobjAdorst = Nothing
End Sub

'******************************************************************************
'�T�u���[�`���FORA_DataBase_Close()
'�����T�v
'******************************************************************************
Sub ORA_DataBase_Close()





'*************
'DB�o�͗}�~�ŁA�I���N�����b�Z�[�W�o�͗}�~�ŁiEXIT SUB����j�y�~�z
'    Exit Sub
'*************

'*************
'DB�o�͎��{�A�I���N�����b�Z�[�W�o�͗}�~�ŁiEXIT SUB���Ȃ��j�y���z
    
'*************





    '******************************************************
    'oo4o�̐ڑ�����������B
    '******************************************************
'    Set ssOra = Nothing
'    Set dbOra = Nothing
    On Error Resume Next
    If Not (gAdoCon Is Nothing) Then
        If gAdoCon.State = adStateOpen Then gAdoCon.Close
    End If
    Set gAdoCon = Nothing
End Sub

'******************************************************************************
'�T�u���[�`���FORA_Message_Out(Place As String, msg As String, Lebel As Long)
'�����T�v�F
'�v�Z�󋵂��c�a�ɏ�������
'******************************************************************************
Sub ORA_Message_Out(Place As String, msg As String, Lebel As Long)
    
    
    
    
    
'*************
'DB�o�͗}�~�ŁA�I���N�����b�Z�[�W�o�͗}�~�ŁiEXIT SUB����j�y�~�z
    Exit Sub
'*************

'*************
'DB�o�͎��{�A�I���N�����b�Z�[�W�o�͗}�~�ŁiEXIT SUB����j�y���z
    Exit Sub
'*************
    
    
    
    
    
    Dim i        As Long
    Dim SQL      As String
    Dim WHERE    As String
    Dim code(2)  As String
    Dim Ndate    As String
    Dim dw       As Date
    code(1) = "1"                                                       '���v�Z�l
    code(2) = "2"                                                       '�v�Z�s��
    LOG_Out "IN  ORA_Message_Out"
    LOG_Out "    msg=" & msg
    If msg = "" Then
        Exit Sub
    End If
    If DBX_ora = False Then
        '�I���N���T�[�o�[�ɃA�b�v���Ȃ�
        Exit Sub
    End If
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    'Exit Sub                                                           '���}���u
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    ORA_DataBase_Connection
    If OraDB_OK = False Then
        '******************************************************
        '�I���N���T�[�o�[�ɐڑ��ł��Ȃ������̂Ń��^�[������B
        '******************************************************
        LOG_Out "IN  ORA_Message_Out  �I���N���T�[�o�[�ɐڑ��ł��Ȃ������̂Ń��^�[������B"
        Exit Sub
    End If
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    'SQL = "SELECT * FROM oracle.CAL_MESSAGE"
    'Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    '******************************************************
    '�t�B�[���h�����擾����
    '******************************************************
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    'Dim Tw, n
    'n = dynOra.Fields.Count
    'For i = 0 To n - 1
    '    Tw = dynOra.Fields(i).name
    '    Debug.Print " Number=" & Format(str(i), "@@@") & " �t�B�[���h��="; Tw
    'Next i
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    On Error GoTo ErrOracle
    Obs_Time = Obs_Time + 1
    If Obs_Time >= 90 Then
        LOG_Out "    Obs_Time��90�ȏ�ɂȂ�܂����̂Ń��b�Z�[�W�o�͂�������߂܂�"
        Exit Sub
    End If
    dw = DateAdd("s", Obs_Time, jgd)
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/07 O.OKADA
    '���ŏI�I��hh24:mi:ss�ɖ߂����B
    '******************************************************
    Ndate = "'" & Format(dw, "yyyy/mm/dd hh:nn:ss") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'Ndate = "'" & Format(dw, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi" & "'"
    '******************************************************
    'Ver1.0.0 �C���I�� 2015/08/07 O.OKADA
    '******************************************************
    Message.Show
    Message.Label1 = "���b�Z�[�W��DB�A�b�v��"
    Message.ZOrder 0
    Message.Label1.Refresh
    SQL = "SELECT * FROM oracle.CAL_MESSAGE WHERE jikoku= TO_DATE(" & Ndate & ") "
    LOG_Out "    SQL=" & SQL
    '******************************************************
    'SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����
    '******************************************************
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    If dynOra.EOF Then
        dynOra.AddNew
    Else
        dynOra.Edit
    End If
    '******************************************************
    '�f�[�^��������
    '******************************************************
    dynOra.Fields("WRITE_TIME").Value = Format(Now, "yyyy/mm/dd hh:nn") '�������ݎ���
    dynOra.Fields("jikoku").Value = Format(dw, "yyyy/mm/dd hh:nn:ss")
    dynOra.Fields("river_no").Value = "85053002"
    dynOra.Fields("RAIN_KIND").Value = isRAIN
    dynOra.Fields("error_area").Value = Place                           '��Q��
    dynOra.Fields("message").Value = msg
    dynOra.Fields("cal_level").Value = 1                                'Lebel
    dynOra.Update
    On Error GoTo 0
    ORA_DataBase_Close
    Exit Sub
ErrOracle:
    '******************************************************
    '��������G���[��������
    '******************************************************
    Dim strMessage As String
    If dbOra.LastServerErr <> 0 Then
        strMessage = dbOra.LastServerErrText                            'DB�����ɂ�����G���[
    Else
        strMessage = Err.Description                                    '�ʏ�̃G���[
    End If
    LOG_Out "###### ERROR ERROR ERROR ERROR ERROR ERROR ERROR ######"
    LOG_Out "IN ORA_Message_Out " & strMessage
    LOG_Out "     SQL=" & SQL
    LOG_Out "###### ERROR ERROR ERROR ERROR ERROR ERROR ERROR ######"
    On Error GoTo 0
End Sub

'******************************************************************************
'�T�u���[�`���FORA_SUII_YOSOKU_KIJYUN_PUT(Return_Code As Boolean)
'�����T�v�F
'�\�����ʂ̏�������
'******************************************************************************
Sub ORA_SUII_YOSOKU_KIJYUN_PUT(Return_Code As Boolean)
    Dim sql_SELECT   As String
    Dim sql_WHERE    As String
    Dim SQL          As String
'    Dim N_rec        As Long
'    Dim n            As Integer
    Dim i            As Long
    Dim j            As Long
    Dim Ndate        As String
    Dim Mdate        As String
    Dim Gdate        As String
    Dim TMAX         As String
    Dim HMAX         As Single
    Dim dw           As Date
'    Dim nf           As Integer
'    Dim buf          As String
    Dim K_CODE       As String
    Dim IYHK(18)     As String
    Dim IPYHK        As String
    Dim strDebugData As String
    Message.Label1 = "�\�����ʂ��c�a�ɓo�^��"
    Message.ZOrder 0
    Message.Label1.Refresh
    
    
    
    
    
'*************
'DB�o�͗}�~�ŁA�I���N�����b�Z�[�W�o�͗}�~�ŁiEXIT SUB����j�y�~�z
'    Exit Sub
'*************

'*************
'DB�o�͎��{�A�I���N�����b�Z�[�W�o�͗}�~�ŁiEXIT SUB���Ȃ��j�y���z

'*************
    
    
    
    
    
    
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    'nf = FreeFile
    'Open App.Path & "\Data\SUII_YOSOKU_KIJYUN.DAT" For Output As #nf
    'jgd = Now                                                          '�e�X�g�p
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    Ndate = Format(Now, "yyyy/mm/dd hh:nn")
    Mdate = Format(jgd, "yyyy/mm/dd hh:nn")
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/07 O.OKADA
    '���ŏI�I��hh24:mi:ss�ɖ߂����B
    '******************************************************
    Gdate = "'" & Format(jgd, "yyyy/mm/dd hh:nn") & "'" '," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'Gdate = "'" & Format(jgd, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi" & "'"
    '******************************************************
    'Ver1.0.0 �C���I�� 2015/08/07 O.OKADA
    '******************************************************
    '******************************************************
    'SELECT
    '******************************************************
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/07 O.OKADA
    '******************************************************
    'sql_SELECT = "SELECT * FROM oracle.SUII_YOSOKU_KIJYUN"
    '******************************************************
    'Ver1.0.0 �C���I�� 2015/08/07 O.OKADA
    '******************************************************
    On Error GoTo ErrOracle:
    '1002=������O����
    '1015=�V�쉺�V��F
    '1016=�厡
    '1017=�����O����
    '1019=�v�n��
    '1020=�t��
    gAdoCon.BeginTrans
    For j = 1 To 5
        Select Case j
            Case 1
                K_CODE = "81"                                           '1015=�V�쉺�V��F
            Case 2
                K_CODE = "201"                                          '1016=�厡
            Case 3
                K_CODE = "91"                                           '1017=�����O����
            Case 4
                K_CODE = "71"                                                                                                                         '1019=�v�n��
            Case 5
                K_CODE = "131"                                          '1020=�t��
        End Select
        strDebugData = ",�ϑ�����;" & Mdate
        strDebugData = strDebugData & ",�ϑ���ID;" & K_CODE
        '******************************************************
        'WHERE
        '******************************************************
'        sql_WHERE = " WHERE  JIKOKU = TO_DATE(" & Gdate & ") and POINT_CODE='" & K_CODE & "' " & _
'                    " AND RAIN_KIND='" & isRAIN & "' AND RIVER_NO ='85053002' AND PONPU='" & isPump & "'"
        sql_WHERE = " WHERE obs_time=" & Gdate & " AND obs_sta_id=" & K_CODE
        sql_SELECT = "SELECT * FROM t_cast_water_level_data"
        SQL = sql_SELECT & sql_WHERE
        '******************************************************
        'SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����
        '******************************************************
'        Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
        Set gAdoRst = New ADODB.Recordset
        gAdoRst.CursorType = adOpenStatic
        gAdoRst.Open SQL, gAdoCon, , adLockPessimistic, adCmdText
        '******************************************************
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        '******************************************************
        'SQL = sql_SELECT
        'Set RST_YB.ActiveConnection = Conn_DB
        'RST_YB.Open SQL, Conn_DB, adOpenDynamic
        '******************************************************
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        '******************************************************
        '******************************************************
        '�t�B�[���h�����擾����
        '******************************************************
        '******************************************************
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        '******************************************************
        'Dim Tw
        'n = dynOra.Fields.Count
        'For i = 0 To n - 1
        '    Tw = dynOra.Fields(i).Name
        '    Debug.Print " Number=" & Format(str(i), "@@@") & " �t�B�[���h��="; Tw
        'Next i
        '******************************************************
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        '******************************************************
        If gAdoRst.EOF Then
            gAdoRst.AddNew
            gAdoRst.Fields("cre_time").Value = Ndate
'        Else
'            dynOra.Edit
        End If
        HMAX = -9999#
        dw = DateAdd("n", 10, jgd)
        For i = 1 To 18
            IYHK(i) = Format(YHK(j, i) * 100#, "00000")
            strDebugData = strDebugData & "," & CStr(i * 10) & "���㐅��;" & IYHK(i)
            If YHK(j, i) > HMAX Then
                HMAX = YHK(j, i)
                TMAX = Format(dw, "yyyy/mm/dd hh:nn")
            End If
            dw = DateAdd("n", 10, dw)
        Next i
        IPYHK = F_2_I(HMAX)
        LOG_Out strDebugData
        gAdoRst.Fields("upd_time").Value = Ndate                       '�������ݎ���
        gAdoRst.Fields("obs_time").Value = Mdate
'        gAdoRst.Fields("RIVER_NO").Value = "85053002"
'        gAdoRst.Fields("RIVER_DIV").Value = "00"
        gAdoRst.Fields("obs_sta_id").Value = K_CODE
        gAdoRst.Fields("cast_data_10").Value = F_2_I(YHK(j, 1))
        gAdoRst.Fields("cast_data_20").Value = F_2_I(YHK(j, 2))
        gAdoRst.Fields("cast_data_30").Value = F_2_I(YHK(j, 3))
        gAdoRst.Fields("cast_data_40").Value = F_2_I(YHK(j, 4))
        gAdoRst.Fields("cast_data_50").Value = F_2_I(YHK(j, 5))
        gAdoRst.Fields("cast_data_60").Value = F_2_I(YHK(j, 6))
        gAdoRst.Fields("cast_data_70").Value = F_2_I(YHK(j, 7))
        gAdoRst.Fields("cast_data_80").Value = F_2_I(YHK(j, 8))
        gAdoRst.Fields("cast_data_90").Value = F_2_I(YHK(j, 9))
        gAdoRst.Fields("cast_data_100").Value = F_2_I(YHK(j, 10))
        gAdoRst.Fields("cast_data_110").Value = F_2_I(YHK(j, 11))
        gAdoRst.Fields("cast_data_120").Value = F_2_I(YHK(j, 12))
        gAdoRst.Fields("cast_data_130").Value = F_2_I(YHK(j, 13))
        gAdoRst.Fields("cast_data_140").Value = F_2_I(YHK(j, 14))
        gAdoRst.Fields("cast_data_150").Value = F_2_I(YHK(j, 15))
        gAdoRst.Fields("cast_data_160").Value = F_2_I(YHK(j, 16))
        gAdoRst.Fields("cast_data_170").Value = F_2_I(YHK(j, 17))
        gAdoRst.Fields("cast_data_180").Value = F_2_I(YHK(j, 18))
        gAdoRst.Fields("peak_time").Value = TMAX
        gAdoRst.Fields("peak_cast_data").Value = IPYHK
        gAdoRst.Fields("rain_kind").Value = isRAIN
        gAdoRst.Fields("pump_status").Value = isPump
        gAdoRst.Update
        Call SQLdbsDeleteRecordset(gAdoRst)
        sql_SELECT = "SELECT * FROM t_cast_water_level_status"
        sql_WHERE = " WHERE obs_sta_id=" & K_CODE
        SQL = sql_SELECT & sql_WHERE
        Set gAdoRst = New ADODB.Recordset
        gAdoRst.CursorType = adOpenStatic
        gAdoRst.Open SQL, gAdoCon, , adLockPessimistic, adCmdText
        If gAdoRst.EOF Then
            gAdoRst.AddNew
            gAdoRst.Fields("obs_sta_id").Value = K_CODE
            gAdoRst.Fields("cre_time").Value = Ndate
        End If
        gAdoRst.Fields("latest_obs_time").Value = Mdate
        gAdoRst.Fields("upd_time").Value = Ndate
        gAdoRst.Update
        Call SQLdbsDeleteRecordset(gAdoRst)
    Next j
'    dynOra.Close
    gAdoCon.CommitTrans
    DoEvents
'    Close #nf
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    'Set dynOra = Nothing
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    On Error GoTo 0
'    ORA_KANSOKU_JIKOKU_PUT "SUII_YOSOKU_KIJYUN", jgd
    Call SQLdbsDeleteRecordset(gAdoRst)
    Exit Sub
ErrOracle:
    '******************************************************
    '��������G���[��������
    '******************************************************
    Dim strMessage As String
'    If dbOra.LastServerErr <> 0 Then
'        strMessage = dbOra.LastServerErrText                            'DB�����ɂ�����G���[
'    Else
        strMessage = Err.Description                                    '�ʏ�̃G���[
'    End If
    LOG_Out "###### ERROR ERROR ERROR ERROR ERROR ERROR ERROR ######"
    LOG_Out "IN ORA_SUII_YOSOKU_KIJYUN_PUT " & strMessage
    LOG_Out "     SQL=" & SQL
    LOG_Out "###### ERROR ERROR ERROR ERROR ERROR ERROR ERROR ######"
    On Error GoTo 0
    Call SQLdbsDeleteRecordset(gAdoRst)
    gAdoCon.RollbackTrans
End Sub

'******************************************************************************
'�T�u���[�`���FORA_YOHOUBUNAN()
'�����T�v�F
'���m���T�[�o�[�ɗ\�񕶂���������
'******************************************************************************
Sub ORA_YOHOUBUNAN()
    Dim sql_SELECT   As String
    Dim sql_WHERE    As String
    Dim SQL          As String
    Dim n            As Integer
    Dim i            As Long
    Dim SDATE        As String
    Dim EDATE        As String
    Dim jssd         As Date
    Dim jeed         As Date
    Dim C2           As String
    Dim C3           As String
    Dim C4           As String
    Dim C5           As String
    jssd = jgd
    jeed = DateAdd("n", 30, jssd)
    C2 = Messag(Pattan_Now).Patn(2)                                     'FORCAST_KIND_CODE    50
    C3 = Messag(Pattan_Now).Patn(16)                                    'FORCAST_KIND         �͂񗔌x�񔭕\
    C4 = Format(jssd, "yyyy/mm/dd hh:nn")                               'ESTIMATE_TIME
    C5 = Format(jeed, "yyyy/mm/dd hh:nn")                               'ANNOUNCE_TIME
    
    
    
    
    
'*************
'DB�o�͗}�~�ŁA�I���N�����b�Z�[�W�o�͗}�~�ŁiEXIT SUB����j�y�~�z
    Exit Sub
'*************

'*************
'DB�o�͎��{�A�I���N�����b�Z�[�W�o�͗}�~�ŁiEXIT SUB����j�y���z
    Exit Sub
'*************
    
    
    
    
    
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/07 O.OKADA
    '���ŏI�I��hh24:mi:ss�ɖ߂����B
    '******************************************************
    SDATE = "'" & Format(jssd, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    EDATE = "'" & Format(jeed, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'SDATE = "'" & Format(jssd, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi" & "'"
    'EDATE = "'" & Format(jeed, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi" & "'"
    '******************************************************
    'Ver1.0.0 �C���I�� 2015/08/07 O.OKADA
    '******************************************************
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/06 O.OKADA�y01�z
    '��Oracle_DB�̃w�b�_�����Q�Ƃ��A�K�v�ӏ����C�����邱�ƁB
    '******************************************************
    Exit Sub
    '******************************************************
    'Ver1.0.0 �C���I�� 2015/08/06 O.OKADA�y01�z
    '******************************************************
    '******************************************************
    'SELECT
    '******************************************************
    sql_SELECT = "SELECT * FROM oracle.C_FUKENKOZUIAN"
    '******************************************************
    'WHERE
    '******************************************************
    sql_WHERE = " WHERE  ESTIMATE_TIME = TO_DATE(" & SDATE & ") AND" & _
                " DATA_KIND_CODE = '�t�P���R�E�Y�C�A��01' AND" & _
                " SENDING_STATION_CODE ='23001' AND" & _
                " RAIN_KIND = " & isRAIN
    SQL = sql_SELECT & sql_WHERE
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    'SQL = sql_SELECT
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    '******************************************************
    '�t�B�[���h�����擾����
    '******************************************************
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    'Dim Tw
    'n = RST_YB.Fields.Count
    'For i = 0 To n - 1
    '    Tw = RST_YB.Fields(i).Name
    '    Debug.Print " Number=" & Format(Str(i), "@@@") & " �t�B�[���h��="; Tw
    'Next i
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    '******************************************************
    'SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����
    '******************************************************
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    Dim buf As String
    If dynOra.EOF Then
        dynOra.AddNew
    Else
        dynOra.Edit
    End If
    dynOra.Fields("WRITE_TIME").Value = TIMEC(Now)                      '�������ݎ���
    dynOra.Fields("DATA_KIND_CODE").Value = "�t�P���R�E�Y�C�A��01"
    dynOra.Fields("DATA_KIND").Value = "�\�񕶈āi���ʕ����j"
    dynOra.Fields("SENDING_STATION_CODE").Value = "23001"
    dynOra.Fields("SENDING_STATION").Value = "���m���������ݎ�����"
    dynOra.Fields("APPOINTED_CODE").Value = ""
    dynOra.Fields("ESTIMATE_TIME").Value = C4
    If AutoDrive.Option2(1).Value Then
        dynOra.Fields("PRACTICE_FLG_CODE").Value = "99"                 '"40"=�\��  "99"=���K
        dynOra.Fields("PRACTICE_FLG").Value = "���K"                    '"���K""�\��"
    End If
    If AutoDrive.Option2(0).Value Then
        dynOra.Fields("PRACTICE_FLG_CODE").Value = "40"                 '"40"=�\��  "99"=���K
        dynOra.Fields("PRACTICE_FLG").Value = "�\��"                    '"���K""�\��"
    End If
    dynOra.Fields("SEQ_NO").Value = ""
    dynOra.Fields("ANNOUNCE_TIME").Value = C5
    dynOra.Fields("RIVER_NAME").Value = "���m�������쐅�n�@�V��"
    dynOra.Fields("RIVER_NO_CODE").Value = "85053002"
    dynOra.Fields("RIVER_NO").Value = "�V��"
    dynOra.Fields("RIVER_DIV_CODE").Value = "00"
    dynOra.Fields("RIVER_DIV").Value = ""
    dynOra.Fields("ANNOUNCE_NO").Value = ""
    '******************************************************
    '�V�K�ǉ��ƌ��������e�[�u���yYOHOU_TARGET_RIVER�i�\�񕶑Ώۉ͐�j�z�������Ɉړ��B
    '******************************************************
    dynOra.Fields("TARGET_RIVER_COUNT").Value = 1
    dynOra.Fields("TARGET_RIVER_NAME").Value = "�V��"
    dynOra.Fields("TARGET_RIVER_NO_CODE").Value = "85053002"
    dynOra.Fields("TARGET_RIVER_NO").Value = "�V��"
    dynOra.Fields("TARGET_RIVER_DIV_CODE").Value = "00"
    '******************************************************
    'dynOra.Fields("TARGET_RIVER_DIV").Value = 1
    '******************************************************
    '�܂�
    dynOra.Fields("FORECAST_KIND").Value = C2
    dynOra.Fields("FORECAST_KIND_CODE").Value = C3
    dynOra.Fields("HEADLINE").Value = Messag(Pattan_Now).Patn(3)        '���o��
    dynOra.Fields("BUNSHO1").Value = �啶1
    dynOra.Fields("BUNSHO2").Value = �啶2
    dynOra.Fields("BUNSHO3").Value = ""
    dynOra.Fields("RAIN_KIND").Value = isRAIN
    dynOra.Update
    dynOra.Close
    DoEvents
    Set dynOra = Nothing
End Sub

'******************************************************************************
'�T�u���[�`���FORA_SUII_YOSOKU_JYUDAN_PUT(Return_Code As Boolean)
'�����T�v�F
'�\���c�f���ʂ̏�������
'���ǃo�[�W���� 2003.03.25
'******************************************************************************
Sub ORA_SUII_YOSOKU_JYUDAN_PUT(Return_Code As Boolean)
    Dim sql_SELECT   As String
    Dim sql_WHERE    As String
    Dim SQL          As String
'    Dim N_rec        As Long
'    Dim n            As Integer
    Dim m            As Long
    Dim i            As Long
'    Dim j            As Long
    Dim Ndate        As String
    Dim Mdate        As String
    Dim Odate        As String
    Dim Gdate        As String
    Dim Ydate        As String
'    Dim TMAX         As String
'    Dim HMAX         As Single
    Dim dw           As Date
'    Dim nf           As Integer
'    Dim buf          As String
'    Dim P_CODE       As String
'    Dim up(53)       As Boolean
    Dim Jyudan(4)    As String
    Dim w            As String
    Dim strDebugData As String
    
    
    
    
    
'*************
'DB�o�͗}�~�ŁA�I���N�����b�Z�[�W�o�͗}�~�ŁiEXIT SUB����j�y�~�z
'    Exit Sub
'*************

'*************
'DB�o�͎��{�A�I���N�����b�Z�[�W�o�͗}�~�ŁiEXIT SUB���Ȃ��j�y���z

'*************
    
    
    
    
    
    Const RIVER_NO As Long = 1
    Message.Label1 = "�\�����ʏc�f���c�a�ɓo�^��"
    Message.ZOrder 0
    Message.Refresh
    LOG_Out " ���ʏc�f�f�[�^�c�a�������݊J�n"
    Jyudan_Hosei
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    'nf = FreeFile
    'Open App.Path & "\Data\SUII_YOSOKU_JYUDAN.DAT" For Output As #nf
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/07 O.OKADA
    '���ŏI�I�Ɍ���hh24:mi:ss�ɖ߂����B
    '******************************************************
    Ndate = Format(Now, "yyyy/mm/dd hh:nn")
    Mdate = "'" & Format(jgd, "yyyy/mm/dd hh:nn") & "'" '," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    Gdate = Format(jgd, "yyyy/mm/dd hh:nn")
    'Ndate = Format(Now, "yyyy/mm/dd hh:nn")
    'Mdate = "'" & Format(jgd, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi" & "'"
    'Gdate = Format(jgd, "yyyy/mm/dd hh:nn")
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/07 O.OKADA
    '******************************************************
    dw = DateAdd("h", 1, jgd)
    On Error GoTo ErrOracle:
    '******************************************************
    'SELECT
    '******************************************************
'    sql_SELECT = "SELECT * FROM oracle.SUII_YOSOKU_JYUDAN "
    sql_SELECT = "SELECT * FROM t_cast_water_level_vertical_data"
    For m = 1 To 4
        w = ""
        For i = 1 To 53
            w = w & Format(YHJ(m - 1, i) * 100#, "00000") & ","
        Next i
        Jyudan(m) = w
    Next m
    strDebugData = ",�ϑ�����;" & Gdate
    strDebugData = strDebugData & ",�͐�ID;" & RIVER_NO
    LOG_Out strDebugData
    strDebugData = ",�c�f�f�[�^(��������)," & Jyudan(1)
    LOG_Out strDebugData
    strDebugData = ",�c�f�f�[�^(60���㐅��)," & Jyudan(2)
    LOG_Out strDebugData
    strDebugData = ",�c�f�f�[�^(120���㐅��)," & Jyudan(3)
    LOG_Out strDebugData
    strDebugData = ",�c�f�f�[�^(180���㐅��)," & Jyudan(4)
    LOG_Out strDebugData
    dw = DateAdd("h", m, jgd)
    Ydate = Format(dw, "yyyy/mm/dd hh:nn")
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/07 O.OKADA
    '���ŏI�I�Ɍ���hh24:mi:ss�ɖ߂����B
    '******************************************************
    Odate = "'" & Format(dw, "yyyy/mm/dd hh:nn") & "'" '," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'Odate = "'" & Format(dw, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi" & "'"
    '******************************************************
    'Ver1.0.0 �C���I�� 2015/08/07 O.OKADA
    '******************************************************
    '******************************************************
    'WHERE
    '******************************************************
'    sql_WHERE = " WHERE RIVER_NO = 85053002 AND RIVER_DIV = '00'" & _
'                " AND JIKOKU = TO_DATE(" & Mdate & ")" & _
'                " AND RAIN_KIND='" & isRAIN & "'" & _
'                " AND PONPU='" & isPump & "'" & _
'                " AND YOSOKU_TIME = TO_DATE(" & Odate & ")"
    sql_WHERE = " WHERE obs_time=" & Mdate & _
                " AND river_id=" & RIVER_NO
    SQL = sql_SELECT & sql_WHERE
    '******************************************************
    'SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����
    '******************************************************
'    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    gAdoCon.BeginTrans
    Set gAdoRst = New ADODB.Recordset
    gAdoRst.CursorType = adOpenStatic
    gAdoRst.Open SQL, gAdoCon, , adLockPessimistic, adCmdText
    LOG_Out " SQL=" & SQL
    DoEvents
    Message.Label1 = "���ʏc�f�f�[�^ �o�^"
    Message.ZOrder 0
    Message.Refresh
    If gAdoRst.EOF Then
        LOG_Out " ���ʏc�f�f�[�^�c�a�������݊J�n   �ǉ����[�h"
        gAdoRst.AddNew
'        dynOra.Fields("WRITE_TIME").Value = Ndate                       '�������ݎ���
        gAdoRst.Fields("obs_time").Value = Gdate
        gAdoRst.Fields("river_id").Value = RIVER_NO
'        dynOra.Fields("RIVER_DIV").Value = "00"
        '******************************************************
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        '******************************************************
        'dynOra.Fields("YOSOKU_TIME").Value = Ydate
'        dynOra.Fields("YOSOKU_TIME").Value = Ydate
        '******************************************************
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        '******************************************************
'        dynOra.Fields("COUNT").Value = 53
'        dynOra.Fields("SUII0").Value = Jyudan(1)
'        dynOra.Fields("SUII1").Value = Jyudan(2)
'        dynOra.Fields("SUII2").Value = Jyudan(3)
'        dynOra.Fields("SUII3").Value = Jyudan(4)
'        dynOra.Fields("RAIN_KIND").Value = isRAIN
'        dynOra.Fields("PONPU").Value = isPump
        gAdoRst.Fields("cre_time").Value = Ndate
'        dynOra.Update
    Else
        LOG_Out " ���ʏc�f�f�[�^�c�a�������݊J�n   �㏑�����[�h"
'        dynOra.Edit
'        dynOra.Fields("WRITE_TIME").Value = Ndate                       '�������ݎ���
'        dynOra.Fields("SUII0").Value = Jyudan(1)
'        dynOra.Fields("SUII1").Value = Jyudan(2)
'        dynOra.Fields("SUII2").Value = Jyudan(3)
'        dynOra.Fields("SUII3").Value = Jyudan(4)
'        dynOra.Update
    End If
    gAdoRst.Fields("cast_time").Value = Ydate
    gAdoRst.Fields("spot_count").Value = 53
    gAdoRst.Fields("now_data_10").Value = Jyudan(1)
    gAdoRst.Fields("cast_data_60").Value = Jyudan(2)
    gAdoRst.Fields("cast_data_120").Value = Jyudan(3)
    gAdoRst.Fields("cast_data_180").Value = Jyudan(4)
    gAdoRst.Fields("rain_kind").Value = isRAIN
    gAdoRst.Fields("pump_status").Value = isPump
    gAdoRst.Fields("upd_time").Value = Ndate
    gAdoRst.Update
    Call SQLdbsDeleteRecordset(gAdoRst)
    sql_SELECT = "SELECT * FROM t_cast_water_level_vertical_status"
    sql_WHERE = " WHERE river_id=" & RIVER_NO
    SQL = sql_SELECT & sql_WHERE
    Set gAdoRst = New ADODB.Recordset
    gAdoRst.CursorType = adOpenStatic
    gAdoRst.Open SQL, gAdoCon, , adLockPessimistic, adCmdText
    If gAdoRst.EOF Then
        gAdoRst.AddNew
        gAdoRst.Fields("river_id").Value = RIVER_NO
        gAdoRst.Fields("cre_time").Value = Ndate
    End If
    gAdoRst.Fields("latest_obs_time").Value = Gdate
    gAdoRst.Fields("upd_time").Value = Ndate
    gAdoRst.Update
    Call SQLdbsDeleteRecordset(gAdoRst)
    gAdoCon.CommitTrans
    DoEvents
'    dynOra.Close
    LOG_Out " ���ʏc�f�f�[�^�c�a�������ݏI��"
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    'Close #nf
    'Set dynOra = Nothing
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    On Error GoTo 0
'    ORA_KANSOKU_JIKOKU_PUT "SUII_YOSOKU_JYUDAN", jgd
    Message.Hide
    Exit Sub
ErrOracle:
    '******************************************************
    '��������G���[��������
    '******************************************************
    Dim strMessage As String
'    If dbOra.LastServerErr <> 0 Then
'        strMessage = dbOra.LastServerErrText                            'DB�����ɂ�����G���[
'    Else
        strMessage = Err.Description                                    '�ʏ�̃G���[
'    End If
    LOG_Out "###### ERROR ERROR ERROR ERROR ERROR ERROR ERROR ######"
    LOG_Out "IN ORA_SUII_YOSOKU_JYUDAN_PUT " & strMessage
    LOG_Out "     SQL=" & SQL
    LOG_Out "###### ERROR ERROR ERROR ERROR ERROR ERROR ERROR ######"
    On Error GoTo 0
    Call SQLdbsDeleteRecordset(gAdoRst)
    On Error Resume Next
    gAdoCon.RollbackTrans
    On Error GoTo 0
End Sub
