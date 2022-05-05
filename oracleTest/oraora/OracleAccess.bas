Attribute VB_Name = "OracleAccess"
'******************************************************************************
'���W���[�����FOracleAccess
'
'�C�������F
'2015/08/04�F�I���N���f�[�^�x�[�X�̃e�[�u���폜�ɑΉ����A���L�e�[�u���ւ�
'�A�N�Z�X�����������R�����g�A�E�g����B
'ARAIZEKI           �F�􉁐��ʁ^�z����
'F_MESSYU_10MIN_1   �F�P�O���ԍ~���ʗ\�����b�V���l�i�\���P���ԁj
'F_MESSYU_10MIN_2   �F�P�O���ԍ~���ʗ\�����b�V���l�i�\���U���ԁj
'F_RADAR            �FFRICS�\�����[�_�J�ʏ��
'P_MESSYU_10MIN     �F�P�O���ԍ~���ʎ������b�V���l
'P_RADAR            �FFRICS�������[�_�[�J�ʏ��
'YOHOU_TARGET_RIVER �F�\�񕶑Ώۉ͐�
'YOHOUBUN           �F�\�񕶍쐬����
'YOHOUBUNAN         �F�\�񕶈č쐬����
'
'OracleAccess.ORA_Araizeki()���C�������B�y01�z
'��OracleDB.Check_Araizeki_Time()���C�����邱�ƁB�y01-01�z
'
'OracleAccess.Dump_F_MESSYU_10MIN_1()���C�������B�y02�z
'��OracleAccess.ORA_F_MESSYU_10MIN_2()���C�����邱�ƁB�y02-01�z
'��OracleAccess.ORA_F_MESSYU_10MIN_1()���C�����邱�ƁB�y02-02�z
'
'OracleAccess.ORA_F_MESSYU_10MIN_1()���C�������B�y03�z
'��OracleAccess.ORA_F_MESSYU_10MIN_1()���C�����邱�ƁB�y03-01�z
'��OracleDB.Check_F_MESSYU_10MIN_1_Time()���C�����邱�ƁB�y03-02�z
'
'OracleAccess.ORA_F_MESSYU_10MIN_2()���C�������B�y04�z
'��OracleDB.Check_F_MESSYU_10MIN_2_Time()���C�����邱�ƁB�y04-01�z
'
'OracleAccess.ORA_F_MESSYU_10MIN_20()���C�������B�y05�z
'�����̏C���ɔ����e���͂Ȃ��B
'
'OracleAccess.ORA_F_RADAR()���C�������B�y06�z
'��OracleDB.Check_F_RADAR_Time()���C�����邱�Ɓy06-01�z�B
'
'OracleAccess.ORA_P_MESSYU_10MIN()���C�������B�y07�z
'��OracleDB.Check_P_MESSYU_10MIN_Time()���C�����邱�ƁB�y07-01�z
'
'OracleAccess.ORA_P_MESSYU_1Hour()���C�������B�y08�z
'��OracleDB.Check_P_MESSYU_1HOUR_Time()���C�����邱�ƁB�y08-01�z
'
'OracleAccess.ORA_P_RADAR()���C�������B�y09�z
'��OracleDB.Check_P_RADAR_Time()���C�����邱�ƁB�y09-01�z
'
'OracleAccess.ORA_YOHOUBUNAN()���C�������B�y10�z
'���\�񕶃e�X�g���M.Command1_Click()���C�����邱�ƁB�y10-01�z
'
'******************************************************************************
Option Explicit
Option Base 1

'******************************************************************************
'���̑��̃O���[�o���ϐ����Z�b�g����B
'******************************************************************************
Global jgd                As Date           '�{�Ԃł̓R�����g�ɂ��邱��
Global LOG_N              As Integer        '���O�o�͗p�ԍ�
Global LOG_File           As String         '���O�o�͗p�t�@�C����
Global Dmp_N              As Integer        '�f�[�^�_���v�t�@�C���ԍ�
'******************************************************************************
'OO4O�֘A�̃O���[�o���ϐ����Z�b�g����B
'******************************************************************************
'Global ssOra              As Object         '
Global dbOra              As Object ' OraDatabase    '
Global dynOra             As Object ' OraDynaset     '
Public gAdoCon As ADODB.Connection
Public gAdoRst As ADODB.Recordset
Global gbool_Start_Set    As Boolean        '�f�[�^��荞�ݒ���True
Global gdate_oraTims      As Date           '�I���N���f�[�^��荞�݊J�n����
Global gdate_oraTime      As Date           '�I���N���f�[�^��荞�ݏI������
'******************************************************************************
'�\�񕶊֌W�̃O���[�o���ϐ����Z�b�g����B
'******************************************************************************
Global B1                 As String         '�啶
Global B2                 As String         '
Global C1                 As String         'WRITE_TIME
Global C2                 As String         'FORECAST_KIND
Global C3                 As String         'FORECAST_KIND_CODE
Global C4                 As String         'ESTIMATE_TIME
Global C5                 As String         'ANNOUNCE_TIME
Global YHK(6, 18)         As Single         '
'******************************************************************************
'�\���̂��Z�b�g����B
'******************************************************************************
Type FRC
     ir As Long                             '
     ic As Long                             '
     m  As Long                             '
End Type

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public gDebugMode As String

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
'�T�u���[�`���FBin2Int()
'�����T�v�F
'******************************************************************************
Sub Bin2Int(b As Variant, i As Integer, rc As Boolean)
    Dim nf    As Long
    Dim rec   As Long
    Dim F     As String
    On Error GoTo BinErorr
    F = App.Path & "\Pump.Bin"
    nf = FreeFile
    Open F For Binary As #nf
    rec = 1
    Put #nf, rec, b
    rec = 1
    Get #nf, rec, i
    Close #nf
    On Error GoTo 0
    rc = True
    Exit Sub
BinErorr:
    On Error GoTo 0
    ORA_LOG "�o�C�i���ϊ��ŃG���[���������܂����B"
End Sub

'******************************************************************************
'�T�u���[�`���FDump_F_MESSYU_10MIN_1()
'�����T�v�F
'******************************************************************************
Sub Dump_F_MESSYU_10MIN_1(m As String, d As Date, data() As String)
    Dim i    As Long
    Dim j    As Long
    Dim k    As Long
    Dim bf   As String
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/04 O.OKADA�y02�z
    '��OracleAccess.ORA_F_MESSYU_10MIN_2()���C�����邱�ƁB�y02-01�z
    '��OracleAccess.ORA_F_MESSYU_10MIN_1()���C�����邱�ƁB�y02-02�z
    '�����ɃR�����g�A�E�g�ς݂ł��邪�A�R�����g�A�E�g��߂��Ȃ��悤�ɃR�����g��ǉ����邱�ƁB
    '******************************************************
    Exit Sub
    '******************************************************
    'Ver1.0.0 �C���I�� 2015/08/04 O.OKADA�y02�z
    '******************************************************
    Print #Dmp_N, m; "  "; d
    k = 0
    For i = 1 To 5
        bf = " "
        For j = 1 To 5
        k = k + 1
           bf = bf & data(k) & " "
        Next j
        Print #Dmp_N, bf
    Next i
End Sub

'******************************************************************************
'�T�u���[�`���FORA_P_WATER()
'�����T�v�F
'���ʃf�[�^���f�[�^�x�[�X���擾����
'�ϑ����ԍ�
'station IN( 1002,1015,1016,1017,1019,1020 )
'1002=������O����
'1015=�V�쉺�V��F
'1016=�厡
'1017=�����O����
'1019=�v�n��
'1020=�t��
'������
'1076=���V��F
'1077=�t��
'1079=�����O����
'"�e�����[�^���ʎ�M"
'�����ʕ�U�̃��b�Z�[�W�͒����Ԏ擾�̏ꍇ�ŏ���30���܂Ń��b�Z�[�W��
'�o�͂��܂��A�ȍ~�͕�U�͂��邪���b�Z�[�W�͏o�܂���B
'******************************************************************************
Sub ORA_P_WATER(d1 As Date, d2 As Date, ic As Boolean)
    Dim sql_SELECT   As String
    Dim sql_WHERE    As String
    Dim SQL          As String
'    Dim n            As Integer
    Dim i            As Long
    Dim SDATE        As String
    Dim EDATE        As String
    Dim Timew        As String
    Dim w(6)         As Single              '�吅�ʌv�f�[�^
    Dim s(3)         As Single              '�����ʌv�f�[�^
    Dim dw           As Date
    Dim dt           As String
    Dim A1
    Dim A2
    Dim A3
    Dim A4
    Dim f1
    Dim nf           As Integer
    Dim buf          As String
    Dim msgD(100)    As Date
    Dim msg(100)     As String
    Dim msg_num      As Long
    Dim hw           As Single
    Dim rc           As Boolean
    Const Ksk = -99#
    On Error GoTo ORA_P_WATER_Error
    ic = False
    ORA_LOG "���ʃf�[�^�擾�J�n"
    SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'" '," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    EDATE = "'" & Format(d2, "yyyy/mm/dd hh:nn") & "'" '," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    OracleDB.Label3 = "���m���͐���V�X�e���f�[�^�x�[�X���u��萅�ʃf�[�^�擾��"
    OracleDB.Label3.Refresh
    '******************************************************
    'SELECT
    '******************************************************
'    sql_SELECT = "SELECT * FROM oracle.P_WATER "
    sql_SELECT = "SELECT"
    sql_SELECT = sql_SELECT & "  obs_time"
    sql_SELECT = sql_SELECT & ", obs_sta_id"
    sql_SELECT = sql_SELECT & ", flag10"
    sql_SELECT = sql_SELECT & ", data10"
    sql_SELECT = sql_SELECT & "  FROM t_water_level_data"
    '******************************************************
    'WHERE
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
'    sql_WHERE = "WHERE station IN( 1002,1015,1016,1017,1019,1020,1076,1077,1079 ) AND jikoku BETWEEN TO_DATE(" & _
'                SDATE & ") AND TO_DATE(" & EDATE & ") ORDER BY jikoku"
    'sql_WHERE = "WHERE station IN( 2,16,17,18,20,21 ) and JIKOKU = TO_DATE(" & Sdate & ")"
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    sql_WHERE = " WHERE obs_sta_id IN(1012, 81, 201, 91, 71, 131, 80, 130, 240)"
    sql_WHERE = sql_WHERE & " AND obs_time BETWEEN " & SDATE
    sql_WHERE = sql_WHERE & " AND " & EDATE
    sql_WHERE = sql_WHERE & " ORDER BY obs_time"
    SQL = sql_SELECT & sql_WHERE
    '******************************************************
    'SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����B
    '******************************************************
'    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    Set gAdoRst = New ADODB.Recordset
    gAdoRst.CursorType = adOpenStatic
    gAdoRst.LockType = adLockReadOnly
    gAdoRst.Open SQL, gAdoCon, , , adCmdText
    If gAdoRst.EOF And gAdoRst.BOF Then
        ORA_LOG "���ʊϑ��f�[�^���f�[�^�x�[�X�ɓo�^����Ă��܂���B" & vbCrLf & _
                 "�������m���߂Ă��������B"
        ORA_LOG "SQL=" & SQL
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        'MsgBox "���ʊϑ��f�[�^���f�[�^�x�[�X�ɓo�^����Ă��܂���B" & vbCrLf & _
                 "�������m���߂Ă��������B"
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        ic = False
        Call SQLdbsDeleteRecordset(gAdoRst)
        OracleDB.Label3 = "���m���͐���V�X�e���f�[�^�x�[�X���u��萅�ʃf�[�^�擾���s"
        OracleDB.Label3.Refresh
        Exit Sub
    End If
    nf = FreeFile
    Open App.Path & "\Data\DB_H.DAT" For Output As #nf
    '******************************************************
    '�t�B�[���h�����擾����B
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'Dim Tw
    'n = dynOra.Fields.Count
    'For i = 0 To n - 1
    '    Tw = dynOra.Fields(i).Name
    '    Print #nf, " Number=" & Format(Str(i), "@@@") & " �t�B�[���h��="; Tw
    'Next i
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    w(1) = Ksk: w(2) = Ksk: w(3) = Ksk: w(4) = Ksk: w(5) = Ksk: w(6) = Ksk
    s(1) = Ksk: s(2) = Ksk: s(3) = Ksk
    gAdoRst.MoveFirst
    i = 0
    Timew = ""
    msg_num = 0
    Do
        buf = ""
        If Not gAdoRst.EOF Then A1 = Format(gAdoRst.Fields("obs_time").Value, "yyyy/mm/dd hh:nn")
        If Timew <> A1 And i > 0 Or gAdoRst.EOF Then
            '******************************************************
            '���V��F���ʃf�[�^�����������`�F�b�N����B
            '******************************************************
            If w(2) = Ksk Then
                If s(1) <> Ksk Then
                    w(2) = s(1)
                    If msg_num < 100 Then
                        msg_num = msg_num + 1
                        msg(msg_num) = "���V��F���ʃf�[�^���������܂����B�����ʌv�f�[�^�ŕ�U���܂����B"
                        msgD(msg_num) = DateAdd("s", 5, jgd)
                    End If
                Else
                    dw = CDate(A1)
                    �����ʎ擾 dw, hw, "���V��F", rc
                    If rc Then
                        w(2) = hw
                        If msg_num < 100 Then
                            msg_num = msg_num + 1
                            msg(msg_num) = "���V��F���ʊϑ��ǃf�[�^�̖����o�R�f�[�^���������܂����B������o�R�̎吅�ʌv�f�[�^�ŕ�U���܂����B"
                            msgD(msg_num) = DateAdd("s", 5, jgd)
                        End If
                    End If
                End If
            End If
            '******************************************************
            '�t�����ʃf�[�^�����������`�F�b�N����B
            '******************************************************
            If w(6) = Ksk Then
                If s(2) <> Ksk Then
                    w(6) = s(2)
                    If msg_num < 100 Then
                        msg_num = msg_num + 1
                        msg(msg_num) = "�t�����ʃf�[�^���������܂����B�����ʌv�f�[�^�ŕ�U���܂����B"
                        msgD(msg_num) = DateAdd("S", 6, jgd)
                    End If
                Else
                    dw = CDate(A1)
                    �����ʎ擾 dw, hw, "�t��", rc
                    If rc Then
                        w(6) = hw
                        If msg_num < 100 Then
                            msg_num = msg_num + 1
                            msg(msg_num) = "�t�����ʊϑ��ǃf�[�^�̖����o�R�f�[�^���������܂����B������o�R�̎吅�ʌv�f�[�^�ŕ�U���܂����B"
                            msgD(msg_num) = DateAdd("s", 6, jgd)
                        End If
                    End If
                End If
            End If
            '******************************************************
            '�����O���ʃf�[�^�����������`�F�b�N����B
            '******************************************************
            If w(4) = Ksk Then
                If s(3) <> Ksk Then
                    w(4) = s(3)
                    If msg_num < 100 Then
                        msg_num = msg_num + 1
                        msg(msg_num) = "�����O���ʃf�[�^���������܂����B�����ʌv�f�[�^�ŕ�U���܂����B"
                        msgD(msg_num) = DateAdd("S", 7, jgd)
                    End If
                Else
                    dw = CDate(A1)
                    �����ʎ擾 dw, hw, "�����O", rc
                    If rc Then
                        w(4) = hw
                        If msg_num < 100 Then
                            msg_num = msg_num + 1
                            msg(msg_num) = "�����O���ʊϑ��ǃf�[�^�̖����o�R�f�[�^���������܂����B������o�R�̎吅�ʌv�f�[�^�ŕ�U���܂����B"
                            msgD(msg_num) = DateAdd("s", 7, jgd)
                        End If
                    End If
                End If
            End If
            '******************************************************
            '�����ʌv�̂Ȃ��ϑ����������ʂŕ�U����B
            '�厡
            '******************************************************
            If w(3) = Ksk Then
                dw = CDate(A1)
                �����ʎ擾 dw, hw, "�厡", rc
                If rc Then
                    w(3) = hw
                    If msg_num < 100 Then
                        msg_num = msg_num + 1
                        msg(msg_num) = "�厡���ʊϑ��ǃf�[�^�̖����o�R�f�[�^���������܂����B������o�R�̎吅�ʌv�f�[�^�ŕ�U���܂����B"
                        msgD(msg_num) = DateAdd("s", 8, jgd)
                    End If
                End If
            End If
            '******************************************************
            '�����ʌv�̂Ȃ��ϑ����������ʂŕ�U����B
            '�v�n��
            '******************************************************
            If w(5) = Ksk Then
                dw = CDate(A1)
                �����ʎ擾 dw, hw, "�v�n��", rc
                If rc Then
                    w(5) = hw
                    If msg_num < 100 Then
                        msg_num = msg_num + 1
                        msg(msg_num) = "�v�n�쐅�ʊϑ��ǃf�[�^�̖����o�R�f�[�^���������܂����B������o�R�̎吅�ʌv�f�[�^�ŕ�U���܂����B"
                        msgD(msg_num) = DateAdd("s", 9, jgd)
                    End If
                End If
            End If
            '******************************************************
            'MDB�ɏ������ށB
            '******************************************************
            dw = CDate(Timew)
            dt = Format(dw, "yyyy/mm/dd hh:nn")
            MDB_Rst_H.Open "select * from .���� where Time = '" & dt & "' ; ", MDB_Con, adOpenDynamic, adLockOptimistic
            If MDB_Rst_H.EOF Or MDB_Rst_H.BOF Then
                MDB_Rst_H.AddNew            '���ʃf�[�^��ǉ�����B
            End If
            MDB_Rst_H.Fields("Time").Value = dt
            MDB_Rst_H.Fields("Minute").Value = Minute(dw)
            MDB_Rst_H.Fields("Tide").Value = w(1)
            MDB_Rst_H.Fields("���V��F").Value = w(2)
            MDB_Rst_H.Fields("�厡").Value = w(3)
            MDB_Rst_H.Fields("�����O").Value = w(4)
            MDB_Rst_H.Fields("�v�n��").Value = w(5)
            MDB_Rst_H.Fields("�t��").Value = w(6)
            MDB_Rst_H.Update
            MDB_Rst_H.Close
            w(1) = Ksk: w(2) = Ksk: w(3) = Ksk: w(4) = Ksk: w(5) = Ksk: w(6) = Ksk
            s(1) = Ksk: s(2) = Ksk: s(3) = Ksk
            Timew = A1
            '******************************************************
            '�֐����R�[������B
            '******************************************************
            Pump_Check dt, dw, w()          '�|���v��~���ʂ̃`�F�b�N���s���B
            If gAdoRst.EOF Then Exit Do
        End If
        If i = 0 Then Timew = A1
        i = i + 1
        A2 = gAdoRst.Fields("obs_sta_id").Value
        A3 = gAdoRst.Fields("obs_time").Value
        A4 = gAdoRst.Fields("data10").Value
        f1 = gAdoRst.Fields("flag10").Value
        buf = buf & Format(A1, "@@@@@@@@@@@@@@@@@@@@,")
        buf = buf & Format(Str(A2), "@@@@@@@@@@,")
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        'buf = buf & Format(Str(A3), "@@@@@@@@@@")
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        buf = buf & Format(Str(A4), "@@@@@@@@@@@@@@@,")
        buf = buf & Format(Str(f1), "@@@@@")
        Print #nf, buf
        jgd = CDate(A3)
        ORA_LOG Format(A1, "yyyy/mm/dd hh:nn") & "  " & A2 & "  H=" & A4 & " f=" & Str(f1)
        If f1 = 0 Or f1 = 10 Or f1 = 40 Or f1 = 50 Or f1 = 60 Or f1 = 70 Then
            Select Case CInt(A2)
                 Case 1012                  '������O����
                    w(1) = CSng(A4) * 0.01
                 Case 81                    '�V�쉺�V��F
                    w(2) = CSng(A4) * 0.01
                 Case 201                   '�厡
                    w(3) = CSng(A4) * 0.01
                 Case 91                    '�����O����
                    w(4) = CSng(A4) * 0.01
                 Case 71                    '�v�n��
                    w(5) = CSng(A4) * 0.01
                 Case 131                   '�t��
                    w(6) = CSng(A4) * 0.01
                 Case 80                    '�V�쉺�V��F�����ʌv
                    s(1) = CSng(A4) * 0.01
                 Case 130                   '�t�������ʌv
                    s(2) = CSng(A4) * 0.01
                 Case 240                   '�����O���ʕ����ʌv
                    s(3) = CSng(A4) * 0.01
            End Select
        End If
        gAdoRst.MoveNext
        DoEvents
    Loop
    If msg_num > 0 Then
        For i = 1 To msg_num
            jgd = msgD(i)
            ORA_LOG msg(i)
            ORA_Message_Out "�e�����[�^���ʎ�M", msg(i), 1
        Next i
    End If
    ic = True
    Close #nf
    Call SQLdbsDeleteRecordset(gAdoRst)
    Set MDB_Rst_H = Nothing
    OracleDB.Label3 = "���m���͐���V�X�e���f�[�^�x�[�X���u��萅�ʃf�[�^�擾�I��"
    OracleDB.Label3.Refresh
    Exit Sub
ORA_P_WATER_Error:
    Dim strMessage As String
    strMessage = Err.Description
    ORA_LOG "���m���͐���V�X�e���f�[�^�x�[�X���u��萅�ʃf�[�^�擾���s"
    ORA_LOG strMessage
    On Error GoTo 0
    Call SQLdbsDeleteRecordset(gAdoRst)
    OracleDB.Label3 = "���m���͐���V�X�e���f�[�^�x�[�X���u��萅�ʃf�[�^�擾���s"
    OracleDB.Label3.Refresh
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
'�����T�v�F
'��n������B
'******************************************************************************
Sub ORA_DataBase_Close()
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
    gDebugMode = vbNullString
End Sub

'******************************************************************************
'�T�u���[�`���FORA_DataBase_Connection()
'�����T�v�F
'******************************************************************************
Sub ORA_DataBase_Connection(ic As Boolean)
    '******************************************************
    'OO4O �� Oracle �ɐڑ�����B
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'On Error GoTo ERRHAND
    On Error Resume Next
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    '�Z�b�V�������쐬����B
    '******************************************************
'    Set ssOra = CreateObject("OracleInProcServer.XOraSession")
'    If Err <> 0 Then
'        'Ver0.0.0 �C���J�n 1900/01/01 00:00
'        'MsgBox "���m���I���N���f�[�^�x�[�X�ɐڑ��o���܂���B" & Chr(10) & _
'               "CreateObject - Oracle oo4o �G���["
'        'Ver0.0.0 �C���I�� 1900/01/01 00:00
'        ORA_LOG "���m���I���N���f�[�^�x�[�X�ɐڑ��o���܂���B" & Chr(10) & _
'                "CreateObject - Oracle oo4o �G���[" & Chr(10) & _
'                "10�b�x�e���܂� " & Now
'        Short_Break 10
'        GoTo ERRHAND
'    End If
    '******************************************************
    '�T�[�r�X���i�T�[�o���j�� ���[�U��/�p�X���[�h ���w�肷��B
    '******************************************************
    
    Dim strConfigFile As String
    Dim strProvider As String
    Dim strServer As String
    Dim strDBS As String
    Dim strUID As String
    Dim strPWD As String
    Dim strConn As String
    gDebugMode = vbNullString
    strConfigFile = App.Path
    If Right(strConfigFile, 1) <> "\" Then strConfigFile = strConfigFile & "\"
    strConfigFile = strConfigFile & "dbsinfo.cfg"
    If Len(Dir(strConfigFile, vbNormal)) < 1 Then
        ORA_LOG "���m���͐���V�X�e���f�[�^�x�[�X���u�̐ڑ����t�@�C��������܂���B" & Chr(10) & _
                "10�b�x�e���܂� " & Now
        Short_Break 10
        GoTo ERRHAND
    End If
    
    strProvider = GetConfigData("databases", "provider", strConfigFile)
    strServer = GetConfigData("databases", "server", strConfigFile)
    strDBS = GetConfigData("databases", "dbs", strConfigFile)
    strUID = GetConfigData("databases", "uid", strConfigFile)
    strPWD = GetConfigData("databases", "pwd", strConfigFile)
    gDebugMode = GetConfigData("databases", "debug", strConfigFile)
    If Len(strServer) < 1 Or Len(strDBS) < 1 Then
        ORA_LOG "���m���͐���V�X�e���f�[�^�x�[�X���u�̐ڑ���񂪂���܂���B" & Chr(10) & _
                "10�b�x�e���܂� " & Now
        Short_Break 10
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
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        'MsgBox "���m���I���N���f�[�^�x�[�X�ɐڑ��o���܂���B" & vbCrLf & _
              Err & ": " & Error
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        ORA_LOG "���m���͐���V�X�e���f�[�^�x�[�X���u�Ɛڑ��ł��܂���B" & Chr(10) & _
                "10�b�x�e���܂� " & Now
        Short_Break 10
        GoTo ERRHAND
    End If
   On Local Error GoTo 0
   ic = True
   Exit Sub
ERRHAND:
    Dim strMessage As String
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'If dbOra.LastServerErr <> 0 Then
'    If dbOra.LastServerErr <> 0 Then
'    'Ver0.0.0 �C���I�� 1900/01/01 00:00
'        strMessage = dbOra.LastServerErrText    'DB�����ŃG���[�����������Ƃ��̏����B
'    Else
        strMessage = Err.Description            'DB�����ȊO�ŃG���[�����������Ƃ��̏����B
'    End If
    ORA_LOG strMessage
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'MsgBox strMessage, vbExclamation
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    ic = False
    On Local Error GoTo 0
    Call ORA_DataBase_Close
End Sub

'******************************************************************************
'�T�u���[�`���F
'�����T�v�F
'�C�ے�10���\�����b�V���f�[�^(10���J�ʁj
'�U�����Z���Ȃ��Ǝ��ԉJ�ʂɂȂ�Ȃ��̂ő����Z������
'�l�c�a�ɂ͎��ԗ��敽�ωJ�ʂƂ��Ċi�[����
'�w�肳�ꂽd1�����Ɍv�Z����������U�O���̂P�O���s�b�`�̗\���J��
'���v�����d1�����̂P���Ԍ�̗\���J�ʂɂȂ�B
'******************************************************************************
Sub ORA_F_MESSYU_10MIN_1(d1 As Date, ic As Boolean)
    Dim sql_SELECT   As String
    Dim sql_WHERE1   As String
    Dim sql_WHERE2   As String
    Dim SQL          As String
    Dim n            As Integer
    Dim i            As Long
    Dim j            As Long
    Dim k            As Long
    Dim SDATE        As String
    Dim EDATE        As String
    Dim Timew        As String
    Dim w1(315)      As Single              '1km���b�V���l
    Dim w2(250)      As Single              '2km���b�V���l
    Dim dw           As String
    Dim df           As Date
    Dim dm           As Date
    Dim MM           As Long
    Dim nf           As Integer
    Dim ks           As Integer
    Dim mes(10)      As String              '2�����b�V���ԍ�
    Dim FM(25)       As String              '�c�a��2�����b�V���ԍ�
    Dim i1           As String
    Dim j1           As String
    Dim buf          As String
    Dim irc          As Boolean
    Dim m            As Integer
    Dim rr(135)      As Single
    Dim Ytime        As String
    Dim tm(10)       As String
    Dim c            As Single
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/04 O.OKADA�y03�z
    '��OracleAccess.ORA_F_MESSYU_10MIN_1()���C�����邱�ƁB�y03-01�z
    '��OracleDB.Check_F_MESSYU_10MIN_1_Time()���C�����邱�ƁB�y03-02�z
    '******************************************************
    Exit Sub
    '******************************************************
    'Ver1.0.0 �C���I�� 2015/08/04 O.OKADA�y03�z
    '******************************************************
    mes(1) = "533606"
    mes(2) = "533607"
    mes(3) = "523676"
    mes(4) = "523677"
    mes(5) = "523770"
    mes(6) = "523666"
    mes(7) = "523667"
    mes(8) = "523760"
    mes(9) = "523656"
    mes(10) = "523646"
    n = 0
    For i = 1 To 5
        i1 = Format(i, "0")
        For j = 1 To 5
            j1 = Format(j, "0")
            n = n + 1
            FM(n) = "data_" & i1 & j1
        Next j
    Next i
    OracleDB.Label3 = "�I���N�����C�ے��P�O���\�����[�_�f�[�^�J�ʎ擾��"
    OracleDB.Label3.Refresh
    '******************************************************
    'SELECT
    '******************************************************
    sql_SELECT = "SELECT * FROM oracle.F_MESSYU_10MIN_1 "
    '******************************************************
    'WHERE
    '******************************************************
    SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    sql_WHERE1 = "WHERE code IN( '533606','533607','523676','523677','523770'," & _
                               "'523666','523667','523760','523656','523646' ) AND " & _
                 "jikoku = TO_DATE(" & SDATE & ") "
    'sql_WHERE2 = " AND \yosoku_time = TO_DATE(" & EDATE & ") "
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    '�t�B�[���h���e���擾����B
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'Dim Tw, sql_WHERE, d2
    'd1 = "2002/06/14 21:10"
    'd2 = "2002/06/14 21:10"
    'SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'EDATE = "'" & Format(d2, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'sql_WHERE = " WHERE  jikoku BETWEEN TO_DATE(" & SDATE & ") AND TO_DATE(" & EDATE & ") " & _
    '    " AND code IN( '533606','533607','523676','523677','523770','523666','523667','523760','523656','523646' )"
    'SQL = sql_SELECT & sql_WHERE & " order by jikoku"
    'Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    'n = dynOra.Fields.Count
    'Do Until dynOra.EOF
    '    Tw = dynOra.Fields("jikoku").Value
    '    Tw = Tw & "  " & dynOra.Fields("code").Value
    '    Tw = Tw & "  " & dynOra.Fields("start_time").Value
    '    Tw = Tw & "  " & dynOra.Fields("end_time").Value
    '    Tw = Tw & "  " & dynOra.Fields("yosoku_time").Value
    '    Debug.Print Tw
    '    dynOra.MoveNext
    '    DoEvents
    'Loop
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    ic = True
    SQL = sql_SELECT & sql_WHERE1           '& sql_WHERE2
    '******************************************************
    'SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����B
    '******************************************************
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    ORA_LOG "SQL=" & SQL
    If dynOra.EOF And dynOra.BOF Then
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        'MsgBox "�ϑ��f�[�^���f�[�^�x�[�X�ɓo�^����Ă��܂���B" & vbCrLf & _
                 "�������m���߂Ă��������B"
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        dynOra.Close
        ORA_LOG "�C�ے����[�_�[�\�� �X�L�b�v" & d1
        ORA_LOG "�C�ے��P�O���\���J�ʃf�[�^�X�L�b�v�����������݊J�n " & d1
        dm = DateAdd("h", 1, d1)
        nf = FreeFile
        Open App.Path & "\data\F_MESSYU_10MIN_1.DAT" For Output As #nf
        Print #nf, Format(dm, "yyyy/mm/dd hh:nn")
        Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
        Close #nf
        ORA_LOG "�C�ے��P�O���\���J�ʃf�[�^�L�b�v�����������ݏI��"
        GoTo SKIP
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        'Set dynOra = Nothing
        'OracleDB.Label3 = "�e�q�h�b�r�P�O���\�����[�_�f�[�^�J�ʎ掸���s"
        'OracleDB.Label3.Refresh
        'Exit Sub
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
    End If
    For i = 1 To 10
        tm(i) = "000000"
    Next i
    Do Until dynOra.EOF
        Ytime = dynOra.Fields("yosoku_time").Value
        dw = dynOra.Fields("jikoku").Value
        df = CDate(dw)
        dw = Format(df, "yyyy/mm/dd hh:nn")
        MM = dynOra.Fields("code").Value
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        'Debug.Print " dw="; dw; "  code="; MM; "  m="; m; "  Ytime="; Ytime
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        Select Case MM
            Case mes(1)
                ks = 1
                m = 1
            Case mes(2)
                ks = 26
                m = 2
            Case mes(3)
                ks = 51
                m = 3
            Case mes(4)
                ks = 76
                m = 4
            Case mes(5)
                ks = 101
                m = 5
            Case mes(6)
                ks = 128
                m = 6
            Case mes(7)
                ks = 151
                m = 7
            Case mes(8)
                ks = 176
                m = 8
            Case mes(9)
                ks = 201
                m = 9
            Case mes(10)
                ks = 226
                m = 10
        End Select
        Select Case Ytime
            Case "010"
                dm = DateAdd("n", 1, df)
                Mid(tm(m), 1, 1) = "1"
            Case "020"
                dm = DateAdd("n", 2, df)
                Mid(tm(m), 2, 1) = "2"
            Case "030"
                dm = DateAdd("n", 3, df)
                Mid(tm(m), 3, 1) = "3"
            Case "040"
                dm = DateAdd("n", 4, df)
                Mid(tm(m), 4, 1) = "4"
            Case "050"
                dm = DateAdd("n", 5, df)
                Mid(tm(m), 5, 1) = "5"
            Case "060"
                dm = DateAdd("n", 6, df)
                Mid(tm(m), 6, 1) = "6"
            Case Else
                ORA_LOG "IN ORA_F_MESSYU_10MIN_1  �ǂ����ČX�ɂ���́H"
                'Ver0.0.0 �C���J�n 1900/01/01 00:00
                'MsgBox " �ǂ����ČX�ɂ���́H"
                'Ver0.0.0 �C���I�� 1900/01/01 00:00
        End Select
        For i = ks To ks + 24
            j = i - ks + 1
            c = CSng(dynOra.Fields(FM(j)).Value)
            If c < 0# Then c = 0#           '�C�ے���-1�𑗂��Ă���\��������B
            w2(i) = w2(i) + c
        Next i
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        'Dump_F_MESSYU_10MIN_1 MES(m), df, dmp()
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        Mesh_2km_to_1km_cvt w2(), w1()
        dynOra.MoveNext
    Loop
    Mesh_To_Ryuiki w1(), rr(), irc
    dm = DateAdd("h", 1, d1)
    ORA_LOG "�C�ے����[�_�[�\��10�� " & dm
    For i = 1 To 10
        ORA_LOG Format(Str(i), "@@@") & " " & mes(i) & " " & tm(i)
    Next i
    '******************************************************
    'MDB��OPEN����B
    '******************************************************
    Set MDB_Rst_H.ActiveConnection = MDB_Con
    '******************************************************
    'MDB�ɏ������ށB
    '******************************************************
    MDB_Rst_H.Open "select * from .�C�ے����[�_�[�\��_1 where Time = '" & Format(dm, "yyyy/mm/dd hh:nn") & "' ; ", MDB_Con, adOpenDynamic, adLockOptimistic
    If MDB_Rst_H.EOF Or MDB_Rst_H.BOF Then
        MDB_Rst_H.AddNew                    '���݂��Ă�����f�[�^��ǉ�����B
    End If
    MDB_Rst_H.Fields("Time").Value = Format(dm, "yyyy/mm/dd hh:nn")
    MDB_Rst_H.Fields("Minute").Value = Minute(dm)
    For i = 1 To 135
        i1 = Format(i, "###")
        MDB_Rst_H.Fields(i1).Value = rr(i)
    Next i
    MDB_Rst_H.Update
    MDB_Rst_H.Close
    ORA_LOG "�C�ے��P�O���\���J�ʃf�[�^�����������݊J�n " & d1
    nf = FreeFile
    Open App.Path & "\data\F_MESSYU_10MIN_1.DAT" For Output As #nf
    Print #nf, Format(d1, "yyyy/mm/dd hh:nn")
    Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
    Close #nf
    ORA_LOG "�C�ے��P�O���\���J�ʃf�[�^�����������ݏI��"
SKIP:
    dynOra.Close
    Set dynOra = Nothing
    Set MDB_Rst_H = Nothing
    OracleDB.Label3 = "�I���N�����C�ے����[�_�[�\���I��"
    OracleDB.Label3.Refresh
End Sub

'******************************************************************************
'�T�u���[�`���FORA_F_MESSYU_10MIN_2()
'�����T�v�F
'�C�ے������\�����b�V���f�[�^(10���J�ʁj�U���ԕ�
'�U�����Z���Ȃ��Ǝ��ԉJ�ʂɂȂ�Ȃ�
'�l�c�a�ɂ͂P�O�����敽�ωJ�ʂƂ��Ċi�[����
'******************************************************************************
Sub ORA_F_MESSYU_10MIN_2(d1 As Date, ic As Boolean)
    Dim sql_SELECT   As String
    Dim sql_WHERE1   As String
    Dim sql_WHERE2   As String
    Dim SQL          As String
    Dim n            As Integer
    Dim i            As Long
    Dim j            As Long
    Dim k            As Long
    Dim SDATE        As String
    Dim EDATE        As String
    Dim Timew        As String
    Dim w1(315)      As Single              '1km���b�V���l
    Dim w2(250)      As Single              '2km���b�V���l
    Dim dw           As String
    Dim df           As Date
    Dim dm           As Date
    Dim dmc          As String
    Dim MM           As Long
    Dim nf           As Integer
    Dim ks           As Integer
    Dim mes(10)      As String              '2�����b�V���ԍ�
    Dim FM(25)       As String              '�c�a��2�����b�V���ԍ�
    Dim i1           As String
    Dim j1           As String
    Dim buf          As String
    Dim irc          As Boolean
    Dim m            As Integer
    Dim rr(135)      As Single
    Dim sr(135)      As Single
    Dim rrr(36, 135) As Single
    Dim Ytime        As String
    Dim tm(36)       As String
    Dim m1           As Long
    Dim m2           As Long
    Dim m3           As Long
    Dim m4           As Long
    Dim dmp(25)      As String
    Dim c            As Single
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/04 O.OKADA�y04�z
    '��OracleDB.Check_F_MESSYU_10MIN_2_Time()���C�����邱�ƁB�y04-01�z
    '******************************************************
    Exit Sub
    '******************************************************
    'Ver1.0.0 �C���I�� 2015/08/04 O.OKADA�y04�z
    '******************************************************
    ic = True
    mes(1) = "533606"
    mes(2) = "533607"
    mes(3) = "523676"
    mes(4) = "523677"
    mes(5) = "523770"
    mes(6) = "523666"
    mes(7) = "523667"
    mes(8) = "523760"
    mes(9) = "523656"
    mes(10) = "523646"
    n = 0
    For i = 1 To 5
        i1 = Format(i, "0")
        For j = 1 To 5
            j1 = Format(j, "0")
            n = n + 1
            FM(n) = "data_" & i1 & j1
        Next j
    Next i
    OracleDB.Label3 = "�I���N�����C�ے������\�����[�_�f�[�^�J�ʎ擾��"
    OracleDB.Label3.Refresh
    '******************************************************
    'SELECT
    '******************************************************
    sql_SELECT = "SELECT * FROM oracle.F_MESSYU_10MIN_2 "
    '******************************************************
    'WHERE
    '******************************************************
    SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    sql_WHERE1 = "WHERE code IN( '533606','533607','523676','523677','523770'," & _
                               "'523666','523667','523760','523656','523646' ) AND " & _
                 "jikoku = TO_DATE(" & SDATE & ") ORDER BY jikoku"
    'sql_WHERE2 = " AND \yosoku_time = TO_DATE(" & EDATE & ") "
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    '�t�B�[���h���e���擾����B
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'Dim Tw, sql_WHERE, d2
    'd1 = "2002/06/11 15:00"
    'd2 = "2002/06/11 15:00"
    'SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'EDATE = "'" & Format(d2, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'sql_WHERE = " WHERE  jikoku BETWEEN TO_DATE(" & SDATE & ") AND TO_DATE(" & EDATE & ") " & _
    '    " AND code IN( '533606','533607','523676','523677','523770','523666','523667','523760','523656','523646' )"
    'SQL = sql_SELECT & sql_WHERE & " order by jikoku"
    'Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    'n = dynOra.Fields.Count
    'Do Until dynOra.EOF
    '    Tw = dynOra.Fields("jikoku").Value
    '    Tw = Tw & "  " & dynOra.Fields("code").Value
    '    Tw = Tw & "  " & dynOra.Fields("start_time").Value
    '    Tw = Tw & "  " & dynOra.Fields("end_time").Value
    '    Tw = Tw & "  " & dynOra.Fields("yosoku_time").Value
    '    Debug.Print Tw
    '    dynOra.MoveNext
    '    DoEvents
    'Loop
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    SQL = sql_SELECT & sql_WHERE1 '& sql_WHERE2
    '******************************************************
    'SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����B
    '******************************************************
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    If dynOra.EOF And dynOra.BOF Then
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        'MsgBox "�ϑ��f�[�^���f�[�^�x�[�X�ɓo�^����Ă��܂���B" & vbCrLf & _
                 "�������m���߂Ă��������B"
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        ic = False
        dynOra.Close
        ORA_LOG "�C�ے����[�_�[�\�� �X�L�b�v" & d1
        ORA_LOG "�C�ے������\���J�ʃf�[�^�X�L�b�v�����������݊J�n " & d1
        nf = FreeFile
        Open App.Path & "\data\F_MESSYU_10MIN_2.DAT" For Output As #nf
        Print #nf, Format(d1, "yyyy/mm/dd hh:nn")
        Close #nf
        ORA_LOG "�C�ے������\���J�ʃf�[�^�L�b�v�����������ݏI��"
        GoTo SKIP
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        'Set dynOra = Nothing
        'OracleDB.Label3 = "�C�ے��P�O���\�����[�_�f�[�^�J�ʎ掸���s"
        'OracleDB.Label3.Refresh
        'Exit Sub
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
    End If
    For j = 1 To 36
        tm(j) = "0000000000"
    Next j
    Do Until dynOra.EOF
        Erase w1
        For m4 = 1 To 10
            If m4 = 1 Then
                Ytime = dynOra.Fields("yosoku_time").Value
            Else
                If dynOra.Fields("yosoku_time").Value <> Ytime Then
                    ORA_LOG "�C�ے������\���f�[�^���s���S"
                    GoTo SKIP
                End If
            End If
            dw = dynOra.Fields("jikoku").Value
            df = CDate(dw)
            dw = Format(df, "yyyy/mm/dd hh")
            MM = dynOra.Fields("code").Value
            'Ver0.0.0 �C���J�n 1900/01/01 00:00
            'Debug.Print " dw="; dw; "  code="; MM; "  m="; m; "  Ytime="; Ytime
            'Ver0.0.0 �C���I�� 1900/01/01 00:00
            Select Case MM
                Case mes(1)
                    ks = 1
                    n = 1
                Case mes(2)
                    ks = 26
                    n = 2
                Case mes(3)
                    ks = 51
                    n = 3
                Case mes(4)
                    ks = 76
                    n = 4
                Case mes(5)
                    ks = 101
                    n = 5
                Case mes(6)
                    ks = 126
                    n = 6
                Case mes(7)
                    ks = 151
                    n = 7
                Case mes(8)
                    ks = 176
                    n = 8
                Case mes(9)
                    ks = 201
                    n = 9
                Case mes(10)
                    ks = 226
                    n = 10
            End Select
            For i = ks To ks + 24
                j = i - ks + 1
                c = CSng(dynOra.Fields(FM(j)).Value)
                If c < 0# Then c = 0#       '�C�ے���-1�𑗂��Ă���\��������̂ł��܂��Ȃ�������B
                w2(i) = c
            Next i
            'Ver0.0.0 �C���J�n 1900/01/01 00:00
            'Dump_F_MESSYU_10MIN_1 MM & " " & Ytime, df, dmp()
            'Ver0.0.0 �C���I�� 1900/01/01 00:00
            dynOra.MoveNext
            DoEvents
            Mesh_2km_to_1km_cvt w2(), w1()
            If IsNumeric(Ytime) Then
                dw = dw & ":" & Mid(Ytime, 1, 2)
                m = CInt(Mid(Ytime, 1, 2))
            Else
                ORA_LOG "IN ORA_F_MESSYU_10MIN_2  �ǂ����ČX�ɂ���́H"
                'Ver0.0.0 �C���J�n 1900/01/01 00:00
                'MsgBox " �ǂ����Ă����ɂ���́H"
                'Ver0.0.0 �C���I�� 1900/01/01 00:00
            End If
            If n < 10 Then
                Mid(tm(m), n, 1) = Trim(Str(n))
            Else
                Mid(tm(m), n, 1) = "A"
            End If
        Next m4
        ORA_LOG "  Date=" & d1 & "  Ytime=" & Ytime
        Mesh_To_Ryuiki w1(), rr(), irc
        For i = 1 To 135
            rrr(m, i) = rr(i)
        Next i
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        'GoTo SKIP
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
    Loop
    '******************************************************
    '
    '******************************************************
    For i = 1 To 36
        If tm(i) <> "123456789A" Then
            ORA_LOG " tm(" & Format(Trim(Str(i)), "@@") & ")= " & tm(i)
        End If
    Next i
    dm = DateAdd("h", 1, d1)
    For m1 = 1 To 31
        Debug.Print " 10MIN_2" & " d1="; Format(d1, "yyyy/mm/dd hh:nn") & " dm=" & Format(dm, "yyyy/mm/dd hh:nn")
        Erase sr
        For m2 = m1 To m1 + 5
            For m3 = 1 To 135
                sr(m3) = sr(m3) + rrr(m2, m3)
            Next m3
        Next m2
        '******************************************************
        'MDB��OPEN����B
        '******************************************************
        Set MDB_Rst_H.ActiveConnection = MDB_Con
        '******************************************************
        'MDB�ɏ������ށB
        '******************************************************
        dmc = Format(dm, "yyyy/mm/dd hh:nn")
        MDB_Rst_H.Open "select * from .�C�ے����[�_�[�\��_2 where Time = '" & dmc & "' ; ", MDB_Con, adOpenDynamic, adLockOptimistic
        If MDB_Rst_H.EOF Or MDB_Rst_H.BOF Then
            MDB_Rst_H.AddNew                '���݂��Ă�����f�[�^��ǉ�����B
        End If
        MDB_Rst_H.Fields("Time").Value = dmc
        MDB_Rst_H.Fields("Minute").Value = Minute(dm)
        For i = 1 To 135
            i1 = Format(i, "###")
            MDB_Rst_H.Fields(i1).Value = sr(i)
        Next i
        MDB_Rst_H.Update
        MDB_Rst_H.Close
        dm = DateAdd("n", 10, dm)
    Next m1
    ic = True
SKIP:
    dynOra.Close
    Set dynOra = Nothing
    OracleDB.Label3 = "�I���N�����C�ے����[�_�[�����\���I��"
    OracleDB.Label3.Refresh
    Set MDB_Rst_H = Nothing
End Sub

'******************************************************************************
'�T�u���[�`���FORA_F_MESSYU_10MIN_20
'�����T�v�F
'�C�ے������\�����b�V���f�[�^(10���J�ʁj�U���ԕ�
'�U�����Z���Ȃ��Ǝ��ԉJ�ʂɂȂ�Ȃ�
'�l�c�a�ɂ͂P�O�����敽�ωJ�ʂƂ��Ċi�[����
'******************************************************************************
Sub ORA_F_MESSYU_10MIN_20(d1 As Date, ic As Boolean)
    Dim sql_SELECT   As String
    Dim sql_WHERE1   As String
    Dim sql_WHERE2   As String
    Dim SQL          As String
    Dim n            As Integer
    Dim i            As Long
    Dim j            As Long
    Dim k            As Long
    Dim SDATE        As String
    Dim EDATE        As String
    Dim Timew        As String
    Dim w1(315)      As Single              '1km���b�V���l
    Dim w2(250)      As Single              '2km���b�V���l
    Dim dw           As String
    Dim df           As Date
    Dim dm           As Date
    Dim dmc          As String
    Dim MM           As Long
    Dim nf           As Integer
    Dim ks           As Integer
    Dim mes(10)      As String              '2�����b�V���ԍ�
    Dim FM(25)       As String              '�c�a��2�����b�V���ԍ�
    Dim i1           As String
    Dim j1           As String
    Dim buf          As String
    Dim irc          As Boolean
    Dim m            As Integer
    Dim rr(135)      As Single
    Dim sr(135)      As Single
    Dim rrr(36, 135) As Single
    Dim Ytime        As String
    Dim tm(36)       As String
    Dim m1           As Long
    Dim m2           As Long
    Dim m3           As Long
    Dim jj           As Long
    Dim jjj          As Long
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/04 O.OKADA�y05�z
    '�����̏C���ɔ����e���͂Ȃ��B
    '******************************************************
    Exit Sub
    '******************************************************
    'Ver1.0.0 �C���I�� 2015/08/04 O.OKADA�y05�z
    '******************************************************
    mes(1) = "533606"
    mes(2) = "533607"
    mes(3) = "523676"
    mes(4) = "523677"
    mes(5) = "523770"
    mes(6) = "523666"
    mes(7) = "523667"
    mes(8) = "523760"
    mes(9) = "523656"
    mes(10) = "523646"
    n = 0
    For i = 1 To 5
        i1 = Format(i, "0")
        For j = 1 To 5
            j1 = Format(j, "0")
            n = n + 1
            FM(n) = "data_" & i1 & j1
        Next j
    Next i
    OracleDB.Label3 = "�I���N�����C�ے������\�����[�_�f�[�^�J�ʎ擾��"
    OracleDB.Label3.Refresh
    '******************************************************
    'SELECT
    '******************************************************
    sql_SELECT = "SELECT * FROM oracle.F_MESSYU_10MIN_2 "
    '******************************************************
    'WHERE
    '******************************************************
    SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    sql_WHERE1 = "WHERE code IN( '533606','533607','523676','523677','523770'," & _
                               "'523666','523667','523760','523656','523646' ) AND " & _
                 "jikoku = TO_DATE(" & SDATE & ") AND DETAIL = 2"
    'sql_WHERE2 = " AND \yosoku_time = TO_DATE(" & EDATE & ") "
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    '�t�B�[���h���e���擾����B
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'Dim Tw, sql_WHERE, d2
    'd1 = "2002/06/11 15:00"
    'd2 = "2002/06/11 15:00"
    'SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'EDATE = "'" & Format(d2, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'sql_WHERE = " WHERE  jikoku BETWEEN TO_DATE(" & SDATE & ") AND TO_DATE(" & EDATE & ") " & _
    '    " AND code IN( '533606','533607','523676','523677','523770','523666','523667','523760','523656','523646' )"
    'SQL = sql_SELECT & sql_WHERE & " order by jikoku"
    'Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    'n = dynOra.Fields.Count
    'Do Until dynOra.EOF
    '    Tw = dynOra.Fields("jikoku").Value
    '    Tw = Tw & "  " & dynOra.Fields("code").Value
    '    Tw = Tw & "  " & dynOra.Fields("start_time").Value
    '    Tw = Tw & "  " & dynOra.Fields("end_time").Value
    '    Tw = Tw & "  " & dynOra.Fields("yosoku_time").Value
    '    Debug.Print Tw
    '    dynOra.MoveNext
    '    DoEvents
    'Loop
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    SQL = sql_SELECT & sql_WHERE1 '& sql_WHERE2
    '******************************************************
    'SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����B
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    'If dynOra.EOF And dynOra.BOF Then
    ''    MsgBox "�ϑ��f�[�^���f�[�^�x�[�X�ɓo�^����Ă��܂���B" & vbCrLf & _
    '             "�������m���߂Ă��������B"
    '    ic = False
    '    dynOra.Close
    '    ORA_LOG "�C�ے����[�_�[�\�� �X�L�b�v" & d1
    '    ORA_LOG "�C�ے������\���J�ʃf�[�^�X�L�b�v�����������݊J�n " & d1
    '    nf = FreeFile
    '    Open App.Path & "\data\F_MESSYU_10MIN_2.DAT" For Output As #nf
    '    Print #nf, Format(d1, "yyyy/mm/dd hh:nn")
    '    Close #nf
    '    ORA_LOG "�C�ے������\���J�ʃf�[�^�L�b�v�����������ݏI��"
    '    GoTo SKIP
    ''    Set dynOra = Nothing
    ''    OracleDB.Label3 = "�C�ے��P�O���\�����[�_�f�[�^�J�ʎ掸���s"
    ''    OracleDB.Label3.Refresh
    ''    Exit Sub
    'End If
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    For j = 1 To 36
        tm(j) = "0000000000"
    Next j
    For jjj = 1 To 36                       'Do Until dynOra.EOF
    For jj = 1 To 10                        'Do Until dynOra.EOF
        Ytime = Format(jjj, "00") & "0"     'Ytime = dynOra.Fields("yosoku_time").Value
        dw = d1                             'dynOra.Fields("jikoku").Value
        df = CDate(dw)
        dw = Format(df, "yyyy/mm/dd hh")
        MM = mes(jj)                        'MM = dynOra.Fields("code").Value
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        'Debug.Print " dw="; dw; "  code="; MM; "  m="; m; "  Ytime="; Ytime
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        Select Case MM
            Case mes(1)
                ks = 1
                n = 1
            Case mes(2)
                ks = 26
                n = 2
            Case mes(3)
                ks = 51
                n = 3
            Case mes(4)
                ks = 76
                n = 4
            Case mes(5)
                ks = 101
                n = 5
            Case mes(6)
                ks = 128
                n = 6
            Case mes(7)
                ks = 151
                n = 7
            Case mes(8)
                ks = 176
                n = 8
            Case mes(9)
                ks = 201
                n = 9
            Case mes(10)
                ks = 226
                n = 10
        End Select
        For i = ks To ks + 24
            j = i - ks + 1
            w2(i) = 1                       'dynOra.Fields(FM(j)).Value * 0.5
        Next i
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        'dynOra.MoveNext
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        DoEvents
        Mesh_2km_to_1km_cvt w2(), w1()
        ORA_LOG "  Date=" & d1 & "  Ytime=" & Ytime
        Mesh_To_Ryuiki w1(), rr(), irc
        If IsNumeric(Ytime) Then
            dw = dw & ":" & Mid(Ytime, 1, 2)
            m = CInt(Mid(Ytime, 1, 2))
        Else
            ORA_LOG "IN ORA_F_MESSYU_10MIN_2  �ǂ����ČX�ɂ���́H"
            'Ver0.0.0 �C���J�n 1900/01/01 00:00
            'MsgBox " �ǂ����Ă����ɂ���́H"
            'Ver0.0.0 �C���I�� 1900/01/01 00:00
        End If
        If n < 10 Then
            Mid(tm(m), n, 1) = Trim(Str(n))
        Else
            Mid(tm(m), n, 1) = "A"
        End If
        For i = 1 To 135
            rrr(m, i) = rr(i)
        Next i
    Next jj                                 'Loop
    Next jjj                                'Loop
    '******************************************************
    '
    '******************************************************
    For i = 1 To 36
        If tm(i) <> "123456789A" Then
            ORA_LOG " tm(" & Format(Trim(Str(i)), "@@") & ")= " & tm(i)
        End If
    Next i
    dm = DateAdd("h", 1, d1)
    For m1 = 1 To 31
        Debug.Print " 10MIN_2" & " d1="; Format(d1, "yyyy/mm/dd hh:nn") & " dm=" & Format(dm, "yyyy/mm/dd hh:nn")
        Erase sr
        For m2 = m1 To m1 + 5
            For m3 = 1 To 135
                sr(m3) = sr(m3) + rrr(m2, m3)
            Next m3
            '******************************************************
            'MDB��OPEN����B
            '******************************************************
            Set MDB_Rst_H.ActiveConnection = MDB_Con
            '******************************************************
            'MDB�ɏ������ށB
            '******************************************************
            dmc = Format(dm, "yyyy/mm/dd hh:nn")
            MDB_Rst_H.Open "select * from .�C�ے����[�_�[�\��_2 where Time = '" & dmc & "' ; ", MDB_Con, adOpenDynamic, adLockOptimistic
            If MDB_Rst_H.EOF Or MDB_Rst_H.BOF Then
                MDB_Rst_H.AddNew            '���݂��Ă�����f�[�^��ǉ�����
            End If
            MDB_Rst_H.Fields("Time").Value = dmc
            MDB_Rst_H.Fields("Minute").Value = Minute(dm)
            For i = 1 To 135
                i1 = Format(i, "###")
                MDB_Rst_H.Fields(i1).Value = sr(i)
            Next i
            MDB_Rst_H.Update
            MDB_Rst_H.Close
        Next m2
        dm = DateAdd("n", 10, dm)
    Next m1
SKIP:
'    dynOra.Close
'    Set dynOra = Nothing
'    OracleDB.Label3 = "�I���N�����C�ے����[�_�[�����\���I��"
'    OracleDB.Label3.Refresh
    Set MDB_Rst_H = Nothing
End Sub

'******************************************************************************
'�T�u���[�`���FORA_F_RADAR()
'�����T�v�F
'FRICS���[�_�[�\���J��
'FRICS���[�_�[�\���J�ʂ̓f�[�^�ʂ������̂ő����Ԏ擾���~�߂�
'�P���Ԏ擾�̃��[�e�B���Ƃ����B
'******************************************************************************
Sub ORA_F_RADAR(d1 As Date, irc As Boolean)
    Dim SQL          As String
    Dim sql_SELECT   As String
    Dim sql_WHERE1   As String
    Dim MDB_SQL      As String
    Dim SDATE        As String
    Dim n            As Integer
    Dim i            As Long
    Dim j            As Long
    Dim ir           As Long
    Dim im           As Long
    Dim Timew        As String
    Dim w1(18, 315)  As Single              '1km���b�V���l
    Dim w2(315)      As Single              '1km���b�V���l
    Dim rr
    Dim MS           As Long
    Dim ruika        As Long
    Dim Dim2         As String
    Dim buf          As String
    Dim rrr(135)     As Single
    Dim dw           As Date
    Dim DC           As String
    Dim i1           As Long
    Dim nf           As Long
    Dim Minutew      As Long
    Dim Mesh         As String
    Dim MMS          As Long
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/04 O.OKADA�y06�z
    '��OracleDB.Check_F_RADAR_Time()���C�����邱�ƁB�y06-01�z
    '******************************************************
    Exit Sub
    '******************************************************
    'Ver1.0.0 �C���I�� 2015/08/04 O.OKADA�y06�z
    '******************************************************
    '�g�p����2�����b�V���ԍ�
    '533607
    '533606
    '523770
    '523677
    '523676
    '523760
    '523667
    '523666
    '523656
    '523646
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'On Error GoTo ERR1
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    Const Ksk = -99#
    Set MDB_Rst_H.ActiveConnection = MDB_Con
    SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    OracleDB.Label3 = "�I���N�����FRICS�\�����[�_�f�[�^�J�ʎ擾��"
    OracleDB.Label3.Refresh
    '******************************************************
    'SELECT
    '******************************************************
    sql_SELECT = "SELECT * FROM oracle.F_RADAR WHERE "
    '******************************************************
    'WHERE1
    '******************************************************
    sql_WHERE1 = "jikoku = TO_DATE(" & SDATE & ") AND "
    SQL = sql_SELECT & sql_WHERE1 & Dim2_WHERE2
    ORA_LOG " SQL= " & SQL
    '******************************************************
    'SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����B
    '******************************************************
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    '******************************************************
    '�t�B�[���h�����擾����B
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'Dim Tw
    'n = dynOra.Fields.Count
    'For i = 0 To n - 1
    '    Tw = dynOra.Fields(i).Name
    '    Debug.Print " Number=" & Format(Str(i), "@@@") & " �t�B�[���h��="; Tw
    'Next i
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    If dynOra.EOF And dynOra.BOF Then
        ORA_LOG " SQL = " & SQL & " ���̃f�[�^��ORACLE-DB�ɂȂ�"
        ORA_LOG "FRICS���[�_�[�\���J�� �X�L�b�v"
        OracleDB.Label3 = "FRICS�\�����[�_�f�[�^�J�ʎ擾���s"
        OracleDB.Label3.Refresh
        GoTo SKIP
    End If
    Erase w1
    MMS = 0
    Do Until dynOra.EOF
        ORA_LOG " jikoku     =" & dynOra.Fields("JIKOKU").Value
        ORA_LOG " DATA_STATUS=" & dynOra.Fields("DATA_STATUS").Value
        Mesh = dynOra.Fields("LATITUDE").Value & dynOra.Fields("LONGITUDE").Value & Format(dynOra.Fields("CODE").Value, "00")
        ORA_LOG " Mesh       =" & Mesh
        Minutew = dynOra.Fields("RUIKA_DATE").Value / 10
        buf = dynOra.Fields("RADAR").Value
        rr = Split(buf, ",")
        Select Case Mesh
            Case "533607"
                 MS = 1
            Case "533606"
                 MS = 2
            Case "523770"
                 MS = 3
            Case "523677"
                 MS = 4
            Case "523676"
                 MS = 5
            Case "523760"
                 MS = 6
            Case "523667"
                 MS = 7
            Case "523666"
                 MS = 8
            Case "523656"
                 MS = 9
            Case "523646"
                 MS = 10
            Case Else
                 GoTo NOP
        End Select
        MMS = MMS + MS
        For i = 1 To Dim2_mesh_Number(MS)
            ir = Dim2_To_315(MS, i).Rn
            im = Dim2_To_315(MS, i).Mn - 1
            If rr(im) > 250 Then
                w1(Minutew, ir) = 0
            Else
                w1(Minutew, ir) = CSng(rr(im))
            End If
        Next i
NOP:
        dynOra.MoveNext
    Loop
SKIP:
    dynOra.Close
    ORA_LOG " MMS        =" & Format(MMS, "#0") & "  = 990 �łȂ��Ƃ��������B"
    DC = Format(d1, "yyyy/mm/dd hh:nn")
    For ruika = 1 To 18
        For i = 1 To 315
            w2(i) = w1(ruika, i)
        Next i
        Mesh_To_Ryuiki w2(), rrr(), irc
        '**************************************************
        'MDB��OPEN����B
        '**************************************************
        
        '**************************************************
        'MDB�ɏ������ށB
        '**************************************************
        MDB_SQL = "select * from FRICS���[�_�[�\�� where Time = '" & DC & "' AND Prediction_Minute =" & Str(ruika * 10)
        MDB_Rst_H.Open MDB_SQL, MDB_Con, adOpenDynamic, adLockOptimistic
        ORA_LOG "MDB FRICS�\�����[�_�\���I�[�v�� SQL=" & MDB_SQL
        ORA_LOG "MDB_Rst_H.BOF=" & MDB_Rst_H.BOF
        ORA_LOG "MDB_Rst_H.EOF=" & MDB_Rst_H.EOF
        If MDB_Rst_H.EOF Or MDB_Rst_H.BOF Then
            MDB_Rst_H.AddNew                '�Ȃ�������f�[�^��ǉ�����B
            MDB_Rst_H.Fields("Time").Value = DC
            MDB_Rst_H.Fields("Prediction_Minute").Value = ruika * 10
        End If
        For i = 1 To 135
            MDB_Rst_H.Fields(i + 1).Value = rrr(i)
        Next
        ORA_LOG "MDB FRICS�\�����[�_�\�l��������"
        MDB_Rst_H.Update
        MDB_Rst_H.Close
        ORA_LOG "MDB FRICS�\�����[�_�\�l�������ݏI�� "
    Next ruika
    ORA_LOG "FRICS�\�����[�_�f�[�^�����������݊J�n " & d1
    nf = FreeFile
    Open App.Path & "\data\F_RADAR.DAT" For Output As #nf
    Print #nf, Format(d1, "yyyy/mm/dd hh:nn")
    Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
    Close #nf
    ORA_LOG "FRICS�\�����[�_�f�[�^�����������ݏI��"
    dynOra.Close
    Set dynOra = Nothing
    Set MDB_Rst_H = Nothing
    OracleDB.Label3 = "�I���N�����FRICS���[�_�[�\���J�ʃf�[�^��I��"
    OracleDB.Label3.Refresh
    irc = True
    On Error GoTo 0
    Exit Sub
ERR1:
    On Error GoTo 0
    On Error Resume Next
    If MDB_Rst_H.State = adStateOpen Then
        MDB_Rst_H.Close
    End If
    irc = False
    dynOra.Close
    Set dynOra = Nothing
    Set MDB_Rst_H = Nothing
    OracleDB.Label3 = "�I���N�����FRICS���[�_�[�\���J�ʃf�[�^��荞�ُ݈�I��"
    OracleDB.Label3.Refresh
End Sub

'******************************************************************************
'�T�u���[�`���FORA_KANSOKU_JIKOKU_GET()
'�����T�v�F
'�e�[�u�� KANSOKU_JIKOU ����ŐV������荞�ށB
'******************************************************************************
Sub ORA_KANSOKU_JIKOKU_GET(TBL As String, dw As Date, ic As Boolean)
    Dim cDw   As String
    Dim SQL   As String
    Dim buf   As String
    Dim n     As Long
    If TBL <> "F_MESSYU_10MIN_2" Then
        SQL = "SELECT * FROM oracle.KANSOKU_JIKOKU WHERE TABLE_NAME='" & TBL & "'"
    Else
        SQL = "SELECT * FROM oracle.KANSOKU_JIKOKU WHERE TABLE_NAME='" & TBL & "' AND DETAIL = 2"
    End If
    '******************************************************
    'SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����B
    '******************************************************
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    '******************************************************
    '�t�B�[���h�����擾����B
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'Dim Tw, i
    'n = dynOra.Fields.Count
    'For i = 0 To n - 1
    '    Tw = dynOra.Fields(i).Name
    '    Debug.Print " Number=" & Format(Str(i), "@@@") & " �t�B�[���h��="; Tw
    '    Debug.Print " Value=" & dynOra.Fields(i).Value
    'Next i
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    '�c�a���e���擾����B
    '******************************************************
    Dim w1, w2, w3
    n = FreeFile
    Open App.Path & "\KANSOKU_JIKOKU.dat" For Output As #n
    w1 = dynOra.Fields(0).Value
    w2 = dynOra.Fields(1).Value
    w3 = dynOra.Fields(2).Value
    If IsNull(w1) Then
        ORA_LOG "Error IN  ORA_KANSOKU_JIKOKU_GET Field=(" & TBL & ")�̃e�[�u���Q�Ǝ���NULL���A���Ă���"
        ORA_LOG "SQL= (" & SQL & ")"
        ic = False
        GoTo JUMP1
    End If
    buf = Format(w1, "yyyy/mm/dd hh:nn") & "  "
    buf = buf & Format(w2, "@@@@@@@@@@@@@@@") & "  "
    buf = buf & Format(w3, "yyyy/mm/dd hh:nn") & "  "
    Print #n, buf
    '******************************************************
    '
    '******************************************************
    ic = True
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'w3 = "2005/08/05 18:20"
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    dw = CDate(w3)
JUMP1:
    DoEvents
    Close #n
    dynOra.Close
    Set dynOra = Nothing
End Sub

'******************************************************************************
'�T�u���[�`���FORA_KANSOKU_JIKOKU_PUT()
'�����T�v�F
'�e�[�u�� KANSOKU_JIKOU �ɍŐV�����������ށB
'******************************************************************************
Sub ORA_KANSOKU_JIKOKU_PUT(TBL As String, dw As Date)
    Dim cDw   As String
    Dim SQL   As String
    Dim buf   As String
    Dim n     As Long
    SQL = "SELECT * FROM oracle.KANSOKU_JIKOKU"     'WHERE TABLE_NAME=" & TBL
    '******************************************************
    'SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����B
    '******************************************************
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    '******************************************************
    '�t�B�[���h�����擾����B
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'Dim Tw, i, n
    'n = dynOra.Fields.Count
    'For i = 0 To n - 1
    '    Tw = dynOra.Fields(i).Name
    '    Debug.Print " Number=" & Format(Str(i), "@@@") & " �t�B�[���h��="; Tw
    'Next i
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    '�c�a�̓��e���擾����B
    '******************************************************
    Dim w1, w2, w3
    n = FreeFile
    Open App.Path & "\KANSOKU_JIKOKU.dat" For Output As #n
    Do Until dynOra.EOF
        w1 = dynOra.Fields(0).Value
        w2 = dynOra.Fields(1).Value
        w3 = dynOra.Fields(2).Value
        buf = Format(w1, "yyyy/mm/dd hh:nn") & "  "
        buf = buf & Format(w2, "@@@@@@@@@@@@@@@") & "  "
        buf = buf & Format(w3, "yyyy/mm/dd hh:nn") & "  "
        Print #n, buf
        dynOra.MoveNext
    Loop
    '******************************************************
    '
    '******************************************************
    If dynOra.EOF Then
        dynOra.AddNew
    Else
        dynOra.Edit
    End If
    cDw = Format(Now, "yyyy/mm/dd hh:nn")
    dynOra.Fields("write_time").Value = cDw
    dynOra.Fields("table_name").Value = TBL
    cDw = Format(jgd, "yyyy/mm/dd hh:nn")
    dynOra.Fields("last_date_time").Value = cDw
    dynOra.Update
    dynOra.Close
    DoEvents
    Close #n
    dynOra.Close
    Set dynOra = Nothing
End Sub

'******************************************************************************
'�T�u���[�`���FORA_LOG()
'�����T�v�F
'******************************************************************************
Sub ORA_LOG(msg As String)
    OracleDB.List1.AddItem Format$(Now, "MM:DD:HH:NN:SS") & " " & msg
    OracleDB.List1.ListIndex = OracleDB.List1.NewIndex
    If OracleDB.List1.ListIndex > 30000 Then
        Close #LOG_N
        LOG_N = FreeFile
        Open App.Path & LOG_File For Output As #LOG_N
        OracleDB.List1.Clear
    End If
    Print #LOG_N, Format(Now, "yyyy/mm/dd hh:nn:ss") & "  " & msg
    OracleDB.Time_Disp = OracleDB.List1.ListIndex
    OracleDB.Time_Disp.Refresh
End Sub

'******************************************************************************
'�T�u���[�`���FORA_Message_Out
'�����T�v�F
'�v�Z�󋵂��c�a�ɏ������ށB
'******************************************************************************
Sub ORA_Message_Out(Place As String, msg As String, Lebel As Long)
    Exit Sub
    Dim i        As Long
    Dim SQL      As String
    Dim WHERE    As String
    Dim Code(2)  As String
    Dim Ndate    As String
    Dim dw       As Date
    Dim rc       As Boolean
    Dim Obs_Time As Long
    Code(1) = "1"                           '���v�Z�l
    Code(2) = "2"                           '�v�Z�s��
    ORA_LOG "IN  ORA_Message_Out"
    ORA_LOG "    msg=" & msg
    If msg = "" Then
        Exit Sub
    End If
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'If DBX_ora = False Then
    '    '�I���N���T�[�o�[�ɃA�b�v���Ȃ�
    '    Exit Sub
    'End If
    'Exit Sub  '���}���u
    'For Obs_Time = 1 To 10
    '    ORA_DataBase_Connection rc
    '    If rc Then GoTo JUMP1
    'Next Obs_Time
    'ORA_LOG "    �I���N���ɂȂ���Ȃ��̂Ń��b�Z�[�W�o�͂�������߂܂�"
    'Exit Sub
'JUMP1:
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    Obs_Time = 0
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'SQL = "SELECT * FROM oracle.CAL_MESSAGE"
    'Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    '�t�B�[���h�����擾����B
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'Dim Tw, n
    'n = dynOra.Fields.Count
    'For i = 0 To n - 1
    '    Tw = dynOra.Fields(i).name
    '    Debug.Print " Number=" & Format(str(i), "@@@") & " �t�B�[���h��="; Tw
    'Next i
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    On Error GoTo ErrOracle
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'dw = DateAdd("s", Obs_Time, jgd)
    dw = jgd
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    Ndate = "'" & Format(dw, "yyyy/mm/dd hh:nn:ss") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'Message.Show
    'Message.Label1 = "���b�Z�[�W��DB�A�b�v��"
    'Message.ZOrder 0
    'Message.Label1.Refresh
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    SQL = "SELECT * FROM oracle.CAL_MESSAGE WHERE jikoku= TO_DATE(" & Ndate & ") "
    ORA_LOG "    SQL=" & SQL
    '******************************************************
    'SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����B
    '******************************************************
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'If dynOra.EOF Then
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
        dynOra.AddNew
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'Else
    '    dynOra.Edit
    'End If
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    '�f�[�^�������ށB
    '******************************************************
    dynOra.Fields("WRITE_TIME").Value = Format(Now, "yyyy/mm/dd hh:nn")     '�������ݎ���
    dynOra.Fields("jikoku").Value = Format(dw, "yyyy/mm/dd hh:nn:ss")
    dynOra.Fields("river_no").Value = "85053002"
    dynOra.Fields("RAIN_KIND").Value = "01"
    dynOra.Fields("error_area").Value = Place                               '��Q��
    dynOra.Fields("message").Value = msg
    dynOra.Fields("cal_level").Value = 1                                    'Lebel
    dynOra.Update
    On Error GoTo 0
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'ORA_DataBase_Close
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    Exit Sub
ErrOracle:
    '******************************************************
    '��������G���[��������
    '******************************************************
    Dim strMessage As String
    If dbOra.LastServerErr <> 0 Then
        strMessage = dbOra.LastServerErrText                                'DB�����ɂ�����G���[�̏����B
    Else
        strMessage = Err.Description                                        'DB�����ȊO�̃G���[�̏����B
    End If
    ORA_LOG "###### ERROR ERROR ERROR ERROR ERROR ERROR ERROR ######"
    ORA_LOG "IN ORA_Message_Out " & strMessage
    ORA_LOG "     SQL=" & SQL
    ORA_LOG "###### ERROR ERROR ERROR ERROR ERROR ERROR ERROR ######"
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'ORA_DataBase_Close
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    On Error GoTo 0
End Sub

'******************************************************************************
'�e�[�u�� KANSOKU_JIKOU �̓��e���o�͂���
'******************************************************************************
Sub ORA_NEW_DATA_TIME()
    Dim cDw   As String
    Dim SQL   As String
    Dim buf   As String
    Dim n     As Long
    SQL = "SELECT * FROM oracle.KANSOKU_JIKOKU"
    '******************************************************
    'SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����B
    '******************************************************
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    '******************************************************
    '�t�B�[���h�����擾����B
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'Dim tw, i
    'n = dynOra.Fields.Count
    'For i = 0 To n - 1
    '    tw = dynOra.Fields(i).Name
    '    Debug.Print " Number=" & Format(Str(i), "@@@") & " �t�B�[���h��="; tw
    'Next i
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    '�c�a�̓��e���擾����B
    '******************************************************
    Dim w1, w2, w3
    Do Until dynOra.EOF
        w1 = dynOra.Fields(0).Value
        w2 = dynOra.Fields(1).Value
        w3 = dynOra.Fields(2).Value
        buf = Format(w1, "yyyy/mm/dd hh:nn") & "  "
        buf = buf & Format(w2, "@@@@@@@@@@@@@@@") & "  "
        buf = buf & Format(w3, "yyyy/mm/dd hh:nn") & "  "
        Debug.Print buf
        dynOra.MoveNext
    Loop
    '******************************************************
    '
    '******************************************************
    dynOra.Close
    DoEvents
    dynOra.Close
    Set dynOra = Nothing
End Sub

'******************************************************************************
'�T�u���[�`���FORA_OWARI_WATER()
'�����T�v�F
'2005/05/16  ���P�[�u���o�R�̓���
'���ʃf�[�^���f�[�^�x�[�X���擾����
'�ϑ����ԍ�
'station IN( 1015,1016,1017,1019,1020 )
'1015=�V�쉺�V��F
'1016=�厡
'1017=�����O����
'1019=�v�n��
'1020=�t��
'******************************************************************************
Sub ORA_OWARI_WATER(d1 As Date, d2 As Date, ic As Boolean)
    Dim sql_SELECT   As String
    Dim sql_WHERE    As String
    Dim SQL          As String
    Dim n            As Integer
    Dim i            As Long
    Dim SDATE        As String
    Dim EDATE        As String
    Dim Timew        As String
    Dim w(5)         As Single
    Dim dw           As Date
    Dim dt           As String
    Dim A1
    Dim A2
    Dim A3
    Dim A4
    Dim f1
    Dim nf           As Integer
    Dim buf          As String
    Const Ksk = -99#
    ORA_LOG "�����ʃf�[�^�擾�J�n"
    SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    EDATE = "'" & Format(d2, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    OracleDB.Label3 = "�I���N���������ʃf�[�^�擾��"
    OracleDB.Label3.Refresh
    '******************************************************
    'SELECT
    '******************************************************
    sql_SELECT = "SELECT * FROM oracle.OWARI_WATER "
    '******************************************************
    'WHERE
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    sql_WHERE = "WHERE station IN( 1015,1016,1017,1019,1020 ) AND jikoku BETWEEN TO_DATE(" & _
                SDATE & ") AND TO_DATE(" & EDATE & ") ORDER BY jikoku"
    'sql_WHERE = "WHERE station IN( 2,16,17,18,20,21 ) and JIKOKU = TO_DATE(" & Sdate & ")"
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    SQL = sql_SELECT & sql_WHERE
    '******************************************************
    'SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����B
    '******************************************************
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    If dynOra.EOF And dynOra.BOF Then
        ORA_LOG "���ʊϑ��f�[�^���f�[�^�x�[�X�ɓo�^����Ă��܂���B" & vbCrLf & _
                 "�������m���߂Ă��������B"
        ORA_LOG "SQL=" & SQL
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        'MsgBox "���ʊϑ��f�[�^���f�[�^�x�[�X�ɓo�^����Ă��܂���B" & vbCrLf & _
                 "�������m���߂Ă��������B"
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        ic = False
        dynOra.Close
        Set dynOra = Nothing
        OracleDB.Label3 = "�I���N����萅�ʃf�[�^�掸�s"
        OracleDB.Label3.Refresh
        Exit Sub
    End If
    '******************************************************
    'MDB��OPEN����B
    '******************************************************
    Set MDB_Rst_H.ActiveConnection = MDB_Con
    nf = FreeFile
    Open App.Path & "\Data\DB_H.DAT" For Output As #nf
    '******************************************************
    '�t�B�[���h�����擾����B
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'Dim Tw
    'n = dynOra.Fields.Count
    'For i = 0 To n - 1
    '    Tw = dynOra.Fields(i).Name
    '    Debug.Print " Number=" & Format(Str(i), "@@@") & " �t�B�[���h��="; Tw
    'Next i
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    w(1) = Ksk: w(2) = Ksk: w(3) = Ksk: w(4) = Ksk: w(5) = Ksk
    dynOra.MoveFirst
    i = 0
    Timew = ""
    Do
        buf = ""
        If Not dynOra.EOF Then A1 = Str(dynOra.Fields("jikoku").Value)
        If Timew <> A1 And i > 0 Or dynOra.EOF Then
            '******************************************************
            'MDB�ɏ������ށB
            '******************************************************
            dw = CDate(Timew)
            dt = Format(dw, "yyyy/mm/dd hh:nn")
            MDB_Rst_H.Open "select * from .������ where Time = '" & dt & "' ; ", MDB_Con, adOpenDynamic, adLockOptimistic
            If MDB_Rst_H.EOF Or MDB_Rst_H.BOF Then
                MDB_Rst_H.AddNew            '���ʃf�[�^��ǉ�����B
            End If
            MDB_Rst_H.Fields("Time").Value = dt
            MDB_Rst_H.Fields("Minute").Value = Minute(dw)
            MDB_Rst_H.Fields("���V��F").Value = w(1)
            MDB_Rst_H.Fields("�厡").Value = w(2)
            MDB_Rst_H.Fields("�����O").Value = w(3)
            MDB_Rst_H.Fields("�v�n��").Value = w(4)
            MDB_Rst_H.Fields("�t��").Value = w(5)
            MDB_Rst_H.Update
            MDB_Rst_H.Close
            w(1) = Ksk: w(2) = Ksk: w(3) = Ksk: w(4) = Ksk: w(5) = Ksk
            Timew = A1
            'Ver0.0.0 �C���J�n 1900/01/01 00:00
            'Pump_Check dt, dw, w()
            'Ver0.0.0 �C���I�� 1900/01/01 00:00
            If dynOra.EOF Then Exit Do
        End If
        If i = 0 Then Timew = A1
        i = i + 1
        A2 = dynOra.Fields("station").Value
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        'A3 = dynOra.Fields("water_flag").Value
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        A4 = dynOra.Fields("water_data").Value
        f1 = dynOra.Fields("flag").Value
        buf = buf & Format(A1, "@@@@@@@@@@@@@@@@@@@@,")
        buf = buf & Format(Str(A2), "@@@@@@@@@@,")
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        'buf = buf & Format(Str(A3), "@@@@@@@@@@")
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        buf = buf & Format(Str(A4), "@@@@@@@@@@@@@@@,")
        buf = buf & Format(Str(f1), "@@@@@")
        Print #nf, buf
        ORA_LOG Format(A1, "yyyy/mm/dd hh:nn") & "  " & A2 & " H(cm)=" & A4
        If f1 = 0 Then
            Select Case CInt(A2)
                 Case 1015                  '�V�쉺�V��F
                    w(1) = CSng(A4) * 0.01
                 Case 1016                  '�厡
                    w(2) = CSng(A4) * 0.01
                 Case 1017                  '�����O����
                    w(3) = CSng(A4) * 0.01
                 Case 1019                  '�v�n��
                    w(4) = CSng(A4) * 0.01
                 Case 1020                  '�t��
                    w(5) = CSng(A4) * 0.01
            End Select
        End If
        dynOra.MoveNext
        DoEvents
    Loop
    ic = True
    Close #nf
    dynOra.Close
    Set dynOra = Nothing
    Set MDB_Rst_H = Nothing
    OracleDB.Label3 = "�I���N���������ʃf�[�^�擾�I��"
    OracleDB.Label3.Refresh
End Sub

'******************************************************************************
'�T�u���[�`���FORA_P_RADAR()
'�����T�v�F
'FRICS���[�_�[���щJ��
'******************************************************************************
Sub ORA_P_RADAR(d1 As Date, d2 As Date, irc As Boolean)
    Dim sql_SELECT   As String
    Dim sql_WHERE1   As String
    Dim sql_WHERE2   As String
    Dim SQL          As String
    Dim n            As Integer
    Dim i            As Long
    Dim j            As Long
    Dim ir           As Long
    Dim ic           As Long
    Dim MM           As Long
    Dim SDATE        As String
    Dim Timew        As String
    Dim w1(315)      As Single              '1km���b�V���l
    Dim dw           As Date
    Dim wk           As Long
    Dim nf           As Integer
    Dim ks           As Integer
    Dim buf
    Dim m            As Integer
    Dim w(5)
    Dim i1           As String
    Dim im           As String
    Dim rc           As Boolean
    Dim rr
    Dim rrr(135)     As Single
    Dim Times        As Long
    Dim Tim          As Long
    Dim MS           As Long
    Dim MMS          As Long
    Dim ds           As String
    Dim Mesh         As String
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/04 O.OKADA�y09�z
    '��OracleDB.Check_P_RADAR_Time()���C�����邱�ƁB�y09-01�z
    '******************************************************
    Exit Sub
    '******************************************************
    'Ver1.0.0 �C���I�� 2015/08/04 O.OKADA�y09�z
    '******************************************************
    '�g�p����2�����b�V���ԍ�
    '533607
    '533606
    '523770
    '523677
    '523676
    '523760
    '523667
    '523666
    '523656
    '523646
    Const Ksk = -99#
    Times = DateDiff("n", d1, d2) / 10 + 1
    dw = d1
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'On Error GoTo ERR1
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    Set MDB_Rst_H.ActiveConnection = MDB_Con
    OracleDB.Label3 = "�I���N�����FRICS���у��[�_�f�[�^�J�ʎ擾��"
    OracleDB.Label3.Refresh
    '******************************************************
    'SELECT
    '******************************************************
    sql_SELECT = "SELECT JIKOKU,DATA_STATUS,LATITUDE,LONGITUDE,CODE,RADAR FROM oracle.P_RADAR "
    For Tim = 1 To Times
        SDATE = "'" & Format(dw, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
        sql_WHERE1 = "WHERE JIKOKU= TO_DATE(" & SDATE & ") AND "
        Erase w1, rrr
        SQL = sql_SELECT & sql_WHERE1 & Dim2_WHERE2
        ORA_LOG " SQL= " & SQL
        '******************************************************
        'SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����B
        '******************************************************
        Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
        '******************************************************
        '�t�B�[���h�����擾����B
        '******************************************************
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        'Dim Tw
        'n = dynOra.Fields.Count
        'For i = 0 To n - 1
        '    Tw = dynOra.Fields(i).Name
        '    Debug.Print " Number=" & Format(Str(i), "@@@") & " �t�B�[���h��="; Tw
        'Next i
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        If dynOra.EOF And dynOra.BOF Then
            ORA_LOG " SQL = " & SQL & " ���̃f�[�^��ORACLE-DB�ɂȂ�"
            ORA_LOG "FRICS���[�_�[���щJ�� �X�L�b�v"
            OracleDB.Label3 = "FRICS���[�_�f�[�^���щJ�ʎ擾���s"
            OracleDB.Label3.Refresh
            GoTo SKIP
        End If
        MMS = 0
        Do Until dynOra.EOF
            ORA_LOG " jikoku     =" & dynOra.Fields("JIKOKU").Value
            ORA_LOG " DATA_STATUS=" & dynOra.Fields("DATA_STATUS").Value
            Mesh = dynOra.Fields("LATITUDE").Value & dynOra.Fields("LONGITUDE").Value & Format(dynOra.Fields("CODE").Value, "00")
            ORA_LOG " Mesh       =" & Mesh
            buf = dynOra.Fields("RADAR").Value
            rr = Split(buf, ",")
            Select Case Mesh
                Case "533607"
                     MS = 1
                Case "533606"
                     MS = 2
                Case "523770"
                     MS = 3
                Case "523677"
                     MS = 4
                Case "523676"
                     MS = 5
                Case "523760"
                     MS = 6
                Case "523667"
                     MS = 7
                Case "523666"
                     MS = 8
                Case "523656"
                     MS = 9
                Case "523646"
                     MS = 10
                Case Else
                     GoTo NOP
            End Select
            MMS = MMS + MS
            For i = 1 To Dim2_mesh_Number(MS)
                ir = Dim2_To_315(MS, i).Rn
                im = Dim2_To_315(MS, i).Mn - 1
                If rr(im) > 250 Then
                    w1(ir) = 0
                Else
                    w1(ir) = rr(im)
                End If
            Next i
NOP:
            dynOra.MoveNext
        Loop
SKIP:
        dynOra.Close
        ORA_LOG " MMS        =" & Format(MMS, "#0")
        Mesh_To_Ryuiki w1(), rrr(), irc
        '**************************************************
        'MDB��OPEN����B
        '**************************************************
        
        '**************************************************
        'MDB�ɏ������ށB
        '**************************************************
        ds = Format(dw, "yyyy/mm/dd hh:nn")
        MDB_Rst_H.Open "select * from .FRICS���[�_�[���� where Time = '" & ds & "' ; ", MDB_Con, adOpenDynamic, adLockOptimistic
        If MDB_Rst_H.EOF Or MDB_Rst_H.BOF Then
            MDB_Rst_H.AddNew                '���݂��Ă�����f�[�^��ǉ�����B
            MDB_Rst_H.Fields("Time").Value = ds
            MDB_Rst_H.Fields("Minute").Value = Minute(ds)
        End If
        For i = 1 To 135
            i1 = Format(i, "###")
            MDB_Rst_H.Fields(i1).Value = rrr(i)
        Next
        MDB_Rst_H.Update
        MDB_Rst_H.Close
        Erase rrr
        ORA_LOG "FRICS���у��[�_�f�[�^�����������݊J�n " & dw
        nf = FreeFile
        Open App.Path & "\data\P_RADAR.DAT" For Output As #nf
        Print #nf, Format(dw, "yyyy/mm/dd hh:nn")
        Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
        Close #nf
        ORA_LOG "FRICS���у��[�_�f�[�^���ю����������ݏI��"
        dw = DateAdd("n", 10, dw)
    Next Tim
    dynOra.Close
    Set dynOra = Nothing
    Set MDB_Rst_H = Nothing
    OracleDB.Label3 = "�I���N����FRICS���у��[�_�f�[�^���ю�荞�ݏI��"
    OracleDB.Label3.Refresh
    On Error GoTo 0
    Exit Sub
ERR1:
    If MDB_Rst_H.State = 1 Then
        MDB_Rst_H.Close
    End If
    dynOra.Close
    Set dynOra = Nothing
    Set MDB_Rst_H = Nothing
    OracleDB.Label3 = "�I���N�����FRICS���у��[�_�f�[�^���шُ�I��"
    OracleDB.Label3.Refresh
    On Error GoTo 0
End Sub

'******************************************************************************
'�T�u���[�`���FORA_Araizeki()
'�����T�v�F
'�􉁃f�[�^���f�[�^�x�[�X���擾����
' Number=  1 �t�B�[���h��=JIKOKU
' Number=  2 �t�B�[���h��=SUII
' Number=  3 �t�B�[���h��=ETURYU_NOW
' Number=  4 �t�B�[���h��=ETURYU_010
' Number=  5 �t�B�[���h��=ETURYU_020
' Number=  6 �t�B�[���h��=ETURYU_030
' Number=  7 �t�B�[���h��=ETURYU_040
' Number=  8 �t�B�[���h��=ETURYU_050
' Number=  9 �t�B�[���h��=ETURYU_060
' Number= 10 �t�B�[���h��=ETURYU_070
' Number= 11 �t�B�[���h��=ETURYU_080
' Number= 12 �t�B�[���h��=ETURYU_090
' Number= 13 �t�B�[���h��=ETURYU_100
' Number= 14 �t�B�[���h��=ETURYU_110
' Number= 15 �t�B�[���h��=ETURYU_120
' Number= 16 �t�B�[���h��=ETURYU_130
' Number= 17 �t�B�[���h��=ETURYU_140
' Number= 18 �t�B�[���h��=ETURYU_150
' Number= 19 �t�B�[���h��=ETURYU_160
' Number= 20 �t�B�[���h��=ETURYU_170
' Number= 21 �t�B�[���h��=ETURYU_180
' Number= 22 �t�B�[���h��=STATION
'******************************************************************************
Sub ORA_Araizeki(d1 As Date, d2 As Date, ic As Boolean)
    Dim sql_SELECT   As String
    Dim sql_WHERE    As String
    Dim SQL          As String
    Dim n            As Integer
    Dim i            As Long
    Dim SDATE        As String
    Dim EDATE        As String
    Dim Timew        As String
    Dim w(6)         As Single
    Dim dw           As Date
    Dim dt           As String
    Dim f1
    Dim nf           As Integer
    Dim buf          As String
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/04 O.OKADA�y01�z
    '��OracleDB.frm Check_Araizeki_Time()���C�����邱�ƁB�y01-01�z
    '******************************************************
    Exit Sub
    '******************************************************
    'Ver1.0.0 �C���I�� 2015/08/04 O.OKADA�y01�z
    '******************************************************
    Const Ksk = -99#
    ORA_LOG "�􉁃f�[�^�擾�J�n"
    SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    EDATE = "'" & Format(d2, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    OracleDB.Label3 = "�I���N�����􉁃f�[�^�擾��"
    OracleDB.Label3.Refresh
    '******************************************************
    'SELECT
    '******************************************************
    sql_SELECT = "SELECT * FROM oracle.ARAIZEKI "
    '******************************************************
    '�t�B�[���h�����擾����B
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'SQL = sql_SELECT & sql_WHERE
    'Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    'Dim Tw
    'n = dynOra.Fields.Count
    'For i = 0 To n - 1
    '    Tw = dynOra.Fields(i).Name
    '    Debug.Print " Number=" & Format(Str(i), "@@@") & " �t�B�[���h��="; Tw
    'Next i
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    '�t�B�[���h���e���擾����B
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'Dim Tw
    'd1 = "2002/06/24 11:10"
    'd2 = "2002/06/24 16:50"
    'SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'EDATE = "'" & Format(d2, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'sql_WHERE = " WHERE  jikoku BETWEEN TO_DATE(" & SDATE & ") AND TO_DATE(" & EDATE & ") "
    'SQL = sql_SELECT & sql_WHERE & " order by jikoku"
    'Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    'n = dynOra.Fields.Count
    'Do Until dynOra.EOF
    '    Tw = ""
    '    For i = 0 To 22
    '        Tw = Tw & "  " & dynOra.Fields(i).Value
    '        DoEvents
    '    Next i
    '    Debug.Print Tw
    '    dynOra.MoveNext
    'Loop
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    'WHERE
    '******************************************************
    sql_WHERE = "WHERE  jikoku BETWEEN TO_DATE(" & _
                SDATE & ") AND TO_DATE(" & EDATE & ") ORDER BY jikoku"
    SQL = sql_SELECT & sql_WHERE
    '******************************************************
    'SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����B
    '******************************************************
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    If dynOra.EOF And dynOra.BOF Then
        ORA_LOG "�􉁊ϑ��f�[�^���f�[�^�x�[�X�ɓo�^����Ă��܂���B" & vbCrLf & _
                 "�������m���߂Ă��������B"
        ORA_LOG "SQL=" & SQL
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        'MsgBox "�􉁊ϑ��f�[�^���f�[�^�x�[�X�ɓo�^����Ă��܂���B" & vbCrLf & _
                 "�������m���߂Ă��������B"
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        ic = False
        dynOra.Close
        Set dynOra = Nothing
        OracleDB.Label3 = "�I���N�����􉁃f�[�^�擾���s"
        OracleDB.Label3.Refresh
        Exit Sub
    End If
    '******************************************************
    'MDB��OPEN����B
    '******************************************************
    Set MDB_Rst_H.ActiveConnection = MDB_Con
    w(1) = Ksk: w(2) = Ksk: w(3) = Ksk: w(4) = Ksk: w(5) = Ksk: w(6) = Ksk
    dynOra.MoveFirst
    Do Until dynOra.EOF
        If Not IsNull(dynOra.Fields("jikoku").Value) Then
            buf = ""
            dw = CDate(dynOra.Fields("jikoku").Value)
            Timew = Format(dw, "yyyy/mm/dd hh:nn")
            If IsNumeric(dynOra.Fields("eturyu_now").Value) Then
                w(1) = CSng(dynOra.Fields("eturyu_now").Value) * 0.001 '�P�ʂ̊m�F 2002/06/24 18:00 Frics
            Else
                w(1) = 0#
            End If
            If IsNumeric(dynOra.Fields("eturyu_060").Value) Then
                w(2) = CSng(dynOra.Fields("eturyu_060").Value) * 0.001
            Else
                w(2) = 0#
            End If
            If IsNumeric(dynOra.Fields("eturyu_120").Value) Then
                w(3) = CSng(dynOra.Fields("eturyu_120").Value) * 0.001
            Else
                w(3) = 0#
            End If
            If IsNumeric(dynOra.Fields("eturyu_180").Value) Then
                w(4) = CSng(dynOra.Fields("eturyu_180").Value) * 0.001
            Else
                w(4) = 0#
            End If
            '******************************************************
            'MDB�ɏ������ށB
            '******************************************************
            MDB_Rst_H.Open "select * from .�� where Time = '" & Timew & "' ; ", MDB_Con, adOpenDynamic, adLockOptimistic
            If MDB_Rst_H.EOF Or MDB_Rst_H.BOF Then
                MDB_Rst_H.AddNew            '�􉁃f�[�^��ǉ�����B
            End If
            MDB_Rst_H.Fields("Time").Value = Timew
            MDB_Rst_H.Fields("Minute").Value = Minute(dw)
            MDB_Rst_H.Fields("Q0").Value = w(1)
            MDB_Rst_H.Fields("Q1").Value = w(2)
            MDB_Rst_H.Fields("Q2").Value = w(3)
            MDB_Rst_H.Fields("Q3").Value = w(4)
            MDB_Rst_H.Update
            MDB_Rst_H.Close
        End If
        dynOra.MoveNext
        DoEvents
    Loop
    dynOra.Close
    Set dynOra = Nothing
    Set MDB_Rst_H = Nothing
    OracleDB.Label3 = "�I���N�����􉁃f�[�^��I��"
    OracleDB.Label3.Refresh
End Sub

'******************************************************************************
'�T�u���[�`���FORA_P_MESSYU_10MIN
'�����T�v�F
'�C�ے�10���������b�V���f�[�^(10���J�ʁj
'�U�����Z���Ȃ��Ǝ��ԉJ�ʂɂȂ�Ȃ�
'�l�c�a�ɂ͂P�O�����敽�ωJ�ʂƂ��Ċi�[����
' Number=  0 �t�B�[���h��=WRITE_TIME
' Number=  1 �t�B�[���h��=JIKOKU
' Number=  2 �t�B�[���h��=COUNT
' Number=  3 �t�B�[���h��=SIZE_LN
' Number=  4 �t�B�[���h��=SIZE_LE
' Number=  5 �t�B�[���h��=SEKISAN_TIME
' Number=  6 �t�B�[���h��=TANI
' Number=  7 �t�B�[���h��=START_TIME
' Number=  9 �t�B�[���h��=TIME_SPAN
' Number= 10 �t�B�[���h��=YOSOKU_TIME
' Number= 11 �t�B�[���h��=CODE
' Number= 12 �t�B�[���h��=DATA_11
' Number= 13 �t�B�[���h��=DATA_12
' Number= 15 �t�B�[���h��=DATA_14
' Number= 16 �t�B�[���h��=DATA_15
' Number= 17 �t�B�[���h��=DATA_21
' Number= 18 �t�B�[���h��=DATA_22
' Number= 19 �t�B�[���h��=DATA_23
' Number= 20 �t�B�[���h��=DATA_24
' Number= 21 �t�B�[���h��=DATA_25
' Number= 22 �t�B�[���h��=DATA_31
' Number= 23 �t�B�[���h��=DATA_32
' Number= 25 �t�B�[���h��=DATA_34
' Number= 26 �t�B�[���h��=DATA_35
' Number= 27 �t�B�[���h��=DATA_41
' Number= 28 �t�B�[���h��=DATA_42
' Number= 29 �t�B�[���h��=DATA_43
' Number= 30 �t�B�[���h��=DATA_44
' Number= 31 �t�B�[���h��=DATA_45
' Number= 32 �t�B�[���h��=DATA_51
' Number= 33 �t�B�[���h��=DATA_52
' Number= 34 �t�B�[���h��=DATA_53
' Number= 35 �t�B�[���h��=DATA_54
' Number= 36 �t�B�[���h��=DATA_55
'******************************************************************************
Sub ORA_P_MESSYU_10MIN(d1 As Date, d2 As Date, ic As Boolean)
    Dim sql_SELECT   As String
    Dim sql_WHERE1   As String
    Dim sql_WHERE2   As String
    Dim SQL          As String
    Dim SDATE        As String
    Dim EDATE        As String
    Dim n            As Integer
    Dim i            As Long
    Dim j            As Long
    Dim k            As Long
    Dim Wdate        As String
    Dim Timew        As String
    Dim w1(315)      As Single              '1km���b�V���l
    Dim w2(250)      As Single              '2km���b�V���l
    Dim dw           As Date
    Dim dt           As String
    Dim Ntime        As Long
    Dim MM           As Long
    Dim nn           As Long
    Dim nf           As Integer
    Dim ks           As Integer
    Dim mes(10)      As String              '2�����b�V���ԍ�
    Dim FM(25)       As String              '�c�a��2�����b�V���ԍ�
    Dim i1           As String
    Dim j1           As String
    Dim buf          As String
    Dim irc          As Boolean
    Dim m            As Integer
    Dim cr           As String
    Dim rr           As Single
    Dim rrr(135)     As Single
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/04 O.OKADA�y07�z
    '��OracleDB.Check_P_MESSYU_10MIN_Time()���C�����邱�ƁB�y07-01�z
    '******************************************************
    Exit Sub
    '******************************************************
    'Ver1.0.0 �C���I�� 2015/08/04 O.OKADA�y07�z
    '******************************************************
    mes(1) = "533606"
    mes(2) = "533607"
    mes(3) = "523676"
    mes(4) = "523677"
    mes(5) = "523770"
    mes(6) = "523666"
    mes(7) = "523667"
    mes(8) = "523760"
    mes(9) = "523656"
    mes(10) = "523646"
    n = 0
    For i = 1 To 5
        i1 = Format(i, "0")
        For j = 1 To 5
            j1 = Format(j, "0")
            n = n + 1
            FM(n) = "data_" & i1 & j1
        Next j
    Next i
    Const Ksk = -99#
    OracleDB.Label3 = "�I���N�����C�ے��P�O���������[�_�f�[�^�J�ʎ擾��"
    OracleDB.Label3.Refresh
    '******************************************************
    'SELECT
    '******************************************************
    sql_SELECT = "SELECT * FROM oracle.P_MESSYU_10MIN "
    '******************************************************
    'WHERE1
    '******************************************************
    sql_WHERE1 = "WHERE code IN( '533606','533607','523676','523677','523770'," & _
                                "'523666','523667','523760','523656','523646' ) AND "
    '******************************************************
    '�t�B�[���h�����擾����B
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'Dim Tw
    'SQL = sql_SELECT
    'Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    'n = dynOra.Fields.Count
    'For i = 0 To n - 1
    '    Tw = dynOra.Fields(i).Name
    '    Debug.Print " Number=" & Format(Str(i), "@@@") & " �t�B�[���h��="; Tw
    'Next i
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    '�t�B�[���h���e���擾����B
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'Dim Tw, sql_WHERE
    'd1 = "2002/06/10 19:00"
    'd2 = "2002/06/10 19:30"
    'SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'EDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'sql_WHERE = " WHERE  jikoku BETWEEN TO_DATE(" & SDATE & ") AND TO_DATE(" & EDATE & ") " & _
    '    "ORDER BY code,jikoku"
    'SQL = sql_SELECT & sql_WHERE
    'Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    'n = dynOra.Fields.Count
    'Do Until dynOra.EOF
    '    Tw = dynOra.Fields("jikoku").Value
    '    Tw = Tw & "  " & dynOra.Fields("code").Value
    '    Debug.Print Tw
    '    dynOra.MoveNext
    '    DoEvents
    'Loop
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    Ntime = DateDiff("n", d1, d2) / 10 + 1
    dw = d1
    For nn = 1 To Ntime
        '******************************************************
        'WHERE2
        '******************************************************
        Wdate = "'" & Format(dw, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
        sql_WHERE2 = "jikoku = TO_DATE(" & Wdate & ") "
        SQL = sql_SELECT & sql_WHERE1 & sql_WHERE2
        '******************************************************
        'SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����B
        '******************************************************
        Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
        If dynOra.EOF And dynOra.BOF Then
            'Ver0.0.0 �C���J�n 1900/01/01 00:00
            'MsgBox "�ϑ��f�[�^���f�[�^�x�[�X�ɓo�^����Ă��܂���B" & vbCrLf & _
                     "�������m���߂Ă��������B" & vbCrLf & dw
            'Ver0.0.0 �C���I�� 1900/01/01 00:00
            ORA_LOG "�C�ے����[�_�[���уf�[�^�X�L�b�v�����������݊J�n " & dt
            nf = FreeFile
            Open App.Path & "\data\P_MESSYU_10MIN.DAT" For Output As #nf
            Print #nf, dt
            Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
            Close #nf
            ORA_LOG "�C�ے����[�_�[���уf�[�^�X�L�b�v�����������ݏI��"
            ic = False
            dynOra.Close
            'Ver0.0.0 �C���J�n 1900/01/01 00:00
            'Set dynOra = Nothing
            'Ver0.0.0 �C���I�� 1900/01/01 00:00
            OracleDB.Label3 = "�C�ے��P�O���������[�_�f�[�^�J�ʎ掸���s"
            OracleDB.Label3.Refresh
            GoTo SKIP
        End If
        Erase w2
        m = 0
        Do Until dynOra.EOF
            m = m + 1
            'Ver0.0.0 �C���J�n 1900/01/01 00:00
            'dm = dynOra.Fields("jikoku").Value
            'Ver0.0.0 �C���I�� 1900/01/01 00:00
            MM = dynOra.Fields("code").Value
            Select Case MM
                Case mes(1)
                    ks = 1
                Case mes(2)
                    ks = 26
                Case mes(3)
                    ks = 51
                Case mes(4)
                    ks = 76
                Case mes(5)
                    ks = 101
                Case mes(6)
                    ks = 128
                Case mes(7)
                    ks = 151
                Case mes(8)
                    ks = 176
                Case mes(9)
                    ks = 201
                Case mes(10)
                    ks = 226
            End Select
            For i = ks To ks + 24
                j = i - ks + 1
                cr = dynOra.Fields(FM(j)).Value
                If IsNumeric(cr) Then
                    rr = CSng(cr)
                Else
                    rr = 0#
                End If
                If rr < 0# Then rr = 0#
                w2(i) = rr
            Next i
            dynOra.MoveNext
        Loop
        If m < 10 Then
            ORA_LOG "2�����b�V���̂ǂꂩ���擾�ł��Ă��Ȃ��B" & dw
        End If
        Mesh_2km_to_1km_cvt w2(), w1()
        Mesh_To_Ryuiki w1(), rrr(), irc
        '**************************************************
        'MDB��OPEN����B
        '**************************************************
        dt = Format(dw, "yyyy/mm/dd hh:nn")
        ORA_LOG " �C�ے����[�_�[���� MDB�ɏ�������" & dw
        Set MDB_Rst_H.ActiveConnection = MDB_Con
        '**************************************************
        'MDB�ɏ������ށB
        '**************************************************
        MDB_Rst_H.Open "select * from .�C�ے����[�_�[���� where Time = '" & dt & "' ; ", MDB_Con, adOpenDynamic, adLockOptimistic
        If MDB_Rst_H.EOF Or MDB_Rst_H.BOF Then
            MDB_Rst_H.AddNew                '���݂��Ă�����f�[�^��ǉ�����B
        End If
        MDB_Rst_H.Fields("Time").Value = dt
        MDB_Rst_H.Fields("Minute").Value = Minute(dw)
        For i = 1 To 135
            i1 = Format(i, "###")
            MDB_Rst_H.Fields(i1).Value = rrr(i)
            'Ver0.0.0 �C���J�n 1900/01/01 00:00
            'Debug.Print dt; " rrr="; rrr(i)
            'Ver0.0.0 �C���I�� 1900/01/01 00:00
        Next i
        MDB_Rst_H.Update
        MDB_Rst_H.Close
        ORA_LOG " �C�ے����[�_�[���� MDB�ɏ������ݏI��" & dw
        ORA_LOG "�C�ے����[�_�[���уf�[�^�����������݊J�n " & dt
        nf = FreeFile
        Open App.Path & "\data\P_MESSYU_10MIN.DAT" For Output As #nf
        Print #nf, dt
        Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
        Close #nf
        ORA_LOG "�C�ے����[�_�[���уf�[�^�����������ݏI��"
SKIP:
        dw = DateAdd("n", 10, dw)
    Next nn
    dynOra.Close
    Set dynOra = Nothing
    Set MDB_Rst_H = Nothing
    OracleDB.Label3 = "�I���N�����C�ے����[�_�[���ю擾�I��"
    OracleDB.Label3.Refresh
End Sub

'******************************************************************************
'�T�u���[�`���FORA_P_MESSYU_1Hour
'�����T�v�F
'�C�ے�10���������b�V���f�[�^(10���J�ʁj
'�U�����Z���Ȃ��Ǝ��ԉJ�ʂɂȂ�Ȃ�
'�l�c�a�ɂ͂P�O�����敽�ωJ�ʂƂ��Ċi�[����
'******************************************************************************
Sub ORA_P_MESSYU_1Hour(d1 As Date, d2 As Date, ic As Boolean)
    Dim sql_SELECT   As String
    Dim sql_WHERE    As String
    Dim SQL          As String
    Dim n            As Integer
    Dim i            As Long
    Dim j            As Long
    Dim k            As Long
    Dim SDATE        As String
    Dim EDATE        As String
    Dim Timew        As String
    Dim w1(315)      As Single              '1km���b�V���l
    Dim w2(250)      As Single              '2km���b�V���l
    Dim dw           As String
    Dim dm           As Date
    Dim MM           As Long
    Dim nf           As Integer
    Dim ks           As Integer
    Dim mes(10)      As String              '2�����b�V���ԍ�
    Dim FM(25)       As String              '�c�a��2�����b�V���ԍ�
    Dim i1           As String
    Dim j1           As String
    Dim buf          As String
    Dim irc          As Boolean
    Dim m            As Integer
    Dim rr(135)     As Single
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/04 O.OKADA�y08�z
    '��OracleDB.Check_P_MESSYU_1HOUR_Time()���C�����邱�ƁB�y08-01�z
    '******************************************************
    Exit Sub
    '******************************************************
    'Ver1.0.0 �C���I�� 2015/08/04 O.OKADA�y08�z
    '******************************************************
    mes(1) = "533606"
    mes(2) = "533607"
    mes(3) = "523676"
    mes(4) = "523677"
    mes(5) = "523770"
    mes(6) = "523666"
    mes(7) = "523667"
    mes(8) = "523760"
    mes(9) = "523656"
    mes(10) = "523646"
    n = 0
    For i = 1 To 5
        i1 = Format(i, "0")
        For j = 1 To 5
            j1 = Format(j, "0")
            n = n + 1
            FM(n) = "data_" & i1 & j1
        Next j
    Next i
    Const Ksk = -99#
    SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    EDATE = "'" & Format(d2, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    OracleDB.Label3 = "�I���N�����C�ے��P�O���������[�_�f�[�^�J�ʎ擾��"
    OracleDB.Label3.Refresh
    '******************************************************
    'SELECT
    '******************************************************
    sql_SELECT = "SELECT * FROM oracle.P_MESSYU_10MIN "
    '******************************************************
    'WHERE
    '******************************************************
    sql_WHERE = "WHERE code IN( '533606','533607','523676','523677','523770'," & _
                               "'523666','523667','523760','523656','523646' ) AND " & _
        "jikoku BETWEEN TO_DATE(" & SDATE & ") AND TO_DATE(" & EDATE & ") " & _
        "ORDER BY jikoku,code"
    SQL = sql_SELECT & sql_WHERE
    '******************************************************
    'SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����B
    '******************************************************
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    If dynOra.EOF And dynOra.BOF Then
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        'MsgBox "�C�ے��P�O���ϑ��f�[�^���f�[�^�x�[�X�ɓo�^����Ă��܂���B" & vbCrLf & _
                 "�������m���߂Ă��������B"
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        ORA_LOG "�C�ے��P�O���ϑ��f�[�^���f�[�^�x�[�X�ɓo�^����Ă��܂���B" & vbCrLf & _
                 "�������m���߂Ă��������B"
        ic = False
        dynOra.Close
        Set dynOra = Nothing
        OracleDB.Label3 = "�C�ے��P���Ԏ������[�_�f�[�^�J�ʎ掸���s"
        OracleDB.Label3.Refresh
        Exit Sub
    End If
    Do Until dynOra.EOF
        For m = 1 To 10
            dw = dynOra.Fields("jikoku").Value
            MM = dynOra.Fields("code").Value
            Select Case MM
                Case mes(1)
                    ks = 1
                Case mes(2)
                    ks = 26
                Case mes(3)
                    ks = 51
                Case mes(4)
                    ks = 76
                Case mes(5)
                    ks = 101
                Case mes(6)
                    ks = 128
                Case mes(7)
                    ks = 151
                Case mes(8)
                    ks = 176
                Case mes(9)
                    ks = 201
                Case mes(10)
                    ks = 226
            End Select
            For i = ks To ks + 24
                j = k - ks + 1
                w2(i) = dynOra.Fields(FM(j)).Value * 0.1
            Next i
        Next m
        dynOra.MoveNext
        Mesh_2km_to_1km_cvt w2(), w1()
        Mesh_To_Ryuiki w1(), rr(), irc
        '**************************************************
        'MDB��OPEN����B
        '**************************************************
        Set MDB_Rst_H.ActiveConnection = MDB_Con
        '**************************************************
        'MDB�ɏ������ށB
        '**************************************************
        MDB_Rst_H.Open "select * from .�C�ے����[�_�[���� where Time = #" & dw & "# ; ", MDB_Con, adOpenDynamic, adLockOptimistic
        If MDB_Rst_H.EOF Or MDB_Rst_H.BOF Then
            MDB_Rst_H.AddNew                '���݂��Ă�����f�[�^��ǉ�����B
        End If
        dm = CDate(dw)
        MDB_Rst_H.Fields("Time").Value = dm
        MDB_Rst_H.Fields("Minute").Value = Minute(dm)
        For i = 1 To 135
            i1 = Format(i, "###")
            MDB_Rst_H.Fields(i1).Value = rr(i)
        Next i
        MDB_Rst_H.Update
        MDB_Rst_H.Close
    Loop
    dynOra.Close
    Set dynOra = Nothing
    Set MDB_Rst_H = Nothing
    OracleDB.Label3 = "�I���N�����C�ے��P���Ԏ��уf�[�^��I��"
    OracleDB.Label3.Refresh
End Sub

'******************************************************************************
'�T�u���[�`���FORA_YOHOUBUNAN()
'�����T�v�F
'���m���T�[�o�[�ɗ\�񕶂��������ށB
'******************************************************************************
'Sub ORA_YOHOUBUNAN(Return_Code As Boolean)
'    Dim sql_SELECT   As String
'    Dim sql_WHERE    As String
'    Dim SQL          As String
'    Dim N_rec        As Long
'    Dim n            As Integer
'    Dim i            As Long
'    Dim SDATE        As String
'    Dim EDATE        As String
'    Dim jssd         As Date
'    Dim jeed         As Date
'    Dim Timew        As String
'    Dim f1           As String
'    Dim f2           As String
'    Dim B11          As String
'    Dim B12          As String
'    '******************************************************
'    'Ver1.0.0 �C���J�n 2015/08/04 O.OKADA�y10�z
'    '���\�񕶃e�X�g���M.Command1_Click()���C�����邱�ƁB�y10-01�z
'    '******************************************************
'    Exit Sub
'    '******************************************************
'    'Ver1.0.0 �C���I�� 2015/08/04 O.OKADA�y10�z
'    '******************************************************
'    jssd = CDate(C4)
'    jeed = DateAdd("n", 30, jssd)
'    SDATE = "'" & Format(jssd, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
'    EDATE = "'" & Format(jeed, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
'    '******************************************************
'    'SELECT
'    '******************************************************
'    sql_SELECT = "SELECT * FROM oracle.YOHOUBUNAN"
'    '******************************************************
'    'WHERE
'    '******************************************************
'    sql_WHERE = " WHERE  ESTIMATE_TIME = TO_DATE(" & SDATE & ")"
'    'Ver0.0.0 �C���J�n 1900/01/01 00:00
'    SQL = sql_SELECT & sql_WHERE
'    'SQL = sql_SELECT
'    'Ver0.0.0 �C���I�� 1900/01/01 00:00
'    '******************************************************
'    '�t�B�[���h�����擾����B
'    '******************************************************
'    'Ver0.0.0 �C���J�n 1900/01/01 00:00
'    'Dim Tw
'    'n = RST_YB.Fields.Count
'    'For i = 0 To n - 1
'    '    Tw = RST_YB.Fields(i).Name
'    '    Debug.Print " Number=" & Format(Str(i), "@@@") & " �t�B�[���h��="; Tw
'    'Next i
'    'Ver0.0.0 �C���I�� 1900/01/01 00:00
'    '******************************************************
'    'SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����B
'    '******************************************************
'    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
'    Dim nf As Integer
'    Dim buf As String
'    nf = FreeFile
'    Open App.Path & "\Data\DB_YB.DAT" For Output As #nf
'    If dynOra.EOF Then
'        dynOra.AddNew
'    Else
'        dynOra.Edit
'    End If
'    f1 = Format(CDate(C4), "d��h��m��")
'    f2 = Format(CDate(C4), "d��h��m��")
'    B12 = "�@�@�V��̐��ʂ�" & f2 & "���ɂͤ���̂悤�Ɍ����܂�܂��"
'    B11 = "�@�@�V��̐��ʂ�" & f1 & "���ݤ���̂Ƃ���ƂȂ��Ă��܂��" & vbLf & _
'          "�@�@�����O���ʐ��ʊϑ����m�V�쒬�厚�����n���n��" & vbLf & _
'          "  �@�@�@�@�@4.52m(�}�㏸��)" & vbLf & _
'          B12 & vbLf & _
'          "�@�@�����O���ʐ��ʊϑ����m�V�쒬�厚�����n���n��" & vbLf & _
'          "�@�@�@�@�@�@5.30m���x"
'    dynOra.Fields("WRITE_TIME").Value = C1                      '�������ݎ���
'    dynOra.Fields("DATA_KIND_CODE").Value = "�t�P���R�E�Y�C�A��01"
'    dynOra.Fields("DATA_KIND").Value = "�\�񕶈āi���ʕ����j"
'    dynOra.Fields("SENDING_STATION_CODE").Value = "23001"
'    dynOra.Fields("SENDING_STATION").Value = "���m���������ݎ�����"
'    dynOra.Fields("APPOINTED_CODE").Value = ""
'    dynOra.Fields("ESTIMATE_TIME").Value = C4
'    dynOra.Fields("PRACTICE_FLG_CODE").Value = "40"             '"40"=�\��  "99"=���K
'    dynOra.Fields("PRACTICE_FLG").Value = "�\��"                '"���K"
'    dynOra.Fields("SEQ_NO").Value = ""
'    dynOra.Fields("ANNOUNCE_TIME").Value = C5
'    dynOra.Fields("RIVER_NAME").Value = "���m�������쐅�n�@�V��"
'    dynOra.Fields("RIVER_NO_CODE").Value = "85053002"
'    dynOra.Fields("RIVER_NO").Value = "�V��"
'    dynOra.Fields("RIVER_DIV_CODE").Value = "00"
'    dynOra.Fields("RIVER_DIV").Value = ""
'    dynOra.Fields("ANNOUNCE_NO").Value = ""
'    dynOra.Fields("FORECAST_KIND").Value = C2
'    dynOra.Fields("FORECAST_KIND_CODE").Value = C3
'    dynOra.Fields("BUNSHO1").Value = B1
'    dynOra.Fields("BUNSHO2").Value = B2
'    dynOra.Fields("BUNSHO3").Value = ""
'    dynOra.Fields("RAIN_KIND").Value = "01"
'    dynOra.Update
'    dynOra.Close
'    '******************************************************
'    '�\�񕶑Ώۉ͐�
'    'SELECT
'    '******************************************************
'    sql_SELECT = "SELECT * FROM oracle.YOHOU_TARGET_RIVER"
'    '******************************************************
'    'WHERE
'    '******************************************************
'    sql_WHERE = " WHERE  ESTIMATE_TIME = TO_DATE(" & SDATE & ")"
'    SQL = sql_SELECT & sql_WHERE
'    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
'    If dynOra.EOF Then
'        dynOra.AddNew
'    Else
'        dynOra.Edit
'    End If
'    dynOra.Fields("WRITE_TIME").Value = C1                      '�������ݎ���
'    dynOra.Fields("BUNAN_CODE").Value = "01"
'    dynOra.Fields("DATA_KIND_CODE").Value = "�t�P���R�E�Y�C�A��01"
'    dynOra.Fields("SENDING_STATION_CODE").Value = "23001"
'    dynOra.Fields("ESTIMATE_TIME").Value = C4
'    dynOra.Fields("TRIVER_NAME").Value = "�V��"
'    dynOra.Fields("TRIVER_NO_CODE").Value = "85053002"
'    dynOra.Fields("TRIVER_NO").Value = "�V��"
'    dynOra.Fields("TRIVER_DIV_CODE").Value = "00"
'    dynOra.Fields("FORECAST_KIND").Value = C2
'    dynOra.Fields("FORECAST_KIND_CODE").Value = C3
'    dynOra.Fields("RAIN_KIND").Value = "01"
'    dynOra.Fields("OUT_NO").Value = 1
'    dynOra.Update
'    dynOra.Close
'    DoEvents
'    Close #nf
'    Set dynOra = Nothing
'End Sub

'******************************************************************************
'�T�u���[�`���FWAIT_Minute()
'�����T�v�F
'******************************************************************************
Sub WAIT_Minute(m As Integer)
    Dim Start_Time, End_Time
    Start_Time = Timer
    End_Time = Start_Time + m
    Do While Timer < End_Time
        DoEvents
        If Timer - Start_Time > 1 Then
            OracleDB.Label3.Caption = "Oracle DB�ɐڑ��ł��܂���A�Q���ԋx�e�� ����" & Format(End_Time - Timer, "###0") & "�b"
            Start_Time = Timer
            OracleDB.Time_Disp.Caption = " " + Format(Now, "yyyy�Nmm��dd���@hh��nn��ss�b")
            OracleDB.Time_Disp.Refresh
        End If
    Loop
    OracleDB.Label3.Caption = "��荞�ݑ҂�"
End Sub

Public Sub WaterDataNewTime(aObsTime As Date, aFlag As Boolean)

    Dim strSQL As String
    Dim strGetMinTime As String
    Dim strGetMaxTime As String
    Dim strNowTime As String
    Dim intMinute As Integer
    
    On Error GoTo WaterDataNewTime_Error
    
    aFlag = False
    strGetMinTime = vbNullString
    strGetMaxTime = vbNullString
    
    strSQL = "SELECT"
    strSQL = strSQL & "  MIN(latest_obs_time) AS min_time"
    strSQL = strSQL & ", MAX(latest_obs_time) AS max_time"
    strSQL = strSQL & "  FROM t_water_level_obs_sta_status"
    strSQL = strSQL & " WHERE obs_sta_id IN(1012, 81, 201, 91, 71, 131, 80, 130, 240)"
    
    Set gAdoRst = New ADODB.Recordset
    gAdoRst.CursorType = adOpenStatic
    gAdoRst.LockType = adLockReadOnly
    gAdoRst.Open strSQL, gAdoCon, , , adCmdText
    If Not gAdoRst.EOF Then
        If IsDate(gAdoRst!min_time) Then strGetMinTime = Format(gAdoRst!min_time, "yyyy/mm/dd hh:nn")
        If IsDate(gAdoRst!max_time) Then strGetMaxTime = Format(gAdoRst!max_time, "yyyy/mm/dd hh:nn")
    End If
    Call SQLdbsDeleteRecordset(gAdoRst)
    
    If Not IsDate(strGetMinTime) And Not IsDate(strGetMaxTime) Then
        ORA_LOG "Error IN  ���ʊϑ��f�[�^�A�ŐV�ϑ��������Ȃ�"
        ORA_LOG "SQL= (" & strSQL & ")"
    Else
        If DateDiff("n", strGetMinTime, strGetMaxTime) = 0 Then
            aObsTime = CDate(strGetMaxTime)
        Else
            strNowTime = Format(Now, "yyyy/mm/dd hh:nn")
            intMinute = DatePart("n", strNowTime) Mod 10
            strNowTime = Format(DateAdd("n", -(intMinute), strNowTime), "yyyy/mm/dd hh:nn")
            If intMinute >= 6 Or DateDiff("n", strGetMaxTime, strNowTime) > 0 Then
                aObsTime = CDate(strGetMaxTime)
            Else
                aObsTime = CDate(strGetMinTime)
            End If
        End If
        aFlag = True
    End If
    
    Exit Sub
WaterDataNewTime_Error:
    Dim strMessage As String
    strMessage = Err.Description
    ORA_LOG strMessage
    On Error GoTo 0

End Sub

Public Sub RadarMeshuDataNewTime(ByVal aTableName As String, aObsTime As Date, aFlag As Boolean)

    Dim strTableName As String
    Dim strSQL As String
    Dim strGetTime As String
    Const intJSTAddHour9 As Long = 540
    
    On Error GoTo RadarMeshuDataNewTime_Error
    
    aFlag = False
    strTableName = vbNullString
    strGetTime = vbNullString
    
    Select Case aTableName
        Case "VDXA70"
            strTableName = "t_excg_vdxa70"
        Case "VDXB70"
            strTableName = "t_excg_vdxb70"
        Case "VCXB70"
            strTableName = "t_excg_vcxb70"
        Case "VCXB71"
            strTableName = "t_excg_vcxb71"
        Case "VCXB75"
            strTableName = "t_excg_vcxb75"
        Case "VCXB76"
            strTableName = "t_excg_vcxb76"
        Case Else
            ORA_LOG "Error IN  ���b�V���f�[�^�̃f�[�^�x�[�X�e�[�u���Ȃ�"
            ORA_LOG "�e�[�u����= (" & aTableName & ")"
            Exit Sub
    End Select
    
    strSQL = "SELECT last_data_time"
    strSQL = strSQL & " FROM t_excg_kansoku_jikoku"
    strSQL = strSQL & " WHERE table_name='" & strTableName & "'"
    
    Set gAdoRst = New ADODB.Recordset
    gAdoRst.CursorType = adOpenStatic
    gAdoRst.LockType = adLockReadOnly
    gAdoRst.Open strSQL, gAdoCon, , , adCmdText
    If Not gAdoRst.EOF Then
        If IsDate(gAdoRst!last_data_time) Then strGetTime = Format(DateAdd("n", intJSTAddHour9, gAdoRst!last_data_time), "yyyy/mm/dd hh:nn")
    End If
    Call SQLdbsDeleteRecordset(gAdoRst)
    
    If Not IsDate(strGetTime) Then
        ORA_LOG "Error IN  ���b�V���f�[�^�A�ŐV�ϑ��������Ȃ�"
        ORA_LOG "SQL= (" & strSQL & ")"
    Else
        aObsTime = CDate(strGetTime)
        aFlag = True
    End If
    
    Exit Sub
RadarMeshuDataNewTime_Error:
    Dim strMessage As String
    strMessage = Err.Description
    ORA_LOG strMessage
    On Error GoTo 0

End Sub
