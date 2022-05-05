Attribute VB_Name = "OraPump"
Option Explicit
Option Base 1
Public Pump()           As String
Public Bit_Num(16)      As Long
Public Pump_Stats(20)   As Pump

Type Pump
    Name    As String  '�|���v��
    P_Code  As Long    '�|���v��R�[�h
    S_Num   As Long    '�v���O�����㏇�ԃR�[�h
    P_Num   As Long    '�|���v�䐔
    sv_Num  As Long    'sv�ԍ�
End Type

Sub Bit_Intial()
    Dim i As Long
    Dim j As Long
    Bit_Num(1) = 32768
    For j = 2 To 16
        i = j - 1
        Bit_Num(j) = Bit_Num(i) / 2
    Next j
End Sub

'�|���v�̃r�b�g��Ԃ𒲂ׂ�
'����
'     n ------- 16�r�b�g�̃|���v���
'�o��
'    na() ----- 16�̃A���[�Ń|���v���I���̏���1������B
Sub Check_Bit(n As Long, na() As Long)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim m As Long
    k = n
    m = 1
    For i = 1 To 16
        j = k And m
        If j > 0 Then na(i) = 1
        m = m * 2
    Next i
End Sub

'�|���v���эŐV�擾�����`�F�b�N
Sub Check_OWARI_PUMP(rc As Boolean)
    Exit Sub
    Dim nf  As Integer
    Dim n   As Long
    Dim d1  As Date
    Dim d2  As Date
    Dim d3  As Date
    Dim ans As Long
    Dim buf As String
    Dim irc As Boolean
    Dim ic  As Boolean
    nf = FreeFile
'    Debug.Print " Freefile="; nf
    Open App.Path & "\data\OWARI_PUMP.DAT" For Input As #nf
    Line Input #nf, buf
    d1 = CDate(buf)
    Close #nf
'    d3 = Format(Now, "yyyy/mm/dd hh") & ":00"
'    n = DateDiff("h", d1, d3)
'    If n > 25 Then
'        d1 = DateAdd("h", -25, d3) '�O��I�����Q�T���Ԃ��O�������̂Ŏ�荞�݊J�n������ύX
'    End If
    ORA_KANSOKU_JIKOKU_GET "OWARI_SV", d2, irc
    If irc = False Then
        rc = irc
        GoTo JUMP
    End If
    If d2 > d1 Then
'        n = DateDiff("h", d1, d2) + 1
'        If n > 100 Then '100�͓K���Ɍ��߂��l�A�v�����if���Ɉ���������Ȃ��悤�ɂ����B2002/08/07 in YOKOHAMA
'            ans = MsgBox("�ǉ��Ŏ擾���悤�Ƃ��Ă���|���v�f�[�^�X�e�b�v���Q�S������" & vbCrLf & _
'                         "�Ԋu������܂��B��Ƃ��p�����܂����H" & vbCrLf & _
'                         "�V�K�̍^���v�Z�ł͂��߂邱�Ƃ����i�߂��܂��B" & vbCrLf & _
'                         "[�͂�]�ł��̃W���u�͏I�����܂��A[������]�Ōp�����܂��B", vbInformation + vbYesNo)
'            If ans = vbYes Then
'                Close
'                End
'            End If
'        End If
        d1 = DateAdd("n", 10, d1)
        ORA_LOG "�|���v�f�[�^��荞�݊J�n " & d1 & " ���� " & d2 & "�܂�"
        ORA_P_PUMP d1, d2, ic
        If Not ic Then
            ORA_LOG "�I���N���f�[�^�x�[�X���|���v�f�[�^���擾���悤�Ƃ�������" & vbCrLf & _
                    "�G���[���������Ă��܂��B"
            GoTo JUMP
        Else
            ORA_LOG "�|���v�f�[�^��荞�ݐ���I��"
            ORA_LOG "�|���v�f�[�^�����������݊J�n " & d2
            nf = FreeFile
            Open App.Path & "\data\OWARI_PUMP.DAT" For Output As #nf
            Print #nf, Format(d2, "yyyy/mm/dd hh:nn")
'            Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
            Print #nf, Format(d1, "yyyy/mm/dd hh:nn")
'            d1 = CDate(buf)
            Close #nf
            ORA_LOG "�|���v�f�[�^�����������ݏI��"
        End If
    End If
JUMP:
End Sub

'�|���v�f�[�^�擾
'����
'     Name ----------- �|���v�ꖼ
'       np ----------- ���̃v���O������̃|���v����
'     Code ----------- �|���v��R�[�h
'       sv ----------- sv�ԍ�
'      N_P ----------- �|���v��
'       d1 ----------- �f�[�^�擾�J�n����
'       d2 ----------- �f�[�^�擾�I������
'�o��
'      Pump() -------- �|���v��Ԃ����� �O���[�o���ϐ�
'          rc -------- �������
Sub Ora_OWARI_PUMP(Name As String, Code As Long, sv As Long, _
                   N_P As Long, np As Long, d1 As Date, d2 As Date, rc As Boolean)
    Dim i       As Long
    Dim j       As Long
    Dim n       As Long
    Dim m
    Dim F       As Long
    Dim P_MAX   As Single
    Dim P(16)   As Long
    Dim SQL     As String
    Dim ODATS   As String
    Dim ODATE   As String
    Dim dw      As Date
    Dim t       As Date
    Dim pw      As String
    Dim flag    As Long
    Dim jj      As Long
    Dim svc     As Long
    Dim sta     As Long
    ORA_LOG "IN    Ora_OWARI_PUMP"
    Const w = "1,1,1,1,1,1,1,1,1,1,"
    OracleDB.Label3 = "�I���N�����" & Name & "�f�[�^�擾��"
    OracleDB.Label3.Refresh
    '�Ƃ肠���������l�Ƃ��đS�|���v�I���ɂ���
    n = DateDiff("n", d1, d2) / 10 + 1
    pw = Mid(w, 1, N_P * 2)
    For j = 1 To n
        Pump(np, j) = pw
    Next j
    ODATS = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    ODATE = "'" & Format(d2, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
'    SQL = "SELECT * FROM oracle.OWARI_SV WHERE jikoku " & _
          "BETWEEN TO_DATE(" & ODATS & ") AND TO_DATE(" & ODATE & ")" & _
          " AND Station=" & Str(Code) & " AND sv_no=" & Str(sv)
    SQL = "SELECT * FROM oracle.OWARI_SV WHERE jikoku= TO_DATE(" & ODATS & ")"
    ORA_LOG " SQL=" & SQL
    ' SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    If dynOra.EOF And dynOra.BOF Then
        ORA_LOG Name & "�|���v�f�[�^���f�[�^�x�[�X�ɓo�^����Ă��܂���B" & vbCrLf & _
                 "�������m���߂Ă��������B"
        ORA_LOG "SQL=" & SQL
        rc = False
        dynOra.Close
        Set dynOra = Nothing
        OracleDB.Label3 = "�I���N�����" & Name & "�|���v�f�[�^�擾���s"
        OracleDB.Label3.Refresh
        Exit Sub
    End If
    dw = d1
    Do Until dynOra.EOF
        t = CDate(dynOra.Fields("jikoku").Value)
        n = DateDiff("n", d1, t) / 10 + 1
        m = AscB(dynOra.Fields("sv_data").Value)
        svc = dynOra.Fields("sv_no").Value
        sta = dynOra.Fields("Station").Value
        flag = dynOra.Fields("flag").Value
        ORA_LOG " Name=" & Name
        ORA_LOG " t=" & Format(t, "yyyy/mm/dd hh:nn")
        ORA_LOG " n=" & Str(n)
        ORA_LOG " m=" & Str(m)
        ORA_LOG " svc=" & Str(svc)
        ORA_LOG " sta=" & Str(sta)
        ORA_LOG " flag=" & Str(flag)
        Debug.Print " Name="; Name
        Debug.Print " t="; t
        Debug.Print " n="; n
        Debug.Print " m="; m
        Debug.Print " svc="; svc
        Debug.Print " sta="; sta
        Debug.Print " flag="; flag
        If flag = 0 Then
            pw = ""
            For i = 1 To N_P
                jj = m And Bit_Num(i)
                If jj > 0 Then
                    pw = pw & "1,"  '�|���v�ғ���
                Else
                    pw = pw & "0,"  '�|���v��~��
                End If
            Next i
            Pump(np, n) = pw
        Else
            pw = ""
            For i = 1 To N_P
                pw = pw & "1,"
            Next i
            Pump(np, n) = pw
        End If
        dynOra.MoveNext
    Loop
    dynOra.Close
    ORA_LOG "OUT   Ora_OWARI_PUMP"
End Sub

'�����|���v��p�ɍ쐬�����T�u���[�`���ł�
'
'
Sub Ora_OWARI_SUIBA_PUMP(d1 As Date, d2 As Date, rc As Boolean)
    Dim i       As Long
    Dim j       As Long
    Dim n       As Long
    Dim nn      As Long
    Dim m
    Dim F       As Long
    Dim P_MAX   As Single
    Dim P(16)   As Long
    Dim SQL     As String
    Dim ODATS   As String
    Dim ODATE   As String
    Dim dw      As Date
    Dim t       As Date
    Dim pw      As String
    Dim flag    As Long
    Dim jj      As Long
    Dim svc     As Long
    Dim sta     As Long
    Dim sv      As Long

    Const Name = "�����"
    Const Code = 1017
    Const np = 18

    ORA_LOG "IN    Ora_OWARI_SUIBA_PUMP"

    OracleDB.Label3 = "�I���N�����" & Name & "�f�[�^�擾��"
    OracleDB.Label3.Refresh

    ODATS = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    ODATE = "'" & Format(d2, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"

    SQL = "SELECT * FROM oracle.OWARI_SV WHERE jikoku " & _
          "BETWEEN TO_DATE(" & ODATS & ") AND TO_DATE(" & ODATE & ")" & _
          " AND Station=" & Str(Code) & " AND sv_no BETWEEN 1 AND 4"

    SQL = "SELECT * FROM oracle.OWARI_SV WHERE jikoku= TO_DATE(" & ODATS & ")"

    Debug.Print " SQL="; SQL

    nn = DateDiff("n", d1, d2) / 10 + 1
    ReDim s_p(nn, 4) As Variant 'nn=�����X�e�b�v��  4=sv�ԍ����Ƃ̃f�[�^
    '������
    For n = 1 To nn
        s_p(n, 1) = 1025
        s_p(n, 2) = 64
        s_p(n, 3) = 4100
        s_p(n, 4) = 256
    Next n

    ' SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)

    If dynOra.EOF And dynOra.BOF Then
        ORA_LOG Name & "�f�[�^���f�[�^�x�[�X�ɓo�^����Ă��܂���B" & vbCrLf & _
                 "�������m���߂Ă��������B"
        ORA_LOG "SQL=" & SQL
        rc = False
        dynOra.Close
        Set dynOra = Nothing
        OracleDB.Label3 = "�I���N�����" & Name & "�|���v�f�[�^�掸�s"
        OracleDB.Label3.Refresh
        GoTo JUMP1
    End If

    dw = d1
    Do Until dynOra.EOF
        t = CDate(dynOra.Fields("jikoku").Value)
        n = DateDiff("n", d1, t) / 10 + 1
        m = AscB(dynOra.Fields("sv_data").Value)
        svc = dynOra.Fields("sv_no").Value
        sta = dynOra.Fields("Station").Value
        flag = dynOra.Fields("flag").Value
        ORA_LOG " Name=" & Name
        ORA_LOG " t=" & Format(t, "yyyy/mm/dd hh:nn")
        ORA_LOG " n=" & Str(n)
        ORA_LOG " m=" & Str(m)
        ORA_LOG " svc=" & Str(svc)
        ORA_LOG " sta=" & Str(sta)
        ORA_LOG " flag=" & Str(flag)
        Debug.Print " Name="; Name
        Debug.Print " t="; t
        Debug.Print " n="; n
        Debug.Print " m="; m
        Debug.Print " svc="; svc
        Debug.Print " sta="; sta
        Debug.Print " flag="; flag
        If flag = 0 And svc < 5 Then
            s_p(n, svc) = m
        End If
        dynOra.MoveNext
    Loop
    dynOra.Close

JUMP1:
    For n = 1 To nn '������
        pw = ""
        For i = 1 To 4 'sv�ԍ���
            Select Case i
                Case 1
                    jj = s_p(n, i) And 1024
                    If jj > 0 Then
                        pw = "1,"
                    Else
                        pw = "0,"
                    End If
                    jj = s_p(n, i) And 1
                    If jj > 0 Then
                        pw = pw & "1,"
                    Else
                        pw = pw & "0,"
                    End If
                Case 2
                    jj = s_p(n, i) And 64
                    If jj > 0 Then
                        pw = pw & "1,"
                    Else
                        pw = pw & "0,"
                    End If
                Case 3
                    jj = s_p(n, i) And 4096
                    If jj > 0 Then
                        pw = pw & "1,"
                    Else
                        pw = pw & "0,"
                    End If
                    jj = s_p(n, i) And 4
                    If jj > 0 Then
                        pw = pw & "1,"
                    Else
                        pw = pw & "0,"
                    End If
                Case 4
                    jj = s_p(n, i) And 256
                    If jj > 0 Then
                        pw = pw & "1,"
                    Else
                        pw = pw & "0,"
                    End If
            End Select
        Next i
        Pump(np, n) = pw
    Next n

    ORA_LOG "OUT   Ora_OWARI_SUIBA_PUMP"

End Sub
'
'�|���v�f�[�^���擾����
'
'
Sub ORA_P_PUMP(d1 As Date, d2 As Date, rc As Boolean)

    Dim i       As Long
    Dim j       As Long
    Dim n       As Long
    Dim m       As Long
    Dim Name    As String   '�|���v����
    Dim Code    As Long     '�|���v���R�[�h
    Dim sv      As Long     'sv�ԍ�
    Dim N_P     As Long     '�|���v��
    Dim np      As Long     '�|���v��̒ʂ��ԍ�
    Dim ic      As Boolean

    n = DateDiff("n", d1, d2) / 10 + 1 '10���f�[�^�̌�

    ReDim Pump(19, n)  '17=�|���v�ꐔ  n=�����X�e�b�v��


    Name = "���c�|���v��"
    Code = 2501
    sv = 1
    N_P = 4
    np = 1
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "���V��F�|���v��"
    Code = 2502
    sv = 2
    N_P = 5
    np = 2
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "�����|���v��"
    Code = 2503
    sv = 3
    N_P = 4
    np = 3
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "�x�c�|���v��"
    Code = 2504
    sv = 1
    N_P = 5
    np = 4
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "�����|���v��"
    Code = 2505
    sv = 1
    N_P = 3
    np = 5
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "����|���v��"
    Code = 2506
    sv = 3
    N_P = 2
    np = 6
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "�����c��|���v��"
    Code = 2507
    sv = 1
    N_P = 4
    np = 7
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "��ꕽ�c�|���v��"
    Code = 2508
    sv = 2
    N_P = 2
    np = 8
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "���c������|���v��"
    Code = 2509
    sv = 3
    N_P = 5
    np = 9
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "��񕽓c�|���v��"
    Code = 2610
    sv = 1
    N_P = 2
    np = 10
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "�㏬�c��|���v��"
    Code = 2611
    sv = 2
    N_P = 8
    np = 11
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "����˃|���v��"
    Code = 2601
    sv = 1
    N_P = 4
    np = 12
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "��c�T�|���v��"
    Code = 2602
    sv = 1
    N_P = 5
    np = 13
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "�L�c��|���v��"
    Code = 2603
    sv = 1
    N_P = 5
    np = 14
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "�x�]�|���v��"
    Code = 2604
    sv = 1
    N_P = 6
    np = 15
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "�y���|���v��"
    Code = 2605
    sv = 1
    N_P = 4
    np = 16
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "�d�Ԑ�|���v��"
    Code = 2606
    sv = 1
    N_P = 3
    np = 17
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "�����|���v��"
    Ora_OWARI_SUIBA_PUMP d1, d2, ic

    Name = "���c��|���v��"
    Code = 1082
    sv = 1
    N_P = 4
    np = 19
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    rc = True
    Pump_TO_mdb d1, d2

End Sub
Sub Pump_Inital()

    Pump_Stats(1).Name = "���c�|���v��"
    Pump_Stats(1).P_Code = 2501
    Pump_Stats(1).sv_Num = 1
    Pump_Stats(1).P_Num = 4
    Pump_Stats(1).S_Num = 1

    Pump_Stats(2).Name = "���V��F�|���v��"
    Pump_Stats(2).P_Code = 2502
    Pump_Stats(2).sv_Num = 2
    Pump_Stats(2).P_Num = 5
    Pump_Stats(2).S_Num = 2

    Pump_Stats(3).Name = "�����|���v��"
    Pump_Stats(3).P_Code = 2503
    Pump_Stats(3).sv_Num = 3
    Pump_Stats(3).P_Num = 4
    Pump_Stats(3).S_Num = 3

    Pump_Stats(4).Name = "�x�c�|���v��"
    Pump_Stats(4).P_Code = 2504
    Pump_Stats(4).sv_Num = 1
    Pump_Stats(4).P_Num = 5
    Pump_Stats(4).S_Num = 4

    Pump_Stats(5).Name = "�����|���v��"
    Pump_Stats(5).P_Code = 2505
    Pump_Stats(5).sv_Num = 1
    Pump_Stats(5).P_Num = 3  '1,3,4 2�͍Ō�ɒǉ��\��
    Pump_Stats(5).S_Num = 5

    Pump_Stats(6).Name = "����|���v��"
    Pump_Stats(6).P_Code = 2506
    Pump_Stats(6).sv_Num = 3
    Pump_Stats(6).P_Num = 2
    Pump_Stats(6).S_Num = 6

    Pump_Stats(7).Name = "�����c��|���v��"
    Pump_Stats(7).P_Code = 2507
    Pump_Stats(7).sv_Num = 1
    Pump_Stats(7).P_Num = 5
    Pump_Stats(7).S_Num = 7

    Pump_Stats(8).Name = "��ꕽ�c�|���v��"
    Pump_Stats(8).P_Code = 2508
    Pump_Stats(8).sv_Num = 2
    Pump_Stats(8).P_Num = 2
    Pump_Stats(8).S_Num = 8

    Pump_Stats(9).Name = "���c��������|���v��"
    Pump_Stats(9).P_Code = 2509
    Pump_Stats(9).sv_Num = 3
    Pump_Stats(9).P_Num = 5
    Pump_Stats(9).S_Num = 9

    Pump_Stats(10).Name = "��񕽓c�|���v��"
    Pump_Stats(10).P_Code = 2610
    Pump_Stats(10).sv_Num = 1
    Pump_Stats(10).P_Num = 2
    Pump_Stats(10).S_Num = 10

    Pump_Stats(11).Name = "�㏬�c��|���v��"
    Pump_Stats(11).P_Code = 2611
    Pump_Stats(11).sv_Num = 2
    Pump_Stats(11).P_Num = 8
    Pump_Stats(11).S_Num = 11

    Pump_Stats(12).Name = "����˃|���v��"
    Pump_Stats(12).P_Code = 2601
    Pump_Stats(12).sv_Num = 1
    Pump_Stats(12).P_Num = 4
    Pump_Stats(12).S_Num = 12

    Pump_Stats(13).Name = "��c�T�|���v��"
    Pump_Stats(13).P_Code = 2602
    Pump_Stats(13).sv_Num = 1
    Pump_Stats(13).P_Num = 5
    Pump_Stats(13).S_Num = 13

    Pump_Stats(14).Name = "�L�c��|���v��"
    Pump_Stats(14).P_Code = 2603
    Pump_Stats(14).sv_Num = 1
    Pump_Stats(14).P_Num = 5
    Pump_Stats(14).S_Num = 14

    Pump_Stats(15).Name = "�x�]�|���v��"
    Pump_Stats(15).P_Code = 2604
    Pump_Stats(15).sv_Num = 1
    Pump_Stats(15).P_Num = 6
    Pump_Stats(15).S_Num = 15

    Pump_Stats(16).Name = "�y���|���v��"
    Pump_Stats(16).P_Code = 2605
    Pump_Stats(16).sv_Num = 1
    Pump_Stats(16).P_Num = 4
    Pump_Stats(16).S_Num = 16

    Pump_Stats(17).Name = "�d�Ԑ�|���v��"
    Pump_Stats(17).P_Code = 2606
    Pump_Stats(17).sv_Num = 1
    Pump_Stats(17).P_Num = 3
    Pump_Stats(17).S_Num = 17

    Pump_Stats(18).Name = "�����|���v��"
    Pump_Stats(18).P_Code = 1017
    Pump_Stats(18).sv_Num = 1
    Pump_Stats(18).P_Num = 6
    Pump_Stats(18).S_Num = 18

    Pump_Stats(19).Name = "���c��|���v��"
    Pump_Stats(19).P_Code = 1082
    Pump_Stats(19).sv_Num = 1
    Pump_Stats(19).P_Num = 4
    Pump_Stats(19).S_Num = 19

End Sub
