Attribute VB_Name = "MDB_Access"
Option Explicit
Option Base 1

Public MDBx              As Boolean              '�\���f�[�^�x�[�X  �ڑ���=True  �ڑ��s��=False
Public ����MDB           As String               '����MDB�̃t���p�X
Public ����MDB           As String               '����MDB�̃t���p�X
Public Con_����          As New ADODB.Connection
Public Con_����          As New ADODB.Connection
Public Rec_����          As New ADODB.Recordset
Public Rec_����          As New ADODB.Recordset

Public Con_�\��        As New ADODB.Connection
Public Rst_�\��        As New ADODB.Recordset
Public DB_�\��         As Boolean

Public H_Pred(500, 5, 4) As Single               '���ʗ\������
Public R_Pred(500, 5, 4) As Single               '�J�ʗ\������
Public T_Pred(500)       As Date                 '�\���v�Z������

Public History           As Boolean              '����\���R���g���[�� �\���L��=True  ����=False
Sub MDB_����_Close()

    Con_����.Close
    Set Rec_����.ActiveConnection = Nothing

End Sub
Sub MDB_����_Connection()

    Dim Con_str As String
    Dim a

    On Error GoTo ER1

    Con_str = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & ����MDB
    Con_����.ConnectionString = Con_str
    Con_����.Open

    Set Rec_����.ActiveConnection = Con_����
    MDBx = True
    On Error GoTo 0
    Exit Sub

ER1:
    a = MsgBox("MDB_������DB�ɃA�N�Z�X�ł��܂���ADB�̗L���AODBC���̐ݒ���m�F���Ă��������B" & vbCrLf & _
           "�v�Z�𑱍s���܂���(���s�̏ꍇ�͗\���l�̗����͕ۑ�����܂���)�H", vbYesNo + vbInformation)
    If a = vbYes Then
        MDBx = False
        Exit Sub
    Else
        End
    End If


End Sub
Sub MDB_����_Close()

    Con_����.Close
    Set Rec_����.ActiveConnection = Nothing

End Sub


Sub MDB_����_Connection()

    Dim Con_str As String
    Dim a

    On Error GoTo ER1

    Con_str = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & ����MDB
    Con_����.ConnectionString = Con_str
    Con_����.Open

    Set Rec_����.ActiveConnection = Con_����
    MDBx = True
    On Error GoTo 0
    Exit Sub

ER1:
    a = MsgBox("�\��������DB�ɃA�N�Z�X�ł��܂���ADB�̗L���AODBC���̐ݒ���m�F���Ă��������B" & vbCrLf & _
           "�v�Z�𑱍s���܂���(���s�̏ꍇ�͗\���l�̗����͕ۑ�����܂���)�H", vbYesNo + vbInformation)
    On Error GoTo 0
    If a = vbYes Then
        MDBx = False
        Exit Sub
    Else
        End
    End If

End Sub
Sub MDB_����_Read()


    Dim i       As Integer
    Dim j       As Integer
    Dim k       As Integer
    Dim n       As Integer
    Dim d       As Date
    Dim sr      As Single
    Dim SQL     As String
    Dim rw      As Variant
    Dim buf     As String
    Dim fn      As String

    fn = Format(Minute(jgd), "00")
    If isRAIN = "02" Then
        SQL = "Select * From FRICS���� Where Time Between  '" & Format(jsd, "yyyy/mm/dd hh:nn") & _
              "' And '" & Format(jgd, "yyyy/mm/dd hh:nn") & "' AND Minute=" & fn
    Else
        SQL = "Select * From �C�ے����� Where Time Between  '" & Format(jsd, "yyyy/mm/dd hh:nn") & _
              "' And '" & Format(jgd, "yyyy/mm/dd hh:nn") & "' AND Minute=" & fn
    End If


    Rec_����.Open SQL, Con_����, adOpenDynamic, adLockReadOnly

    For i = 1 To 500
        For j = 1 To 5
            For k = 1 To 4
                H_Pred(i, j, k) = -99#
                R_Pred(i, j, k) = -99#
            Next k
        Next j
    Next i

    If Rec_����.BOF Or Rec_����.EOF Then
        Rec_����.Close
        Exit Sub
    End If

    Do Until Rec_����.EOF

        d = CDate(Rec_����.Fields("Time").Value)
        n = DateDiff("h", jsd, d) + 1
        T_Pred(n) = d

'���V��F
        buf = Rec_����.Fields("���V��F").Value
        rw = Split(buf, ",")
        k = UBound(rw) - 1
        For j = 0 To k
            sr = rw(j)
            H_Pred(n, 1, j + 1) = sr
        Next j
'�厡
        buf = Rec_����.Fields("�厡").Value
        rw = Split(buf, ",")
        k = UBound(rw) - 1
        For j = 0 To k
            sr = rw(j)
            H_Pred(n, 2, j + 1) = sr
        Next j
'����O����
        buf = Rec_����.Fields("�����O����").Value
        rw = Split(buf, ",")
        k = UBound(rw) - 1
        For j = 0 To k
            sr = rw(j)
            H_Pred(n, 3, j + 1) = sr
        Next j
'�v�n��
        buf = Rec_����.Fields("�v�n��").Value
        rw = Split(buf, ",")
        k = UBound(rw) - 1
        For j = 0 To k
            sr = rw(j)
            H_Pred(n, 4, j + 1) = sr
        Next j
'�t��
        buf = Rec_����.Fields("�t��").Value
        rw = Split(buf, ",")
        k = UBound(rw) - 1
        For j = 0 To k
            sr = rw(j)
            H_Pred(n, 5, j + 1) = sr
        Next j
'�\���J��
        buf = Rec_����.Fields("�\���~�J").Value
        rw = Split(buf, ",")
        k = UBound(rw) - 1
        For j = 0 To k
            sr = rw(j)
            R_Pred(n, 1, j + 1) = sr
        Next j

        Rec_����.MoveNext
    Loop

    Rec_����.Close

End Sub
'
'�\���l��MDB�ɕۑ�����B
'
'
Sub MDB_����_Write()

    Dim i   As Integer
    Dim j   As Integer
    Dim ns  As Integer
    Dim buf As String
    Dim SQL As String

    Const f2 = "##0.00"
    Const f1 = "###0.0"

    If isRAIN = "02" Then
        SQL = "Select * From FRICS���� Where Time = '" & Format(jgd, "yyyy/mm/dd hh:nn") & "'"
    Else
        SQL = "Select * From �C�ے����� Where Time = '" & Format(jgd, "yyyy/mm/dd hh:nn") & "'"
    End If

    Rec_����.Open SQL, Con_����, adOpenDynamic, adLockOptimistic

    If Rec_����.BOF Or Rec_����.EOF Then
        Rec_����.AddNew
        Rec_����.Fields("Time").Value = Format(jgd, "yyyy/mm/dd hh:nn")
    End If
    Rec_����.Fields("Minute").Value = Format(Minute(jgd), "00")

'������O����
    buf = Format(DH_Tide, f2) & ","                   '�������V�����ʂƎ��ѐ��ʂƂ̍�
    buf = buf & Format(HO(1, Now_Step), f2) & ","     '����������(����)
    buf = buf & Format(HO(1, Now_Step + 1), f2) & "," '1���Ԍ�
    buf = buf & Format(HO(1, Now_Step + 2), f2) & "," '2���Ԍ�
    buf = buf & Format(HO(1, Now_Step + 3), f2) & "," '3���Ԍ�
    Rec_����.Fields("������O����").Value = buf
'���V��F
    ns = V_Sec_Num(1)
    buf = Format(HO(3, Now_Step), f2) & ","           '����������(����)
    buf = buf & Format(HQ(1, ns, NT - 12), f2) & ","  '1���Ԍ�
    buf = buf & Format(HQ(1, ns, NT - 6), f2) & ","   '2���Ԍ�
    buf = buf & Format(HQ(1, ns, NT), f2) & ","       '3���Ԍ�
    Rec_����.Fields("���V��F").Value = buf
'�厡
    ns = V_Sec_Num(2)
    buf = Format(HO(4, Now_Step), f2) & ","           '����������(����)
    buf = buf & Format(HQ(1, ns, NT - 12), f2) & ","  '1���Ԍ�
    buf = buf & Format(HQ(1, ns, NT - 6), f2) & ","   '2���Ԍ�
    buf = buf & Format(HQ(1, ns, NT), f2) & ","       '3���Ԍ�
    Rec_����.Fields("�厡").Value = buf
'����O����
    ns = V_Sec_Num(3)
    buf = Format(HO(5, Now_Step), f2) & ","           '����������(����)
    buf = buf & Format(HQ(1, ns, NT - 12), f2) & ","  '1���Ԍ�
    buf = buf & Format(HQ(1, ns, NT - 6), f2) & ","   '2���Ԍ�
    buf = buf & Format(HQ(1, ns, NT), f2) & ","       '3���Ԍ�
    Rec_����.Fields("�����O����").Value = buf
'�v�n��
    ns = V_Sec_Num(4)
    buf = Format(HO(6, Now_Step), f2) & ","           '����������(����)
    buf = buf & Format(HQ(1, ns, NT - 12), f2) & ","  '1���Ԍ�
    buf = buf & Format(HQ(1, ns, NT - 6), f2) & ","   '2���Ԍ�
    buf = buf & Format(HQ(1, ns, NT), f2) & ","       '3���Ԍ�
    Rec_����.Fields("�v�n��").Value = buf
'�t��
    ns = V_Sec_Num(5)
    buf = Format(HO(7, Now_Step), f2) & ","           '����������(����)
    buf = buf & Format(HQ(1, ns, NT - 12), f2) & ","  '1���Ԍ�
    buf = buf & Format(HQ(1, ns, NT - 6), f2) & ","   '2���Ԍ�
    buf = buf & Format(HQ(1, ns, NT), f2) & ","       '3���Ԍ�
    Rec_����.Fields("�t��").Value = buf
'�\���J��
    buf = Format(RO(1, Now_Step), f1) & ","           '���������敽�ωJ��
    buf = buf & Format(RO(1, Now_Step + 1), f1) & "," '1���Ԍ㗬�敽�ωJ��
    buf = buf & Format(RO(1, Now_Step + 2), f1) & "," '2���Ԍ㗬�敽�ωJ��
    buf = buf & Format(RO(1, Now_Step + 3), f1) & "," '3���Ԍ㗬�敽�ωJ��
    Rec_����.Fields("�\���~�J").Value = buf

'DB��������
    Rec_����.Update
    Rec_����.Close

End Sub
'
'���݌v�Z�Ɏg���Ă���\���J�ʂ̏�Ԃ��Z�[�u����B
'
'
'
Sub RAIN_SELECT_READ()

    Dim Con    As String
    Dim R_Con  As New ADODB.Connection
    Dim R_Rst  As New ADODB.Recordset

    LOG_Out "IN    RAIN_SELECT_READ"

    Con = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & ����MDB

    R_Con.ConnectionString = Con
    R_Con.Open

    Set R_Rst.ActiveConnection = R_Con
    R_Rst.Open "SELECT * FROM RAIN_SELECT", R_Con, adOpenDynamic, adLockOptimistic
    
    If R_Rst.Fields("�C�ے�").Value Then
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
'���݌v�Z�Ɏg���Ă���\���J�ʂ̏�Ԃ��Z�[�u����B
'
'
'
Sub RAIN_SELECT_SAVE()

    Dim Con    As String
    Dim R_Con  As New ADODB.Connection
    Dim R_Rst  As New ADODB.Recordset

    LOG_Out "IN    RAIN_SELECT_SAVE"

    Con = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & ����MDB

    R_Con.ConnectionString = Con
    R_Con.Open

    Set R_Rst.ActiveConnection = R_Con
    R_Rst.Open "SELECT * FROM RAIN_SELECT", R_Con, adOpenDynamic, adLockOptimistic

    R_Rst.Fields("�C�ے�").Value = KISYO
    R_Rst.Fields("FRICS").Value = FRICS

    R_Rst.Update
    R_Rst.Close
    R_Con.Close

    Set R_Rst = Nothing
    Set R_Con = Nothing

    LOG_Out "OUT   RAIN_SELECT_SAVE"

End Sub
Sub �\�񕶗���DB_Close()

    On Error Resume Next

    If Rst_�\��.State = 1 Then
        Rst_�\��.Close
    End If
    Set Rst_�\�� = Nothing
    Set Con_�\�� = Nothing

End Sub
Sub �\�񕶗���DB_Connection()

    Dim Con  As String

    LOG_Out "IN    �\�񕶗���DB_Connection"

    On Error GoTo ERH1
    
    Con = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & ����MDB

    Con_�\��.ConnectionString = Con
    Con_�\��.Open

    Set Rst_�\��.ActiveConnection = Con_�\��
    DB_�\�� = True

    LOG_Out "OUT   �\������DB_Connection Normal Return"

    On Error GoTo 0
    Exit Sub

ERH1:

    DB_�\�� = False
    MsgBox "�\�񕶗����f�[�^�x�[�X�ɐڑ��ł��܂���ł����A�����͎c��܂���B"
    LOG_Out "�\�񕶗����f�[�^�x�[�X�ɐڑ��ł��܂���ł����A�����͎c��܂���B"

    LOG_Out "OUT   �\�񕶗���DB_Connection ABNormal Return"

    On Error GoTo 0

End Sub
Sub �\�񕶗���DB_Read()

    Dim SQL     As String
    Dim dw      As String
    Dim T_Last  As Date
    Dim n       As Long

    LOG_Out "IN    �\�񕶗���DB_Read"

    �\�񕶗���DB_Connection

    If DB_�\�� = False Then
        LOG_Out "OT   �\������DB_Read DB_�\�� = False"
        Exit Sub
    End If

    SQL = "Select MAX(TIME) From �\�񕶗��� Where RAIN_KIND = '" & isRAIN & "'"

    Rst_�\��.Open SQL, Con_�\��, adOpenDynamic, adLockOptimistic

    If Rst_�\��.BOF Or Rst_�\��.EOF Then
       '�����ɂ͂��Ȃ��͂���������������
        BP = 0
        Rst_�\��.Close
        �\�񕶗���DB_Close
        LOG_Out "OUT   �\������DB_Read �����ɂ͂��Ȃ��͂���������������"
        Exit Sub
    End If

    dw = Rst_�\��.Fields(0).Value
    T_Last = CDate(dw)

    Rst_�\��.Close

    n = DateDiff("h", T_Last, jgd) + 1
    If n > 25 Then
        BP = 0
        LOG_Out "                   jgd=" & TIMEC(jgd)
        LOG_Out "                T_Last=" & TIMEC(T_Last)
        �\�񕶗���DB_Close
        LOG_Out "OUT   �\�񕶗���DB_Read n=" & str(n)
        Exit Sub
    End If

    SQL = "Select * From �\�񕶗��� Where TIME = '" & dw & "' AND  RAIN_KIND = '" & isRAIN & "'"
    Rst_�\��.Open SQL, Con_�\��, adOpenDynamic, adLockOptimistic
    If Rst_�\��.EOF Then
        BP = 0
        Wng_Last_Time = 0
    Else
        BP = Rst_�\��.Fields("�\��t���O").Value
        Wng_Last_Time = Rst_�\��.Fields("Course").Value
    End If

    Rst_�\��.Close

    �\�񕶗���DB_Close

    LOG_Out "OUT   �\�񕶗���DB_Read SQL=" & SQL

End Sub
Sub �\�񕶗���DB_Write()

    Dim SQL    As String

    LOG_Out "IN    �\�񕶗���DB_Write"

    �\�񕶗���DB_Connection

    If DB_�\�� = False Then
        Exit Sub
    End If

    SQL = "Select * From �\�񕶗��� Where TIME = '" & Format(jgd, "yyyy/mm/dd hh:nn") & "' AND RAIN_KIND = '" & isRAIN & "'"

    Rst_�\��.Open SQL, Con_�\��, adOpenDynamic, adLockOptimistic

    If Rst_�\��.BOF Or Rst_�\��.EOF Then
        Rst_�\��.AddNew
        Rst_�\��.Fields("Time").Value = Format(jgd, "yyyy/mm/dd hh:nn")
        Rst_�\��.Fields("RAIN_KIND").Value = isRAIN
    End If

    If Pattan_Now <> 4 Then
        Rst_�\��.Fields("�\��t���O").Value = Pattan_Now  '�\�񕶃p�^�[��
        Rst_�\��.Fields("�\���ʃR�[�h").Value = Messag(Pattan_Now).Patn(16) 'Kind_N
        Rst_�\��.Fields("�\����").Value = Messag(Pattan_Now).Patn(2) 'Kind_S
    Else
        '�\�񕶔����I��
        Rst_�\��.Fields("�\��t���O").Value = 0                    '�\�񕶃p�^�[��
        Rst_�\��.Fields("�\���ʃR�[�h").Value = "0"            'Kind_N
        Rst_�\��.Fields("�\����").Value = "�^�����ӏ�����"   'Kind_S
    End If
    Rst_�\��.Fields("Course").Value = Wng_Last_Time

    If isRAIN = "01" Then
        Rst_�\��.Fields("RAIN_NAME").Value = "�C�ے�"
    Else
        Rst_�\��.Fields("RAIN_NAME").Value = "FRICS"
    End If

    If PRACTICE_FLG_CODE = "40" Then
        Rst_�\��.Fields("PRACTICE").Value = "�\��"
    Else
        Rst_�\��.Fields("PRACTICE").Value = "���K"
    End If

    Rst_�\��.Update

    Rst_�\��.Close
    
    �\�񕶗���DB_Close

    LOG_Out "OUT   �\�񕶗���DB_Write"

End Sub
