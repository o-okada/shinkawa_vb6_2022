Attribute VB_Name = "�C�ے�1km"
'
'Option BAse 1���͂����Ă��Ȃ����Ƃɒ���
'
Option Explicit
Public Rst               As New ADODB.Recordset
Public m1sd              As Date
Public m1ed              As Date
Public Process           As String
Public M_Link(315)       As Item
Type Item
    Cod3    As String
    id      As Long
    kd      As Long
End Type
Public SQL               As String
Public R_1km(107, 99)    As Long    '0�`107,0�`99���g�p����B
Public R_315(315)        As Single
Public R_135()           As Single
Public Five_Section      As N5
Type N5 '                  Byte
    L_Size(3) As Byte    ' 1�` 4
    No        As Byte    ' 5�` 5
    Num(3)    As Byte    ' 6�` 9
    Bit(1)    As Byte    '10�`11
    M_LVL     As Byte    '12�`12
    L_MAX(1)  As Byte    '13�`14
    M_MAX(1)  As Byte    '15�`16
    P         As Byte    '17�`17 �f�[�^/10**P
End Type
Public Five_Section_Num  As N5_Num
Type N5_Num
    L_Size    As Long
    No        As Long
    Num       As Long
    Bit       As Long
    M_LVL     As Long
    L_MAX     As Long
    M_MAX     As Long
    P         As Long
End Type
Public R_Lank()             As Single 'M_MAX ��Redim���邱��!!!!!!!!!!

Private gstrFiveSecLogFilenm As String
'
'dd1=���2�o�C�g
'dd2=����2�o�C�g
'
Private Function Byte2Long(dd1 As Byte, dd2 As Byte) As Long
        Byte2Long = CLng(dd1) * 256 + CLng(dd2)
End Function
'
' �C�ے�1km���b�V���J�ʃf�[�^�`�F�b�N
'
'
'
Sub Check_1kmMesh_Time(Cat As String, ic As Boolean)

    Dim nf   As Long
    Dim n    As Long
    Dim d1   As Date
    Dim d2   As Date
'    Dim d3   As Date
'    Dim ans  As Long
    Dim buf  As String
    Dim irc  As Boolean
    Dim F    As String
    Dim msg  As String

    Select Case Cat
        Case "VDXA70"
            '���щJ��
            msg = "10�������~�JVDXA70"
            F = App.Path & "\data\" & msg

        Case "VCXB70"
            '�~�J�Z���ԗ\��(1-3)
            msg = "�~�J�Z���ԗ\��(1-3)VCXB70"
            F = App.Path & "\data\" & msg

        Case "VCXB71"
            '�~�J�Z���ԗ\��(4-6)
            msg = "�~�J�Z���ԗ\��(4-6)VCXB71"
            F = App.Path & "\data\" & msg

        Case "VCXB75"
            '�~�J�Z���ԗ\��(1-3)30��
            msg = "�~�J�Z���ԗ\��(1-3)30��VCXB75"
            F = App.Path & "\data\" & msg

        Case "VCXB76"
            '�~�J�Z���ԗ\��(4-6)30��
            msg = "�~�J�Z���ԗ\��(4-6)30��VCXB76"
            F = App.Path & "\data\" & msg

        Case "VDXB70"
            '�i�E�L���X�g
            msg = "�i�E�L���X�gVDXB70"
            F = App.Path & "\data\" & msg
    End Select

    nf = FreeFile
    F = F & ".dat"
    Open F For Input As #nf
    Line Input #nf, buf
    d1 = CDate(buf)
    Close #nf

    Call RadarMeshuDataNewTime(Cat, d2, irc)
    If irc = False Then
        ic = irc
        GoTo JUMP
    End If

    n = DateDiff("h", d1, d2) + 1
    If n > 5 Then
        d1 = DateAdd("h", -5, d2) '�O��I�����T���Ԃ��O�������̂Ŏ�荞�݊J�n������ύX
    End If

    If d2 > d1 Then

        Select Case Cat
            Case "VDXA70"
                '���щJ��

                m1sd = DateAdd("n", 10, d1) '��荞�ݍς݂�10����f�[�^�����荞��
                m1ed = d2
                ORA_LOG msg & "�f�[�^��荞�݊J�n " & d1 & " ���� " & d2 & "�܂�"

                ORA_VDXA70_Rain ic

            Case "VDXB70"
                '�i�E�L���X�g

                m1sd = DateAdd("n", 10, d1) '��荞�ݍς݂�10����f�[�^�����荞��
                m1ed = d2
                ORA_LOG msg & "�f�[�^��荞�݊J�n " & d1 & " ���� " & d2 & "�܂�"

                ORA_VDXB70_Rain ic

            Case "VCXB70"
                '�~�J�Z���ԗ\��(1-3)

                m1sd = DateAdd("h", 1, d1) '��荞�ݍς݂�1���Ԍ�f�[�^�����荞��
                m1ed = d2
                ORA_LOG msg & "�f�[�^��荞�݊J�n " & d1 & " ���� " & d2 & "�܂�"

                ORA_VCXB70_Rain ic

            Case "VCXB71"
                '�~�J�Z���ԗ\��(4-6)

                m1sd = DateAdd("h", 1, d1) '��荞�ݍς݂�1���Ԍ�f�[�^�����荞��
                m1ed = d2
                ORA_LOG msg & "�f�[�^��荞�݊J�n " & d1 & " ���� " & d2 & "�܂�"

                ORA_VCXB71_Rain ic

            Case "VCXB75"
                '�~�J�Z���ԗ\��(1-3)30��

                m1sd = DateAdd("h", 1, d1) '��荞�ݍς݂�1���Ԍ�f�[�^�����荞��
                m1ed = d2
                ORA_LOG msg & "�f�[�^��荞�݊J�n " & d1 & " ���� " & d2 & "�܂�"

                ORA_VCXB75_Rain ic

            Case "VCXB76"
                '�~�J�Z���ԗ\��(4-6)30��

                m1sd = DateAdd("h", 1, d1) '��荞�ݍς݂�1���Ԍ�f�[�^�����荞��
                m1ed = d2
                ORA_LOG msg & "�f�[�^��荞�݊J�n " & d1 & " ���� " & d2 & "�܂�"

                ORA_VCXB76_Rain ic

        End Select
        Call SQLdbsDeleteRecordset(gAdoRst)

        If Not ic Then
            ORA_LOG msg & "�f�[�^���擾���悤�Ƃ�������" & vbCrLf & _
                    "�G���[���������Ă��܂��B"
            GoTo JUMP
        Else
            ORA_LOG msg & "�f�[�^��荞�ݐ���I��"
        End If
    End If


JUMP:

End Sub
'
'
'5�߃f�[�^��荞��
'
'
Sub Five_Section_Get(Field_Name As String)

'    Dim d1           As String
    Dim msg          As String
'    Dim PartImage    As OraBlob
'    Dim chunksize    As Long
'    Dim AmountRead   As Long
'    Dim buffer       As Variant
'    Dim SQL          As String
    Dim buf()        As Byte
    Dim FNum         As Long
    Dim i            As Long
    Dim m            As Long
    Dim s(1)         As Byte
    Dim pp           As Single
    Dim ii(8)        As Long
    Dim strSrcFile As String
    Dim strDestFile As String

    'Get OraBlob from OraDynaset
'    Set PartImage = dynOra.Fields(Field_Name).Value
    buf() = gAdoRst.Fields(Field_Name).GetChunk(gAdoRst.Fields(Field_Name).ActualSize)

    'Set Offset and PollingAmount property for piecewise Read operation
'    PartImage.offset = 1
'    PartImage.PollingAmount = PartImage.Size
'    chunksize = 1000
    'Get a free file number
    FNum = FreeFile
    
    If Len(Dir(App.Path & "\Rep.bin", vbNormal)) > 0 Then
        Kill App.Path & "\Rep.bin"
    End If

    'Open the file
    Open App.Path & "\Rep.bin" For Binary As #FNum

    'Do the first read on 'PartImage, buffer must be a variant
'    AmountRead = PartImage.Read(buffer, chunksize)
    
    'put will not allow Variant type
'    buf = buffer
    
    Put #FNum, , buf

    ' Check for the Status property for polling read operation
'    While PartImage.Status = ORALOB_NEED_DATA
'        AmountRead = PartImage.Read(buffer, chunksize)
'        buf = buffer
'        Put #FNum, , buf
'    Wend

    Close FNum

    msg = "Get Complete"
    
    If gDebugMode = "ON" Then
        strDestFile = App.Path
        If Right(strDestFile, 1) <> "\" Then strDestFile = strDestFile & "\"
        strSrcFile = strDestFile
        strSrcFile = strSrcFile & "Rep.bin"
        strDestFile = strDestFile & gstrFiveSecLogFilenm
        strDestFile = strDestFile & Field_Name
        strDestFile = strDestFile & ".bin"
        If Len(Dir(strSrcFile, vbNormal)) > 0 Then Call FileCopy(strSrcFile, strDestFile)
    End If

    FNum = FreeFile
    Open App.Path & "\Rep.bin" For Binary As #FNum

    Get #FNum, , Five_Section

    With Five_Section
        ii(1) = Byte2Long(.L_Size(2), .L_Size(3)) 'L_Size
        ii(2) = CLng(.No)                         'No
        ii(3) = Byte2Long(.Num(2), .Num(3))       'Num
        ii(4) = Byte2Long(.Bit(0), .Bit(1))       'Bit
        ii(5) = CLng(.M_LVL)                      'M_LVL
        ii(6) = Byte2Long(.L_MAX(0), .L_MAX(1))   'L_MAX
        ii(7) = Byte2Long(.M_MAX(0), .M_MAX(1))   'M_MAX
        ii(8) = CLng(.P)                          'P
    End With
    With Five_Section_Num
        .L_Size = ii(1)
        .No = ii(2)
        .Num = ii(3)
        .Bit = ii(4)
        .M_LVL = ii(5)
        .L_MAX = ii(6)
        .M_MAX = ii(7)
        .P = ii(8)
        pp = 1# / 10# ^ .P
        m = .M_MAX
    End With

    ReDim R_Lank(m)

    For i = 1 To m
        Get #FNum, , s
        R_Lank(i) = Byte2Long(s(0), s(1))
        R_Lank(i) = R_Lank(i) * pp
    Next i


    Close #FNum

End Sub
Sub M_Link_Read()

    Dim i      As Long
    Dim buf    As String
    Dim nf     As Long

    ORA_LOG "IN   Sub M_Link_Read"

    nf = FreeFile
    Open App.Path & "\Data\M_Link.dat" For Input As #nf

    For i = 1 To 315 '�g�p1�����b�V����
        Line Input #nf, buf
        M_Link(i).Cod3 = Mid(buf, 13, 8)
        M_Link(i).id = Mid(buf, 24, 2)
        M_Link(i).kd = Mid(buf, 29, 2)
    Next i

    Close #nf

    ORA_LOG "OUT  Sub M_Link_Read"

End Sub
'**************************************************
'
'���[�_�[���b�V����135����Ɍv�Z����
'
'
'
'**************************************************
Sub Mesh_To_Ryuiki_JWA(w() As Single, RY() As Single)

    Dim i               As Integer
    Dim k               As Integer
    Dim m               As Integer
    Dim rw              As Single
    Dim r               As Single

    ORA_LOG "IN   Sub Mesh_To_Ryuiki_JWA"

    For i = 1 To RRYU '135����
        rw = 0#
        RY(i) = 0#
        For k = 1 To 20
            m = R_Meshu(k, i)  '���[�_�[���b�V���̔ԍ�
            If m = 0 Then Exit For
            r = w(m)
            rw = rw + r * R_Thissen(k, i)
'            Debug.Print " k="; k; " R_Meshu(k, i)="; m; " r = "; r; "   R_Thissen(k, i)="; R_Thissen(k, i); "  Rw="; Rw
        Next k
        If rw < 0# Then rw = 0#
        RY(i) = rw
    Next i

    ORA_LOG "OUT  Sub Mesh_To_Ryuiki_JWA"

End Sub
'******************************************************************
'
'
'�e�[�u�� �ŐV�C�ے�1km���b�V���J�ʃf�[�^���f�[�^�x�[�X�������荞��
'
'
'
'
'
'
'******************************************************************
Sub ORA_KANSOKU_JIKOKU_GET_1kmMesh(R_Lank As String, ic As Boolean)

    Dim cDw     As String
    Dim buf     As String
    Dim SQL     As String
    Dim n       As Long
    Dim nf      As Long
    Dim mj      As String
    Dim dw      As Date
    Dim d1      As Date
    Dim d2      As Date


    Process = "ORA_KANSOKU_JIKOKU_GET_1kmMesh"

    ic = True
'�擾�ς݂̎����𓾂�
    Select Case R_Lank
        Case "VDXA70"
            '�����J��
            mj = "10�����э~�JVDXA70.DAT"

        Case "VCXB75"
            '�~�J�Z���ԗ\��(1-3)30��
            mj = "�~�J�Z���ԗ\��(1-3)30��VCXB75.DAT"

        Case "VCXB76"
            '�~�J�Z���ԗ\��(4-6)30��
            mj = "�~�J�Z���ԗ\��(4-6)30��VCXB76.DAT"

        Case "VCXB70"
            '�~�J�Z���ԗ\��(1-3)����
            mj = "�~�J�Z���ԗ\��(1-3)VCXB70.DAT"

        Case "VCXB71"
            '�~�J�Z���ԗ\��(4-6)����
            mj = "�~�J�Z���ԗ\��(4-6)VCXB71.DAT"

        Case "VDXB70"
            '�i�E�L���X�g
            mj = "�i�E�L���X�gVDXB70.DAT"

    End Select

    nf = FreeFile
    Open App.Path & "\DATA\" & mj For Input As #nf
    Line Input #nf, buf
    dw = CDate(buf)
    Close #nf



    SQL = "SELECT * FROM oracle.KANSOKU_JIKOKU WHERE TABLE_NAME='" & R_Lank & "'"
'    SQL = "SELECT * FROM oracle.KANSOKU_JIKOKU WHERE TABLE_NAME='" & R_Lank & "' AND DETAIL = 2"
    ' SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)


'------------ �t�B�[���h�����擾���� -----------------
'    Dim Tw, i
'    n = dynOra.Fields.Count
'    For i = 0 To n - 1
'        Tw = dynOra.Fields(i).Name
'        Debug.Print " Number=" & Format(Str(i), "@@@") & " �t�B�[���h��="; Tw
'        Debug.Print " Value=" & dynOra.Fields(i).Value
'    Next i
'---------------------------------------------------

'------------ �c�a���e���擾���� -----------------

    Dim w1, w2, w3

    w1 = dynOra.Fields("write_time").Value
    w2 = dynOra.Fields("table_name").Value
    w3 = dynOra.Fields("last_data_time").Value
    If IsNull(w1) Then
        ORA_LOG "Error IN  ORA_KANSOKU_JIKOKU_GET Field=(" & R_Lank & ")�̃e�[�u���Q�Ǝ���NULL���A���Ă���"
        ORA_LOG "SQL= (" & SQL & ")"
        ic = False
        GoTo SKIP
    End If

    If w3 <= dw Then
        GoTo SKIP
    End If

    n = DateDiff("h", dw, w3)
    If n > 25 Then
        d1 = DateAdd("h", -23, w3)
        d2 = w3
    Else
        d1 = DateAdd("n", 10, dw)
        d2 = w3
    End If

'�f�[�^�x�[�X����J�ʃf�[�^�𓾂�
    Select Case R_Lank
        Case "VDXA70"
            '���щJ��
            ORA_VDXA70_Rain ic

        Case "VDXB70"
            '�i�E�L���X�g
            ORA_VDXB70_Rain ic

        Case "VCXB70"
            '�~�J�Z���ԗ\��(1-3)
            ORA_VCXB70_Rain ic

        Case "VCXB71"
            '�~�J�Z���ԗ\��(4-6)
            ORA_VCXB71_Rain ic

        Case "VCXB75"
            '�~�J�Z���ԗ\��(1-3)30��
            ORA_VCXB75_Rain ic

        Case "VCXB76"
            '�~�J�Z���ԗ\��(4-6)30��
            ORA_VCXB76_Rain ic

    End Select


    nf = FreeFile
    Open App.Path & "\DATA\" & mj For Output As #nf
    Print #nf, TIMEC(d2)
    Close #nf





'---------------------------------------------------
SKIP:
    DoEvents
    Close #n
    dynOra.Close
    Set dynOra = Nothing

End Sub
'
'�~�J�Z���ԗ\��(1-3)����
'
'
'
Sub ORA_VCXB70_Rain(ir As Boolean)

    Dim rc     As Boolean
    Dim dw     As Date
    Dim i      As Long
    Dim dx     As Date
    Dim k      As Long
    Dim n      As Long
    Dim nf     As Long
    Dim buf    As String
    Dim setData As Boolean

    On Error GoTo ER1

    n = DateDiff("h", m1sd, m1ed) + 1
    dw = m1sd
    For k = 1 To n
        OracleDB.Label3 = "�~�J�Z���ԗ\��(1-3)����(VCXB70)��荞�ݒ� " & TIMEC(dw)
        OracleDB.Label3.Refresh
        For i = 1 To 18
            Seven_Section_Get "VCXB70", i, dw 'dw=�擾����
            
            setData = False
            If Not (gAdoRst Is Nothing) Then
                If gAdoRst.State = adStateOpen Then
                    If Not gAdoRst.EOF Then
                        setData = True
                    End If
                End If
            End If
            
            If setData = True Then

            dx = DateAdd("n", i * 10, dw)

            Regist_PRD_1km_Rain dx   'dx=�\������

            End If
        Next i


        ORA_LOG "�C�ے��~�J�Z���ԗ\���i1-3�j���� MDB�ɏ������ݏI��" & dw


        ORA_LOG "�C�ے��~�J�Z���ԗ\���i1-3�j���� �����������݊J�n " & dw
        nf = FreeFile
        Open App.Path & "\data\�~�J�Z���ԗ\��(1-3)VCXB70.DAT" For Output As #nf
        Print #nf, TIMEC(dw)
        Print #nf, TIMEC(Now)
        Close #nf
        ORA_LOG "�C�ے��~�J�Z���ԗ\���i1-3�j���� �����������ݏI��"

        dw = DateAdd("h", 1, dw)

    Next k

    ir = True
    OracleDB.Label3 = ""
    OracleDB.Label3.Refresh
    On Error GoTo 0
    Exit Sub

ER1:
    Dim strMessage As String
    
'    If dbOra.LastServerErr <> 0 Then
'        strMessage = dbOra.LastServerErrText ' DB�����ɂ�����G���[
'    Else
        strMessage = Err.Description ' �ʏ�̃G���[
'    End If

    ORA_LOG "�C�ے��~�J�Z���ԗ\���i1-3�j���� �擾�G���[�����������݊J�n " & dw
    nf = FreeFile
    Open App.Path & "\data\�~�J�Z���ԗ\��(1-3)VCXB70.DAT" For Output As #nf
    Print #nf, TIMEC(dw)
    Print #nf, TIMEC(Now)
    Close #nf
    ORA_LOG "�C�ے��~�J�Z���ԗ\���i1-3�j���� �擾�G���[�����������ݏI��"

    ORA_LOG "ERROR IN  ORA_VCXB70_Rain"
    ORA_LOG "ERROR=" & strMessage

    ir = False
    On Local Error GoTo 0

End Sub
'
'�~�J�Z���ԗ\��(4-6)����
'
'
'
Sub ORA_VCXB71_Rain(ir As Boolean)

    Dim rc     As Boolean
    Dim dw     As Date
    Dim i      As Long
    Dim dx     As Date
    Dim k      As Long
    Dim n      As Long
    Dim nf     As Long
    Dim buf    As String
    Dim setData As Boolean

    On Error GoTo ER1

    n = DateDiff("h", m1sd, m1ed) + 1
    dw = m1sd
    For k = 1 To n
        OracleDB.Label3 = "�~�J�Z���ԗ\��(4-6)����(VCXB71)��荞�ݒ� " & TIMEC(dw)
        OracleDB.Label3.Refresh
        For i = 1 To 18
            Seven_Section_Get "VCXB71", i, dw
            
            setData = False
            If Not (gAdoRst Is Nothing) Then
                If gAdoRst.State = adStateOpen Then
                    If Not gAdoRst.EOF Then
                        setData = True
                    End If
                End If
            End If
            
            If setData = True Then

            dx = DateAdd("n", 180 + i * 10, dw) '180=3Hr

            Regist_PRD_1km_Rain dx

            End If
        Next i

        ORA_LOG "�C�ے��~�J�Z���ԗ\���i4-6�j���� MDB�ɏ������ݏI��" & dw

        ORA_LOG "�C�ے��~�J�Z���ԗ\���i4-6�j���� �����������݊J�n " & dw
        nf = FreeFile
        Open App.Path & "\data\�~�J�Z���ԗ\��(4-6)VCXB71.DAT" For Output As #nf
        Print #nf, TIMEC(m1sd)
        Print #nf, TIMEC(Now)
        Close #nf
        ORA_LOG "�C�ے��~�J�Z���ԗ\���i4-6�j���� �����������ݏI��"

        dw = DateAdd("h", 1, dw)

    Next k

    ir = True
    OracleDB.Label3 = ""
    OracleDB.Label3.Refresh
    On Error GoTo 0
    Exit Sub

ER1:
    Dim strMessage As String
    
'    If dbOra.LastServerErr <> 0 Then
'        strMessage = dbOra.LastServerErrText ' DB�����ɂ�����G���[
'    Else
        strMessage = Err.Description ' �ʏ�̃G���[
'    End If

    ORA_LOG "�C�ے��~�J�Z���ԗ\���i4-6�j���� �擾�G���[�����������݊J�n " & dw
    nf = FreeFile
    Open App.Path & "\data\�~�J�Z���ԗ\��(4-6)VCXB71.DAT" For Output As #nf
    Print #nf, TIMEC(m1sd)
    Print #nf, TIMEC(Now)
    Close #nf
    ORA_LOG "�C�ے��~�J�Z���ԗ\���i4-6�j���� �擾�G���[�����������ݏI��"

    ORA_LOG "ERROR IN  ORA_VCXB71_Rain"
    ORA_LOG "ERROR=" & strMessage

    ir = False
    On Local Error GoTo 0

End Sub
'
'�~�J�Z���ԗ\��(1-3)30��
'
'
'
Sub ORA_VCXB75_Rain(ir As Boolean)

    Dim rc     As Boolean
    Dim dw     As Date
    Dim i      As Long
    Dim dx     As Date
    Dim k      As Long
    Dim n      As Long
    Dim nf     As Long
    Dim buf    As String
    Dim setData As Boolean

    On Error GoTo ER1

    n = DateDiff("h", m1sd, m1ed) + 1
    dw = m1sd
    For k = 1 To n
        OracleDB.Label3 = "�~�J�Z���ԗ\��(1-3)30��(VCXB75)��荞�ݒ� " & TIMEC(dw)
        OracleDB.Label3.Refresh
        For i = 1 To 18
            Seven_Section_Get "VCXB75", i, dw 'dw=�擾����
            
            setData = False
            If Not (gAdoRst Is Nothing) Then
                If gAdoRst.State = adStateOpen Then
                    If Not gAdoRst.EOF Then
                        setData = True
                    End If
                End If
            End If
            
            If setData = True Then

            dx = DateAdd("n", i * 10, dw)

            Regist_PRD_1km_Rain dx   'dw=�\������

            End If
        Next i


        ORA_LOG "�C�ے��~�J�Z���ԗ\���i1-3�j30�� MDB�ɏ������ݏI��" & dw


        ORA_LOG "�C�ے��~�J�Z���ԗ\���i1-3�j30�� �����������݊J�n " & dw
        nf = FreeFile
        Open App.Path & "\data\�~�J�Z���ԗ\��(1-3)30��VCXB75.DAT" For Output As #nf
        Print #nf, TIMEC(dw)
        Print #nf, TIMEC(Now)
        Close #nf
        ORA_LOG "�C�ے��~�J�Z���ԗ\���i1-3�j30�� �����������ݏI��"

        dw = DateAdd("h", 1, dw)

    Next k

    ir = True
    OracleDB.Label3 = ""
    OracleDB.Label3.Refresh
    On Error GoTo 0
    Exit Sub

ER1:
    Dim strMessage As String
    
'    If dbOra.LastServerErr <> 0 Then
'        strMessage = dbOra.LastServerErrText ' DB�����ɂ�����G���[
'    Else
        strMessage = Err.Description ' �ʏ�̃G���[
'    End If

    ORA_LOG "�C�ے��~�J�Z���ԗ\���i1-3�j30�� �擾�G���[�����������݊J�n " & dw
    nf = FreeFile
    Open App.Path & "\data\�~�J�Z���ԗ\��(1-3)30��VCXB75.DAT" For Output As #nf
    Print #nf, TIMEC(dw)
    Print #nf, TIMEC(Now)
    Close #nf
    ORA_LOG "�C�ے��~�J�Z���ԗ\���i1-3�j30�� �擾�G���[�����������ݏI��"

    ORA_LOG "ERROR IN  ORA_VCXB75_Rain"
    ORA_LOG "ERROR=" & strMessage

    ir = False
    On Local Error GoTo 0

End Sub
'
'�~�J�Z���ԗ\��(4-6)30��
'
'
'
Sub ORA_VCXB76_Rain(ir As Boolean)

    Dim rc     As Boolean
    Dim dw     As Date
    Dim i      As Long
    Dim dx     As Date
    Dim k      As Long
    Dim n      As Long
    Dim nf     As Long
    Dim buf    As String
    Dim setData As Boolean

    On Error GoTo ER1

    n = DateDiff("h", m1sd, m1ed) + 1
    dw = m1sd
    For k = 1 To n
        OracleDB.Label3 = "�~�J�Z���ԗ\��(4-6)30��(VCXB76)��荞�ݒ� " & TIMEC(dw)
        OracleDB.Label3.Refresh
        For i = 1 To 18
            Seven_Section_Get "VCXB76", i, dw 'dw=�擾����
            
            setData = False
            If Not (gAdoRst Is Nothing) Then
                If gAdoRst.State = adStateOpen Then
                    If Not gAdoRst.EOF Then
                        setData = True
                    End If
                End If
            End If
            
            If setData = True Then

            dx = DateAdd("n", i * 10, dw)

            Regist_PRD_1km_Rain dx   'dw=�\������

            End If
        Next i


        ORA_LOG "�C�ے��~�J�Z���ԗ\���i4-6�j30�� MDB�ɏ������ݏI��" & dw


        ORA_LOG "�C�ے��~�J�Z���ԗ\���i4-6�j30�� �����������݊J�n " & dw
        nf = FreeFile
        Open App.Path & "\data\�~�J�Z���ԗ\��(4-6)30��VCXB76.DAT" For Output As #nf
        Print #nf, TIMEC(dw)
        Print #nf, TIMEC(Now)
        Close #nf
        ORA_LOG "�C�ے��~�J�Z���ԗ\���i4-6�j30�� �����������ݏI��"

        dw = DateAdd("h", 1, dw)

    Next k

    ir = True
    OracleDB.Label3 = ""
    OracleDB.Label3.Refresh
    On Error GoTo 0
    Exit Sub

ER1:
    Dim strMessage As String

'    If dbOra.LastServerErr <> 0 Then
'        strMessage = dbOra.LastServerErrText ' DB�����ɂ�����G���[
'    Else
        strMessage = Err.Description ' �ʏ�̃G���[
'    End If

    ORA_LOG "�C�ے��~�J�Z���ԗ\���i4-6�j30�� �擾�G���[�����������݊J�n " & dw
    nf = FreeFile
    Open App.Path & "\data\�~�J�Z���ԗ\��(4-6)30��VCXB76.DAT" For Output As #nf
    Print #nf, TIMEC(dw)
    Print #nf, TIMEC(Now)
    Close #nf
    ORA_LOG "�C�ے��~�J�Z���ԗ\���i4-6�j30�� �擾�G���[�����������ݏI��"

    ORA_LOG "ERROR IN  ORA_VCXB76_Rain"
    ORA_LOG "ERROR=" & strMessage

    ir = False
    On Local Error GoTo 0

End Sub
'
'�C�ے������J��
'
'
'
Sub ORA_VDXA70_Rain(ir As Boolean)

    Dim rc     As Boolean
    Dim dw     As Date
    Dim i      As Long
    Dim dx     As Date
    Dim k      As Long
    Dim n      As Long
    Dim nf     As Long
    Dim buf    As String
    Dim setData As Boolean

    On Error GoTo ER1

    n = DateDiff("n", m1sd, m1ed) / 10 + 1
    dw = m1sd
    For k = 1 To n

        OracleDB.Label3 = "�C�ے������J��(VDXA70)��荞�ݒ� " & TIMEC(dw)
        OracleDB.Label3.Refresh

        Seven_Section_Get "VDXA70", 0, dw
            
        setData = False
        If Not (gAdoRst Is Nothing) Then
            If gAdoRst.State = adStateOpen Then
                If Not gAdoRst.EOF Then
                    setData = True
                End If
            End If
        End If
        
        If setData = True Then


        Regist_OBS_1km_Rain dw

        End If


        ORA_LOG " �C�ے����[�_�[���� MDB�ɏ������ݏI��" & dw


        ORA_LOG "�C�ے����[�_�[���уf�[�^�����������݊J�n " & dw
        nf = FreeFile
        Open App.Path & "\data\10�������~�JVDXA70.DAT" For Output As #nf
        Print #nf, TIMEC(dw)
        Print #nf, TIMEC(Now)
        Close #nf
        ORA_LOG "�C�ے����[�_�[���уf�[�^�����������ݏI��"



        dw = DateAdd("n", 10, dw)
    Next k

    ir = True
    OracleDB.Label3 = ""
    OracleDB.Label3.Refresh
    On Error GoTo 0
    Exit Sub

ER1:
    Dim strMessage As String
    
'    If dbOra.LastServerErr <> 0 Then
'        strMessage = dbOra.LastServerErrText ' DB�����ɂ�����G���[
'    Else
        strMessage = Err.Description ' �ʏ�̃G���[
'    End If
        
    ORA_LOG "ERROR IN  ORA_VDXA70_Rain"
    ORA_LOG "ERROR=" & strMessage

    ORA_LOG "�C�ے����[�_�[���уf�[�^�擾�G���[�����������݊J�n " & dw
    
'�擾�G���[�ɂȂ�����G���[�ɂȂ��������̃f�[�^�͎擾�������Ƃɂ��Ď���҂悤�ɂ���
'�C�� 2009/12/28
    nf = FreeFile
    Open App.Path & "\data\10�������~�JVDXA70.DAT" For Output As #nf
    Print #nf, TIMEC(dw)
    Print #nf, TIMEC(Now)
    Close #nf
    ORA_LOG "�C�ے����[�_�[���уf�[�^�擾�G���[�����������ݏI��"

    ir = False
    On Local Error GoTo 0

End Sub
'
'�i�E�L���X�g
'
'
'
Sub ORA_VDXB70_Rain(ir As Boolean)

    Dim rc     As Boolean
    Dim dw     As Date
    Dim i      As Long
    Dim dx     As Date
    Dim k      As Long
    Dim n      As Long
    Dim nf     As Long
    Dim buf    As String
    Dim setData As Boolean

    On Error GoTo ER1

    n = DateDiff("n", m1sd, m1ed) / 10 + 1
    dw = m1sd
    For k = 1 To n
        OracleDB.Label3 = "�i�E�L���X�g(VDXB70)��荞�ݒ� " & TIMEC(dw)
        OracleDB.Label3.Refresh
        For i = 1 To 6
            Seven_Section_Get "VDXB70", i, dw
            
            setData = False
            If Not (gAdoRst Is Nothing) Then
                If gAdoRst.State = adStateOpen Then
                    If Not gAdoRst.EOF Then
                        setData = True
                    End If
                End If
            End If
            
            If setData = True Then

            dx = DateAdd("n", i * 10, dw)

            Regist_PRD_1km_Rain dx

            End If
        Next i


        ORA_LOG "�C�ے����[�_�[�i�E�L���X�g MDB�ɏ������ݏI��" & dw


        ORA_LOG "�C�ے����[�_�[�i�E�L���X�g�f�[�^�����������݊J�n " & dw
        nf = FreeFile
        Open App.Path & "\data\�i�E�L���X�gVDXB70.DAT" For Output As #nf
        Print #nf, TIMEC(dw)
        Print #nf, TIMEC(Now)
        Close #nf
        ORA_LOG "�C�ے����[�_�[�i�E�L���X�f�[�^�����������ݏI��"

        dw = DateAdd("n", 10, dw)
    Next k

    ir = True
    OracleDB.Label3 = ""
    OracleDB.Label3.Refresh
    On Error GoTo 0
    Exit Sub

ER1:
    Dim strMessage As String
    
'    If dbOra.LastServerErr <> 0 Then
'        strMessage = dbOra.LastServerErrText ' DB�����ɂ�����G���[
'    Else
        strMessage = Err.Description ' �ʏ�̃G���[
'    End If
        
    ORA_LOG "ERROR IN  ORA_VDXB70_Rain"
    ORA_LOG "ERROR=" & strMessage

    ORA_LOG "�C�ے����[�_�[�i�E�L���X�g�f�[�^�擾�G���[�����������݊J�n " & dw

'�擾�G���[�ɂȂ�����G���[�ɂȂ��������̃f�[�^�͎擾�������Ƃɂ��Ď���҂悤�ɂ���
'�C�� 2009/12/28
    nf = FreeFile
    Open App.Path & "\data\�i�E�L���X�gVDXB70.DAT" For Output As #nf
    Print #nf, TIMEC(dw)
    Print #nf, TIMEC(Now)
    Close #nf
    ORA_LOG "�C�ے����[�_�[�i�E�L���X�f�[�^�擾�G���[�����������ݏI��"

    ir = False
    On Local Error GoTo 0

End Sub
'
'�I���N�������荞��1km���b�V�������N�f�[�^
'��315���b�V����mm�J�ʂɂ���B
'
'
Sub R_RANK2mm()

    Dim i      As Long
    Dim id     As Long
    Dim kd     As Long
    Dim Lank    As Long

    For i = 1 To 315
        id = M_Link(i).id
        kd = M_Link(i).kd
        Lank = R_1km(id, kd)
        If (Lank >= 0) And (Lank < 315) Then
            R_315(i) = R_Lank(Lank)
        Else
            R_315(i) = 0#
        End If
    Next i

End Sub
Sub Regist_OBS_1km_Rain(dw As Date)

    Dim i      As Long
    Dim ir     As Long
    Dim nf     As Long
    Dim ic     As String
    Dim F      As String
    Dim SQL    As String
    Dim DC     As String
    Dim msg    As String
    Dim mnt    As Long
    Dim rc     As Boolean

'    On Error GoTo ERRHAND1

    MDB_Connection rc

    mnt = Minute(dw)
    DC = TIMEC(dw)
    SQL = "SELECT * FROM �C�ے����[�_�[���� WHERE TIME='" & DC & "'"

    Rst.Open SQL, MDB_Con, adOpenDynamic, adLockOptimistic

    If Rst.EOF Then
        Rst.AddNew
        Rst.Fields("TIME").Value = DC
    End If

    Rst.Fields("Minute").Value = mnt
    For i = 1 To 135
        ic = Format(i, "##0")
        ir = Round((R_135(i) + 0.02) * 10#)  '0.02�͂��܂��Ȃ�
        Rst.Fields(ic).Value = ir
    Next i
    Rst.Update
    Rst.Close

    On Error GoTo 0
    Exit Sub

ERRHAND1:
    msg = "ERROR IN Regist_OBS_1km_Rain" & vbCrLf
    msg = msg & "  ERROR NO=" & Str(Err.Number) & vbCrLf & _
           "  ERROR   =" & Err.Description
    ORA_LOG msg
    On Error GoTo 0

End Sub
'
'�C�ے��\���l�ۑ�
'
'
'
Sub Regist_PRD_1km_Rain(dw As Date)

    Dim i      As Long
    Dim ir     As Long
    Dim nf     As Long
    Dim ic     As String
    Dim F      As String
    Dim SQL    As String
    Dim DC     As String
    Dim msg    As String
    Dim mnt    As Long

    On Error GoTo ERRHAND1

'    MDB_Connection

    mnt = Minute(dw)
    DC = TIMEC(dw)
    SQL = "SELECT * FROM �C�ے����[�_�[�\��_1 WHERE TIME='" & DC & "'"

    Rst.Open SQL, MDB_Con, adOpenDynamic, adLockOptimistic

    If Rst.EOF Then
        Rst.AddNew
        Rst.Fields("TIME").Value = DC
        Rst.Fields("Minute").Value = mnt
    End If

    For i = 1 To 135
        ic = Format(i, "##0")
        ir = Round(R_135(i) * 10#)
        Rst.Fields(ic).Value = ir
    Next i
    Rst.Update
    Rst.Close

    On Error GoTo 0
    Exit Sub

ERRHAND1:
    msg = "ERROR IN Regist_PRD_1km_Rain" & vbCrLf
    msg = msg & "  ERROR NO=" & Str(Err.Number) & vbCrLf & _
           "  ERROR   =" & Err.Description
    ORA_LOG msg
    On Error GoTo 0

End Sub
'
'
'7�߃f�[�^��荞�݌�5�߂���荞��135�J�ʂ����B
'
'ii=�e�[�u�����ԍ�
'dw=�擾�f�[�^����
'
'
Sub Seven_Section_Get(Table_Name As String, ii As Long, dw As Date)

    Dim i            As Long
    Dim j            As Long
    Dim k            As Long
    Dim d1           As String
    Dim msg          As String
'    Dim PartImage    As OraBlob
'    Dim chunksize    As Long
'    Dim AmountRead   As Long
'    Dim buffer       As Variant
    Dim SQL          As String
    Dim buf()        As Byte
    Dim FNum         As Long
    Dim s            As Byte
    Dim Field_Name   As String
    Dim w
    Dim strTableName As String
    Dim strSrcFile As String
    Dim strDestFile As String
    Const intJSTAddHour9 As Long = 540

    Field_Name = "unpack_data" & Format(ii, "##")
    If ii <= 1 Then
        'Create the OraDynaset Object.
        d1 = "'" & Format(DateAdd("n", -(intJSTAddHour9), dw), "yyyy/mm/dd hh:nn") & "'" '," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
'        SQL = "SELECT * FROM " & Table_Name & " WHERE DATA_TIME=TO_DATE(" & d1 & ")"
        Select Case Table_Name
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
        End Select
        SQL = "SELECT * FROM " & strTableName & " WHERE data_time=" & d1
        Debug.Print " SQL=" & SQL
        If gDebugMode = "ON" Then
            ORA_LOG "�e�[�u����= (" & strTableName & ")" & "JST=" & Format(dw, "yyyy/mm/dd hh:nn") & "DATA_TIME=" & d1
        End If
'        Set dynOra = dbOra.CreateDynaset(SQL, 0&)
        Call SQLdbsDeleteRecordset(gAdoRst)
        Set gAdoRst = New ADODB.Recordset
        gAdoRst.CursorType = adOpenStatic
        gAdoRst.LockType = adLockReadOnly
        gAdoRst.Open SQL, gAdoCon, , , adCmdText
        If gAdoRst.EOF Then
            Call SQLdbsDeleteRecordset(gAdoRst)
            Exit Sub
        End If
    Else
        If Not (gAdoRst Is Nothing) Then
            If gAdoRst.State = adStateOpen Then
                If gAdoRst.EOF Then
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If

'data_time�̃t�B�[���h���`�F�b�N���邱��
'    w = dynOra.Fields("data_time").Value
'    w = dynOra.Fields("write_time").Value
'    w = dynOra.EOF
'    Debug.Print "       dw=" & TIMEC(dw)
'    Debug.Print " data_time=" & TIMEC(CDate(w))
'    Stop
'------------ �t�B�[���h�����擾���� -----------------
'    Dim Tw
'    Dim n            As Long
'    n = dynOra.Fields.Count
'    For i = 0 To n - 1
'        Tw = dynOra.Fields(i).Name
'        Debug.Print " Number=" & Format(Str(i), "@@@") & " �t�B�[���h��="; Tw
'    Next i
'---------------------------------------------------


    'Get OraBlob from OraDynaset
'    Set PartImage = dynOra.Fields(Field_Name).Value
    buf() = gAdoRst.Fields(Field_Name).GetChunk(gAdoRst.Fields(Field_Name).ActualSize)

    'Set Offset and PollingAmount property for piecewise Read operation
'    PartImage.offset = 1
'    PartImage.PollingAmount = PartImage.Size
'    chunksize = 10800
    'Get a free file number
    FNum = FreeFile
    
    If Len(Dir(App.Path & "\Level.bin", vbNormal)) > 0 Then
        Kill App.Path & "\Level.bin"
    End If

    'Open the file
    Open App.Path & "\Level.bin" For Binary As #FNum

    'Do the first read on 'PartImage, buffer must be a variant
'    AmountRead = PartImage.Read(buffer, chunksize)

    'put will not allow Variant type
'    buf = buffer

    Put #FNum, , buf

    ' Check for the Status property for polling read operation
'    While PartImage.Status = ORALOB_NEED_DATA
'        AmountRead = PartImage.Read(buffer, chunksize)
'        buf = buffer
'        Put #FNum, , buf
'    Wend

    Close FNum

    msg = "Table_Name=" & Table_Name & " Field_Name=" & Field_Name
    Debug.Print msg & "  Get Complete"
    
    If gDebugMode = "ON" Then
        strDestFile = App.Path
        If Right(strDestFile, 1) <> "\" Then strDestFile = strDestFile & "\"
        strSrcFile = strDestFile
        strSrcFile = strSrcFile & "Level.bin"
        strDestFile = strDestFile & Format(dw, "yyyymmdd_hhnn")
        strDestFile = strDestFile & "_"
        strDestFile = strDestFile & Table_Name
        strDestFile = strDestFile & "_"
        strDestFile = strDestFile & Field_Name
        strDestFile = strDestFile & ".bin"
        gstrFiveSecLogFilenm = Format(dw, "yyyymmdd_hhnn")
        gstrFiveSecLogFilenm = gstrFiveSecLogFilenm & "_"
        gstrFiveSecLogFilenm = gstrFiveSecLogFilenm & Table_Name
        gstrFiveSecLogFilenm = gstrFiveSecLogFilenm & "_"
        If Len(Dir(strSrcFile, vbNormal)) > 0 Then Call FileCopy(strSrcFile, strDestFile)
    End If

    DoEvents

    FNum = FreeFile
    Open App.Path & "\Level.bin" For Binary As #FNum
    i = 0
    j = 0
    k = 0
    Do
        k = k + 1
        Get #FNum, , s
        R_1km(j, i) = CLng(s)
        If i >= 99 Then
            i = 0
            j = j + 1
            If j > 107 Then
                Exit Do
            End If
        Else
            i = i + 1 '�o�x����
        End If
    Loop Until EOF(FNum)
    Close #FNum
    Debug.Print " ���b�V����=" & Str(k)
    If k < 10800 Then
        msg = "Table_Name=" & Table_Name & " Filed_Name=" & Table_Name
        msg = msg & " �̃f�[�^��������Ȃ�" & Str(i)
        ORA_LOG msg
    End If

'5�ߎ�荞��

    If ii > 0 Then
        Field_Name = "section5_" & Format(ii, "##")
    Else
        Field_Name = "section5"
    End If
    ORA_LOG "IN   Seven_Section_Get Table=" & Table_Name & " " & Field_Name
    Five_Section_Get Field_Name

'135�J�ʂ����

    Dim rw     As Single
    Dim id     As Long
    Dim kd     As Long
    ReDim R_135(135)

    '�܂�315���b�V�����o��
    For i = 1 To 315
        id = M_Link(i).id
        kd = M_Link(i).kd
        rw = R_Lank(R_1km(id, kd)) '�����N���J�ʂɕϊ�
        R_315(i) = rw
    Next i

    Mesh_To_Ryuiki_JWA R_315, R_135  '315���b�V����135����ɂ���

End Sub
Function TIMEC(d As Date) As String
    TIMEC = Format(d, "yyyy/mm/dd hh:nn")
End Function
