Attribute VB_Name = "Radar_Rain"
Option Explicit
Option Base 1

Public Const RRYU = 135                     '���搔
Public Radar_YRain_File        As String    'RSHINK�p���[�_�[�\�����敽�ωJ��
Public Radar_Rain_File         As String    'RSHINK�p���[�_�[���ї��敽�ωJ��
Public IRADAR                  As Integer   '���[�_�[�J�ʂ����鎞=1 �Ȃ���=0
Public JRADAR                  As Integer   '���[�_�[�J�ʂ��g����=1 �g��Ȃ���=0
Public Radar_File              As String    '���[�_�[�J�ʃt�@�C����
Public R_Thissen(20, 140)      As Single    '���[�_�[�J�ʗp�e�B�[�Z���W��(�P����ő�Q�O���b�V��)
Public R_Meshu(20, 140)        As Integer   '����J�ʗp���b�V���ԍ�(�P����ő�Q�O���b�V��)
Public R_T_Name(140)           As String    '���於�L��
Public rr()                    As Single    '���[�_�[���ї��敽�ωJ��
Public RY()                    As Single    '���[�_�[�\�����敽�ωJ��
Public RhY(5, 18)              As Single    '���[�_�[�\�����敽�ωJ��HANS�p

Public R_Ave(5, 500)           As Single    '��n�X�㗬���敽�ωJ��
Public R_Ave_N(135)            As Long      '��n�X�㗬���敽�ωJ�ʍ쐬�R���g���[��
Public R_Ave_Num               As Integer   '���搔

Public JMA_Num                 As Long
'
'�o�C�i���T�[�`
'
'
Sub Find_Rname(Xname As String, num As Long)

    Dim i1   As Long
    Dim i2   As Long
    Dim i3   As Long
    Dim j    As Long

    i1 = 1
    i2 = RRYU

f1:
    i3 = Int(i1 + i2) / 2
    If Xname > R_T_Name(i3) Then
        i1 = i3
    Else
        i2 = i3
    End If
    If Xname = R_T_Name(i3) Then
        num = i3
        Exit Sub
    End If
    If i2 - i1 <= 1 Then
        If Xname = R_T_Name(i1) Then
            num = i1
            Exit Sub
        Else
            num = i2
            Exit Sub
        End If
    End If
    GoTo f1

End Sub
'
'�C�ے��J�ʎ擾�f�[�^�`�F�b�N
'
'
Sub JMA_File_Open()

    Dim File    As String
    Dim L       As Long

    JMA_Num = FreeFile
    File = App.Path & "\data\�C�ے��J�ʃf�[�^�擾��ԃ`�F�b�N.dat"
    If Len(Dir(File)) > 0 Then
        L = FileLen(File)
        If L < 3000000 Then
            Open File For Append As #JMA_Num
        Else
            Open File For Output As #JMA_Num
        End If
    Else
        Open File For Output As #JMA_Num
    End If

End Sub
Sub JMA_OUT(msg As String)

    If LOF(JMA_Num) > 3000000 Then
        Close #JMA_Num
        JMA_File_Open
    End If

    Print #JMA_Num, Format(Now, "yyyy/mm/dd hh:nn:ss") & "|" & msg

End Sub
'**************************************************************
'�􉁉z���f�[�^�擾
'
'ds=��]�J�n����
'de=��]�I������
'
'
'dw=�擾�ŏI����
'
'*************************************************************
Sub MDB_��(ds As Date, de As Date, irc As Long)

    Dim i      As Long
    Dim j      As Long
    Dim k      As Long
    Dim b      As String
    Dim SQL    As String
    Dim mi     As String
    Dim nd1    As String
    Dim dw     As String
    Dim dew    As Date
    Dim w      As Single

    Dim ConR   As New ADODB.Recordset

    LOG_Out "IN  MDB_��"

    k = DateDiff("h", ds, de) + 1

    For i = 1 To k + 3
        HO(2, i) = 0#
    Next i

    mi = Format(Minute(de), "00")

    SQL = "select * from �� where Time between '" & Format(ds, "yyyy/mm/dd hh:nn") & "' and '" & _
           Format(de, "yyyy/mm/dd hh:nn") & "' and Minute = " & mi & " order by Time"

    ConR.Open SQL, Con_����, adOpenKeyset, adLockReadOnly

    If ConR.EOF Then
        LOG_Out "IN MB_�􉁃f�[�^�擾�ł���"
        LOG_Out Format(ds, "yyyy/mm/dd hh:nn") & " �` " & Format(de, "yyyy/mm/dd hh:nn")
        ConR.Close
        irc = False
        Exit Sub
    End If

    nd1 = Format(jgd, "yyyy/mm/dd hh:nn")
    j = 0
    Do Until ConR.EOF
        dw = ConR.Fields("Time").Value
        i = DateDiff("h", ds, dw) + 1
        w = ConR.Fields("Q0").Value    '���щz����
        HO(2, i) = w * 0.01
        If dw = nd1 Then
            w = ConR.Fields("Q1").Value  '1���Ԍ�\���z����
            HO(2, i + 1) = w * 0.01
            w = ConR.Fields("Q2").Value  '2���Ԍ�\���z����
            HO(2, i + 2) = w * 0.01
            w = ConR.Fields("Q3").Value  '3���Ԍ�\���z����
            HO(2, i + 3) = w * 0.01
            j = 1
            Exit Do
        End If
        ConR.MoveNext
    Loop
    ConR.Close

    If j = 1 Then
        irc = 0
    Else
        '�������̃f�[�^�����������̂�10���O�����ɍs�� 2006/03/31 15:31 In FRICS YOKOHAMA DC
        dew = DateAdd("n", -10, de)
        SQL = "select * from �� where Time ='" & Format(dew, "yyyy/mm/dd hh:nn") & "'"
        ConR.Open SQL, Con_����, adOpenKeyset, adLockReadOnly
        i = Now_Step
        If ConR.EOF Then
            '10���O����������
            irc = 2
            HO(2, i + 1) = 0#  '�\���z����
            HO(2, i + 2) = 0#  '�\���z����
            HO(2, i + 3) = 0#  '�\���z����
            ConR.Close
            LOG_Out "Out MDB_�� ����"
            Exit Sub
        Else
            '10���O��������
            i = DateDiff("h", ds, de) + 1
            HO(2, i + 1) = ConR.Fields("Q1").Value  '�\���z����
            HO(2, i + 2) = ConR.Fields("Q2").Value  '�\���z����
            HO(2, i + 3) = ConR.Fields("Q3").Value  '�\���z����
            irc = 1
            ConR.Close
        End If
    End If

    LOG_Out "Out MDB_��"

End Sub
'**************************************************************
'FRICS���щJ�ʎ擾
'
'ds=��]�J�n����
'de=��]�I������
'
'
'dw=�擾�ŏI����
'
'*************************************************************
Sub MDB_FRICS���[�_�[����(ds As Date, de As Date, dw As Date, irc As Boolean)

    Dim i      As Long
    Dim j      As Long
    Dim b      As String
    Dim SQL    As String
    Dim mi     As String
    Dim ConR   As New ADODB.Recordset

    LOG_Out "IN MB_FRICS���[�_�[���� " & ds & "�` " & de

    ReDim rr(500, 140)

    mi = Format(Minute(de), "00")

    SQL = "select * from FRICS���[�_�[���� where Time between '" & Format(ds, "yyyy/mm/dd hh:nn") & "' and '" & _
           Format(de, "yyyy/mm/dd hh:nn") & "' and Minute = " & mi & " order by Time"

    ConR.Open SQL, Con_����, adOpenKeyset, adLockReadOnly

    If ConR.EOF Then
        LOG_Out "IN MB_FRICS���[�_�[���юw��f�[�^�擾�ł���"
        LOG_Out Format(ds, "yyyy/mm/dd hh:nn") & " �` " & Format(de, "yyyy/mm/dd hh:nn")
        ConR.Close
        irc = False
        Exit Sub
    End If

    i = 0
    Do Until ConR.EOF
        dw = ConR.Fields("Time").Value
        i = DateDiff("h", ds, dw) + 1
        Debug.Print "FRICS i=" & Format(i, "000") & "  dw=" & Format(dw, "yyyy/mm/dd hh:nn")
        For j = 1 To 135
            b = Format(j, "###")
            rr(i, j) = ConR.Fields(b).Value    '����J��
        Next j
        ConR.MoveNext
    Loop
    irc = True
    ConR.Close


End Sub
'**************************************************************
'FRICS�\���~�J���擾
'
'ds=��]�J�n����
'
'
' RY(3, 140)
'dw=�擾�ŏI����
'
'*************************************************************
Sub MDB_FRICS���[�_�[�\��(ds As Date, irc As Boolean)

    Dim i      As Long
    Dim j      As Long
    Dim k      As Long
    Dim m      As Long
    Dim n      As Long
    Dim dur    As Date
    Dim dwr    As Date
    Dim dw     As Date
    Dim b      As String
    Dim SQL    As String
    Dim ConR   As New ADODB.Recordset
    Dim r
    Dim rw     As Single

    ReDim RY(3, 140)

    dur = ds
    SQL = "select * from FRICS���[�_�[�\�� where Time ='" & Format(dur, "yyyy/mm/dd hh:nn") & "' and " & _
          " Prediction_Minute IN( 60, 120, 180)"

    ConR.Open SQL, Con_����, adOpenKeyset, adLockReadOnly

    If ConR.EOF Then
        LOG_Out "IN MB_FRICS���[�_�[�\���w��f�[�^�擾�ł���"
        LOG_Out "SQL=" & SQL
        ORA_Message_Out "FRICS���[�_�J�ʎ�M", "FRICS�~�J�\�����v�Z���ɍ^���\���V�X�e���Ɏ�荞�܂�܂���ł����B�v�Z���X�L�b�v���܂��B", 1
        irc = False
        ConR.Close
        Exit Sub
    End If

    Do Until ConR.EOF
        dw = ConR.Fields("Time").Value
        k = DateDiff("h", jgd, dw)
        m = CLng(ConR.Fields("Prediction_Minute").Value / 60 + 0.4)
    Debug.Print "  m=" & Format(m, "##0") & " k=" & Format(k, "000") & "  dw=" & Format(dw, "yyyy/mm/dd hh:nn")
    Debug.Print "FRICS i=" & Format(i, "000") & " Now_Step=" & Format(Now_Step, "##0")
        For j = 1 To 135
            b = Format(j, "###")
            rw = CSng(ConR.Fields(b).Value)
            If rw > 250 Then
                RY(m, j) = 0                       '����J��
            Else
                RY(m, j) = rw                      '����J��
            End If
        Next j
NOP:
       ConR.MoveNext
    Loop

    irc = True
    ConR.Close


End Sub
'**************************************************************
'HANS��ʗpFRICS�\���~�J���擾
'
'ds=��]�J�n����
'
'
' RY(18, 140)  '10���s�b�`��3���Ԍ�܂�
'dw=�擾�ŏI����
'
'*************************************************************
Sub MDB_FRICS���[�_�[�\��_For_HANS(ds As Date, irc As Boolean)

    Dim i      As Long
    Dim j      As Long
    Dim k      As Long
    Dim m      As Long
    Dim n      As Long
    Dim dur    As Date
    Dim dwr    As Date
    Dim dw     As Date
    Dim b      As String
    Dim SQL    As String
    Dim ConR   As New ADODB.Recordset
    Dim r
    Dim rw     As Single

    Dim RYg(18, 140)  As Single

    LOG_Out "IN MDB_FRICS���[�_�[�\��_For_HANS " & ds

    Erase RhY '�o�͗p�G���A���N��������

    dur = ds
    SQL = "select * from FRICS���[�_�[�\�� where Time ='" & Format(dur, "yyyy/mm/dd hh:nn") & "'"

    ConR.Open SQL, Con_����, adOpenKeyset, adLockReadOnly

    If ConR.EOF Then
        irc = False
        ConR.Close
        Exit Sub
    End If

    Do Until ConR.EOF
        dw = ConR.Fields("Time").Value
        m = Int(ConR.Fields("Prediction_Minute").Value / 10 + 0.4)
    Debug.Print "  m=" & Format(m, "##0") & "  dw=" & Format(dw, "yyyy/mm/dd hh:nn")
    Debug.Print "FRICS i=" & Format(i, "000") & " Now_Step=" & Format(Now_Step, "##0")
        For j = 1 To 135
            b = Format(j, "###")
            rw = CSng(ConR.Fields(b).Value)
            If rw > 250 Then
                RYg(m, j) = 0             '����J�� ���v�Z���̂Ƃ���0�Ƃ���B
            Else
                RYg(m, j) = rw * 0.16667  '����J�� mm/hr�Ȃ̂�1/6����B
            End If
        Next j
       ConR.MoveNext
    Loop

    irc = True
    ConR.Close


'�\�����敽�ωJ��
    For j = 1 To 5      '5����
        For m = 1 To 18 '10���s�b�`��3���ԕ�
            r = 0
            For k = 1 To R_Ave_Num
                i = R_Ave_N(k)
                r = r + R_Ave(j, k) * RYg(m, i)
            Next k
            RhY(j, m) = r
        Next m
    Next j


    LOG_Out "OUT MDB_FRICS���[�_�[�\��_For_HANS "


End Sub

Sub ���[�_�[�J�ʍ�}�p���敽�ωJ�ʌv�Z()

    Dim i     As Long
    Dim j     As Long
    Dim k     As Long
    Dim m     As Long
    Dim r     As Single

    LOG_Out "  In  ���[�_�[�J�ʍ�}�p���敽�ωJ�ʌv�Z"

'���щJ��
    For i = 1 To Now_Step
        For j = 1 To 5
            r = 0
            For k = 1 To R_Ave_Num
                m = R_Ave_N(k)
                r = r + R_Ave(j, k) * rr(i, m)
            Next k
            RO(j, i) = r
        Next j
    Next i

'�\���J��
    For i = 1 To 3
        For j = 1 To 5
            r = 0
            For k = 1 To R_Ave_Num
                m = R_Ave_N(k)
                r = r + R_Ave(j, k) * RY(i, m)
            Next k
            RO(j, Now_Step + i) = r
        Next j
    Next i

    LOG_Out " Out  ���[�_�[�J�ʍ�}�p���敽�ωJ�ʌv�Z"

End Sub

Sub ���於�ǂݍ���()

    Dim k        As Integer
    Dim buf      As String
    Dim nf       As Integer

    LOG_Out "IN   Sub ���於�ǂݍ���"

    nf = FreeFile
    Open App.Path & "\data\���[�_�[�e�B�[�Z��.dat" For Input As #nf

    k = 0
    Do Until EOF(nf)
        Line Input #nf, buf
        k = k + 1
        R_T_Name(k) = Trim(Mid(buf, 6, 5))
    Loop

    Close #nf


    LOG_Out "Out  Sub ���於�ǂݍ���"

End Sub

'*************************************************
'���؎��̃��[�e�B��
'
'���ݎg�p����
'
'************************************************
Sub ���[�_�[�J�ʃZ�b�g()

    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim m               As Integer
    Dim ii              As Integer
    Dim jj              As Integer
    Dim nf              As Integer
    Dim nh              As Integer
    Dim r               As Single
    Dim rw              As Single
    Dim buf             As String
    Dim dw              As Date
    Dim RM(140)         As Single
    Dim pfct            As Integer

    If MAIN.Check3 Then
        pfct = 3
    Else
        pfct = 0
    End If

    LOG_Out "IN   Sub ���[�_�[�J�ʃZ�b�g"
    LOG_Out App.Path & "\���э^��\" & Radar_File & "  �ǂݍ���"

    nh = FreeFile
    Open App.Path & "\���э^��\" & Radar_File For Input As #nh        '���[�_�[���b�V���f�[�^
    Line Input #nh, buf

    For j = 1 To Now_Step + pfct
        Line Input #nh, buf
        For i = 1 To RRYU '135����
            rw = 0#
            For k = 1 To 20
                m = R_Meshu(k, i)  '���[�_�[���b�V���̔ԍ�
                If m = 0 Then Exit For
                r = CSng(Mid(buf, 17 + (m - 1) * 5, 5))
                rw = rw + r * R_Thissen(k, i)
            Next k
            If rw < 0# Then rw = 0#
            rr(j, i) = rw
        Next i

'��������\��
        Line Input #nh, buf      '1���Ԍ�\��
        For i = 1 To RRYU '135����
            rw = 0#
            For k = 1 To 20
                m = R_Meshu(k, i) '���[�_�[���b�V���̔ԍ�
                If m = 0 Then Exit For
                r = CSng(Mid(buf, 17 + (m - 1) * 5, 5))
                rw = rw + r * R_Thissen(k, i)
            Next k
            If rw < 0# Then rw = 0#
            RY(1, i) = rw
        Next i
        Line Input #nh, buf      '2���Ԍ�\��
        For i = 1 To RRYU '135����
            rw = 0#
            For k = 1 To 20
                m = R_Meshu(k, i)  '���[�_�[���b�V���̔ԍ�
                If m = 0 Then Exit For
                r = CSng(Mid(buf, 17 + (m - 1) * 5, 5))
                rw = rw + r * R_Thissen(k, i)
            Next k
            If rw < 0# Then rw = 0#
            RY(2, i) = rw
        Next i
        Line Input #nh, buf      '3���Ԍ�\��
        For i = 1 To RRYU '135����
            rw = 0#
            For k = 1 To 20
                m = R_Meshu(k, i) '���[�_�[���b�V���̔ԍ�
                If m = 0 Then Exit For
                r = CSng(Mid(buf, 17 + (m - 1) * 5, 5))
                rw = rw + r * R_Thissen(k, i)
            Next k
            If rw < 0# Then rw = 0#
            RY(3, i) = rw
        Next i
    Next j

    Close #nh


End Sub
'
'**************************************************
'
'RSHINKAWA�p�Ƀ��[�_�[�J�ʂ��o�͂���
'
'���ؗp�\���J�ʂȂ��i���э~�J�Ōv�Z�j
'
'
'
'**************************************************
'
Sub ���[�_�[�J�ʏo��_Veri()

    Dim i               As Long
    Dim j               As Long
    Dim k               As Long
    Dim ii              As Long
    Dim jj              As Long
    Dim nf              As Long
    Dim k1              As Long
    Dim k2              As Long
    Dim Steps           As Long
    Dim buf             As String
    Dim dw              As Date


    If Verification2.Check1 <> vbChecked Then
        For i = 1 To 3
            For ii = 1 To RRYU
                If rr(Now_Step + i, ii) = 0# Then RY(i, ii) = 0.1
            Next ii
        Next i
        Steps = All_Step
    End If

    jj = Fix((All_Step - 1) / 12) + 1

'���тƗ\�����[�_�[�f�[�^�o��
    nf = FreeFile
    Open App.Path & "\work\���敽�ωJ��.dat" For Output As #nf  '���[�_�[���ї��敽�ωJ�ʏo��

    For ii = 1 To RRYU
        buf = Format(Format(ii, "####0"), "@@@@@") & "     1.E-1    1"
        Print #nf, buf

        For j = 1 To jj
            k1 = (j - 1) * 12 + 1
            k2 = k1 + 11
            If k2 > All_Step Then k2 = All_Step
            buf = ""
            For k = k1 To k2
                If rr(k, ii) > 0# Then
                    buf = buf & Format(Format(rr(k, ii) * 10, "####0"), "@@@@@")
                Else
                    buf = buf & "1.E-0"
                End If
            Next k
            Print #nf, Space(10) & buf
        Next j
    Next ii

    Close #nf

    ���[�_�[�J�ʍ�}�p���敽�ωJ�ʌv�Z

End Sub
'**************************************************
'RSHINKAWA�p�Ƀ��[�_�[�J�ʂ��o�͂���
'
'2003/10/01 �o�͌`����ύX
'
'
'
'**************************************************
Sub ���[�_�[�J�ʏo��()

    Dim i               As Long
    Dim j               As Long
    Dim k               As Long
    Dim k1              As Long
    Dim k2              As Long
    Dim ii              As Long
    Dim jj              As Long
    Dim nf              As Long
    Dim buf             As String
    Dim dw              As Date


    For i = 1 To 3
        For ii = 1 To RRYU
            rr(Now_Step + i, ii) = RY(i, ii)
        Next ii
    Next i

    jj = Fix((All_Step - 1) / 12) + 1

'���тƗ\�����[�_�[�f�[�^�o��
    nf = FreeFile
    Open App.Path & "\work\���敽�ωJ��.dat" For Output As #nf  '���[�_�[���ї��敽�ωJ�ʏo��

    For ii = 1 To RRYU
        buf = Format(Format(ii, "####0"), "@@@@@") & "     1.E-1    1"
        Print #nf, buf

        For j = 1 To jj
            k1 = (j - 1) * 12 + 1
            k2 = k1 + 11
            If k2 > All_Step Then k2 = All_Step
            buf = ""
            For k = k1 To k2
                If rr(k, ii) > 0# Then
                    buf = buf & Format(Format(rr(k, ii) * 10, "####0"), "@@@@@")
                Else
                    buf = buf & "1.E-0"
                End If
            Next k
            Print #nf, Space(10) & buf
        Next j
    Next ii

    Close #nf

    ���[�_�[�J�ʍ�}�p���敽�ωJ�ʌv�Z

End Sub

Sub ��n�_�Ɨ���Ή���ǂ�()

    Dim i     As Long
    Dim j     As Long
    Dim n     As Long
    Dim buf   As String
    Dim nf    As Integer
    Dim a(5)  As Single
    Dim b     As Single
    Dim c(5)  As Integer
    Dim S     As String

    LOG_Out "  In  ��n�_�Ɨ���Ή���ǂ�"

    nf = FreeFile
    Open App.Path & "\data\��n�_�Ɨ���Ή�.txt" For Input As #nf

    Line Input #nf, buf
    Line Input #nf, buf

    i = 0
    Do Until EOF(nf)
        Line Input #nf, buf
        i = i + 1
        c(1) = 1                                 '���V��F
        c(2) = IIf(Mid(buf, 1, 1) <> " ", 1, 0)  '�厡
        c(3) = IIf(Mid(buf, 31, 1) <> " ", 1, 0) '�����O����
        c(4) = IIf(Mid(buf, 21, 1) <> " ", 1, 0) '�v�n��
        c(5) = IIf(Mid(buf, 11, 1) <> " ", 1, 0) '�t��
        b = CSng(Mid(buf, 46, 10))
        For j = 1 To 5
            If c(j) > 0 Then
                R_Ave(j, i) = b
                a(j) = a(j) + b
            End If
        Next j
        S = Trim(Mid(buf, 40, 5))
        Find_Rname S, n
        R_Ave_N(i) = n
    Loop
    R_Ave_Num = i
    For i = 1 To R_Ave_Num
        For j = 1 To 5
'            If j = 3 Then Debug.Print " i="; i; "  a(j)="; a(j); "  R_Ave="; R_Ave(j, i); "  R_Ave(j, i) / a(j)="; R_Ave(j, i) / a(j)
            R_Ave(j, i) = R_Ave(j, i) / a(j)
        Next j
    Next i

    Close #nf

    LOG_Out " Out  ��n�_�Ɨ���Ή���ǂ�"

End Sub

'**************************************************************
'�C�ے����[�_�[�f�[�^�\�����擾
'
'�J�ʂ�mm/Hour�œo�^����Ă���
'
'2km���b�V���Ή��Ȃ̂Ŏg�p��~ 2007/05/02
'
'**************************************************************
Sub MDB_�C�ے����[�_�[�\��(ds As Date, de As Date, irc As Boolean)

    Dim i      As Long
    Dim j      As Long
    Dim k      As Long
    Dim buf    As String
    Dim du     As Date
    Dim dl     As Date
    Dim dw     As Date
    Dim Conn   As String
    Dim ConS   As New ADODB.Connection
    Dim ConR   As New ADODB.Recordset
    Dim a
    Dim SQL    As String
    Dim mi     As String
    Dim rw     As Single

    LOG_Out "IN   Sub MDB_�C�ے����[�_�[�\��"

'    Con = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source= " & App.Path & "\data\����.mdb"
'    Conn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=D:\SHINKAWA\OracleTest\oraDB\Data\����.mdb"
    Conn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & ����MDB
    ConS.ConnectionString = Conn
    ConS.Open

    irc = False

    ReDim RY(3, 140)

    If Err <> 0 Then
        MsgBox "����.MDB�ɃA�N�Z�X�ł��܂���A����.MDB�̗L�����m�F���Ă��������B" & vbCrLf & _
               "�v�Z�ł��܂���̂Ńv���u�����͏I�����܂��B", vbExclamation
        End
    End If

    Set ConR.ActiveConnection = ConS

    mi = Fix(Minute(de) / 10) * 10

'�P���Ԍ�
    SQL = "select * from �C�ے����[�_�[�\��_1 where Time= '" & Format(ds, "yyyy/mm/dd hh:nn") & "' and " & _
          "  Minute = " & mi    '& " order by Time"

    ConR.Open SQL, ConS, adOpenDynamic, adLockOptimistic

    If ConR.EOF Then
        LOG_Out "IN MB_�C�ے����[�_�[�\���w��f�[�^�擾�ł���"
        LOG_Out Format(ds, "yyyy/mm/dd hh:nn") & " �` " & Format(de, "yyyy/mm/dd hh:nn")
        ConR.Close
        Exit Sub
    End If

    dw = ConR.Fields("Time").Value
    k = DateDiff("h", jgd, dw)
'    Debug.Print " k=" & Format(k, "000") & "  dw=" & Format(dw, "yyyy/mm/dd hh:nn")
    If k <> 1 Then
        LOG_Out "IN MDB_�C�ے����[�_�[�\�� �����ɂ��Ă͂����܂���B"
        LOG_Out " jgd=" & Format(jgd, "yyyy/mm/dd hh:nn")
        LOG_Out "  ds=" & Format(ds, "yyyy/mm/dd hh:nn")
        LOG_Out "  dw=" & Format(dw, "yyyy/mm/dd hh:nn")
    End If
    Debug.Print "  dw="; dw; "  ";
    For j = 1 To 135
        a = Format(j, "###")
        RY(k, j) = ConR.Fields(a).Value * 0.1  '����J��
        Debug.Print Format(RY(k, j), "##0.0 ");
    Next j
    Debug.Print ""
    ConR.Close

'�Q�`�R���Ԍ�
    For k = 2 To 3
        dl = DateAdd("h", k, jgd)
        SQL = "select * from �C�ے����[�_�[�\��_2 where Time= '" & Format(dl, "yyyy/mm/dd hh:nn") & "'"

        ConR.Open SQL, ConS, adOpenDynamic, adLockOptimistic

        If ConR.EOF Then
            LOG_Out "IN MB_�C�ے����[�_�[�\��_2�w��f�[�^�擾�ł��� ����=" & Format(dl, "yyyy/mm/dd hh:nn")
            ORA_Message_Out "�C�ے����[�_�J�ʎ�M", "�C�ے��Z���ԍ~�J�\�����擾�ł��Ă��܂���B", 1
            irc = False
            ConR.Close
            Exit Sub
        Else
            For j = 1 To 135
                a = Format(j, "###")
                RY(k, j) = RY(k, j) + ConR.Fields(a).Value * 0.1 '����J��
            Next j
            Debug.Print "  dl="; dl; "  ";
            For j = 1 To 135
                Debug.Print Format(RY(k, j), "##0.0 ");
            Next j
            Debug.Print ""
        End If
        irc = True
        ConR.Close
    Next k

End Sub
'**************************************************************
'�C�ے����[�_�[�f�[�^�\�����擾
'
'�J�ʂ�mm/10�œo�^����Ă���
'
'2007/05/02 22:37 �V�K�쐬
'
'**************************************************************
Sub MDB_�C�ے����[�_�[�\��2(ds As Date, de As Date, irc As Boolean)

    Dim i      As Integer
    Dim j      As Integer
    Dim k      As Integer
    Dim n      As Integer
    Dim m      As Integer
    Dim a
    Dim SQL    As String
    Dim d1     As Date
    Dim d2     As Date
    Dim dw     As Date
    Dim d1c    As String
    Dim dsc    As String
    Dim dec    As String

    Dim RM(140)         As Single

    ReDim RY(3, 140)

    dsc = TIMEC(ds)
    dec = TIMEC(de)

    LOG_Out "IN   Sub MDB_�C�ے����[�_�[�\��2 " & dsc & "�` " & dec
    JMA_OUT "IN   Sub MDB_�C�ے����[�_�[�\��2 " & dsc & "�` " & dec

    n = DateDiff("h", ds, de) + 1
    d1 = DateAdd("n", -50, ds)
    d2 = de
    SQL = "select * from �C�ے����[�_�[�\��_1 where Time between '" & TIMEC(d1) & "' and '" & _
           TIMEC(d2) & "' ORDER BY Time"

    Rec_����.Open SQL, Con_����, adOpenDynamic, adLockOptimistic
    d1 = DateAdd("n", -50, ds)
    d2 = d1
    dw = ds
    For k = 1 To n
        Erase RM
        JMA_OUT "                    ���ԉJ�ʍ쐬 " & TIMEC(dw)
        For m = 1 To 6 '6�X�e�b�v�𑫂��Ď��ԉJ�ʂɂ���
            d1c = TIMEC(d1)
            Rec_����.Find "Time = '" & d1c & "'"
            If Rec_����.EOF Then
                LOG_Out "IN MDB_�C�ے����[�_�[�\��2 �w��f�[�^�擾�ł���"
                LOG_Out d1c
                JMA_OUT "                 �����J�ʎ擾 " & d1c & " �擾�ł���"
            Else
                JMA_OUT "                 �����J�ʎ擾 " & d1c
                For j = 1 To 135
                    a = Format(j, "###")
                    RM(j) = RM(j) + Rec_����.Fields(a).Value * 0.1  '����J��
                Next j
            End If
            Rec_����.MoveFirst
            d1 = DateAdd("n", 10, d1)
        Next m

        For j = 1 To 135
            RY(k, j) = RM(j)    '����J��
        Next j

        d1 = DateAdd("h", k, d2)
        dw = DateAdd("h", 1, dw)

    Next k
    irc = True
    Rec_����.Close

    LOG_Out "Out  Sub MDB_�C�ے����[�_�[�\��2 " & dsc & "�` " & dec
    JMA_OUT "Out  Sub MDB_�C�ے����[�_�[�\��2 " & dsc & "�` " & dec

End Sub
'**************************************************************
'�C�ے����[�_�[�f�[�^���т��擾
'
'���ݎg���Ă��Ȃ�
'
'**************************************************************
Sub MDB_�C�ے����[�_�[����(ds As Date, de As Date, dw As Date, irc As Boolean)


    Dim Conn   As String
    Dim ConS   As New ADODB.Connection
    Dim ConR   As New ADODB.Recordset

    Dim i      As Integer
    Dim j      As Integer
    Dim k      As Integer
    Dim n      As Integer
    Dim a
    Dim SQL    As String
    Dim d1     As Date
    Dim d2     As Date

    ReDim rr(500, 140)

    LOG_Out "IN   Sub MDB_�C�ے����[�_�[���� " & ds & "�` " & de

'    Con = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source= " & App.Path & "\data\����.mdb"
'    Conn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=D:\SHINKAWA\OracleTest\oraDB\Data\����.mdb"
    Conn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & ����MDB
    ConS.ConnectionString = Conn
    ConS.Open

    If Err <> 0 Then
        MsgBox "����.MDB�ɃA�N�Z�X�ł��܂���A����.MDB�̗L�����m�F���Ă��������B" & vbCrLf & _
               "�v�Z�ł��܂���̂Ńv���u�����͏I�����܂��B", vbExclamation
        End
    End If

    Set ConR.ActiveConnection = ConS

    n = DateDiff("h", ds, de) + 1
    d1 = DateAdd("n", -50, ds)
    d2 = ds
    For k = 1 To n

        ReDim RM(140) As Single

        SQL = "select * from �C�ے����[�_�[���� where Time between '" & Format(d1, "yyyy/mm/dd hh:nn") & "' and '" & _
               Format(d2, "yyyy/mm/dd hh:nn") & "' "

        ConR.Open SQL, ConS, adOpenDynamic, adLockOptimistic
        If ConR.EOF Then
            LOG_Out "IN MB_�C�ے����[�_�[���юw��f�[�^�擾�ł���"
            LOG_Out Format(d1, "yyyy/mm/dd hh:nn") & " �` " & Format(d2, "yyyy/mm/dd hh:nn")
        Else
            Do Until ConR.EOF
                dw = ConR.Fields("Time").Value
                i = DateDiff("h", ds, dw) + 1
'     Debug.Print " i=" & Format(i, "000") & "  dw=" & Format(dw, "yyyy/mm/dd hh:nn")
                For j = 1 To 135
                    a = Format(j, "###")
                    RM(j) = RM(j) + ConR.Fields(a).Value * 0.1  '����J��
                Next j
                ConR.MoveNext
            Loop
        End If

'        Debug.Print "  k="; k; "  d2="; d2; "  RM="; RM(1)
        For j = 1 To 135
            rr(k, j) = RM(j)    '����J��
        Next j

        ConR.Close
        d1 = DateAdd("h", 1, d1)
        d2 = DateAdd("h", 1, d2)

    Next k
    irc = True
    dw = de

    LOG_Out "Out  Sub MDB_�C�ے����[�_�[���� " & ds & "�` " & de

End Sub
'
'**************************************************************
'
'�C�ے����[�_�[�f�[�^���т��擾
'���̃T�u���[�e�B���͌��ؗp�ł��B                     ????????2010/03/09
'�\���J�ʂ��g��Ȃ��Ŏ��уf�[�^�݂̂Ōv�Z�ł���悤�� ????????2010/03/09
'�ݒ�Ȃ��Ă��܂��A�\���J�ʂ͓ǂ݂܂���B             ????????2010/03/09
'
'�J�ʂ�mm/10min�œo�^����Ă���
'
'
'**************************************************************
'
Sub MDB_�C�ے����[�_�[����2(ds As Date, de As Date, dw1 As Date, irc As Boolean)

    Dim i      As Integer
    Dim j      As Integer
    Dim k      As Integer
    Dim n      As Integer
    Dim m      As Integer
    Dim a
    Dim SQL    As String
    Dim d1     As Date
    Dim d2     As Date
    Dim dw     As Date
    Dim d1c    As String
    Dim dsc    As String
    Dim dec    As String

    Dim RM(140)         As Single

    dsc = TIMEC(ds)
    dec = TIMEC(de)

    LOG_Out "IN   Sub MDB_�C�ے����[�_�[����2 " & dsc & "�` " & dec
    JMA_OUT "IN   Sub MDB_�C�ے����[�_�[����2 " & dsc & "�` " & dec

    ReDim rr(500, 140)

    n = DateDiff("h", ds, de) + 1
    d1 = DateAdd("n", -50, ds)
    d2 = de
    SQL = "select * from �C�ے����[�_�[���� where Time between '" & TIMEC(d1) & "' and '" & _
           TIMEC(d2) & "' ORDER BY Time"

    Rec_����.Open SQL, Con_����, adOpenDynamic, adLockOptimistic
    Do
        Debug.Print Rec_����.Fields("Time").Value
        Rec_����.MoveNext
    Loop Until Rec_����.EOF
    Rec_����.MoveFirst
    JMA_OUT SQL
    d1 = DateAdd("n", -50, ds)
    d2 = d1
    dw = ds
    For k = 1 To n
        Erase RM
        JMA_OUT "                    ���ԉJ�ʍ쐬 " & TIMEC(dw)
        For m = 1 To 6 '6�X�e�b�v�𑫂��Ď��ԉJ�ʂɂ���
            d1c = TIMEC(d1)
            Rec_����.Find "Time ='" & d1c & "'"
            If Rec_����.EOF Then
                LOG_Out "IN MB_�C�ے����[�_�[���юw��f�[�^�擾�ł���"
                LOG_Out d1c
                JMA_OUT "                 �����J�ʎ擾 " & d1c & " �擾�ł���"
            Else
                JMA_OUT "                 �����J�ʎ擾 " & d1c
                For j = 1 To 135
                    a = Format(j, "###")
                    RM(j) = RM(j) + Rec_����.Fields(a).Value * 0.1  '����J��
                Next j
            End If
            Rec_����.MoveFirst
            d1 = DateAdd("n", 10, d1)
        Next m

        For j = 1 To 135
            rr(k, j) = RM(j)    '����J��
        Next j

        d1 = DateAdd("h", k, d2)
        dw = DateAdd("h", 1, dw)

    Next k
    irc = True
    dw1 = de
    Rec_����.Close

    LOG_Out "Out  Sub MDB_�C�ے����[�_�[����2 " & dsc & "�` " & dec
    JMA_OUT "Out  Sub MDB_�C�ے����[�_�[����2 " & dsc & "�` " & dec

End Sub
