Attribute VB_Name = "Make_Pump_Data"
Option Explicit
Option Base 1
Public Pump_Data(21)  As Pump_Stats
Type Pump_Stats
     name        As String '�|���v����
     ability(8)  As Single '�e�|���v�̔\��
     s_num       As Long   '�|���v����
     p_num       As Long   '�|���v�䐔
     p_base      As Single '�Œ�r����(��荞�݂��o���Ȃ��|���v�{�݂����v�����l)
     max         As Single '�|���v�f�[�^�������́A�����̎���max���g��
End Type
Public P_Hina(183) As String  '�|���v���^�f�[�^ ����183=�f�[�^������
Public P_Ctl(19)   As P_C     '�ϑ��|���v�f�[�^����|���v�f�[�^���쐬���鎞�̃R���g���[���f�[�^
Type P_C
    op    As Long    '�{�v���O������̃|���v���ԍ�
    np    As Long    '�|���v���^�f�[�^�̃|���v�ԍ�
    pp    As Long    '���^�f�[�^��̏��ԍ�
End Type
Public P_Hina_Flag As Boolean  'True=�|���v���^���擾�ł�����
Sub FULL_PUMP_OUT()

    Dim nf      As Long
    Dim i       As Long
    Dim Out_F   As String
    Dim msg(3)  As String
    Dim msgD(3) As Date
    Dim dw      As Date

    msg(1) = "": msg(2) = "": msg(3) = ""

    LOG_Out "In  FULL_PUMP_OUT"

    If HO(3, Now_Step) < H_Stand2(1, 1) And HO(3, Now_Step) > -20# Then
        msg(1) = "�|���v�f�[�^���擾�ł��܂���ł����A���V��F�ϑ��ǐ��ʂ��u�|���v��~���ʁv�ȉ��Ȃ̂ŁA����|���v��̔\�͂ƒX���ʂ̊֌W����f�[�^��ݒ肵�܂��B"
        LOG_Out "In  FULL_PUMP_OUT" & msg(1)
        msgD(1) = DateAdd("s", 21, jgd)
    End If
    If HO(5, Now_Step) < H_Stand2(3, 1) And HO(5, Now_Step) > -20# Then
        msg(2) = "�|���v�f�[�^���擾�ł��܂���ł����A�����O���ʊϑ��ǂ��u�|���v��~���ʁv�ȉ��Ȃ̂ŁA����|���v��̔\�͂ƒX���ʂ̊֌W����f�[�^��ݒ肵�܂��B"
        LOG_Out "In  FULL_PUMP_OUT" & msg(2)
        msgD(2) = DateAdd("s", 22, jgd)
    End If
    If HO(7, Now_Step) < H_Stand2(5, 1) And HO(7, Now_Step) > -20# Then
        msg(3) = "�|���v�f�[�^���擾�ł��܂���ł����A�t���ϑ��ǂ��u�|���v��~���ʁv�ȉ��Ȃ̂ŁA����|���v��̔\�͂ƒX���ʂ̊֌W����f�[�^��ݒ肵�܂��B"
        LOG_Out "In  FULL_PUMP_OUT" & msg(3)
        msgD(3) = DateAdd("s", 23, jgd)
    End If
    dw = jgd
    For i = 1 To 3
        If msg(i) <> "" Then
            jgd = msgD(i)
            ORA_Message_Out "�|���v�f�[�^��M", msg(i), 1
        End If
    Next i
    jgd = dw
    '�|���v���^�f�[�^�����̂܂܎g��
    nf = FreeFile
    Out_F = App.Path & "\WORK\Pump.dat"
    Open Out_F For Output As #nf
    For i = 1 To 183
        Print #nf, P_Hina(i)
    Next i
    Close #nf

    LOG_Out "Out FULL_PUMP_OUT"

End Sub
'
'
'G_Time ---- �|���v�f�[�^���쐬���鎞�̌�����
'
'
'
Sub �|���v�f�[�^�쐬(G_Time As Date)

    Dim i         As Long
    Dim j         As Long
    Dim k         As Long
    Dim m         As Long
    Dim n         As Long
    Dim buf       As String
    Dim dat(185)  As String
    Dim Out_F     As String
    Dim nf        As Long
    Dim ds        As Date
    Dim dw        As Date
    Dim mn        As Long
    Dim ip        As Long
    Dim Pump_W(19, 16) As Single '19=�|���v�� 16=���ԃX�e�b�v
    Dim Pump_Obs(16)   As Single
    Dim SQL       As String
    Dim T         As String
    Dim p_name    As String
    Dim w
    Dim pt        As Single
    Dim base      As Single
    Dim maxp      As Single
    Dim d1        As String
    Dim d2        As String

    Dim Pump_Dat(22, 17) As Single

'    On Error GoTo ERHr10

    LOG_Out "In  �|���v�f�[�^�쐬"

    Const N_Pump = 19  '�|���v�ꐔ

    '�|���v���}�b�N�X�ŏ�����
    For i = 1 To N_Pump
        pt = Pump_Data(i).max
        For j = 1 To 16
            Pump_W(i, j) = pt
        Next j
    Next i

    If P_Hina_Flag = False Then
        GoTo FULL_PUMP
    End If

    ds = DateAdd("h", -12, G_Time)
    mn = Minute(G_Time)
    d1 = Format(ds, "yyyy/mm/dd hh:nn")          '�����
    d2 = Format(G_Time, "yyyy/mm/dd hh:nn")      '����̒�`�����΂������̂ŏC��2008/0/29
    dw = ds
    SQL = "SELECT * FROM �|���v���� WHERE " & _
          "TIME BETWEEN '" & d1 & "' AND '" & d2 & "' AND Minute=" & Format(mn, "#0")
    Rec_����.Open SQL, Con_����, adOpenDynamic, adLockOptimistic

    LOG_Out "In  �|���v�f�[�^�쐬 SQL=" & SQL

    If Rec_����.EOF Then
        LOG_Out " �f�[�^���擾�ł��Ȃ������̂Ńt���|���v�f�[�^���o�́B"
        FULL_PUMP_OUT '�f�[�^���擾�ł��Ȃ������̂Ő��^�f�[�^���o��
        Rec_����.Close
        Exit Sub
    End If

'--------- �t�B�[���h����������� -------------------------------
'    n = Rec_����.Fields.Count
'    For i = 0 To n - 1
'        buf = Rec_����.Fields(i).name
'        Debug.Print " Number=" & str(i) & " �t�B�[���h��=" & buf
'    Next i
'---------------------------------------------------------------

    Do Until Rec_����.EOF
        T = Rec_����.Fields("Time").Value
        j = DateDiff("h", ds, T) + 1
        '�|���v�ꐔ��
        For i = 1 To N_Pump
            p_name = Pump_Data(i).name '�e�|���v��̖��O
            buf = Rec_����.Fields(p_name).Value
'buf = "0,1,1,0,0,0,0,0,0,0,0,0,0,"
            base = Pump_Data(i).p_base
            w = Split(buf, ",")
            pt = base
            For m = 1 To Pump_Data(i).p_num '�e�|���v��̃|���v��
                pt = pt + Pump_Data(i).ability(m) * w(m - 1)
            Next m
            Pump_W(i, j) = pt
        Next i
        Rec_����.MoveNext
    Loop
    Rec_����.Close
    '�������p����3���ԗ\��
    For i = 1 To N_Pump
        pt = Pump_W(i, 13)
        For j = 14 To 16
            Pump_W(i, j) = pt
        Next j
    Next i

FULL_PUMP:

    '�|���v�f�[�^�쐬
    For i = 1 To N_Pump
        n = P_Ctl(i).np
        For j = 1 To 16
            k = j + 1
            Pump_Dat(n, k) = Pump_Dat(n, k) + Pump_W(i, j)
        Next j
    Next i

    For i = 1 To N_Pump
        n = P_Ctl(i).np
        m = P_Ctl(i).pp
        k = InStr(P_Hina(m), "99") - 1
        d1 = Mid(P_Hina(m), 1, k) & "9999"
        For j = 2 To 10
            d1 = d1 & Format(Format(Pump_Dat(n, j), "#0.00"), "@@@@@")
        Next j
        d2 = Space(10)
        For j = 11 To 17
            d2 = d2 & Format(Format(Pump_Dat(n, j), "#0.00"), "@@@@@")
        Next j

        P_Hina(m) = d1
        P_Hina(m + 1) = d2

    Next i

    Out_F = App.Path & "\WORK\Pump.dat"
    n = FreeFile
    Open Out_F For Output As #n
    For i = 1 To 183
        Print #n, P_Hina(i)
    Next i
    Close #n

    LOG_Out "Out �|���v�f�[�^�쐬"
    On Error GoTo 0
    Exit Sub

ERHr10:
    On Error GoTo 0
    LOG_Out " �|���v�f�[�^�쐬���ɃG���[�����������̂Ńt���|���v�f�[�^���g���B"
    FULL_PUMP_OUT

End Sub
'
'�|���v���^�f�[�^�ƃR���g���[���f�[�^��ǂ�
'
'
'
Sub �|���v���^�f�[�^�ǂݍ���()

    Dim i      As Long
    Dim j      As Long
    Dim k
    Dim buf    As String
    Dim nf     As Long
    Dim F      As String

    On Error GoTo ERH3

    LOG_Out "In  �|���v���^�f�[�^�ǂݍ���"

    F = App.Path & "\Data\�|���v���^.txt"
    nf = FreeFile
    Open F For Input As #nf
    For i = 1 To 183  '183=�f�[�^��
        Line Input #nf, buf
        P_Hina(i) = buf
    Next i

    LOG_Out "�f�[�^�쐬�R���g���[���f�[�^�ǂݍ���"
    Line Input #nf, buf
    Line Input #nf, buf
    For i = 1 To 19
        Line Input #nf, buf
            P_Ctl(i).op = CLng(Mid(buf, 1, 3))    '�{�v���O������̃|���v���ԍ�
            P_Ctl(i).np = CLng(Mid(buf, 14, 3))   '�|���v���^�f�[�^�̃|���v�ԍ�
            P_Ctl(i).pp = CLng(Mid(buf, 29, 3))   '���^�f�[�^��̏��ԍ�
    Next i
    Close #nf
'    P_Hina_Flag = True
    P_Hina_Flag = False
    LOG_Out "Out �|���v���^�f�[�^�ǂݍ��ݑ听��"
    On Error GoTo 0
    Exit Sub

ERH3:
    On Error GoTo 0
    ORA_Message_Out "�|���v�f�[�^��M", "�|���v���^�f�[�^�̓ǂݍ��݂Ɏ��s���Ă��܂��A�t���|���v�Ōv�Z���܂��B", 1
    P_Hina_Flag = False

End Sub
Sub �|���v�\�͕\�ǂݍ���()

    Dim i        As Long
    Dim j        As Long
    Dim k        As Long
    Dim buf      As String
    Dim File     As String
    Dim nf       As Long
    Dim p        As Single

    LOG_Out "In   �|���v�\�͕\�ǂݍ���"

'    On Error GoTo EHR1

    File = App.Path & "\data\�|���v�\�͕\.dat"
    nf = FreeFile
    Open File For Input As #nf

    Line Input #nf, buf '�f�[�^�^�C�g��1
    Line Input #nf, buf '�f�[�^�^�C�g��2

    For i = 1 To 19 '19�|���v��
        Line Input #nf, buf
        j = CLng(Mid(buf, 1, 2))
        Pump_Data(j).name = Mid(buf, 72, 7)         '�|���v��
        Pump_Data(j).s_num = CLng(Mid(buf, 1, 2))   '�|���v����
        Pump_Data(j).p_num = CLng(Mid(buf, 6, 5))   '�|���v��
        Pump_Data(j).p_base = CSng(Mid(buf, 56, 5)) '�|���v�x�[�X�r����
        Pump_Data(j).max = CSng(Mid(buf, 61, 5))    '�|���v�ő�r���� �������f�[�^�������Ȃ��������Ɏg��
Debug.Print " s_num="; Pump_Data(j).s_num
Debug.Print " Name="; Pump_Data(j).name
Debug.Print " p_num="; Pump_Data(j).p_num
Debug.Print " p_base="; Pump_Data(j).p_base
Debug.Print " max="; Pump_Data(j).max
        '�e�|���v���@���̔\��
        For k = 1 To Pump_Data(j).p_num  '�|���v��
            p = CSng(Mid(buf, (k - 1) * 5 + 16, 5))
            Pump_Data(i).ability(k) = p
'            Debug.Print " k="; k; " ability="; p
        Next k
    Next i

    Close #nf

    P_Hina_Flag = True
    LOG_Out "Out  �|���v�\�͕\�ǂݍ���"

    On Error GoTo 0
    Exit Sub
EHR1:
    On Error GoTo 0
    ORA_Message_Out "�|���v�f�[�^�ǂݍ���", "�|���v�\�̓f�[�^�̓ǂݍ��݂Ɏ��s���Ă��܂��A�t���|���v�Ōv�Z���܂��B", 1
    
    P_Hina_Flag = False

End Sub
