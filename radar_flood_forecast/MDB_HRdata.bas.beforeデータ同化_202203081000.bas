Attribute VB_Name = "MDB_HRdata"
Option Explicit
Option Base 1
Public wH(6, 25)   As Single
Public DH_Tide     As Single
'
'�����c�a���f�[�^����荞��
'
'�C���L�^
'�����̕�U�͌����������Ƃ��� 2004/03/24
'
Sub Data_IN(ds As Date, de As Date, irc As Boolean)

    Dim i      As Long
    Dim j      As Integer
    Dim k      As Integer
    Dim m      As Integer
    Dim b      As String
    Dim du     As Date
    Dim dw     As Date
    Dim dur    As Date
    Dim dwr    As Date
    Dim ConR   As New ADODB.Recordset
    Dim a
    Dim SQL    As String
    Dim mi     As String
    Dim C0     As Single
    Dim C1     As Single
    Dim C2     As Single
    Dim C3     As Single
    Dim ch     As Boolean
    Dim uh     As Boolean
    Dim hw(4)  As Single
    Dim er     As Boolean

    If Err <> 0 Then
        MsgBox "����.MDB�ɃA�N�Z�X�ł��܂���A����.MDB�̗L�����m�F���Ă��������B" & vbCrLf & _
               "�v�Z�ł��܂���̂Ńv���u�����͏I�����܂��B", vbExclamation
        End
    End If

    mi = Fix(Minute(de) / 10) * 10

'���ʎ擾
    SQL = "select * from ���� where Time between '" & Format(ds, "yyyy/mm/dd hh:nn") & "' and '" & _
           Format(de, "yyyy/mm/dd hh:nn") & "' and Minute = " & mi & " order by Time"
    Short_Break 4
    ConR.Open SQL, Con_����, adOpenKeyset, adLockReadOnly
    i = 0
    Do Until ConR.EOF
        dw = ConR.Fields("Time").Value
        If i = 0 Then
            du = dw
        End If
        i = DateDiff("h", du, dw) + 1
        HO(1, i) = ConR.Fields("Tide").Value       'Tide ������O����
        HO(2, i) = 0#                              '�􉁉z����
        HO(3, i) = ConR.Fields("���V��F").Value   '���V��F
        HO(4, i) = ConR.Fields("�厡").Value       '�厡
        HO(5, i) = ConR.Fields("�����O").Value   '����O
        HO(6, i) = ConR.Fields("�v�n��").Value     '�v�n��
        HO(7, i) = ConR.Fields("�t��").Value       '�t��
        ConR.MoveNext
    Loop
    ConR.Close

    If HO(1, Now_Step) < -50# Then
        Tide_Special
        ORA_Message_Out "���ʃf�[�^��M", "������O���ʃf�[�^���������܂����B�V�����ʂɒ��O�̎����l�Ƃ̍����������āA�����E�\���l�Ƃ��܂��B", 1
    Else
        DH_Tide = 0#
    End If

'�\�����ʗՎ�
'    TidalY dw, C0, C1, C2, C3      '�C�ے����ʕ\����V�����ʂ���}����
    Cal_Tide dw, C0, C1, C2, C3    '60��������V�����ʂ��v�Z����
    If HO(1, Now_Step) < -50# Then
        HO(1, Now_Step) = C0 + DH_Tide
    Else
        DH_Tide = HO(1, Now_Step) - C0
    End If
    HO(1, Now_Step + 1) = C1 + DH_Tide
    HO(1, Now_Step + 2) = C2 + DH_Tide
    HO(1, Now_Step + 3) = C3 + DH_Tide

    If i = 0 Then
'        MsgBox "���[�J��DB�ɐ��ʃf�[�^������܂���B"
        LOG_Out "���[�J��DB�ɐ��ʃf�[�^������܂���B"
        ds = CDate("1900/01/01 01:00")
        de = CDate("1900/01/01 01:00")
        Exit Sub
    End If


'
'���ѐ��ʍŏI�f�[�^���t�̗\���f�[�^��
'���ɍs��
'

    Set ConR = Nothing

    jsd = du
    js(1) = Year(jsd)
    js(2) = Month(jsd)
    js(3) = Day(jsd)
    js(4) = Hour(jsd)
    js(5) = Minute(jsd)
    js(6) = 0
    jgd = dw
    jg(1) = Year(jgd)
    jg(2) = Month(jgd)
    jg(3) = Day(jgd)
    jg(4) = Hour(jgd)
    jg(5) = Minute(jgd)
    jg(6) = 0
    Now_Step = DateDiff("h", jsd, jgd) + 1
    All_Step = Now_Step + Yosoku_Step

    If Now_Step <= 4 Then
        'LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        '�C���J�n�@2016/09/23�@O.OKADA�@��������R�����g�A�E�g����B
        '�C�����R�@�v�Z���������15�����x�x��Ă��邽�߁B
        'Exit Sub
        '�C���I���@2016/09/23�@O.OKADA�@�����܂ŃR�����g�A�E�g����B
    End If
    If ds = de Or All_Step < 3 Then
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        '�C���J�n�@2016/09/23�@O.OKADA�@��������R�����g�A�E�g����B
        '�C�����R�@�v�Z���������15�����x�x��Ă��邽�߁B
        'Exit Sub
        '�C���I���@2016/09/23�@O.OKADA�@�����܂ŃR�����g�A�E�g����B
    End If

'�����[���E����������O���ʂ̕�U
    '�C���P********************************
    For i = 1 To Now_Step
        If HO(1, i) < -50# Then
            j = 1
            Select Case i
                Case 1
                    HO(1, 1) = 1.5
                Case Is > 1
                     HO(1, i) = HO(1, i - 1)
            End Select
        End If
    Next i

'������U
    er = False
    For i = 1 To 7
        If i <> 2 Then
            ch = False
'            For j = Now_Step - 3 To Now_Step�@2004/03/24
            For j = Now_Step To Now_Step
                a = HO(i, j)
                If a < -50# Then
                    ch = True
                    GoTo J1
                End If
            Next j
        End If
    Next i
J1:
    If ch Then
        Pre_���ʌ�����U
        For i = 1 To 7
            If i <> 2 Then
                uh = True
'                For j = Now_Step - 3 To Now_Step  2004/03/24
                For j = Now_Step To Now_Step
                    a = HO(i, j)
                    If a < -50# Then
                        uh = True
'                        For k = Now_Step - 3 To Now_Step  2004/03/24
                        For k = Now_Step To Now_Step
'                            m = k - (Now_Step - 3) + 1  2004/03/24
                            m = k - Now_Step + 1
                            hw(m) = HO(i, k)
                        Next k
                        If hw(m) < -50# Then
                            er = True
                            ORA_Message_Out "�e�����[�^���ʎ�M", Name_H(i) & "�́A���ʃf�[�^���������܂����B�^���\���V�X�e���ɂ�錋�ʂ�p���Đ��ʗ\���v�Z���s���܂��B", 1
                        End If
                        Exit For
                    End If
                Next j
            End If
        Next i
    End If
    If HO(1, Now_Step) < -50# Then
        er = True
    End If
    irc = True
    If (AutoDrive.Check6 = vbChecked) And er Then '������U������͂���
        Load Data_Edit
        Unload Data_Edit
    End If
    If (AutoDrive.Check6 = vbUnchecked) And er Then '�����Ȃ̂Ōv�Z���X�L�b�v����
'        irc = False '�����ł��v�Z����悤�ɏC�� 2004/4/26
'        Exit Sub
    End If

'    Dim nf As Long
'
'    nf = FreeFile
'    open app.Path & "\data\���ʃX���C�h��.dat" for output
'    LOG_Out "IN  Data_IN  ���ʃX���C�h�� CX=" & Format(cx, "###0.000")

'    MDB_�� jsd, jgd, er

End Sub
Sub Pre_���ʌ�����U()

    Dim ConR        As New ADODB.Recordset
    Dim SQL         As String
    Dim ds          As Date
    Dim de          As Date
    Dim i           As Long
    Dim j           As Long

    ds = DateAdd("h", -4, jgd)
    de = jgd

    SQL = "select * from ���� where Time between '" & Format(ds, "yyyy/mm/dd hh:nn") & "' and '" & _
           Format(de, "yyyy/mm/dd hh:nn") & "' order by Time"

    ConR.Open SQL, Con_����, adOpenKeyset, adLockReadOnly

    j = 1
    Do Until ConR.EOF
        For i = 1 To 6 '�U���ʊϑ���
            wH(i, j) = ConR.Fields(i + 1).Value
        Next i
        j = j + 1
        ConR.MoveNext
    Loop

    ConR.Close

End Sub
Sub Tide_Special()

    Dim SQL    As String
    Dim buf    As String
    Dim dw     As Date
    Dim w

    LOG_Out "IN   Tide_Special"

    On Error GoTo ER1

    DH_Tide = 0#

'    MDB_����_Connection

    dw = DateAdd("n", -10, jgd)

    If isRAIN = "02" Then
        SQL = "SELECT ������O���� FROM FRICS���� WHERE TIME='" & Format(dw, "yyyy/mm/dd hh:nn") & "'"
    Else
        SQL = "SELECT ������O���� FROM �C�ے����� WHERE TIME='" & Format(dw, "yyyy/mm/dd hh:nn") & "'"
    End If


    Rec_����.Open SQL, Con_����, adOpenDynamic, adLockReadOnly

    If Rec_����.EOF Then
        DH_Tide = 0#
    Else
        buf = Rec_����.Fields(0).Value
        w = Split(buf, ",")
        DH_Tide = w(0)
    End If

    Rec_����.Close

'    MDB_����_Close

    LOG_Out "OUT  Tide_Special DH_Tide=" & Format(DH_Tide, "###0.000")

    Exit Sub

ER1:
    LOG_Out "OUT  Tide_Special ABend DH_Tide=" & Format(DH_Tide, "###0.000")
    Rec_����.Close
    On Error GoTo 0

End Sub


