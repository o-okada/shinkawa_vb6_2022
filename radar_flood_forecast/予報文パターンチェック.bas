Attribute VB_Name = "�\��"
'******************************************************************************
'���W���[�����F�\�񕶃p�^�[���`�F�b�N
'
'******************************************************************************
Option Explicit
Option Base 1
Public BP              As Long                                          '�O��\�񕶃R�[�h
Public Patan           As Long                                          '����p�^�[��
Public Xnum            As String                                        '�`�F�b�N����p�^�[���ԍ�
Public rch             As Boolean
Public PRACTICE_FLG_CODE  As String                                     '"40"=�\��  "99"=���K
Public Const �͂񗔒��Ӑ��� = 3#
Public Const ���f���� = 4.4
Public Const �͂񗔊댯���� = 5.2

'******************************************************************************
'�T�u���[�`���FPattern_Check()
'�����T�v�F
'******************************************************************************
Sub Pattern_Check()
    rch = False
    Select Case BP
        Case 0
            Xnum = "1,5,10"
        Case 1
            Xnum = "2,4,5,6,7,10"
        Case 2
            Xnum = "4,5,6,7,10"
        Case 3
            Xnum = "4,5,6,7,10,11"
        Case 4
            Patan = 0
            BP = 0
            Wng_Last_Time = 0                                           '���ӕ��̃����N��������
            rch = True
            Xnum = "0"
        Case 5
            Xnum = "3,4,8,10"
        Case 6
            Xnum = "3,4,8,10"
        Case 7
            Xnum = "3,4,8,10"
        Case 8
            Xnum = "3,4,10,12"
        Case 9
            Xnum = "3,4,10,12"
        Case 10
            Xnum = "3,4,9,13"
        Case 11
            Xnum = "4,5,6,7,10"
        Case 12
            Xnum = "3,4,10"
        Case 13
            Xnum = "3,4,9"
    End Select
    ����_Check
    If rch Then                                                         '�ȉ��̕��͖{�Ԃł͊֌W�Ȃ��e�X�g���̂ݗL�� 2008/08/30 check
        BP = Patan
    End If
End Sub

'******************************************************************************
'�T�u���[�`���F�^���\�񕶏�����()
'�����T�v�F
'******************************************************************************
Sub �^���\�񕶏�����()
    Dim nf   As Integer
    Dim j    As Integer
    Dim buf  As String
    Dim a
    LOG_Out "IN  �^���\�񕶏�����"
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    'nf = FreeFile
    'Open App.Path & "\data\�\�񕶏o��.txt" For Input As #nf
    'Input #nf, buf
    'j = CInt(Mid(buf, 1, 5))
    'If j = 1 Then
    '    DBX_ora = True
    '    AutoDrive.Option1(0).Value = True
    'Else
    '    DBX_ora = False
    '    AutoDrive.Option1(1).Value = True
    'End If
    'Input #nf, buf '���ʃ^�C�g��
    'Input #nf, buf
    'a = Mid(buf, 1, 10)
    'If IsNumeric(a) Then
    '    �댯���� = CSng(a)
    'Else
    '    MsgBox "���͂����댯���ʂ͐��l�ł͂���܂���" & vbLf & _
    '           "�I���N���c�a�ɂ͏o�͂��Ȃ����[�h�Ōv�Z�܂��B" & vbLf & _
    '           "�v�Z�𒆎~���܂��B"
    '    End
    'End If
    'a = Mid(buf, 11, 10)
    'If IsNumeric(a) Then
    '    �x������ = CSng(a)
    'Else
    '    MsgBox "���͂����x�����ʂ͐��l�ł͂���܂���" & vbLf & _
    '           "�I���N���c�a�ɂ͏o�͂��Ȃ����[�h�Ōv�Z�܂��B" & vbLf & _
    '           "�v�Z�𒆎~���܂��B"
    '    End
    'End If
    'a = Mid(buf, 20, 10)
    'If IsNumeric(a) Then
    '    �w�萅�� = CSng(a)
    'Else
    '    MsgBox "���͂����w�萅�ʂ͐��l�ł͂���܂���" & vbLf & _
    '           "�I���N���c�a�ɂ͏o�͂��Ȃ����[�h�Ōv�Z�܂��B" & vbLf & _
    '           "�v�Z�𒆎~���܂��B"
    '    End
    'End If
    'Close #nf
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    PRACTICE_FLG_CODE = "40"                                            '�\�񕶖{����񃂁[�h�������l�Ƃ���
    AutoDrive.Option2(0).Value = True
    LOG_Out "OUT �^���\�񕶏�����"
End Sub

'******************************************************************************
'�T�u���[�`���F����_Check()
'�����T�v�F
'******************************************************************************
Sub ����_Check()
    Dim i      As Long
    Dim n      As Long
    Dim m      As Long
    Dim w
    Dim HOM2   As Single                                                '����2���ԑO����
    Dim HOM1   As Single                                                '����1���ԑO����
    Dim HON    As Single                                                '���� ����������
    Dim HC1    As Single                                                '�\��1���Ԍ㐅��
    Dim HC2    As Single                                                '�\��2���Ԍ㐅��
    Dim HC3    As Single                                                '�\��3���Ԍ㐅��
    '******************************************************
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '******************************************************
    'HOM2 = ����.hm2
    'HOM1 = ����.hm1
    'HON = ����.h
    'HC1 = ����.hy1
    'HC2 = ����.hy2
    'HC3 = ����.hy3
    '******************************************************
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    '******************************************************
    ����.hm2 = HO(5, Now_Step - 2)
    ����.hm1 = HO(5, Now_Step - 1)
    ����.H = HO(5, Now_Step)
    ����.hy1 = HQ(1, 41, NT - 12)
    ����.hy2 = HQ(1, 41, NT - 6)
    ����.hy3 = HQ(1, 41, NT)
    HOM2 = ����.hm2
    HOM1 = ����.hm1
    HON = ����.H
    HC1 = ����.hy1
    HC2 = ����.hy2
    HC3 = ����.hy3
    w = Split(Xnum, ",")
    n = UBound(w)
    For i = 0 To n
        m = w(i)                                                        '�p�^�[���ԍ�
        Select Case m
            Case 1
                If (HON < ���f����) And (HON >= �͂񗔒��Ӑ���) Then
                    If (HC3 < �͂񗔊댯����) Then
                        If (HC1 < �͂񗔊댯����) And (HC1 >= �͂񗔒��Ӑ���) Then
                            If (HC2 < �͂񗔊댯����) And (HC2 >= �͂񗔒��Ӑ���) Then
                                Patan = m
                                rch = True
                            End If
                        End If
                    End If
                End If
            Case 2
                If (�͂񗔊댯���� > HON) And (HON >= ���f����) Then
                    If (���f���� > HC3) And (HC3 >= �͂񗔒��Ӑ���) Then
                        Patan = m
                        rch = True
                    End If
                End If
            Case 3
                If (���f���� > HON) And (HON >= �͂񗔒��Ӑ���) Then
                    If (���f���� > HC3) Then
                        Patan = m
                        rch = True
                    End If
                End If
            Case 4
                If (HON < �͂񗔒��Ӑ���) Then
                    If (HC1 < �͂񗔒��Ӑ���) Then
                        If (HC2 < �͂񗔒��Ӑ���) Then
                            If (HC3 < �͂񗔒��Ӑ���) Then
                                Patan = m
                                rch = True
                            End If
                        End If
                    End If
                End If
            Case 5
                If (���f���� > HON) Then
                    If (�͂񗔊댯���� <= HC3) Then
                        Patan = m
                        rch = True
                    End If
                End If
            Case 6
                If (�͂񗔊댯���� > HON) And (HON >= ���f����) Then
                    If (�͂񗔊댯���� <= HC3) Then
                        Patan = m
                        rch = True
                    End If
                End If
            Case 7
                If (�͂񗔊댯���� > HON) And (HON >= ���f����) Then
                    If (�͂񗔊댯���� > HC3) And (HC3 >= ���f����) Then
                        Patan = m
                        rch = True
                    End If
                End If
            Case 8
                If (�͂񗔊댯���� > HON) And (HON >= ���f����) Then
                    If (�͂񗔊댯���� <= HC3) Then
                        Patan = m
                        rch = True
                    End If
                End If
            Case 9
                If (�͂񗔊댯���� > HON) And (HON >= ���f����) Then
                    If (�͂񗔊댯���� > HC3) Then
                        Patan = m
                        rch = True
                    End If
                End If
            Case 10
                If (�͂񗔊댯���� <= HON) And (�͂񗔊댯���� > HOM1) And (�͂񗔊댯���� > HOM2) Then
                    Patan = m
                    rch = True
                End If
            Case 11
                If (���f���� > HON) And (�͂񗔒��Ӑ��� <= HON) Then
                    If (���f���� > HOM1) And (�͂񗔒��Ӑ��� <= HOM1) Then
                        If (���f���� > HOM2) And (�͂񗔒��Ӑ��� <= HOM2) Then
                            If (���f���� > HC1) And (�͂񗔒��Ӑ��� <= HC1) Then
                                If (���f���� > HC2) And (�͂񗔒��Ӑ��� <= HC2) Then
                                    If (���f���� > HC3) And (�͂񗔒��Ӑ��� <= HC3) Then
                                        Patan = m
                                        rch = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Case 12
                If (�͂񗔊댯���� > HON) And (HON >= ���f����) Then
                    If (�͂񗔊댯���� > HOM1) And (HOM1 >= ���f����) Then
                        If (�͂񗔊댯���� > HOM2) And (HOM2 >= ���f����) Then
                            If (�͂񗔊댯���� > HC1) And (HC1 >= ���f����) Then
                                If (�͂񗔊댯���� > HC2) And (HC2 >= ���f����) Then
                                    If (�͂񗔊댯���� > HC3) And (HC3 >= ���f����) Then
                                        Patan = m
                                        rch = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Case 13
                If (�͂񗔊댯���� <= HON) And (�͂񗔊댯���� <= HOM1) And (�͂񗔊댯���� <= HOM2) Then
                    If (�͂񗔊댯���� <= HC1) And (�͂񗔊댯���� <= HC2) And (�͂񗔊댯���� <= HC3) Then
                        Patan = m
                        rch = True
                    End If
                End If
        End Select
    Next i
    Exit Sub
jump:
    Patan = m
End Sub

'******************************************************************************
'�T�u���[�`���F�\�񕶃`�F�b�N()
'�����T�v�F
'******************************************************************************
Sub �\�񕶃`�F�b�N()
    '******************************************************
    'Ver1.0.0 �C���J�n 2015/08/06 O.OKADA�y01-01�z
    '���I���N���f�[�^�x�[�X�̃e�[�u���u�v�̍폜�ɑΉ����A���L�̂Ƃ���C������B
    '******************************************************
    Exit Sub
    '******************************************************
    'Ver1.0.0 �C���I�� 2015/08/06 O.OKADA�y01-01�z
    '******************************************************
    �\�񕶗���DB_Read
    Pattern_Check
    If Patan > 0 Then
        '******************************************************
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        '******************************************************
        '�p�^�[�����W_Read                                              'AutoDrive��Form Load�œǂނ悤�ɂ���
        '******************************************************
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        '******************************************************
        Pattan_Now = Patan
        �啶�쐬2
        If DBX_ora Then
            ORA_YOHOUBUNAN
        End If
        �\�񕶗���DB_Write
    End If
    Patan = 0
End Sub
