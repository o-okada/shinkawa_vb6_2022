Attribute VB_Name = "�^���\��"
Option Explicit
Option Base 1

Public B1                 As String
Public B2                 As String
Public Y_FLAG             As Integer ' 0=�v�Z�J�n�� 1=�^�����ӕ� 2=�^���x�� 3=�^�����Ӊ�����
Public Kind_S             As String  '�啶���
Public Kind_N             As String  '�啶��ʃR�[�h
Public hx                 As Single  '(�x�����ʁ{�댯����)*0.5
Public SYUBN              As String
Public Course             As String
Public Wng_Last_Time      As Integer '�O�X�e�b�v�̒��ӕ��ԍ�


Public �댯����           As Single  '= 5.2
Public �x������           As Single  '= 3#
Public �w�萅��           As Single  '= 2#


Public Const �啶A = "�@�@�V�쐅���O���ʐ��ʊϑ����ł́A" & vbLf & _
                     "�@�@�͂񗔒��Ӑ��ʂ�啝�ɒ�����o���ƂȂ錩���݂ł��̂�" & vbLf & _
                     "�@�@�e�n�Ƃ����d�Ȍx�������ĉ������B"

Public Const �啶B = "�@�@�V�쐅���O���ʐ��ʊϑ����ł́A" & vbLf & _
                     "�@�@�͂񗔒��Ӑ��ʂ𒴂���o���ƂȂ錩���݂ł��̂Ŋe�n" & vbLf & _
                     "�@�@�Ƃ��\���Ȓ��ӂ����ĉ������B"

Public Const �啶C = "�@�@�V�쐅���O���ʐ��ʊϑ����ł́A" & vbLf & _
                     "�@�@�����̊Ԃ͂񗔒��Ӑ��ʈȏ�̐��ʂ����������݂ł��̂Ŋe�n�Ƃ�" & vbLf & _
                     "�@�@�\���Ȓ��ӂ����ĉ������B"

Public Const �啶D = "�@�@�V��^�����ӕ���^���x��ɐ؊����܂��B" & vbCrLf & _
                     "�@�@�V�쐅���O���ʐ��ʊϑ����ł́A" & vbLf & _
                     "�@�@�͂񗔊댯���ʂ𒴂���o���ƂȂ錩���݂ł��̂Ŋe�n�Ƃ����d��" & vbLf & _
                     "�@�@�x�������ĉ������B"

Public Const �啶E = "�@�@�V�쐅���O���ʐ��ʊϑ����ł́A" & vbLf & _
                     "�@�@�͂񗔊댯���ʂ𒴂���o���ƂȂ錩���݂ł��̂Ŋe�n�Ƃ����d��" & vbLf & _
                     "�@�@�x�������ĉ������B"

Public Const �啶F = "�@�@�V�쐅���O���ʐ��ʊϑ����ł́A" & vbLf & _
                     "�@�@�͂񗔊댯���ʂ�啝�ɒ�����o���ƂȂ錩���݂ł��̂Ŋe�n�Ƃ�" & vbLf & _
                     "�@�@���d�Ȍx�������ĉ������B"

Public Const �啶G = "�@�@�V�쐅���O���ʐ��ʊϑ����ł́A" & vbLf & _
                     "�@�@�����̊Ԃ͂񗔊댯���ʈȏ�̏o�������������݂ł��̂Ŋe�n�Ƃ�" & vbLf & _
                     "�@�@���d�Ȍx�������ĉ������B"

Public Const �啶H = "�@�@�V��^���x����^�����ӕ�ɐ؊����܂��B" & vbLf & _
                     "�@�@�V�쐅���O���ʐ��ʊϑ����ł́A" & vbLf & _
                     "�@�@�����̊Ԃ͂񗔒��Ӑ��ʈȏ�̐��ʂ����������ł��̂Ŋe�n�Ƃ�" & vbLf & _
                     "�@�@�\���Ȓ��ӂ����ĉ������B"

Public Const �啶I = "�@�@�V�쐅���O���ʐ��ʊϑ����ł́A" & vbLf & _
                     "�@�@�͂񗔒��Ӑ��ʂ������댯�͂Ȃ��Ȃ������̂Ǝv���܂��B"

Public Const �啶J = "�@�@�V�쐅���O���ʐ��ʊϑ����ł́A" & vbLf & _
                     "�@�@�����̊Ԃ͂񗔒��Ӑ��ʈȏ�̐��ʂ����������ł��̂Ŋe�n�Ƃ�" & vbLf & _
                     "�@�@�\���Ȓ��ӂ����ĉ������B"

Public Const CYUBN_1 = "�@�@����̏o���́A����3�N9���̑䕗17�E18���ɕC�G" & vbLf & _
                       "�@�@����K�͂ƌ����܂�܂��B"

Public Const CYUBN_2 = "�@�@����̏o���́A����3�N9���̑䕗17�E18�������" & vbLf & _
                       "�@�@��K�͂ƌ����܂�܂��B"

Public Const CYUBN_3 = "�@�@����̏o���́A����12�N9���̓��C���J�ɕC�G����" & vbLf & _
                       "�@�@�K�͂ƌ����܂�܂��B"



Sub H2Z(strH As String, strZ As String)

    Dim ZN(0 To 9)    As String
    Dim i        As Long
    Dim j        As Long
    Dim L        As Long
    Dim w
    Dim ww

    ZN(0) = "�O": ZN(1) = "�P": ZN(2) = "�Q": ZN(3) = "�R": ZN(4) = "�S"
    ZN(5) = "�T": ZN(6) = "�U": ZN(7) = "�V": ZN(8) = "�W": ZN(9) = "�X"

    strZ = ""
    L = Len(strH)
    For i = 1 To L
        w = Mid(strH, i, 1)
        Select Case w
         Case "0" To "9" 'IsNumeric(w)
            j = CInt(w)
            ww = ZN(j)
        Case "."
            ww = "�D"
        Case " "
            ww = "�@"
        Case Else
            ww = w
        End Select
        strZ = strZ & ww
    Next i

End Sub
Function Raise(S As Single) As Single

    Dim c As Long
    Dim d As Double

    d = S
    c = Fix(d * 10.00001)
    Raise = c + 1#
    Raise = Raise / 10#

End Function
Sub ST1(H2 As Single)

    If H2 >= hx Then
        SYUBN = �啶A
        Course = "1"
    Else
        SYUBN = �啶B
        Course = "2"
    End If
    Y_FLAG = 1
    Kind_S = "�^�����ӕ񔭕\"
    Kind_N = "10"

End Sub
Sub ST2()

    SYUBN = �啶I
    Course = Course & "O"
    Kind_S = "�^�����ӕ����"
    Kind_N = "30"
    Y_FLAG = 0

End Sub
Sub ST3(H2 As Single, hm1 As Single, hm2 As Single, c1 As Integer)

    If �x������ <= hm1 And �x������ <= hm2 Then
        SYUBN = �啶C
        Y_FLAG = 2
        Course = Course & "5"
        Kind_S = "�^����񔭕\"
        Kind_N = "30"
    Else               '�B
        If c1 = 1 Then
            Course = Course & "�"
            ST1 H2
        End If
    End If

End Sub
Sub ST4(H2 As Single, H3 As Single, hm1 As Single, hm2 As Single, c1 As Integer)

    If �댯���� <= H3 Then
        SYUBN = �啶D
        Y_FLAG = 3
        Course = Course & "6"
        Kind_S = "�^���x�񔭕\"
        Kind_N = "20"
    Else
        ST3 H2, hm1, hm2, c1
    End If

End Sub
Sub ST5(h0 As Single, H1 As Single, H2 As Single, H3 As Single)

    Kind_S = "�^����񔭕\"
    Kind_N = "30"
    If �댯���� <= H1 Or �댯���� <= H2 Or �댯���� <= H3 Then '�D
        Course = "9"
        If h0 <= �댯���� And H3 > �댯���� Then
            SYUBN = �啶E   '�E
            Y_FLAG = 4
            Course = Course & "A"
            Exit Sub
        End If
        If �댯���� < h0 And �댯���� < H3 Then
            SYUBN = �啶F   '�F
            Y_FLAG = 4
            Course = Course & "B"
            Exit Sub
        End If
        If �댯���� < h0 And H3 < �댯���� Then
            SYUBN = �啶G   '�G
            Y_FLAG = 4
            Course = Course & "C"
            Exit Sub
        End If
        If h0 < �댯���� And H3 < �댯���� Then
            SYUBN = �啶G   '�G
            Y_FLAG = 4
            Course = Course & "Ca"
            Exit Sub
        End If
    End If

    If �댯���� <= h0 Then
        SYUBN = �啶G   '�G
        Y_FLAG = 4
        Course = Course & "Cb"
        Exit Sub
    Else
        ST6
    End If

End Sub
Sub ST6()

    SYUBN = �啶H
    Y_FLAG = 5
    Course = Course & "G"
    Kind_S = "�^�����ӕ񔭕\"
    Kind_N = "10"
    
End Sub
Sub ST7()

    SYUBN = �啶J
    Y_FLAG = 6
    Course = Course & "L"
    Kind_S = "�^�����ӕ񔭕\"
    Kind_N = "10"

End Sub
Sub ST8()

    SYUBN = �啶I
    Y_FLAG = 7
    Course = Course & "La"
    Kind_S = "�^�����ӕ�������\"
    Kind_N = "01"

End Sub
'**************************************************
'�����O���ʂ𔻒肵�^���\�񕶂��쐬����
'
'
'
'
'
'
'
'**************************************************
Sub �^���\�񕶈č쐬()

    Dim i           As Long
    Dim j           As Long
    Dim hm2         As Single   '���ѐ���
    Dim hm1         As Single   '���ѐ���
    Dim h0          As Single   '���ѐ���
    Dim H1          As Single   '1���Ԍ�\������
    Dim H2          As Single   '2���Ԍ�\������
    Dim H3          As Single   '3���Ԍ�\������
    Dim HM          As Single
    Dim H2r         As Single   '2���Ԍ�\���؂�グ����
    Dim H3r         As Single   '3���Ԍ�\���؂�グ����
    Dim HV          As Single
    Dim c1          As Integer
    Dim CYUBN       As String
    Dim CYUBN1      As String
    Dim Wng         As Integer
    Dim nf          As Integer
    Dim buf         As String
    Dim Kind(6, 2)  As String  '��ʃR�[�h�ƕ���
    Dim Bun1        As String
    Dim Bun2        As String
    Dim ���ʏ�     As String
    Dim jsx         As Date
    Dim Bunw        As String
    Dim w           As Single
    Dim m1          As String
    Dim mw          As String
    Dim irc         As Boolean
    Dim Kind_M      As String

    LOG_Out "IN    �^���\�񕶈č쐬"

    Const LF = vbLf

'    Kind(1, 1) = "10": Kind(1, 2) = "�^�����ӕ񔭕\"
'    Kind(2, 1) = "11": Kind(2, 2) = "�^�����ӏ�񔭕\�i�؊��j"
'    Kind(3, 1) = "20": Kind(3, 2) = "�^���x�񔭕\"
'    Kind(4, 1) = "21": Kind(4, 2) = "�^���x�񔭕\�i�؊��j"
'    Kind(5, 1) = "30": Kind(5, 2) = "�^����񔭕\"
'    Kind(6, 1) = "01": Kind(6, 2) = "�^�����ӕ����"

    SYUBN = ""
    Kind_M = ""
    Kind_S = ""
    Kind_N = ""
    Course = ""
    CYUBN = ""
    CYUBN1 = ""

    �\������DB_Read

    hx = (�댯���� + �x������) * 0.5
    hm2 = HO(5, Now_Step - 2)
    hm1 = HO(5, Now_Step - 1)
    h0 = HO(5, Now_Step)
    H1 = HQ(1, 41, NT - 12)
    H2 = HQ(1, 41, NT - 6)
    H3 = HQ(1, 41, NT)
    HM = H1
    If H2 > HM Then HM = H2
    If H3 > HM Then HM = H3
    H2r = Raise(HQ(1, 41, NT - 6))
    H3r = Raise(HQ(1, 41, NT))

    c1 = Y_FLAG

    Select Case Y_FLAG

        Case 0
            If H2 < �x������ Then
                Exit Sub
            Else
                ST1 H2
            End If

        Case 1, 2                             '�A
            If H1 < �x������ And H2 < �x������ And H3 < �x������ Then
                If h0 < �x������ Then
                    ST8
                    Course = "�"
                Else
                    Course = "3"
                End If
            Else
                If H1 < �댯���� And H2 < �댯���� And H3 < �댯���� Then
                    ST3 H2, hm1, hm2, c1
                    Course = Course & "4"
                Else
                    ST4 H2, H3, hm1, hm2, c1
                End If
            End If

        Case 3, 4
            ST5 h0, H1, H2, H3

        Case 5, 6
            If h0 >= �댯���� Then '�I
                Course = Course & "H"
                ST4 H2, H3, hm1, hm2, c1
                GoTo J1
            End If
            If H1 >= �댯���� Or H2 >= �댯���� Or H3 >= �댯���� Then '�J
                Course = Course & "I"
                ST4 H2, H3, hm1, hm2, c1
                GoTo J1
            End If
            If H1 < �x������ And H2 < �x������ Or H3 < �x������ Then '�K
                If h0 < �x������ Then
                    ST8
                    GoTo J1
                Else
                    Course = Course & "K"
                    GoTo J1
                End If
            Else
                If �x������ <= h0 Then
                    If �x������ <= hm1 And �x������ <= hm2 Then
                        ST7
                        GoTo J1
                    End If
                Else
                    If c1 = 5 Then  '�L
                        ST6
                        GoTo J1
                    Else
                        Course = Course & "N"
                        GoTo J1
                    End If
                End If
            End If

        Case 7
            ST8
            GoTo J1

    End Select

J1:

'���ӕ�
    If Y_FLAG = 3 Or Y_FLAG = 4 Then
        HV = H3 - H2
        If Y_FLAG >= 2 Then
            Select Case HV
                Case Is < 0.5
                    Wng = 1
                Case Is < 1#
                    Wng = 2
                Case Is >= 1#
                    Wng = 3
            End Select
'           ���ӎ������ʔԍ��ۑ�
            nf = FreeFile
            Open App.Path & "\Data\���ӎ���.dat" For Output As #nf
            Print #nf, Format(jgd, "yyyy/mm/dd hh;nn")
            Print #nf, Wng
            Close #nf
'        Else
'            nf = FreeFile
'            Open App.Path & "\Data\���ӎ���.dat" For Input As #nf
'            Line Input #nf, buf
'            If IsDate(buf) Then
'                j = DateDiff("h", CDate(buf), jgd) + 1
'                Input #nf, Wng
'                Select Case Wng
'                    Case 3
'                        If j > 2 Then
'                            CYUBN = "�@�@����̏o���́A����3�N9���̑䕗17�E18��������K�͂ƌ����܂�܂��B"
'                        Else
'                            CYUBN = "�@�@����̏o���́A����3�N9���̑䕗17�E18���ɕC�G����K�͂ƌ����܂�܂��B"
'                        End If
'                    Case 2
'                        If j > 6 Then
'                            CYUBN = "�@�@����̏o���́A����12�N9���̓��C���J�ɕC�G����K�͂ƌ����܂�܂��B"
'                        Else
'                            CYUBN = "�@�@����̏o���́A����3�N9���̑䕗17�E18��������K�͂ƌ����܂�܂��B"
'                        End If
'                    Case 1
'                        CYUBN = "�@�@����̏o���́A����12�N9���̓��C���J�ɕC�G����K�͂ƌ����܂�܂��B"
'                End Select
'            End If
'            Close #nf
        End If
        If Wng_Last_Time > Wng Then Wng = Wng_Last_Time
        Wng_Last_Time = Wng
        Select Case Wng
            Case 1
                CYUBN = CYUBN_1
            Case 2
                CYUBN = CYUBN_2
            Case 3
                CYUBN = CYUBN_3
        End Select
        If h0 >= 6.2 Or H3 >= 6.2 Then '�v���h��(T.P 6.2m)�𒴂���
            CYUBN1 = "�@�@�܂��A�z���̋��ꂪ����܂��̂Ō��d�Ȍx�����K�v�ł��B"
        End If
    End If

'�^���󋵔��\��
    Select Case Y_FLAG
        Case 1, 2, 5, 6
           Kind_M = "�^�����ӕ񔭕\��"
        Case 3, 4
           Kind_M = "�^���x�񔭕\��"
        Case 7
           Kind_M = " "
   End Select



'���ʏ�
    w = h0 - hm2
    If w <= -0.1 Then ���ʏ� = "���~��"
    If -0.1 < w And w <= 0.1 Then ���ʏ� = "���΂�"
    If 0.1 < w And w <= 0.3 Then ���ʏ� = "�㏸��"
    If 0.3 < w Then ���ʏ� = "�}�㏸��"

    Print #Log_Repo, ""
    Print #Log_Repo, Format(jgd, "yyyy/mm/dd hh:nn") & "  " & Kind_S
    Print #Log_Repo, SYUBN
    Print #Log_Repo, "�������O�Q���Ԑ��� " & Format(Format(hm2, "##0.00"), "@@@@@@@") & " " & IIf((hm2 - �x������) < 0#, "<", ">=") & " �x������  " & IIf((hm2 - �댯����) < 0#, "<", ">=") & " �댯����"
    Print #Log_Repo, "�������O�P���Ԑ��� " & Format(Format(hm1, "##0.00"), "@@@@@@@") & " " & IIf((hm1 - �x������) < 0#, "<", ">=") & " �x������  " & IIf((hm1 - �댯����) < 0#, "<", ">=") & " �댯����"
    Print #Log_Repo, "���������� �@�@�@�@" & Format(Format(h0, "##0.00"), "@@@@@@@") & " " & IIf((h0 - �x������) < 0#, "<", ">=") & " �x������  " & IIf((h0 - �댯����) < 0#, "<", ">=") & " �댯����"
    Print #Log_Repo, "�������{�P���Ԑ��� " & Format(Format(H1, "##0.00"), "@@@@@@@") & " " & IIf((H1 - �x������) < 0#, "<", ">=") & " �x������  " & IIf((H1 - �댯����) < 0#, "<", ">=") & " �댯����"
    Print #Log_Repo, "�������{�Q���Ԑ��� " & Format(Format(H2, "##0.00"), "@@@@@@@") & " " & IIf((H2 - �x������) < 0#, "<", ">=") & " �x������  " & IIf((H2 - �댯����) < 0#, "<", ">=") & " �댯����"
    Print #Log_Repo, "�������{�R���Ԑ��� " & Format(Format(H3, "##0.00"), "@@@@@@@") & " " & IIf((H3 - �x������) < 0#, "<", ">=") & " �x������  " & IIf((H3 - �댯����) < 0#, "<", ">=") & " �댯����"
    Print #Log_Repo, "�\���ő吅��       " & Format(Format(HM, "##0.00"), "@@@@@@@") & " " & IIf((HM - �x������) < 0#, "<", ">=") & " �x������  " & IIf((HM - �댯����) < 0#, "<", ">=") & " �댯����"
    Print #Log_Repo, "�^������=" & Y_FLAG
    Print #Log_Repo, "Course=" & Course

    If SYUBN = "" Then
        Exit Sub
    End If

    Bun1 = "�啶" & LF & SYUBN & LF
    If CYUBN <> "" Then
        Bun1 = Bun1 & "���ӁE�x������" & LF & CYUBN & LF
        If CYUBN1 <> "" Then
            Bun1 = Bun1 & CYUBN1 & LF
        End If
    End If
    Bun1 = Bun1 & " " & LF
    Bun1 = Bun1 & "�����E�\�z" & LF
    
    Bun2 = ""
    If Y_FLAG <> 1 Then
        jsx = DateAdd("h", 3, jgd)
    Else
        jsx = DateAdd("h", 2, jgd)
    End If
    buf = "�@�@�@�@"
    m1 = Format(Day(jgd), "##") & "��" & _
         Format(Hour(jgd), "#0") & "��" & _
         Format(Minute(jgd), "#0") & "��"
'    H2Z M1, Mw
    mw = m1
    buf = "�@�@�V��̐��ʂ�" & mw & "���݁A���̂Ƃ���ƂȂ��Ă��܂��B" & LF
    buf = buf & "�����O���ʐ��ʊϑ����m���{�s�V�쒬�厚�����n���n��" & LF
    m1 = Format(Format(h0, "##0.00"), "@@@@@@")
'    H2Z M1, Mw
    mw = m1
    buf = buf & "�@�@�@�@�@�@" & mw & "���[�g���i" & ���ʏ� & "�j" & LF
    If Y_FLAG <> 7 Then
        m1 = Format(Day(jsx), "##") & "��" & _
             Format(Hour(jsx), "#0") & "��" & _
             Format(Minute(jsx), "#0") & "��"
'        H2Z M1, Mw
        mw = m1
        buf = buf & "�@�@�V��̐��ʂ�" & mw & "���ɂ́A���̂悤�Ɍ����܂�܂��B" & LF
        buf = buf & "�@�@�����O���ʐ��ʊϑ����m���{�s�V�쒬�厚�����n���n��" & LF
        If Y_FLAG <> 1 Then
            m1 = Format(Format(H3r, "###0.00"), "@@@@@@")
        Else
            m1 = Format(Format(H2r, "###0.00"), "@@@@@@")
        End If
'        H2Z M1, Mw
        mw = m1
        buf = buf & "�@�@�@�@�@�@" & mw & "���[�g�����x" & LF & " " & LF
    Else
        buf = buf & "�@�@�@�@�@�@" & LF
        buf = buf & "�@�@�@�@�@�@" & LF
        buf = buf & "�@�@�@�@�@�@" & LF
        buf = buf & "�@�@�@�@�@�@" & LF

    End If
'    H2Z buf, Bunw
    Bunw = buf
    Bun2 = Bun2 & Bunw

'Bunw������  ' 2007/05/08 14:38 �x�m�ʓ��������d�b�������̂Ŗ��m�F�ŏC��
'    Bunw = "�@�@�y�Q�l�z" & LF & _
'           "�@�@�����O���ʐ��ʊϑ����m�V�쒬�厚�����n���n" & LF & _
'           "�@�@��h�� 6.20m  �댯���� 5.20m  �x������ 3.00m  �w�萅�� 2.00m" & LF
    Bunw = "�@�@�y�Q�l�z" & LF & _
           "�@�@�����O���ʐ��ʊϑ����m���{�s�V�쒬�厚�����n���n" & LF & _
           "�@�@��h���@�@�@�@�@�@�@�@�@�@ 6.20m            �͂񗔊댯���ʁi�댯���ʁj   5.20m" & LF & _
           "�@�@�͂񗔒��Ӑ��ʁi�x�����ʁj 3.20m            ���h�c�ҋ@���ʁi�w�萅�ʁj   2.00m" & LF
    Bun2 = Bun2 & Bunw & " " & LF

    Bunw = "�@�@�y�V��̍^���\�񔭕\�󋵁z" & LF
    Bunw = Bunw & "�@�@�@�@�@" & Kind_M & LF

    Bun2 = Bun2 & Bunw & " " & LF

    Bunw = "�@�@�₢���킹��" & LF & _
           "�@�@���ʊ֌W  �@���m���������ݎ������@�@�ێ��Ǘ��ہ@�s�d�k052(961)4421" & LF & _
           "�@�@�C�ۊ֌W�@  �C�ے����É��n���C�ۑ�@�ϑ��\��ہ@�s�d�k052(751)0909" & LF & " "

    Bun2 = Bun2 & Bunw

    Print #Log_Repo, Bun1
    Print #Log_Repo, Bun2
    B1 = Bun1
    B2 = Bun2

    If DBX_ora Then   '�\�񕶏o�͂��w������Ă�����
       ORA_YOHOUBUNAN irc
    End If

    If Y_FLAG = 7 And c1 = 7 Then
        Y_FLAG = 0
    End If

    �\������DB_Write

    LOG_Out "IN    �^���\�񕶈č쐬"

End Sub


