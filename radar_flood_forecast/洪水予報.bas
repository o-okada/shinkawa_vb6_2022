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

Public PRACTICE_FLG_CODE  As String  '"40"=�\��  "99"=���K

Public �댯����           As Single  '= 5.2
Public �x������           As Single  '= 3#
Public �w�萅��           As Single  '= 2#

Public Con_�\��         As New ADODB.Connection
Public Rst_�\��         As New ADODB.Recordset
Public DB_�\��          As Boolean

Public Const �啶A = "�@�@�V�쐅���O���ʐ��ʊϑ����ł́A" & vbLf & _
                     "�@�@�x�����ʂ�啝�ɒ�����o���ƂȂ錩���݂ł��̂�" & vbLf & _
                     "�@�@�e�n�Ƃ����d�Ȍx�������ĉ������B"

Public Const �啶B = "�@�@�V�쐅���O���ʐ��ʊϑ����ł́A" & vbLf & _
                     "�@�@�x�����ʂ𒴂���o���ƂȂ錩���݂ł��̂Ŋe�n" & vbLf & _
                     "�@�@�Ƃ��\���Ȓ��ӂ����ĉ������B"

Public Const �啶C = "�@�@�V�쐅���O���ʐ��ʊϑ����ł́A" & vbLf & _
                     "�@�@�����̊Ԍx�����ʈȏ�̐��ʂ����������݂ł��̂Ŋe�n�Ƃ�" & vbLf & _
                     "�@�@�\���Ȓ��ӂ����ĉ������B"

Public Const �啶D = "�@�@�V��^�����ӕ���^���x��ɐ؊����܂��B" & vbCrLf & _
                     "�@�@�V�쐅���O���ʐ��ʊϑ����ł́A" & vbLf & _
                     "�@�@�댯���ʂ𒴂���o���ƂȂ錩���݂ł��̂Ŋe�n�Ƃ����d��" & vbLf & _
                     "�@�@�x�������ĉ������B"

Public Const �啶E = "�@�@�V�쐅���O���ʐ��ʊϑ����ł́A" & vbLf & _
                     "�@�@�댯���ʂ𒴂���o���ƂȂ錩���݂ł��̂Ŋe�n�Ƃ����d��" & vbLf & _
                     "�@�@�x�������ĉ������B"

Public Const �啶F = "�@�@�V�쐅���O���ʐ��ʊϑ����ł́A" & vbLf & _
                     "�@�@�댯���ʂ�啝�ɒ�����o���ƂȂ錩���݂ł��̂Ŋe�n�Ƃ�" & vbLf & _
                     "�@�@���d�Ȍx�������ĉ������B"

Public Const �啶G = "�@�@�V�쐅���O���ʐ��ʊϑ����ł́A" & vbLf & _
                     "�@�@�����̊Ԋ댯���ʈȏ�̏o�������������݂ł��̂Ŋe�n�Ƃ�" & vbLf & _
                     "�@�@���d�Ȍx�������ĉ������B"

Public Const �啶H = "�@�@�V��^���x����^�����ӕ�ɐ؊����܂��B" & vbLf & _
                     "�@�@�V�쐅���O���ʐ��ʊϑ����ł́A" & vbLf & _
                     "�@�@�����̊Ԍx�����ʈȏ�̐��ʂ����������ł��̂Ŋe�n�Ƃ�" & vbLf & _
                     "�@�@�\���Ȓ��ӂ����ĉ������B"

Public Const �啶I = "�@�@�V�쐅���O���ʐ��ʊϑ����ł́A" & vbLf & _
                     "�@�@�x�����ʂ������댯�͂Ȃ��Ȃ������̂Ǝv���܂��B"

Public Const �啶J = "�@�@�V�쐅���O���ʐ��ʊϑ����ł́A" & vbLf & _
                     "�@�@�����̊Ԍx�����ʈȏ�̐��ʂ����������ł��̂Ŋe�n�Ƃ�" & vbLf & _
                     "�@�@�\���Ȓ��ӂ����ĉ������B"

Public Const CYUBN_1 = "�@�@����̏o���́A����3�N9���̑䕗17�E18���ɕC�G" & vbLf & _
                       "�@�@����K�͂ƌ����܂�܂��B"

Public Const CYUBN_2 = "�@�@����̏o���́A����3�N9���̑䕗17�E18�������" & vbLf & _
                       "�@�@��K�͂ƌ����܂�܂��B"

Public Const CYUBN_3 = "�@�@����̏o���́A����12�N9���̓��C���J�ɕC�G����" & vbLf & _
                       "�@�@�K�͂ƌ����܂�܂��B"



Public Log_Repo           As Integer   '���|�t�@�C���ɏ����o���t�@�C���ԍ�
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
'******************************************************************
'
'
'
'���m���T�[�o�[�ɗ\�񕶂���������
'
'
'
'
'
'
'
'
'
'******************************************************************
Sub ORA_YOHOUBUNAN(Return_Code As Boolean)

    Dim sql_SELECT   As String
    Dim sql_WHERE    As String
    Dim SQL          As String
    Dim N_rec        As Long
    Dim n            As Integer
    Dim i            As Long
    Dim SDATE        As String
    Dim Edate        As String
    Dim jssd         As Date
    Dim NTim         As Date
    Dim dw           As Date
    Dim Timew        As String
    Dim c1           As String
    Dim c2           As String
    Dim c3           As String
    Dim c4           As String
    Dim c5           As String

    LOG_Out "IN    ORA_YOHOUBUNAN"

    NTim = Now
    dw = DateAdd("n", 30, jgd)
    c1 = Format(NTim, "yyyy/mm/dd hh:nn") 'DB�������ݎ���
    c2 = ""
    c3 = ""
    c4 = Format(jgd, "yyyy/mm/dd hh:nn")  '�����f�[�^�̌����� ESTIMATE_TIME
    c5 = Format(dw, "yyyy/mm/dd hh:nn")   '���\���� �f�[�^�̌�����+30�� ANNOUNCE_TIME

    jssd = jgd

    SDATE = "'" & Format(jssd, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"

'SELECT
    sql_SELECT = "SELECT * FROM oracle.YOHOUBUNAN"

'WHERE
    sql_WHERE = " WHERE  ESTIMATE_TIME = TO_DATE(" & SDATE & ") AND" & _
                " DATA_KIND_CODE = '�t�P���R�E�Y�C�A��01' AND" & _
                " SENDING_STATION_CODE ='23001' AND" & _
                " RAIN_KIND = '" & isRAIN & "'"

    SQL = sql_SELECT & sql_WHERE

'    SQL = sql_SELECT
'
'------------ �t�B�[���h�����擾���� -----------------
'    Dim Tw
'    n = RST_YB.Fields.Count
'    For i = 0 To n - 1
'        Tw = RST_YB.Fields(i).Name
'        Debug.Print " Number=" & Format(Str(i), "@@@") & " �t�B�[���h��="; Tw
'    Next i
'---------------------------------------------------

    ' SQL�X�e�[�g�����g���w�肵�ă_�C�i�Z�b�g���擾����
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)

    Dim nf As Integer
    Dim buf As String

    nf = FreeFile
    Open App.Path & "\Data\DB_YB.DAT" For Output As #nf
    If dynOra.EOF Then
        dynOra.AddNew
    Else
        dynOra.Edit
    End If
    dynOra.Fields("WRITE_TIME").Value = c1                      '�������ݎ���
    dynOra.Fields("DATA_KIND_CODE").Value = "�t�P���R�E�Y�C�A��01"
    dynOra.Fields("DATA_KIND").Value = "�\�񕶈āi���ʕ����j"
    dynOra.Fields("SENDING_STATION_CODE").Value = "23001"
    dynOra.Fields("SENDING_STATION").Value = "���m���������ݎ�����"
    dynOra.Fields("APPOINTED_CODE").Value = ""
    dynOra.Fields("ESTIMATE_TIME").Value = c4
    dynOra.Fields("PRACTICE_FLG_CODE").Value = PRACTICE_FLG_CODE  '"40"=�\��  "99"=���K
    If PRACTICE_FLG_CODE = "40" Then
        dynOra.Fields("PRACTICE_FLG").Value = "�\��"
    Else
        dynOra.Fields("PRACTICE_FLG").Value = "���K"
    End If
    dynOra.Fields("SEQ_NO").Value = ""
    dynOra.Fields("ANNOUNCE_TIME").Value = c5
    dynOra.Fields("RIVER_NAME").Value = "���m�������쐅�n�@�V��"
    dynOra.Fields("RIVER_NO_CODE").Value = "85053002"
    dynOra.Fields("RIVER_NO").Value = "�V��"
    dynOra.Fields("RIVER_DIV_CODE").Value = "00"
    dynOra.Fields("RIVER_DIV").Value = ""
    dynOra.Fields("ANNOUNCE_NO").Value = ""
    dynOra.Fields("FORECAST_KIND").Value = Kind_S
    dynOra.Fields("FORECAST_KIND_CODE").Value = Kind_N
    dynOra.Fields("BUNSHO1").Value = B1
    dynOra.Fields("BUNSHO2").Value = B2
    dynOra.Fields("BUNSHO3").Value = ""
    dynOra.Fields("RAIN_KIND").Value = isRAIN '01=�C�ے�  02=FRICS

    dynOra.Update
    dynOra.Close

'�\�񕶑Ώۉ͐�

'SELECT
    sql_SELECT = "SELECT * FROM oracle.YOHOU_TARGET_RIVER"
'WHERE

    sql_WHERE = " WHERE  ESTIMATE_TIME = TO_DATE(" & SDATE & ") AND" & _
                " BUNAN_CODE = '01' AND" & _
                " DATA_KIND_CODE = '�t�P���R�E�Y�C�A��01' AND" & _
                " SENDING_STATION_CODE ='23001' AND" & _
                " TRIVER_NO_CODE = '85053002' AND" & _
                " RAIN_KIND = '" & isRAIN & "'"

    SQL = sql_SELECT & sql_WHERE

    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)

    If dynOra.EOF Then
        dynOra.AddNew
    Else
        dynOra.Edit
    End If

    dynOra.Fields("WRITE_TIME").Value = c1                      '�������ݎ���
    dynOra.Fields("BUNAN_CODE").Value = "01"
    dynOra.Fields("DATA_KIND_CODE").Value = "�t�P���R�E�Y�C�A��01"
    dynOra.Fields("SENDING_STATION_CODE").Value = "23001"
    dynOra.Fields("ESTIMATE_TIME").Value = c4
    dynOra.Fields("TRIVER_NAME").Value = "�V��"
    dynOra.Fields("TRIVER_NO_CODE").Value = "85053002"
    dynOra.Fields("TRIVER_NO").Value = "�V��"
    dynOra.Fields("TRIVER_DIV_CODE").Value = "00"
    dynOra.Fields("FORECAST_KIND").Value = Kind_S              'c2
    dynOra.Fields("FORECAST_KIND_CODE").Value = Kind_N         'c3
    dynOra.Fields("RAIN_KIND").Value = isRAIN         '02=FRICS   01=�C�ے�
    dynOra.Fields("OUT_NO").Value = 1

    dynOra.Update
    dynOra.Close

    DoEvents
    Close #nf
    Set dynOra = Nothing

    LOG_Out "OUT   ORA_YOHOUBUNAN"

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
    Dim a


    LOG_Out "IN    RAIN_SELECT_READ"

    a = Dir(����MDB)

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
Sub ST3(H2 As Single, Hm1 As Single, Hm2 As Single, c1 As Integer)

    If �x������ <= Hm1 And �x������ <= Hm2 Then
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
Sub ST4(H2 As Single, H3 As Single, Hm1 As Single, Hm2 As Single, c1 As Integer)

    If �댯���� <= H3 Then
        SYUBN = �啶D
        Y_FLAG = 3
        Course = Course & "6"
        Kind_S = "�^���x�񔭕\"
        Kind_N = "20"
    Else
        ST3 H2, Hm1, Hm2, c1
    End If

End Sub
Sub ST5(H0 As Single, H1 As Single, H2 As Single, H3 As Single)

    Kind_S = "�^����񔭕\"
    Kind_N = "30"
    If �댯���� <= H1 Or �댯���� <= H2 Or �댯���� <= H3 Then '�D
        Course = "9"
        If H0 <= �댯���� And H3 > �댯���� Then
            SYUBN = �啶E   '�E
            Y_FLAG = 4
            Course = Course & "A"
            Exit Sub
        End If
        If �댯���� < H0 And �댯���� < H3 Then
            SYUBN = �啶F   '�F
            Y_FLAG = 4
            Course = Course & "B"
            Exit Sub
        End If
        If �댯���� < H0 And H3 < �댯���� Then
            SYUBN = �啶G   '�G
            Y_FLAG = 4
            Course = Course & "C"
            Exit Sub
        End If
        If H0 < �댯���� And H3 < �댯���� Then
            SYUBN = �啶G   '�G
            Y_FLAG = 4
            Course = Course & "Ca"
            Exit Sub
        End If
    End If

    If �댯���� <= H0 Then
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
Sub �^���\�񕶏�����()

    Dim nf   As Integer
    Dim j    As Integer
    Dim buf  As String
    Dim a

    LOG_Out "IN  �^���\�񕶏�����"

    nf = FreeFile
    Open App.Path & "\data\�\�񕶏o��.txt" For Input As #nf
    Input #nf, buf
    j = CInt(Mid(buf, 1, 5))
    If j = 1 Then
        DBX_ora = True
        AutoDrive.Option1(0).Value = True
    Else
        DBX_ora = False
        AutoDrive.Option1(1).Value = True
    End If

    Input #nf, buf '���ʃ^�C�g��
    Input #nf, buf
    a = Mid(buf, 1, 10)
    If IsNumeric(a) Then
        �댯���� = CSng(a)
    Else
        MsgBox "���͂����댯���ʂ͐��l�ł͂���܂���" & vbLf & _
               "�I���N���c�a�ɂ͏o�͂��Ȃ����[�h�Ōv�Z�܂��B" & vbLf & _
               "�v�Z�𒆎~���܂��B"
        End
    End If
    a = Mid(buf, 11, 10)
    If IsNumeric(a) Then
        �x������ = CSng(a)
    Else
        MsgBox "���͂����x�����ʂ͐��l�ł͂���܂���" & vbLf & _
               "�I���N���c�a�ɂ͏o�͂��Ȃ����[�h�Ōv�Z�܂��B" & vbLf & _
               "�v�Z�𒆎~���܂��B"
        End
    End If
    a = Mid(buf, 20, 10)
    If IsNumeric(a) Then
        �w�萅�� = CSng(a)
    Else
        MsgBox "���͂����w�萅�ʂ͐��l�ł͂���܂���" & vbLf & _
               "�I���N���c�a�ɂ͏o�͂��Ȃ����[�h�Ōv�Z�܂��B" & vbLf & _
               "�v�Z�𒆎~���܂��B"
        End
    End If

    Close #nf

    PRACTICE_FLG_CODE = "40" '�\�񕶖{����񃂁[�h�������l�Ƃ���
    AutoDrive.Option2(0).Value = True

    LOG_Out "OUT �^���\�񕶏�����"


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
    Dim Hm2         As Single   '���ѐ���
    Dim Hm1         As Single   '���ѐ���
    Dim H0          As Single   '���ѐ���
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
    Dim M1          As String
    Dim Mw          As String
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
    Hm2 = HO(5, Now_Step - 2)
    Hm1 = HO(5, Now_Step - 1)
    H0 = HO(5, Now_Step)
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
                If H0 < �x������ Then
                    ST8
                    Course = "�"
                Else
                    Course = "3"
                End If
            Else
                If H1 < �댯���� And H2 < �댯���� And H3 < �댯���� Then
                    ST3 H2, Hm1, Hm2, c1
                    Course = Course & "4"
                Else
                    ST4 H2, H3, Hm1, Hm2, c1
                End If
            End If

        Case 3, 4
            ST5 H0, H1, H2, H3

        Case 5, 6
            If H0 >= �댯���� Then '�I
                Course = Course & "H"
                ST4 H2, H3, Hm1, Hm2, c1
                GoTo J1
            End If
            If H1 >= �댯���� Or H2 >= �댯���� Or H3 >= �댯���� Then '�J
                Course = Course & "I"
                ST4 H2, H3, Hm1, Hm2, c1
                GoTo J1
            End If
            If H1 < �x������ And H2 < �x������ Or H3 < �x������ Then '�K
                If H0 < �x������ Then
                    ST8
                    GoTo J1
                Else
                    Course = Course & "K"
                    GoTo J1
                End If
            Else
                If �x������ <= H0 Then
                    If �x������ <= Hm1 And �x������ <= Hm2 Then
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
        If H0 >= 6.2 Or H3 >= 6.2 Then '�v���h��(T.P 6.2m)�𒴂���
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
    w = H0 - Hm2
    If w <= -0.1 Then ���ʏ� = "���~��"
    If -0.1 < w And w <= 0.1 Then ���ʏ� = "���΂�"
    If 0.1 < w And w <= 0.3 Then ���ʏ� = "�㏸��"
    If 0.3 < w Then ���ʏ� = "�}�㏸��"

    Print #Log_Repo, ""
    Print #Log_Repo, Format(jgd, "yyyy/mm/dd hh:nn") & "  " & Kind_S
    Print #Log_Repo, SYUBN
    Print #Log_Repo, "�������O�Q���Ԑ��� " & Format(Format(Hm2, "##0.00"), "@@@@@@@") & " " & IIf((Hm2 - �x������) < 0#, "<", ">=") & " �x������  " & IIf((Hm2 - �댯����) < 0#, "<", ">=") & " �댯����"
    Print #Log_Repo, "�������O�P���Ԑ��� " & Format(Format(Hm1, "##0.00"), "@@@@@@@") & " " & IIf((Hm1 - �x������) < 0#, "<", ">=") & " �x������  " & IIf((Hm1 - �댯����) < 0#, "<", ">=") & " �댯����"
    Print #Log_Repo, "���������� �@�@�@�@" & Format(Format(H0, "##0.00"), "@@@@@@@") & " " & IIf((H0 - �x������) < 0#, "<", ">=") & " �x������  " & IIf((H0 - �댯����) < 0#, "<", ">=") & " �댯����"
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
    M1 = Format(Day(jgd), "##") & "��" & _
         Format(Hour(jgd), "#0") & "��" & _
         Format(Minute(jgd), "#0") & "��"
'    H2Z M1, Mw
    Mw = M1
    buf = "�@�@�V��̐��ʂ�" & Mw & "���݁A���̂Ƃ���ƂȂ��Ă��܂��B" & LF
    buf = buf & "�@�@�����O���ʐ��ʊϑ����m�V�쒬�厚�����n���n��" & LF
    M1 = Format(Format(H0, "##0.00"), "@@@@@@")
'    H2Z M1, Mw
    Mw = M1
    buf = buf & "�@�@�@�@�@�@" & Mw & "���[�g���i" & ���ʏ� & "�j" & LF
    If Y_FLAG <> 7 Then
        M1 = Format(Day(jsx), "##") & "��" & _
             Format(Hour(jsx), "#0") & "��" & _
             Format(Minute(jsx), "#0") & "��"
'        H2Z M1, Mw
        Mw = M1
        buf = buf & "�@�@�V��̐��ʂ�" & Mw & "���ɂ́A���̂悤�Ɍ����܂�܂��B" & LF
        buf = buf & "�@�@�����O���ʐ��ʊϑ����m�V�쒬�厚�����n���n��" & LF
        If Y_FLAG <> 1 Then
            M1 = Format(Format(H3r, "###0.00"), "@@@@@@")
        Else
            M1 = Format(Format(H2r, "###0.00"), "@@@@@@")
        End If
'        H2Z M1, Mw
        Mw = M1
        buf = buf & "�@�@�@�@�@�@" & Mw & "���[�g�����x" & LF & " " & LF
    Else
        buf = buf & "�@�@�@�@�@�@" & LF
        buf = buf & "�@�@�@�@�@�@" & LF
        buf = buf & "�@�@�@�@�@�@" & LF
        buf = buf & "�@�@�@�@�@�@" & LF

    End If
'    H2Z buf, Bunw
    Bunw = buf
    Bun2 = Bun2 & Bunw

    Bunw = "�@�@�y�Q�l�z" & LF & _
           "�@�@�����O���ʐ��ʊϑ����m�V�쒬�厚�����n���n" & LF & _
           "�@�@��h�� 6.20m  �댯���� 5.20m  �x������ 3.00m  �w�萅�� 2.00m" & LF
    Bun2 = Bun2 & Bunw & " " & LF

    Bunw = "�@�@�y�V��̍^���\�񔭕\�󋵁z" & LF
    Bunw = Bunw & "�@�@�@�@�@" & Kind_M & LF

    Bun2 = Bun2 & Bunw & " " & LF

    Bunw = "�@�@�₢���킹��" & LF & _
           "�@�@���ʊ֌W�@���m���������ݎ������@�@�ێ��Ǘ��ہ@�s�d�k052(961)4421" & LF & _
           "�@�@�C�ۊ֌W�@�C�ے����É��n���C�ۑ�@�ϑ��\��ہ@�s�d�k052(763)2449" & LF & " "

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
Sub �\������DB_Close()

    If Rst_�\��.State = 1 Then
        Rst_�\��.Close
    End If
    Set Rst_�\�� = Nothing
    Set Con_�\�� = Nothing

End Sub
Sub �\������DB_Connection()

    Dim Con  As String

    LOG_Out "IN    �\������DB_Connection"

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

    LOG_Out "OUT   �\������DB_Connection ABNormal Return"

    On Error GoTo 0

End Sub
Sub �\������DB_Read()

    Dim SQL     As String
    Dim dw      As String
    Dim T_Last  As Date
    Dim n       As Long

    LOG_Out "IN    �\������DB_Read"

    �\������DB_Connection

    If DB_�\�� = False Then
        LOG_Out "OT   �\������DB_Read DB_�\�� = False"
        Exit Sub
    End If

    SQL = "Select MAX(TIME) From �\�񕶗��� Where RAIN_KIND = '" & isRAIN & "'"

    Rst_�\��.Open SQL, Con_�\��, adOpenDynamic, adLockOptimistic

    If Rst_�\��.BOF Or Rst_�\��.EOF Then
       '�����ɂ͂��Ȃ��͂���������������
        Y_FLAG = 0
        Rst_�\��.Close
        �\������DB_Close
        LOG_Out "OUT   �\������DB_Read �����ɂ͂��Ȃ��͂���������������"
        Exit Sub
    End If

    dw = Rst_�\��.Fields(0).Value
    T_Last = CDate(dw)

    Rst_�\��.Close

    n = DateDiff("h", T_Last, jgd) + 1
    If n > 6 Then
        Y_FLAG = 0
        �\������DB_Close
        LOG_Out "OUT   �\������DB_Read n=" & str(n)
        Exit Sub
    End If

    SQL = "Select * From �\�񕶗��� Where TIME = '" & dw & "' AND  RAIN_KIND = '" & isRAIN & "'"
    Rst_�\��.Open SQL, Con_�\��, adOpenDynamic, adLockOptimistic

    Y_FLAG = Rst_�\��.Fields("�\��t���O").Value

    Rst_�\��.Close

    �\������DB_Close

    LOG_Out "OUT   �\������DB_Read SQL=" & SQL

End Sub
Sub �\������DB_Write()

    Dim SQL    As String

    LOG_Out "IN    �\������DB_Write"

    �\������DB_Connection

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
    Rst_�\��.Fields("�\��t���O").Value = Y_FLAG
    Rst_�\��.Fields("�\���ʃR�[�h").Value = Kind_N
    Rst_�\��.Fields("�\����").Value = Kind_S
    Rst_�\��.Fields("Course").Value = Course

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
    
    �\������DB_Close

    LOG_Out "OUT   �\������DB_Write"

End Sub
