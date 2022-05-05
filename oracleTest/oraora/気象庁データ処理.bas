Attribute VB_Name = "�C�ے��f�[�^����"
Option Explicit
Option Base 1

Global Const RRYU = 135                   '���搔
Global m(250)                  As Single  '2�����b�V�����܂Ƃ߂�����
Global MR(315, 2)              As Integer '���[�_�[���b�V����1k���b�V���̑Ή��\
Global R_Thissen(20, 140)      As Single  '���[�_�[�J�ʗp�e�B�[�Z���W��(�P����ő�Q�O���b�V��)
Global R_Meshu(20, 140)        As Integer '����J�ʗp���b�V���ԍ�(�P����ő�Q�O���b�V��)
Global R_T_Name(140)           As String  '���於�L��
Function FMT1(n As Integer) As String
    FMT1 = Format(Format(n, "####0"), "@@@@@")
End Function
Function FMT2(a As Single) As String
    FMT2 = Format(Format(a, "##0.000"), "@@@@@@@")
End Function
'
'�C�ے��Q�������b�V���o�����P�������b�V���f�[�^�ɕϊ�����
'
'
'
Sub Mesh_2km_to_1km_cvt(w2() As Single, w1() As Single)

    Dim i As Long
    Dim j As Long
    Dim k As Long

    For i = 1 To 315
        j = MR(i, 1)
        k = MR(i, 2)
        w1(j) = w2(k)
'        If w1(j) > 0# Then
'            Debug.Print " i="; i; "  w1(j)="; w1(j)
'        End If
    Next i


End Sub
Sub Mesh_2km_to_1km_data()

    Dim nf    As Integer
    Dim buf   As String
    Dim i     As Integer

    ORA_LOG "IN Mesh_2km_to_1km_data  (" & _
            App.Path & "\data\�C�ے����b�V������1k���b�V���Ή��\.dat) �ǂݍ���"

    nf = FreeFile
    Open App.Path & "\data\�C�ے����b�V������1k���b�V���Ή��\.dat" For Input As #nf

    For i = 1 To 315
        Line Input #nf, buf
        MR(i, 1) = CInt(Mid(buf, 1, 5)) '1km���b�V���ԍ�
        MR(i, 2) = CInt(Mid(buf, 6, 5)) '2km���b�V���ԍ�
    Next i

    Close #nf

    ORA_LOG "OUT Mesh_2km_to_1km_data"

End Sub
'**************************************************
'
'���[�_�[���b�V����135����Ɍv�Z����
'
'
'
'**************************************************
Sub Mesh_To_Ryuiki(w() As Single, RY() As Single, irc As Boolean)

    Dim i               As Integer
    Dim k               As Integer
    Dim m               As Integer
    Dim Rw              As Single
    Dim r               As Single

    ORA_LOG "IN   Sub Mesh_To_Ryuiki"

    For i = 1 To RRYU '135����
        Rw = 0#
        RY(i) = 0#
        For k = 1 To 20
            m = R_Meshu(k, i)  '���[�_�[���b�V���̔ԍ�
            If m = 0 Then Exit For
            r = w(m)
            Rw = Rw + r * R_Thissen(k, i)
'            Debug.Print " k="; k; " R_Meshu(k, i)="; m; " r = "; r; "   R_Thissen(k, i)="; R_Thissen(k, i); "  Rw="; Rw
        Next k
        If Rw < 0# Then Rw = 0#
        RY(i) = Rw
    Next i

    ORA_LOG "OUT  Sub Mesh_To_Ryuiki"

End Sub
Sub ���[�_�[�e�B�[�Z���ǂݍ���()

    Dim i        As Integer
    Dim j        As Integer
    Dim k        As Integer
    Dim m        As Integer
    Dim F        As Single
    Dim x        As String
    Dim y        As String
    Dim buf      As String
    Dim nf       As Integer
    Dim t        As Single

    Dim R_Thissen_A(20, 140)      As Single

    ORA_LOG "IN   Sub ���[�_�[�e�B�[�Z���ǂݍ���"

    nf = FreeFile
    Open App.Path & "\data\���[�_�[�e�B�[�Z��.dat" For Input As #nf

    k = 0
    Do Until EOF(nf)
        Line Input #nf, buf
        k = k + 1
        R_T_Name(k) = Trim(Mid(buf, 6, 5))
        t = 0#
        m = 0
        For i = 1 To 20
            x = Mid(buf, 11 + (i - 1) * 10, 5)
            If IsNumeric(x) Then
                m = m + 1
                y = Mid(buf, 16 + (i - 1) * 10, 5)
                R_Meshu(i, k) = CInt(x)
                R_Thissen_A(i, k) = CSng(y)
                t = t + CSng(y)
            Else
                Exit For
            End If
        Next i
        t = 1# / t
        For i = 1 To m
            R_Thissen(i, k) = R_Thissen_A(i, k) * t
        Next i
    Loop

    Close #nf


    ORA_LOG "Out  Sub ���[�_�[�e�B�[�Z���ǂݍ���"

'    Dim k0 As Long
'    Dim k1 As Long
'    Dim k2 As Long
'    Dim kk As Long
'
'    nf = FreeFile
'    Open App.Path & "\data\���[�_�[�e�B�[�Z��.csv" For Output As #nf
'
'    For kk = 1 To k Step 6
'        k1 = kk
'        k2 = k1 + 5
'        If k2 > k Then k2 = k
'        Print #nf, "N,";
'        For k0 = k1 To k2
'            Print #nf, R_T_Name(k0); ",";
'            Print #nf, R_T_Name(k0); ",";
'            Print #nf, R_T_Name(k0); ",";
'        Next k0
'        Print #nf, ""
'        For j = 1 To 15
'            buf = Str(j) & ","
'            For k0 = k1 To k2
'                If R_Meshu(j, k0) > 0 Then
'                    buf = buf & FMT1(R_Meshu(j, k0)) & ","
'                    buf = buf & FMT2(R_Thissen_A(j, k0)) & ","
'                    buf = buf & FMT2(R_Thissen(j, k0)) & ","
'                Else
'                    buf = buf & ","
'                    buf = buf & ","
'                    buf = buf & ","
'                End If
'            Next k0
'            Print #nf, buf
'        Next j
'        Print #nf, ","
'        Print #nf, ","
'        Print #nf, ","
'        Print #nf, ","
'    Next kk
'    Close #nf

End Sub
