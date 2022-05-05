Attribute VB_Name = "MDB"
Option Explicit
Public MDB_Con         As New ADODB.Connection
Public MDB_Rst_H       As New ADODB.Recordset
Sub MDB_Close()

  MDB_Con.Close


End Sub

Sub MDB_Connection(rc As Boolean)

    Dim Con As String

    On Error GoTo ER1

    Con = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\data\����.mdb"

    MDB_Con.ConnectionString = Con
    MDB_Con.Open
    
    rc = True
    On Error GoTo 0
    Exit Sub

ER1:
'    MsgBox "�����D�l�c�a�ɐڑ��ł��܂���A��Ƃ𒆎~���܂�"

    rc = False


End Sub

Sub MDB_�ŐV����()

    Dim t(6) As String
    Dim d(6) As Date
    Dim SQL  As String
    Dim Rst  As New ADODB.Recordset

    SQL = "select MAX(Time) from ����"
    Rst.Open SQL, MDB_Con, adOpenDynamic, adLockReadOnly
    t(1) = Rst.Fields(0).Value
    d(1) = CDate(t(1))
    Rst.Close

    SQL = "select MAX(Time) from FRICS���[�_�[����"
    Rst.Open SQL, MDB_Con, adOpenDynamic, adLockReadOnly
    t(2) = Rst.Fields(0).Value
    d(2) = CDate(t(2))
    Rst.Close

    SQL = "select MAX(Time) from FRICS���[�_�[�\��"
    Rst.Open SQL, MDB_Con, adOpenDynamic, adLockReadOnly
    t(3) = Rst.Fields(0).Value
    d(3) = CDate(t(3))
    d(3) = Format(d(3), "yyyy/mm/dd hh:") & Format(Int(Minute(d(3)) / 10) * 10, "00")
    Rst.Close

    SQL = "select MAX(Time) from �C�ے����[�_�[����"
    Rst.Open SQL, MDB_Con, adOpenDynamic, adLockReadOnly
    t(4) = Rst.Fields(0).Value
    d(4) = CDate(t(4))
    Rst.Close

    SQL = "select MAX(Time) from �C�ے����[�_�[�\��_1"
    Rst.Open SQL, MDB_Con, adOpenDynamic, adLockReadOnly
    t(5) = Rst.Fields(0).Value
    d(5) = CDate(t(5))
    d(5) = DateAdd("h", -1, d(5))
    Rst.Close

    SQL = "select MAX(Time) from �C�ے����[�_�[�\��_2"
    Rst.Open SQL, MDB_Con, adOpenDynamic, adLockReadOnly
    t(6) = Rst.Fields(0).Value
    d(6) = CDate(t(6))
    d(6) = DateAdd("h", -6, d(6))
    Rst.Close

End Sub
Sub Pump_Check(ds As String, dd As Date, w() As Single)

    Dim i      As Long
    Dim j      As Long
    Dim db     As Date
    Dim SQL    As String
    Dim RstP   As New ADODB.Recordset
    Dim wb(4)  As Long

    Const Suiba_Stop = 5.2
    Const Suiba_reStart = 5#
    Const Shimo_Stop = 2.9
    Const Shimo_reStart = 2.7
    Const Haru_Stop = 5.4
    Const Haru_reStart = 5.2

    If w(2) < Shimo_reStart And w(4) < Suiba_reStart And w(6) < Haru_reStart Then
        '
        '�R�ϑ������ʂ��|���v�ĊJ���ʂ��Ⴂ�̂Ń|���v�ғ���DB�ɓo�^
        '
        SQL = "select * from .�|���v���� where Time='" & ds & "'"
        RstP.Open SQL, MDB_Con, adOpenDynamic, adLockOptimistic
        If RstP.EOF Then
            RstP.AddNew
            RstP.Fields("Time").Value = ds
        End If
        RstP.Fields("���V��F").Value = 0
        RstP.Fields("�����O����").Value = 0
        RstP.Fields("�t��").Value = 0
        RstP.Update
        RstP.Close
        Exit Sub
    End If

    If w(2) > Shimo_Stop And w(4) > Suiba_Stop And w(6) > Haru_Stop Then
        '
        '�R�ϑ������ʂ��|���v��~���ʂ�荂���̂Ń|���v��~��DB�ɓo�^
        '
        SQL = "select * from .�|���v���� Time='" & ds & "'"
        RstP.Open SQL, MDB_Con, adOpenDynamic, adLockOptimistic
        If RstP.EOF Then
            RstP.AddNew
            RstP.Fields("Time").Value = ds
        End If
        RstP.Fields("���V��F").Value = 1
        RstP.Fields("�����O����").Value = 1
        RstP.Fields("�t��").Value = 1
        RstP.Update
        RstP.Close
        Exit Sub
    End If

    db = DateAdd("n", -10, dd)
    SQL = "select * from .�|���v���� Time='" & Format(db, "yyyy/mm/dd hh:nn") & "'"
    RstP.Open SQL, MDB_Con, adOpenDynamic, adLockOptimistic
    If RstP.EOF Then
        RstP.Close
        '
        '�P�O���O�f�[�^���o�^����Ă��Ȃ��̂Ō������f�[�^�݂̂Ŕ��f
        '
        SQL = "select * from .�|���v���� Time='" & ds & "'"
        RstP.Open SQL, MDB_Con, adOpenDynamic, adLockOptimistic
        If RstP.EOF Then
            RstP.AddNew
            RstP.Fields("Time").Value = ds
        End If
        If w(2) > Shimo_Stop Then
            RstP.Fields("���V��F").Value = 1
        Else
            RstP.Fields("���V��F").Value = 0
        End If
        If w(4) > Suiba_Stop Then
            RstP.Fields("�����O����").Value = 1
        Else
            RstP.Fields("�����O����").Value = 0
        End If
        If w(6) > Haru_Stop Then
            RstP.Fields("�t��").Value = 1
        Else
            RstP.Fields("�t��").Value = 0
        End If
        RstP.Update
        RstP.Close
        Exit Sub
    End If

    '�P�O���O�̃|���v�f�[�^
    wb(2) = RstP.Fields("���V��F").Value
    wb(4) = RstP.Fields("�����O����").Value
    wb(6) = RstP.Fields("�t��").Value
    RstP.Close

    SQL = "select * from .�|���v���� Time='" & ds & "'"
    RstP.Open SQL, MDB_Con, adOpenDynamic, adLockOptimistic
    If RstP.EOF Then
        RstP.AddNew
        RstP.Fields("Time").Value = ds
    End If
    If w(2) > Shimo_Stop Or (w(2) > Shimo_reStart And wb(2) = 1) Then
        RstP.Fields("���V��F").Value = 1
    Else
        RstP.Fields("���V��F").Value = 0
    End If
    If w(4) > Suiba_Stop Or (w(4) > Suiba_reStart And wb(4) = 1) Then
        RstP.Fields("�����O����").Value = 1
    Else
        RstP.Fields("�����O����").Value = 0
    End If
    If w(6) > Haru_Stop Or (w(6) > Haru_reStart And wb(6) = 1) Then
        RstP.Fields("�t��").Value = 1
    Else
        RstP.Fields("�t��").Value = 0
    End If
    RstP.Update
    RstP.Close

End Sub

Sub Pump_TO_mdb(d1 As Date, d2 As Date)

    Dim i     As Long
    Dim j     As Long
    Dim k     As Long
    Dim n     As Long
    Dim Mn    As Long
    Dim SQLM  As String
    Dim SQL   As String
    Dim dw    As Date
    Dim Tw    As String
    Dim pw    As String
    Dim pnam  As String

    Const Pump_Num = 19

'MDB OPEN
    Set MDB_Rst_H.ActiveConnection = MDB_Con

    SQLM = "SELECT * FROM �|���v���� WHERE Time="

    n = DateDiff("n", d1, d2) / 10 + 1

    If n <= 0 Then
        Exit Sub
    End If

    dw = d1
    For i = 1 To n
        Tw = Format(dw, "yyyy/mm/dd hh:nn")
        Mn = Minute(dw)
        SQL = SQLM & "'" & Tw & "'"
        MDB_Rst_H.Open SQL, MDB_Con, adOpenDynamic, adLockOptimistic
        If MDB_Rst_H.EOF Then
            MDB_Rst_H.AddNew
            MDB_Rst_H.Fields("Time") = Tw
            MDB_Rst_H.Fields("Minute") = Mn
        End If
        For j = 1 To Pump_Num
            pw = Pump_Stats(j).Name
            k = InStr(1, pw, "�|") - 1
            pnam = Mid(pw, 1, k)
            MDB_Rst_H.Fields(pnam) = Pump(j, i)
        Next j
        MDB_Rst_H.Update
        MDB_Rst_H.Close
        dw = DateAdd("n", 10, dw)
    Next i

End Sub
Sub �����ʎ擾(d As Date, h As Single, Name As String, rc As Boolean)

    Dim SQL     As String
    Dim t       As String
    Dim Rst     As New ADODB.Recordset

    rc = False

    t = Format(d, "yyyy/mm/dd hh:nn")

'MDB OPEN
    Set Rst.ActiveConnection = MDB_Con

    SQL = "SELECT " & Name & " FROM ������ WHERE TIME='" & t & "'"

    Rst.Open SQL, MDB_Con, adOpenDynamic, adLockOptimistic

    If Rst.EOF Then
        Rst.Close
        Set Rst.ActiveConnection = Nothing
    Else
        h = Rst.Fields(0).Value
        rc = True
        Rst.Close
        Set Rst.ActiveConnection = Nothing
    End If

End Sub
