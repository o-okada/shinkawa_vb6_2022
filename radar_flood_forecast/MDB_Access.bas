Attribute VB_Name = "MDB_Access"
Option Explicit
Option Base 1

Public MDBx              As Boolean              '—\‘ªƒf[ƒ^ƒx[ƒX  Ú‘±‰Â=True  Ú‘±•s‰Â=False
Public …•¶MDB           As String               '…•¶MDB‚Ìƒtƒ‹ƒpƒX
Public —š—ğMDB           As String               '—š—ğMDB‚Ìƒtƒ‹ƒpƒX
Public Con_…•¶          As New ADODB.Connection
Public Con_—š—ğ          As New ADODB.Connection
Public Rec_…•¶          As New ADODB.Recordset
Public Rec_—š—ğ          As New ADODB.Recordset

Public Con_—\•ñ•¶        As New ADODB.Connection
Public Rst_—\•ñ•¶        As New ADODB.Recordset
Public DB_—\•ñ•¶         As Boolean

Public H_Pred(500, 5, 4) As Single               '…ˆÊ—\‘ª—š—ğ
Public R_Pred(500, 5, 4) As Single               '‰J—Ê—\‘ª—š—ğ
Public T_Pred(500)       As Date                 '—\‘ªŒvZŒ»

Public History           As Boolean              '—š—ğ•\¦ƒRƒ“ƒgƒ[ƒ‹ •\¦—L‚è=True  –³‚µ=False
Sub MDB_…•¶_Close()

    Con_…•¶.Close
    Set Rec_…•¶.ActiveConnection = Nothing

End Sub
Sub MDB_…•¶_Connection()

    Dim Con_str As String
    Dim a

    On Error GoTo ER1

    Con_str = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & …•¶MDB
    Con_…•¶.ConnectionString = Con_str
    Con_…•¶.Open

    Set Rec_…•¶.ActiveConnection = Con_…•¶
    MDBx = True
    On Error GoTo 0
    Exit Sub

ER1:
    a = MsgBox("MDB_…•¶‚ÌDB‚ÉƒAƒNƒZƒX‚Å‚«‚Ü‚¹‚ñADB‚Ì—L–³AODBC“™‚Ìİ’è‚ğŠm”F‚µ‚Ä‚­‚¾‚³‚¢B" & vbCrLf & _
           "ŒvZ‚ğ‘±s‚µ‚Ü‚·‚©(‘±s‚Ìê‡‚Í—\‘ª’l‚Ì—š—ğ‚Í•Û‘¶‚³‚ê‚Ü‚¹‚ñ)H", vbYesNo + vbInformation)
    If a = vbYes Then
        MDBx = False
        Exit Sub
    Else
        End
    End If


End Sub
Sub MDB_—š—ğ_Close()

    Con_—š—ğ.Close
    Set Rec_—š—ğ.ActiveConnection = Nothing

End Sub


Sub MDB_—š—ğ_Connection()

    Dim Con_str As String
    Dim a

    On Error GoTo ER1

    Con_str = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & —š—ğMDB
    Con_—š—ğ.ConnectionString = Con_str
    Con_—š—ğ.Open

    Set Rec_—š—ğ.ActiveConnection = Con_—š—ğ
    MDBx = True
    On Error GoTo 0
    Exit Sub

ER1:
    a = MsgBox("—\‘ª—š—ğ‚ÌDB‚ÉƒAƒNƒZƒX‚Å‚«‚Ü‚¹‚ñADB‚Ì—L–³AODBC“™‚Ìİ’è‚ğŠm”F‚µ‚Ä‚­‚¾‚³‚¢B" & vbCrLf & _
           "ŒvZ‚ğ‘±s‚µ‚Ü‚·‚©(‘±s‚Ìê‡‚Í—\‘ª’l‚Ì—š—ğ‚Í•Û‘¶‚³‚ê‚Ü‚¹‚ñ)H", vbYesNo + vbInformation)
    On Error GoTo 0
    If a = vbYes Then
        MDBx = False
        Exit Sub
    Else
        End
    End If

End Sub
Sub MDB_—š—ğ_Read()


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
        SQL = "Select * From FRICS—š—ğ Where Time Between  '" & Format(jsd, "yyyy/mm/dd hh:nn") & _
              "' And '" & Format(jgd, "yyyy/mm/dd hh:nn") & "' AND Minute=" & fn
    Else
        SQL = "Select * From ‹CÛ’¡—š—ğ Where Time Between  '" & Format(jsd, "yyyy/mm/dd hh:nn") & _
              "' And '" & Format(jgd, "yyyy/mm/dd hh:nn") & "' AND Minute=" & fn
    End If


    Rec_—š—ğ.Open SQL, Con_—š—ğ, adOpenDynamic, adLockReadOnly

    For i = 1 To 500
        For j = 1 To 5
            For k = 1 To 4
                H_Pred(i, j, k) = -99#
                R_Pred(i, j, k) = -99#
            Next k
        Next j
    Next i

    If Rec_—š—ğ.BOF Or Rec_—š—ğ.EOF Then
        Rec_—š—ğ.Close
        Exit Sub
    End If

    Do Until Rec_—š—ğ.EOF

        d = CDate(Rec_—š—ğ.Fields("Time").Value)
        n = DateDiff("h", jsd, d) + 1
        T_Pred(n) = d

'‰º”VˆêF
        buf = Rec_—š—ğ.Fields("‰º”VˆêF").Value
        rw = Split(buf, ",")
        k = UBound(rw) - 1
        For j = 0 To k
            sr = rw(j)
            H_Pred(n, 1, j + 1) = sr
        Next j
'‘å¡
        buf = Rec_—š—ğ.Fields("‘å¡").Value
        rw = Split(buf, ",")
        k = UBound(rw) - 1
        For j = 0 To k
            sr = rw(j)
            H_Pred(n, 2, j + 1) = sr
        Next j
'…êŠO…ˆÊ
        buf = Rec_—š—ğ.Fields("…êìŠO…ˆÊ").Value
        rw = Split(buf, ",")
        k = UBound(rw) - 1
        For j = 0 To k
            sr = rw(j)
            H_Pred(n, 3, j + 1) = sr
        Next j
'‹v’n–ì
        buf = Rec_—š—ğ.Fields("‹v’n–ì").Value
        rw = Split(buf, ",")
        k = UBound(rw) - 1
        For j = 0 To k
            sr = rw(j)
            H_Pred(n, 4, j + 1) = sr
        Next j
't“ú
        buf = Rec_—š—ğ.Fields("t“ú").Value
        rw = Split(buf, ",")
        k = UBound(rw) - 1
        For j = 0 To k
            sr = rw(j)
            H_Pred(n, 5, j + 1) = sr
        Next j
'—\‘ª‰J—Ê
        buf = Rec_—š—ğ.Fields("—\‘ª~‰J").Value
        rw = Split(buf, ",")
        k = UBound(rw) - 1
        For j = 0 To k
            sr = rw(j)
            R_Pred(n, 1, j + 1) = sr
        Next j

        Rec_—š—ğ.MoveNext
    Loop

    Rec_—š—ğ.Close

End Sub
'
'—\‘ª’l‚ğMDB‚É•Û‘¶‚·‚éB
'
'
Sub MDB_—š—ğ_Write()

    Dim i   As Integer
    Dim j   As Integer
    Dim ns  As Integer
    Dim buf As String
    Dim SQL As String

    Const f2 = "##0.00"
    Const f1 = "###0.0"

    If isRAIN = "02" Then
        SQL = "Select * From FRICS—š—ğ Where Time = '" & Format(jgd, "yyyy/mm/dd hh:nn") & "'"
    Else
        SQL = "Select * From ‹CÛ’¡—š—ğ Where Time = '" & Format(jgd, "yyyy/mm/dd hh:nn") & "'"
    End If

    Rec_—š—ğ.Open SQL, Con_—š—ğ, adOpenDynamic, adLockOptimistic

    If Rec_—š—ğ.BOF Or Rec_—š—ğ.EOF Then
        Rec_—š—ğ.AddNew
        Rec_—š—ğ.Fields("Time").Value = Format(jgd, "yyyy/mm/dd hh:nn")
    End If
    Rec_—š—ğ.Fields("Minute").Value = Format(Minute(jgd), "00")

'“úŒõìŠO…ˆÊ
    buf = Format(DH_Tide, f2) & ","                   'Œ»“V•¶’ªˆÊ‚ÆÀÑ…ˆÊ‚Æ‚Ì·
    buf = buf & Format(HO(1, Now_Step), f2) & ","     'Œ»…ˆÊ(ÀÑ)
    buf = buf & Format(HO(1, Now_Step + 1), f2) & "," '1ŠÔŒã
    buf = buf & Format(HO(1, Now_Step + 2), f2) & "," '2ŠÔŒã
    buf = buf & Format(HO(1, Now_Step + 3), f2) & "," '3ŠÔŒã
    Rec_—š—ğ.Fields("“úŒõìŠO…ˆÊ").Value = buf
'‰º”VˆêF
    ns = V_Sec_Num(1)
    buf = Format(HO(3, Now_Step), f2) & ","           'Œ»…ˆÊ(ÀÑ)
    buf = buf & Format(HQ(1, ns, NT - 12), f2) & ","  '1ŠÔŒã
    buf = buf & Format(HQ(1, ns, NT - 6), f2) & ","   '2ŠÔŒã
    buf = buf & Format(HQ(1, ns, NT), f2) & ","       '3ŠÔŒã
    Rec_—š—ğ.Fields("‰º”VˆêF").Value = buf
'‘å¡
    ns = V_Sec_Num(2)
    buf = Format(HO(4, Now_Step), f2) & ","           'Œ»…ˆÊ(ÀÑ)
    buf = buf & Format(HQ(1, ns, NT - 12), f2) & ","  '1ŠÔŒã
    buf = buf & Format(HQ(1, ns, NT - 6), f2) & ","   '2ŠÔŒã
    buf = buf & Format(HQ(1, ns, NT), f2) & ","       '3ŠÔŒã
    Rec_—š—ğ.Fields("‘å¡").Value = buf
'…êŠO…ˆÊ
    ns = V_Sec_Num(3)
    buf = Format(HO(5, Now_Step), f2) & ","           'Œ»…ˆÊ(ÀÑ)
    buf = buf & Format(HQ(1, ns, NT - 12), f2) & ","  '1ŠÔŒã
    buf = buf & Format(HQ(1, ns, NT - 6), f2) & ","   '2ŠÔŒã
    buf = buf & Format(HQ(1, ns, NT), f2) & ","       '3ŠÔŒã
    Rec_—š—ğ.Fields("…êìŠO…ˆÊ").Value = buf
'‹v’n–ì
    ns = V_Sec_Num(4)
    buf = Format(HO(6, Now_Step), f2) & ","           'Œ»…ˆÊ(ÀÑ)
    buf = buf & Format(HQ(1, ns, NT - 12), f2) & ","  '1ŠÔŒã
    buf = buf & Format(HQ(1, ns, NT - 6), f2) & ","   '2ŠÔŒã
    buf = buf & Format(HQ(1, ns, NT), f2) & ","       '3ŠÔŒã
    Rec_—š—ğ.Fields("‹v’n–ì").Value = buf
't“ú
    ns = V_Sec_Num(5)
    buf = Format(HO(7, Now_Step), f2) & ","           'Œ»…ˆÊ(ÀÑ)
    buf = buf & Format(HQ(1, ns, NT - 12), f2) & ","  '1ŠÔŒã
    buf = buf & Format(HQ(1, ns, NT - 6), f2) & ","   '2ŠÔŒã
    buf = buf & Format(HQ(1, ns, NT), f2) & ","       '3ŠÔŒã
    Rec_—š—ğ.Fields("t“ú").Value = buf
'—\‘ª‰J—Ê
    buf = Format(RO(1, Now_Step), f1) & ","           'Œ»—¬ˆæ•½‹Ï‰J—Ê
    buf = buf & Format(RO(1, Now_Step + 1), f1) & "," '1ŠÔŒã—¬ˆæ•½‹Ï‰J—Ê
    buf = buf & Format(RO(1, Now_Step + 2), f1) & "," '2ŠÔŒã—¬ˆæ•½‹Ï‰J—Ê
    buf = buf & Format(RO(1, Now_Step + 3), f1) & "," '3ŠÔŒã—¬ˆæ•½‹Ï‰J—Ê
    Rec_—š—ğ.Fields("—\‘ª~‰J").Value = buf

'DB‘‚«‚İ
    Rec_—š—ğ.Update
    Rec_—š—ğ.Close

End Sub
'
'Œ»İŒvZ‚Ég‚í‚ê‚Ä‚¢‚é—\‘ª‰J—Ê‚Ìó‘Ô‚ğƒZ[ƒu‚·‚éB
'
'
'
Sub RAIN_SELECT_READ()

    Dim Con    As String
    Dim R_Con  As New ADODB.Connection
    Dim R_Rst  As New ADODB.Recordset

    LOG_Out "IN    RAIN_SELECT_READ"

    Con = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & —š—ğMDB

    R_Con.ConnectionString = Con
    R_Con.Open

    Set R_Rst.ActiveConnection = R_Con
    R_Rst.Open "SELECT * FROM RAIN_SELECT", R_Con, adOpenDynamic, adLockOptimistic
    
    If R_Rst.Fields("‹CÛ’¡").Value Then
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
'Œ»İŒvZ‚Ég‚í‚ê‚Ä‚¢‚é—\‘ª‰J—Ê‚Ìó‘Ô‚ğƒZ[ƒu‚·‚éB
'
'
'
Sub RAIN_SELECT_SAVE()

    Dim Con    As String
    Dim R_Con  As New ADODB.Connection
    Dim R_Rst  As New ADODB.Recordset

    LOG_Out "IN    RAIN_SELECT_SAVE"

    Con = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & —š—ğMDB

    R_Con.ConnectionString = Con
    R_Con.Open

    Set R_Rst.ActiveConnection = R_Con
    R_Rst.Open "SELECT * FROM RAIN_SELECT", R_Con, adOpenDynamic, adLockOptimistic

    R_Rst.Fields("‹CÛ’¡").Value = KISYO
    R_Rst.Fields("FRICS").Value = FRICS

    R_Rst.Update
    R_Rst.Close
    R_Con.Close

    Set R_Rst = Nothing
    Set R_Con = Nothing

    LOG_Out "OUT   RAIN_SELECT_SAVE"

End Sub
Sub —\•ñ•¶—š—ğDB_Close()

    On Error Resume Next

    If Rst_—\•ñ•¶.State = 1 Then
        Rst_—\•ñ•¶.Close
    End If
    Set Rst_—\•ñ•¶ = Nothing
    Set Con_—\•ñ•¶ = Nothing

End Sub
Sub —\•ñ•¶—š—ğDB_Connection()

    Dim Con  As String

    LOG_Out "IN    —\•ñ•¶—š—ğDB_Connection"

    On Error GoTo ERH1
    
    Con = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & —š—ğMDB

    Con_—\•ñ•¶.ConnectionString = Con
    Con_—\•ñ•¶.Open

    Set Rst_—\•ñ•¶.ActiveConnection = Con_—\•ñ•¶
    DB_—\•ñ•¶ = True

    LOG_Out "OUT   —\‘ª—š—ğDB_Connection Normal Return"

    On Error GoTo 0
    Exit Sub

ERH1:

    DB_—\•ñ•¶ = False
    MsgBox "—\•ñ•¶—š—ğƒf[ƒ^ƒx[ƒX‚ÉÚ‘±‚Å‚«‚Ü‚¹‚ñ‚Å‚µ‚½A—š—ğ‚Íc‚è‚Ü‚¹‚ñB"
    LOG_Out "—\•ñ•¶—š—ğƒf[ƒ^ƒx[ƒX‚ÉÚ‘±‚Å‚«‚Ü‚¹‚ñ‚Å‚µ‚½A—š—ğ‚Íc‚è‚Ü‚¹‚ñB"

    LOG_Out "OUT   —\•ñ•¶—š—ğDB_Connection ABNormal Return"

    On Error GoTo 0

End Sub
Sub —\•ñ•¶—š—ğDB_Read()

    Dim SQL     As String
    Dim dw      As String
    Dim T_Last  As Date
    Dim n       As Long

    LOG_Out "IN    —\•ñ•¶—š—ğDB_Read"

    —\•ñ•¶—š—ğDB_Connection

    If DB_—\•ñ•¶ = False Then
        LOG_Out "OT   —\‘ª—š—ğDB_Read DB_—\•ñ•¶ = False"
        Exit Sub
    End If

    SQL = "Select MAX(TIME) From —\•ñ•¶—š—ğ Where RAIN_KIND = '" & isRAIN & "'"

    Rst_—\•ñ•¶.Open SQL, Con_—\•ñ•¶, adOpenDynamic, adLockOptimistic

    If Rst_—\•ñ•¶.BOF Or Rst_—\•ñ•¶.EOF Then
       '‚±‚±‚É‚Í‚±‚È‚¢‚Í‚¸‚¾‚ª‚à‚µ‚«‚½‚ç
        BP = 0
        Rst_—\•ñ•¶.Close
        —\•ñ•¶—š—ğDB_Close
        LOG_Out "OUT   —\‘ª—š—ğDB_Read ‚±‚±‚É‚Í‚±‚È‚¢‚Í‚¸‚¾‚ª‚à‚µ‚«‚½‚ç"
        Exit Sub
    End If

    dw = Rst_—\•ñ•¶.Fields(0).Value
    T_Last = CDate(dw)

    Rst_—\•ñ•¶.Close

    n = DateDiff("h", T_Last, jgd) + 1
    If n > 25 Then
        BP = 0
        LOG_Out "                   jgd=" & TIMEC(jgd)
        LOG_Out "                T_Last=" & TIMEC(T_Last)
        —\•ñ•¶—š—ğDB_Close
        LOG_Out "OUT   —\•ñ•¶—š—ğDB_Read n=" & str(n)
        Exit Sub
    End If

    SQL = "Select * From —\•ñ•¶—š—ğ Where TIME = '" & dw & "' AND  RAIN_KIND = '" & isRAIN & "'"
    Rst_—\•ñ•¶.Open SQL, Con_—\•ñ•¶, adOpenDynamic, adLockOptimistic
    If Rst_—\•ñ•¶.EOF Then
        BP = 0
        Wng_Last_Time = 0
    Else
        BP = Rst_—\•ñ•¶.Fields("—\•ñƒtƒ‰ƒO").Value
        Wng_Last_Time = Rst_—\•ñ•¶.Fields("Course").Value
    End If

    Rst_—\•ñ•¶.Close

    —\•ñ•¶—š—ğDB_Close

    LOG_Out "OUT   —\•ñ•¶—š—ğDB_Read SQL=" & SQL

End Sub
Sub —\•ñ•¶—š—ğDB_Write()

    Dim SQL    As String

    LOG_Out "IN    —\•ñ•¶—š—ğDB_Write"

    —\•ñ•¶—š—ğDB_Connection

    If DB_—\•ñ•¶ = False Then
        Exit Sub
    End If

    SQL = "Select * From —\•ñ•¶—š—ğ Where TIME = '" & Format(jgd, "yyyy/mm/dd hh:nn") & "' AND RAIN_KIND = '" & isRAIN & "'"

    Rst_—\•ñ•¶.Open SQL, Con_—\•ñ•¶, adOpenDynamic, adLockOptimistic

    If Rst_—\•ñ•¶.BOF Or Rst_—\•ñ•¶.EOF Then
        Rst_—\•ñ•¶.AddNew
        Rst_—\•ñ•¶.Fields("Time").Value = Format(jgd, "yyyy/mm/dd hh:nn")
        Rst_—\•ñ•¶.Fields("RAIN_KIND").Value = isRAIN
    End If

    If Pattan_Now <> 4 Then
        Rst_—\•ñ•¶.Fields("—\•ñƒtƒ‰ƒO").Value = Pattan_Now  '—\•ñ•¶ƒpƒ^[ƒ“
        Rst_—\•ñ•¶.Fields("—\•ñí•ÊƒR[ƒh").Value = Messag(Pattan_Now).Patn(16) 'Kind_N
        Rst_—\•ñ•¶.Fields("—\•ñí•Ê").Value = Messag(Pattan_Now).Patn(2) 'Kind_S
    Else
        '—\•ñ•¶”­¶I—¹
        Rst_—\•ñ•¶.Fields("—\•ñƒtƒ‰ƒO").Value = 0                    '—\•ñ•¶ƒpƒ^[ƒ“
        Rst_—\•ñ•¶.Fields("—\•ñí•ÊƒR[ƒh").Value = "0"            'Kind_N
        Rst_—\•ñ•¶.Fields("—\•ñí•Ê").Value = "^…’ˆÓî•ñ‰ğœ"   'Kind_S
    End If
    Rst_—\•ñ•¶.Fields("Course").Value = Wng_Last_Time

    If isRAIN = "01" Then
        Rst_—\•ñ•¶.Fields("RAIN_NAME").Value = "‹CÛ’¡"
    Else
        Rst_—\•ñ•¶.Fields("RAIN_NAME").Value = "FRICS"
    End If

    If PRACTICE_FLG_CODE = "40" Then
        Rst_—\•ñ•¶.Fields("PRACTICE").Value = "—\•ñ"
    Else
        Rst_—\•ñ•¶.Fields("PRACTICE").Value = "‰‰K"
    End If

    Rst_—\•ñ•¶.Update

    Rst_—\•ñ•¶.Close
    
    —\•ñ•¶—š—ğDB_Close

    LOG_Out "OUT   —\•ñ•¶—š—ğDB_Write"

End Sub
