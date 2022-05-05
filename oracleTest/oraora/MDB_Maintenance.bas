Attribute VB_Name = "MDB_Maintenance"
Option Explicit
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function CloseHandle Lib "Kernel32.dll" (ByVal Handle As Long) As Long
Public Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Public Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer

Public Type PROCESSENTRY32
   dwSize As Long
   cntUsage As Long
   th32ProcessID As Long
   th32DefaultHeapID As Long
   th32ModuleID As Long
   cntThreads As Long
   th32ParentProcessID As Long
   pcPriClassBase As Long
   dwFlags As Long
   szExeFile As String * 260    ' MAX_PATH
End Type

Public Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long         '1 = Windows 95.
                                '2 = Windows NT
   szCSDVersion As String * 128
End Type

Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_VM_READ = 16
Public Const MAX_PATH = 260
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SYNCHRONIZE = &H100000
'STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF
Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const TH32CS_SNAPPROCESS = &H2&
Public Const hNull = 0

Public Prun        As Boolean

Function StrZToStr(s As String) As String
   StrZToStr = Left$(s, Len(s) - 1)
End Function

Public Function GetVersion() As Long
Dim osinfo As OSVERSIONINFO
Dim retvalue As Integer

osinfo.dwOSVersionInfoSize = 148
osinfo.szCSDVersion = Space$(128)
retvalue = GetVersionExA(osinfo)
GetVersion = osinfo.dwPlatformId
End Function
Sub Check_Run()

    Select Case GetVersion()
    
    Case 1 ' Windows 95/98の場合

        Dim F      As Long
        Dim sname  As String
        Dim hSnap  As Long
        Dim proc   As PROCESSENTRY32

        hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)

        If hSnap = hNull Then Exit Sub

        proc.dwSize = Len(proc)

        ' プロセスを繰り返し取得します
        F = Process32First(hSnap, proc)
        Do While F
            sname = StrZToStr(proc.szExeFile)
            If InStr(1, sname, "RSHINKAWA") > 0 Then
                Prun = True
                Exit Sub
            End If
            F = Process32Next(hSnap, proc)
        Loop
    
    Case 2 ' Windows NTの場合
    
        Dim cb                As Long
        Dim cbNeeded          As Long
        Dim NumElements       As Long
        Dim ProcessIDs()      As Long
        Dim cbNeeded2         As Long
        Dim NumElements2      As Long
        Dim Modules(1 To 300) As Long
        Dim lRet              As Long
        Dim ModuleName        As String
        Dim nSize             As Long
        Dim hProcess          As Long
        Dim i                 As Long

        Prun = False

        ' 各プロセスのIDを含む配列を取得します
        cb = 8
        cbNeeded = 96
        Do While cb <= cbNeeded
            cb = cb * 2
            ReDim ProcessIDs(cb / 4) As Long
            lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
        Loop

        NumElements = cbNeeded / 4

        For i = 1 To NumElements
            ' プロセスのハンドルを取得します
            hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
                Or PROCESS_VM_READ, 0, ProcessIDs(i))
            ' プロセスのハンドルを取得した場合
            If hProcess <> 0 Then
                ' 指定のプロセスのモジュールハンドルの配列を取得します
                lRet = EnumProcessModules(hProcess, Modules(1), 200, _
                                                     cbNeeded2)
                ' モジュール配列が見つかったらモジュールのファイル名を取得します
                If lRet <> 0 Then
                    ModuleName = Space(MAX_PATH)
                    nSize = 500
                    lRet = GetModuleFileNameExA(hProcess, Modules(1), _
                            ModuleName, nSize)
                    sname = Left(ModuleName, lRet)
                    If InStr(1, sname, "RSHINKAWA") > 0 Then
                        Prun = True
                        Exit Sub
                    End If
                End If
            End If
        ' プロセスのハンドルを閉じます
            lRet = CloseHandle(hProcess)
        Next i

    End Select

End Sub

Public Sub CompactMDB()

    Dim ConA    As String
    Dim ConB    As String
    Dim DateS   As String
    Dim FileA   As String '現使用中MDB
    Dim FileB   As String '新MDB
    Dim FileC   As String '保存用MDB
    Dim rc      As Boolean

    DateS = Format(Now, "yyyy-mm-dd-hh")

    ORA_LOG "ローカルMDBの圧縮開始  "

    MDB_Close

    FileA = App.Path & "\data\水文.mdb"
    FileB = App.Path & "\data\Temp.mdb"
    FileC = App.Path & "\data\" & DateS & "_水文.mdb"

    ConA = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\data\水文.mdb"
    ConB = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\data\Temp.mdb"

    ' JRO を使用して Access2000 形式のファイルを最適化する
    Dim jroJET As New JRO.JetEngine
    
    ' 構文 CompactDatabase 最適化前の接続 , 最適化後の接続
    jroJET.CompactDatabase ConA, ConB

    ORA_LOG "Temp.mdb作成終了  "

    '圧縮後ファイルをワークにコピー
    FileCopy FileB, FileC

    ORA_LOG "保存ファイル(" & FileC & ")作成終了  "

    '旧MDBファイルを削除
    Kill FileA

    ORA_LOG "元ファイル(" & FileA & ")削除終了  "

    Dim Cn     As New ADODB.Connection
    Dim Rs     As New ADODB.Recordset
    Dim SQL    As String
    Dim Timew  As String
    Dim dw     As Date
    Dim Dz     As Date


    Cn.ConnectionString = ConB
    Cn.Open

    ORA_LOG "MDBファイル(" & FileB & ")接続終了  "

'テレメータ水位を処理
'最新時刻を取得
    SQL = "select Max(Time) From .水位"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    Timew = Rs.Fields(0).Value
    dw = CDate(Timew)
    Dz = DateAdd("h", -25, dw)             'この時刻未満をDBより削除する
    Timew = Format(Dz, "yyyy/mm/dd hh:nn")
    Rs.Close
'データ削除
    SQL = "DELETE  From .水位  Where Time < '" & Timew & "'"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    ORA_LOG "水位データ処理終了  "

'光水位を処理
'光水位最新時刻を取得
    SQL = "select Max(Time) From .光水位"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    Timew = Rs.Fields(0).Value
    dw = CDate(Timew)
    Dz = DateAdd("h", -25, dw)             'この時刻未満をDBより削除する
    Timew = Format(Dz, "yyyy/mm/dd hh:nn")
    Rs.Close
'データ削除
    SQL = "DELETE  From .光水位  Where Time < '" & Timew & "'"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    ORA_LOG "水位データ処理終了  "

'気象庁レーダー実績
'最新時刻を取得
    SQL = "select Max(Time) From .気象庁レーダー実績"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    Timew = Rs.Fields(0).Value
    dw = CDate(Timew)
    Dz = DateAdd("h", -25, dw)             'この時刻未満をDBより削除する
    Timew = Format(Dz, "yyyy/mm/dd hh:nn")
    Rs.Close
'データ削除
    SQL = "DELETE  From .気象庁レーダー実績  Where Time < '" & Timew & "'"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    ORA_LOG "気象庁レーダー実績処理終了  "

'気象庁レーダー予測_1
'最新時刻を取得
    SQL = "select Max(Time) From .気象庁レーダー予測_1"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    Timew = Rs.Fields(0).Value
    dw = CDate(Timew)
    Dz = DateAdd("h", -25, dw)             'この時刻未満をDBより削除する
    Timew = Format(Dz, "yyyy/mm/dd hh:nn")
    Rs.Close
'データ削除
    SQL = "DELETE  From .気象庁レーダー予測_1  Where Time < '" & Timew & "'"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    ORA_LOG "気象庁レーダー予測_1処理終了  "

'気象庁レーダー予測_2
'最新時刻を取得
    SQL = "select Max(Time) From .気象庁レーダー予測_2"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    Timew = Rs.Fields(0).Value
    dw = CDate(Timew)
    Dz = DateAdd("h", -25, dw)             'この時刻未満をDBより削除する
    Timew = Format(Dz, "yyyy/mm/dd hh:nn")
    Rs.Close
'データ削除
    SQL = "DELETE  From .気象庁レーダー予測_2  Where Time < '" & Timew & "'"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    ORA_LOG "気象庁レーダー予測_2 処理終了  "

'洗堰
'最新時刻を取得
    SQL = "select Max(Time) From .洗堰"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    Timew = Rs.Fields(0).Value
    dw = CDate(Timew)
    Dz = DateAdd("h", -25, dw)             'この時刻未満をDBより削除する
    Timew = Format(Dz, "yyyy/mm/dd hh:nn")
    Rs.Close
'データ削除
    SQL = "DELETE  From .洗堰  Where Time < '" & Timew & "'"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    ORA_LOG "洗堰処理終了  "

'FRICSレーダー実績
'最新時刻を取得
    SQL = "select Max(Time) From .FRICSレーダー実績"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    Timew = Rs.Fields(0).Value
    dw = CDate(Timew)
    Dz = DateAdd("h", -25, dw)             'この時刻未満をDBより削除する
    Timew = Format(Dz, "yyyy/mm/dd hh:nn")
    Rs.Close
'データ削除
    SQL = "DELETE  From .FRICSレーダー実績  Where Time < '" & Timew & "'"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    ORA_LOG "FRICSレーダー実績処理終了  "

'FRICSレーダー予測
'最新時刻を取得
    SQL = "select Max(Time) From .FRICSレーダー予測"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    Timew = Rs.Fields(0).Value
    dw = CDate(Timew)
    Dz = DateAdd("h", -25, dw)             'この時刻未満をDBより削除する
    Timew = Format(Dz, "yyyy/mm/dd hh:nn")
    Rs.Close
'データ削除
    SQL = "DELETE  From .FRICSレーダー予測  Where Time < '" & Timew & "'"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    ORA_LOG "FRICSレーダー予測処理終了  "

'ポンプ実績
'最新時刻を取得
    SQL = "select Max(Time) From .ポンプ実績"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    Timew = Rs.Fields(0).Value
    dw = CDate(Timew)
    Dz = DateAdd("h", -25, dw)             'この時刻未満をDBより削除する
    Timew = Format(Dz, "yyyy/mm/dd hh:nn")
    Rs.Close
'データ削除
    SQL = "DELETE  From .ポンプ実績  Where Time < '" & Timew & "'"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    ORA_LOG "ポンプ実績処理終了  "

'ポンプ履歴
'最新時刻を取得
    SQL = "select Max(Time) From .ポンプ履歴"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    Timew = Rs.Fields(0).Value
    dw = CDate(Timew)
    Dz = DateAdd("h", -25, dw)             'この時刻未満をDBより削除する
    Timew = Format(Dz, "yyyy/mm/dd hh:nn")
    Rs.Close
'データ削除
    SQL = "DELETE  From .ポンプ履歴  Where Time < '" & Timew & "'"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    ORA_LOG "ポンプ履歴処理終了  "

    Set Rs = Nothing
    Set Cn = Nothing

    ORA_LOG "データ削除後の再圧縮処理開始  "
    'データ削除後の再圧縮
    ' 構文 CompactDatabase 最適化前の接続 , 最適化後の接続
    jroJET.CompactDatabase ConB, ConA
    ORA_LOG "データ削除後の再圧縮処理終了  "

    'ワークファイル削除
    Kill FileB
    ORA_LOG "ワークMDB削除終了  "

    MDB_Connection rc
    If rc = False Then
        MsgBox "ローカルDBの圧縮に失敗している可能性があります。" & vbCrLf & _
               "ログファイルを参照してください。" & vbCrLf & _
               "このジョブは終了します。"
        End
    End If

    ORA_LOG "ローカルMDBの圧縮終了  "

End Sub
Sub Pre_Compact(rc As Boolean, RSHINKAWA_RUN As Boolean)

    Dim i             As Long
    Dim c             As String
    Dim nf            As Integer

    rc = False

    Check_Run
    RSHINKAWA_RUN = Prun
    If RSHINKAWA_RUN = False Then
        rc = True  '洪水予測プログラムが起動されていない時はファイルを
        Exit Sub   '確認する必要はないのでここでリターンする。
    End If

    On Error GoTo ERHAND

'ローカルDBをメインテナンスするため洪水予測が実行中なら一時的にストップする
    nf = FreeFile
    Open "D:\SHINKAWA\OracleTest\OraOra\Data\DB.Check" For Output As #nf
    Print #nf, "STOP"
    Close #nf

'洪水予測が止まった事を確認する
    For i = 1 To 10
        Short_Break 5
        nf = FreeFile
        Open "D:\SHINKAWA\OracleTest\OraOra\Data\Yosoku.Check" For Input As #nf
        Line Input #nf, c
        Close #nf
        If c = "OK" Then
            nf = FreeFile
            Open "D:\SHINKAWA\OracleTest\OraOra\Data\DB.Check" For Output As #nf
            Print #nf, "    "
            Close #nf
            nf = FreeFile
            Open "D:\SHINKAWA\OracleTest\OraOra\Data\Yosoku.Check" For Output As #nf
            Print #nf, "    "
            Close #nf
            rc = True
            Exit Sub
        End If
    Next i

ERHAND:
    ORA_LOG "In Pre_Compact 何らかのエラーでローカルDBの圧縮が出来ない"
    nf = FreeFile
    Open "D:\SHINKAWA\OracleTest\OraOra\Data\DB.Check" For Output As #nf
    Print #nf, "    "
    Close #nf
    nf = FreeFile
    Open "D:\SHINKAWA\OracleTest\OraOra\Data\Yosoku.Check" For Output As #nf
    Print #nf, "    "
    Close #nf
    On Error GoTo 0

End Sub
Public Sub Short_Break(s As Long)

'    Sleep s * 1000   'システムが止めるので反応しなくなるからだめ
'                      ただし、ＣＰＵを使わなくなる。
'                      下の方法は反応するがＣＰＵを１００％使う事に
'                      なる。
'
    Dim i   As Date
    Dim j   As Date
    Dim k   As Long

    i = Now

    Do
        j = Now
        k = DateDiff("s", i, j)
        If k >= s Then
            Exit Do
        End If
        DoEvents
    Loop

    Exit Sub

End Sub
