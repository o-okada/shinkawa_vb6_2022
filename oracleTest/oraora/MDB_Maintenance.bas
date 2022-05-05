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
    
    Case 1 ' Windows 95/98�̏ꍇ

        Dim F      As Long
        Dim sname  As String
        Dim hSnap  As Long
        Dim proc   As PROCESSENTRY32

        hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)

        If hSnap = hNull Then Exit Sub

        proc.dwSize = Len(proc)

        ' �v���Z�X���J��Ԃ��擾���܂�
        F = Process32First(hSnap, proc)
        Do While F
            sname = StrZToStr(proc.szExeFile)
            If InStr(1, sname, "RSHINKAWA") > 0 Then
                Prun = True
                Exit Sub
            End If
            F = Process32Next(hSnap, proc)
        Loop
    
    Case 2 ' Windows NT�̏ꍇ
    
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

        ' �e�v���Z�X��ID���܂ޔz����擾���܂�
        cb = 8
        cbNeeded = 96
        Do While cb <= cbNeeded
            cb = cb * 2
            ReDim ProcessIDs(cb / 4) As Long
            lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
        Loop

        NumElements = cbNeeded / 4

        For i = 1 To NumElements
            ' �v���Z�X�̃n���h�����擾���܂�
            hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
                Or PROCESS_VM_READ, 0, ProcessIDs(i))
            ' �v���Z�X�̃n���h�����擾�����ꍇ
            If hProcess <> 0 Then
                ' �w��̃v���Z�X�̃��W���[���n���h���̔z����擾���܂�
                lRet = EnumProcessModules(hProcess, Modules(1), 200, _
                                                     cbNeeded2)
                ' ���W���[���z�񂪌��������烂�W���[���̃t�@�C�������擾���܂�
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
        ' �v���Z�X�̃n���h������܂�
            lRet = CloseHandle(hProcess)
        Next i

    End Select

End Sub

Public Sub CompactMDB()

    Dim ConA    As String
    Dim ConB    As String
    Dim DateS   As String
    Dim FileA   As String '���g�p��MDB
    Dim FileB   As String '�VMDB
    Dim FileC   As String '�ۑ��pMDB
    Dim rc      As Boolean

    DateS = Format(Now, "yyyy-mm-dd-hh")

    ORA_LOG "���[�J��MDB�̈��k�J�n  "

    MDB_Close

    FileA = App.Path & "\data\����.mdb"
    FileB = App.Path & "\data\Temp.mdb"
    FileC = App.Path & "\data\" & DateS & "_����.mdb"

    ConA = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\data\����.mdb"
    ConB = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\data\Temp.mdb"

    ' JRO ���g�p���� Access2000 �`���̃t�@�C�����œK������
    Dim jroJET As New JRO.JetEngine
    
    ' �\�� CompactDatabase �œK���O�̐ڑ� , �œK����̐ڑ�
    jroJET.CompactDatabase ConA, ConB

    ORA_LOG "Temp.mdb�쐬�I��  "

    '���k��t�@�C�������[�N�ɃR�s�[
    FileCopy FileB, FileC

    ORA_LOG "�ۑ��t�@�C��(" & FileC & ")�쐬�I��  "

    '��MDB�t�@�C�����폜
    Kill FileA

    ORA_LOG "���t�@�C��(" & FileA & ")�폜�I��  "

    Dim Cn     As New ADODB.Connection
    Dim Rs     As New ADODB.Recordset
    Dim SQL    As String
    Dim Timew  As String
    Dim dw     As Date
    Dim Dz     As Date


    Cn.ConnectionString = ConB
    Cn.Open

    ORA_LOG "MDB�t�@�C��(" & FileB & ")�ڑ��I��  "

'�e�����[�^���ʂ�����
'�ŐV�������擾
    SQL = "select Max(Time) From .����"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    Timew = Rs.Fields(0).Value
    dw = CDate(Timew)
    Dz = DateAdd("h", -25, dw)             '���̎���������DB���폜����
    Timew = Format(Dz, "yyyy/mm/dd hh:nn")
    Rs.Close
'�f�[�^�폜
    SQL = "DELETE  From .����  Where Time < '" & Timew & "'"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    ORA_LOG "���ʃf�[�^�����I��  "

'�����ʂ�����
'�����ʍŐV�������擾
    SQL = "select Max(Time) From .������"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    Timew = Rs.Fields(0).Value
    dw = CDate(Timew)
    Dz = DateAdd("h", -25, dw)             '���̎���������DB���폜����
    Timew = Format(Dz, "yyyy/mm/dd hh:nn")
    Rs.Close
'�f�[�^�폜
    SQL = "DELETE  From .������  Where Time < '" & Timew & "'"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    ORA_LOG "���ʃf�[�^�����I��  "

'�C�ے����[�_�[����
'�ŐV�������擾
    SQL = "select Max(Time) From .�C�ے����[�_�[����"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    Timew = Rs.Fields(0).Value
    dw = CDate(Timew)
    Dz = DateAdd("h", -25, dw)             '���̎���������DB���폜����
    Timew = Format(Dz, "yyyy/mm/dd hh:nn")
    Rs.Close
'�f�[�^�폜
    SQL = "DELETE  From .�C�ے����[�_�[����  Where Time < '" & Timew & "'"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    ORA_LOG "�C�ے����[�_�[���я����I��  "

'�C�ے����[�_�[�\��_1
'�ŐV�������擾
    SQL = "select Max(Time) From .�C�ے����[�_�[�\��_1"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    Timew = Rs.Fields(0).Value
    dw = CDate(Timew)
    Dz = DateAdd("h", -25, dw)             '���̎���������DB���폜����
    Timew = Format(Dz, "yyyy/mm/dd hh:nn")
    Rs.Close
'�f�[�^�폜
    SQL = "DELETE  From .�C�ے����[�_�[�\��_1  Where Time < '" & Timew & "'"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    ORA_LOG "�C�ے����[�_�[�\��_1�����I��  "

'�C�ے����[�_�[�\��_2
'�ŐV�������擾
    SQL = "select Max(Time) From .�C�ے����[�_�[�\��_2"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    Timew = Rs.Fields(0).Value
    dw = CDate(Timew)
    Dz = DateAdd("h", -25, dw)             '���̎���������DB���폜����
    Timew = Format(Dz, "yyyy/mm/dd hh:nn")
    Rs.Close
'�f�[�^�폜
    SQL = "DELETE  From .�C�ے����[�_�[�\��_2  Where Time < '" & Timew & "'"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    ORA_LOG "�C�ے����[�_�[�\��_2 �����I��  "

'��
'�ŐV�������擾
    SQL = "select Max(Time) From .��"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    Timew = Rs.Fields(0).Value
    dw = CDate(Timew)
    Dz = DateAdd("h", -25, dw)             '���̎���������DB���폜����
    Timew = Format(Dz, "yyyy/mm/dd hh:nn")
    Rs.Close
'�f�[�^�폜
    SQL = "DELETE  From .��  Where Time < '" & Timew & "'"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    ORA_LOG "�􉁏����I��  "

'FRICS���[�_�[����
'�ŐV�������擾
    SQL = "select Max(Time) From .FRICS���[�_�[����"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    Timew = Rs.Fields(0).Value
    dw = CDate(Timew)
    Dz = DateAdd("h", -25, dw)             '���̎���������DB���폜����
    Timew = Format(Dz, "yyyy/mm/dd hh:nn")
    Rs.Close
'�f�[�^�폜
    SQL = "DELETE  From .FRICS���[�_�[����  Where Time < '" & Timew & "'"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    ORA_LOG "FRICS���[�_�[���я����I��  "

'FRICS���[�_�[�\��
'�ŐV�������擾
    SQL = "select Max(Time) From .FRICS���[�_�[�\��"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    Timew = Rs.Fields(0).Value
    dw = CDate(Timew)
    Dz = DateAdd("h", -25, dw)             '���̎���������DB���폜����
    Timew = Format(Dz, "yyyy/mm/dd hh:nn")
    Rs.Close
'�f�[�^�폜
    SQL = "DELETE  From .FRICS���[�_�[�\��  Where Time < '" & Timew & "'"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    ORA_LOG "FRICS���[�_�[�\�������I��  "

'�|���v����
'�ŐV�������擾
    SQL = "select Max(Time) From .�|���v����"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    Timew = Rs.Fields(0).Value
    dw = CDate(Timew)
    Dz = DateAdd("h", -25, dw)             '���̎���������DB���폜����
    Timew = Format(Dz, "yyyy/mm/dd hh:nn")
    Rs.Close
'�f�[�^�폜
    SQL = "DELETE  From .�|���v����  Where Time < '" & Timew & "'"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    ORA_LOG "�|���v���я����I��  "

'�|���v����
'�ŐV�������擾
    SQL = "select Max(Time) From .�|���v����"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    Timew = Rs.Fields(0).Value
    dw = CDate(Timew)
    Dz = DateAdd("h", -25, dw)             '���̎���������DB���폜����
    Timew = Format(Dz, "yyyy/mm/dd hh:nn")
    Rs.Close
'�f�[�^�폜
    SQL = "DELETE  From .�|���v����  Where Time < '" & Timew & "'"
    Rs.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    ORA_LOG "�|���v���������I��  "

    Set Rs = Nothing
    Set Cn = Nothing

    ORA_LOG "�f�[�^�폜��̍Ĉ��k�����J�n  "
    '�f�[�^�폜��̍Ĉ��k
    ' �\�� CompactDatabase �œK���O�̐ڑ� , �œK����̐ڑ�
    jroJET.CompactDatabase ConB, ConA
    ORA_LOG "�f�[�^�폜��̍Ĉ��k�����I��  "

    '���[�N�t�@�C���폜
    Kill FileB
    ORA_LOG "���[�NMDB�폜�I��  "

    MDB_Connection rc
    If rc = False Then
        MsgBox "���[�J��DB�̈��k�Ɏ��s���Ă���\��������܂��B" & vbCrLf & _
               "���O�t�@�C�����Q�Ƃ��Ă��������B" & vbCrLf & _
               "���̃W���u�͏I�����܂��B"
        End
    End If

    ORA_LOG "���[�J��MDB�̈��k�I��  "

End Sub
Sub Pre_Compact(rc As Boolean, RSHINKAWA_RUN As Boolean)

    Dim i             As Long
    Dim c             As String
    Dim nf            As Integer

    rc = False

    Check_Run
    RSHINKAWA_RUN = Prun
    If RSHINKAWA_RUN = False Then
        rc = True  '�^���\���v���O�������N������Ă��Ȃ����̓t�@�C����
        Exit Sub   '�m�F����K�v�͂Ȃ��̂ł����Ń��^�[������B
    End If

    On Error GoTo ERHAND

'���[�J��DB�����C���e�i���X���邽�ߍ^���\�������s���Ȃ�ꎞ�I�ɃX�g�b�v����
    nf = FreeFile
    Open "D:\SHINKAWA\OracleTest\OraOra\Data\DB.Check" For Output As #nf
    Print #nf, "STOP"
    Close #nf

'�^���\�����~�܂��������m�F����
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
    ORA_LOG "In Pre_Compact ���炩�̃G���[�Ń��[�J��DB�̈��k���o���Ȃ�"
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

'    Sleep s * 1000   '�V�X�e�����~�߂�̂Ŕ������Ȃ��Ȃ邩�炾��
'                      �������A�b�o�t���g��Ȃ��Ȃ�B
'                      ���̕��@�͔������邪�b�o�t���P�O�O���g������
'                      �Ȃ�B
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
