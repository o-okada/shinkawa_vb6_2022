VERSION 5.00
Begin VB.Form Control_Form 
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "Control"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Control_Form.frx":0000
   ScaleHeight     =   3705
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   180
      Top             =   3195
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00404000&
      Caption         =   "���@�~"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1620
      TabIndex        =   0
      Top             =   2475
      Width           =   2625
   End
   Begin VB.Label Label4 
      Alignment       =   2  '��������
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  '����
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2610
      TabIndex        =   4
      Top             =   1890
      Width           =   780
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "C 2001 - 2002 Ver.2.00 NIKKEN CONSULTANS.,INC."
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1260
      TabIndex        =   3
      Top             =   3465
      Width           =   4380
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  '����
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1125
      TabIndex        =   2
      Top             =   1215
      Width           =   3750
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  '����
      Caption         =   "�@�V��^�\������v���O�����@"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   945
      TabIndex        =   1
      Top             =   450
      Width           =   4065
   End
End
Attribute VB_Name = "Control_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Run_Drive   As String '�^���\���v���O�����̃C���X�g�[������Ă���h���C�u
Dim Run_Path    As String '�^���\���v���O�����̃C���X�g�[������Ă���t�H���_
Dim Run_Prog    As String '"RSHINK.EXE"
Dim Prun        As Boolean
Dim Flag        As Boolean
' OO4O�̃I�u�W�F�N�g�ϐ���錾����
Dim ssOra       As Object
Dim dbOra       As OraDatabase
Dim dynOra      As OraDynaset
Dim Sdate       As Date
Dim Edate       As Date

Sub Check_Run()

    Select Case GetVersion()
    
    Case 1 ' Windows 95/98�̏ꍇ

        Dim f As Long, sname As String
        Dim hSnap As Long, proc As PROCESSENTRY32
        hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)

        If hSnap = hNull Then Exit Sub

        proc.dwSize = Len(proc)

        ' �v���Z�X���J��Ԃ��擾���܂�
        f = Process32First(hSnap, proc)
        Do While f
            sname = StrZToStr(proc.szExeFile)
            If InStr(1, sname, "RSHINKAWA") > 0 Then
                Prun = True
                Exit Sub
            End If
            f = Process32Next(hSnap, proc)
        Loop
    
    Case 2 ' Windows NT�̏ꍇ
    
        Dim cb As Long
        Dim cbNeeded As Long
        Dim NumElements As Long
        Dim ProcessIDs() As Long
        Dim cbNeeded2 As Long
        Dim NumElements2 As Long
        Dim Modules(1 To 200) As Long
        Dim lRet As Long
        Dim ModuleName As String
        Dim nSize As Long
        Dim hProcess As Long
        Dim i As Long

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
Sub Flag_Check()

    Dim SQL  As String
    Dim i    As Integer
    Dim m    As Long
    Dim nf   As Integer
    Dim dw   As Date
    Dim dc   As String
    Dim w
    Dim j
    Dim d

    SQL = "SELECT * FROM oracle.START_FLAG WHERE river_no = 85053002"  '�V��

    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)


    j = dynOra.Fields("JIKOKU").Value
    w = dynOra.Fields("START_FLAG").Value
    d = dynOra.Fields("WRITE_TIME").Value

    If w = 1 Then '�t���O�������Ȃ��̂ł��̔��ʂ���ꂽ
        m = DateDiff("H", j, Now) + 1
        If m > 48 Then
            w = 0 '�܂�Q���O�̃t���O�͖������Ă���
        End If
    End If

    If w = 1 Then
        Flag = True
        Debug.Print "�J�n�t���O�� " & Format(j, "yyyy/mm/dd hh:ss") & " �ɃI���ɂȂ�܂����B" & Now
        Edate = CDate(j)
        Sdate = DateAdd("d", -1, Edate)
        dc = Format(Sdate, "yyyy/mm/dd") & " 10:00"
        nf = FreeFile
        Open "D:\SHINKAWA\���[�_�[�^���\��\data\BaseDate.dat" For Output As #nf
        Print #nf, dc
        Close #nf
    Else
        Flag = False
        Debug.Print "�J�n�t���O�f�[�^�� " & Format(Now, "yyyy/mm/dd hh:nn:ss") & " ���݃I�t�ł��B" & vbCrLf & _
                   " ���t�f�[�^ = " & Format(j, "yyyy/mm/dd hh:nn:ss") & vbCrLf & _
                   " �t���O    = " & w
    End If

    ORA_DataBase_Close

End Sub


Sub Oracle_Connection(ic As Boolean)

' OO4O �� Oracle �ɐڑ�����
    On Error Resume Next
    ' �Z�b�V�����̍쐬
    Set ssOra = CreateObject("OracleInProcServer.XOraSession")
    If Err <> 0 Then
        MsgBox "���m���I���N���f�[�^�x�[�X�ɐڑ��o���܂���B" & Chr(10) & _
               "CreateObject - Oracle oo4o �G���["
      GoTo ERRHAND
    End If

    ' �T�[�r�X���i�T�[�o���j�� ���[�U��/�p�X���[�h ���w�肷��
    Set dbOra = ssOra.OpenDatabase("ORACLE", "oracle/oracle", 0&)
    If Err <> 0 Then
      MsgBox "���m���I���N���f�[�^�x�[�X�ɐڑ��o���܂���B" & vbCrLf & _
              Err & ": " & Error
      GoTo ERRHAND
    End If

   On Local Error GoTo 0
   ic = True
   Exit Sub

ERRHAND:
   ic = False
   On Local Error GoTo 0

End Sub
'=======================================================================
'  ��n��
'=======================================================================
Sub ORA_DataBase_Close()

'** oo4o �ڑ�����
  Set ssOra = Nothing
  Set dbOra = Nothing

End Sub

Sub Read_Path()

    Dim nf  As Long

    nf = FreeFile
    Open App.Path & "\Path_File.txt" For Input As #nf

    Input #nf, Run_Drive
    Input #nf, Run_Path
    Input #nf, Run_Prog

    Close #nf

End Sub


Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    Prun = False

    If App.PrevInstance Then
        MsgBox "���̃v���O�����͂��łɋN������Ă��܂��B"
        Unload Me

    Else

        Me.Left = (Screen.Width - Me.Width) * 0.5
        Me.Top = (Screen.Height - Me.Height) * 0.3
        Me.Label4.Caption = "      "

        Read_Path

        Check_Run

    End If

End Sub
Private Sub Timer1_Timer()

    Dim i    As Long
    Dim ret  As Long
    Dim ic   As Boolean

    Label2.Caption = Format(Now, "yyyy/mm/dd hh:nn:ss")
    Label2.Refresh
    i = Second(Now)

    If i <> 20 Then Exit Sub

    Check_Run

    If Prun Then
        Label4.Caption = " �^���\���v���O�����͋N�����ł��B"
        Label4.Refresh
        Exit Sub  '�V��^���\���v���O���������ݎ��s��
    Else
        Label4.Caption = " �^���\���v���O�����͑ҋ@���ł��B"
        Label4.Refresh
    End If

    Oracle_Connection ic

    Flag_Check
    
    If Flag Then
        ChDrive Run_Drive
        ChDir Run_Path
        ret = Shell(Run_Prog, 1)
'        MsgBox "�^���\�����J�n����܂����A���͂���ŋA��܂��B"
        End
    End If

End Sub


