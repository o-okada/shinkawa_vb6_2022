VERSION 5.00
Begin VB.Form OracleDB 
   BackColor       =   &H00404040&
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "OraDB"
   ClientHeight    =   4590
   ClientLeft      =   3075
   ClientTop       =   2220
   ClientWidth     =   9435
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   9.75
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "OlacleDB.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   9435
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   195
      TabIndex        =   8
      Top             =   2085
      Width           =   6735
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   750
      Top             =   3855
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   165
      Top             =   3840
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�I�@��"
      Height          =   495
      Left            =   7380
      TabIndex        =   6
      Top             =   3480
      Width           =   1795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DBҲ���ݽ"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7380
      TabIndex        =   5
      Top             =   2797
      Visible         =   0   'False
      Width           =   1795
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�\�񕶏�������"
      Height          =   495
      Left            =   7380
      TabIndex        =   4
      Top             =   2115
      Visible         =   0   'False
      Width           =   1795
   End
   Begin VB.Label Label4 
      Alignment       =   2  '��������
      Caption         =   "Label4"
      Height          =   255
      Left            =   1230
      TabIndex        =   7
      Top             =   870
      Width           =   6810
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  '����
      Caption         =   "��荞�ݑ҂�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1725
      TabIndex        =   3
      Top             =   1305
      Width           =   6150
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "���݂̏��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   495
      TabIndex        =   2
      Top             =   1290
      Width           =   1125
   End
   Begin VB.Label Time_Disp 
      BorderStyle     =   1  '����
      Caption         =   "Label4"
      Height          =   255
      Left            =   -15
      TabIndex        =   1
      Top             =   4320
      Width           =   9435
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      AutoSize        =   -1  'True
      BorderStyle     =   1  '����
      Caption         =   " �V��^���\�������f�[�^�擾�V�X�e���@"
      BeginProperty Font 
         Name            =   "�l�r ����"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   870
      TabIndex        =   0
      Top             =   315
      Width           =   7590
   End
End
Attribute VB_Name = "OracleDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'���W���[�����FOraDB
'
'OracleDB.Check_Araizeki_Time()���C�������B�y01-01�z
'��OracleDB.Time1_Timer()�̏������C�����邱�ƁB�y01-01-01�z
'
'OracleDB.Check_F_MESSYU_10MIN_1_Time()���A�A�A�y03-01�z
'��Timer1_Timer()�̏������C�����邱�ƁB�y03-01-01�z
'
'OracleDB.Check_F_MESSYU_10MIN_2_Time()���A�A�A�y04-01�z
'��OracleDB.Timer1_Timer()�̏������C�����邱�ƁB�y04-01-01�z
'
'OracleDB.Check_F_RADAR_TIME()���A�A�A�y06-01�z
'��OracleDB.Timer1_Timer()�̏������C�����邱�ƁB�y06-01-01�z
'
'OracleDB.Check_P_MESSYU_10MIN_Time()���A�A�A�y07-01�z
'��OracleDB.Check_P_MESSYU_10MIN_Time()���C�����邱�ƁB�y07-01-01�z
'
'OracleDB.Check_P_MESSYU_1Hour_Time()���A�A�A�y08-01�z
'
'OracleDB.Check_P_RADAR_Time()���A�A�A�y09-01�z
'��OracleDB.Timer1_Timer()�̏������C�����邱�ƁB�y09-01-01�z
'
'******************************************************************************
Option Explicit
Option Base 1
Dim jobg As Boolean

'******************************************************************************
'�T�u���[�`���FCheck_Araizeki_Time()
'�����T�v�F
'�􉁉z���ʃf�[�^���`�F�b�N����B
'******************************************************************************
'Sub Check_Araizeki_Time(ic As Boolean)
'    Dim nf    As Integer
'    Dim n     As Long
'    Dim d1    As Date
'    Dim d2    As Date
'    Dim d3    As Date
'    Dim ans   As Long
'    Dim buf   As String
'    Dim irc   As Boolean
'    Dim d1st  As String
'    Dim d2st  As String
'    nf = FreeFile
'    'Ver0.0.0 �C���J�n 1900/01/01 00:00
'    'Debug.Print " Freefile="; nf
'    'Ver0.0.0 �C���I�� 1900/01/01 00:00
'    Open App.Path & "\data\Araizeki.DAT" For Input As #nf
'    Line Input #nf, buf
'    d1 = CDate(buf)
'    Close #nf
'    d3 = Format(Now, "yyyy/mm/dd hh") & ":00"
'    n = DateDiff("h", d1, d3)
'    If n > 25 Then
'        d1 = DateAdd("h", -25, d3)          '�O��I�����Q�T���Ԃ��O�������̂Ŏ�荞�݊J�n������ύX
'    End If
'    '******************************************************
'    'Ver1.0.0 �C���J�n 2015/08/04 O.OKADA�y01-01�z
'    '��OracleDB.Time1_Timer()�̏������C�����邱�ƁB�y01-01-01�z
'    '******************************************************
'    'ORA_KANSOKU_JIKOKU_GET "ARAIZEKI", d2, irc
'    Exit Sub
'    '******************************************************
'    'Ver1.0.0 �C���I�� 2015/08/04 O.OKADA�y01-01�z
'    '******************************************************
'    If irc = False Then
'        ic = irc
'        GoTo JUMP
'    End If
'    d1st = Format(d1, "yyyy/mm/dd hh:nn")
'    d2st = Format(d2, "yyyy/mm/dd hh:nn")
'    If d2st > d1st Then
'        n = DateDiff("h", d1, d2) + 1
'        If n > 100 Then                     '100�͓K���Ɍ��߂��l�A�v�����if���Ɉ���������Ȃ��悤�ɂ����B2002/08/07 in YOKOHAMA
'            ans = MsgBox("�ǉ��Ŏ擾���悤�Ƃ��Ă���􉁃f�[�^�X�e�b�v���Q�S������" & vbCrLf & _
'                         "�Ԋu������܂��B��Ƃ��p�����܂����H" & vbCrLf & _
'                         "�V�K�̍^���v�Z�ł͂��߂邱�Ƃ����i�߂��܂��B" & vbCrLf & _
'                         "[�͂�]�ł��̃W���u�͏I�����܂��A[������]�Ōp�����܂��B", vbInformation + vbYesNo)
'            If ans = vbYes Then
'                Close
'                End
'            End If
'        End If
'        d1 = DateAdd("n", 10, d1)
'        ORA_LOG "�􉁃f�[�^��荞�݊J�n " & d1 & " ���� " & d2 & "�܂�"
'        ORA_Araizeki d1, d2, ic
'        If Not ic Then
'            ORA_LOG "�I���N���f�[�^�x�[�X���􉁃f�[�^���擾���悤�Ƃ�������" & vbCrLf & _
'                    "�G���[���������Ă��܂��B"
'            GoTo JUMP
'        Else
'            ORA_LOG "�􉁃f�[�^��荞�ݐ���I��"
'            ORA_LOG "�􉁃f�[�^�����������݊J�n " & d2
'            nf = FreeFile
'            Open App.Path & "\data\Araizeki.DAT" For Output As #nf
'            Print #nf, Format(d2, "yyyy/mm/dd hh:nn")
'            Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
'            Close #nf
'            ORA_LOG "�􉁃f�[�^�����������ݏI��"
'        End If
'    End If
'JUMP:
'End Sub

'******************************************************************************
'�T�u���[�`���FCheck_P_RADAR_Time()
'�����T�v�F
'FRICS���щJ�ʃf�[�^���`�F�b�N����B
'******************************************************************************
'Sub Check_P_RADAR_Time(ic As Boolean)
'    Dim nf     As Integer
'    Dim n      As Long
'    Dim d1     As Date
'    Dim d2     As Date
'    Dim d3     As Date
'    Dim ans    As Long
'    Dim buf    As String
'    Dim irc    As Boolean
'    Dim d1st   As String
'    Dim d2st   As String
'    nf = FreeFile
'    Open App.Path & "\data\P_RADAR.DAT" For Input As #nf
'    Line Input #nf, buf
'    d1 = CDate(buf)
'    Close #nf
'    d3 = Format(Now, "yyyy/mm/dd hh") & ":00"
'    n = DateDiff("h", d1, d3)
'    If n > 25 Then
'        d1 = DateAdd("h", -25, d3)          '�O��I�����Q�T���Ԃ��O�������̂Ŏ�荞�݊J�n������ύX
'    End If
'    ORA_KANSOKU_JIKOKU_GET "P_RADAR", d2, irc
'    If irc = False Then
'        GoTo JUMP
'    End If
'    d1st = Format(d1, "yyyy/mm/dd hh:nn")
'    d2st = Format(d2, "yyyy/mm/dd hh:nn")
'    If d2st > d1st Then
'        n = DateDiff("h", d1, d2) + 1
'        If n > 100 Then                     '100�͓K���Ɍ��߂��l�A�v�����if���Ɉ���������Ȃ��悤�ɂ����B2002/08/07 in YOKOHAMA
'            ans = MsgBox("�ǉ��Ŏ擾���悤�Ƃ��Ă���FRICS���[�_�f�[�^�X�e�b�v���Q�S������" & vbCrLf & _
'                         "�Ԋu������܂��B��Ƃ��p�����܂����H" & vbCrLf & _
'                         "�V�K�̍^���v�Z�ł͂��߂邱�Ƃ����i�߂��܂��B" & vbCrLf & _
'                         "[�͂�]�ł��̃W���u�͏I�����܂��A[������]�Ōp�����܂��B", vbInformation + vbYesNo)
'            If ans = vbYes Then
'                Close
'                End
'            End If
'        End If
'        d1 = DateAdd("n", 10, d1)
'        ORA_LOG "FRICS���у��[�_�f�[�^��荞�݊J�n " & d1 & " ���� " & d2 & "�܂�"
'        '******************************************************
'        'Ver1.0.0 �C���J�n 2015/08/05 O.OKADA�y09-01�z
'        '��OracleDB.Timer1_Timer()�̏������C�����邱�ƁB�y09-01-01�z
'        '******************************************************
'        'ORA_P_RADAR d1, d2, ic
'        Exit Sub
'        '******************************************************
'        'Ver1.0.0 �C���I�� 2015/08/05 O.OKADA�y09-01�z
'        '******************************************************
'        If Not ic Then
'            ORA_LOG "�I���N���f�[�^�x�[�X���FRICS���у��[�_�f�[�^���擾���悤�Ƃ�������" & vbCrLf & _
'                    "�G���[���������Ă��܂��B"
'            GoTo JUMP
'        Else
'            ORA_LOG "FRICS���у��[�_�f�[�^��荞�ݐ���I��"
'        End If
'    End If
'JUMP:
'End Sub

'******************************************************************************
'�T�u���[�`���FCheck_Suii_Time()
'�����T�v�F
'���ʃf�[�^���`�F�b�N����B
'******************************************************************************
Sub Check_Suii_Time(ic As Boolean)
    Dim nf  As Integer
    Dim n   As Long
    Dim d1  As Date
    Dim d2  As Date
    Dim d3  As Date
    Dim ans As Long
    Dim buf As String
    Dim irc As Boolean
    nf = FreeFile
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'Debug.Print " Freefile="; nf
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    Open App.Path & "\data\P_WATER.DAT" For Input As #nf
    Line Input #nf, buf
    d1 = CDate(buf)
    Close #nf
'    d3 = Format(Now, "yyyy/mm/dd hh") & ":00"
'    n = DateDiff("h", d1, d3)
'    If n > 25 Then
'        d1 = DateAdd("h", -25, d3)          '�O��I�����Q�T���Ԃ��O�������̂Ŏ�荞�݊J�n������ύX
'    End If
    Call WaterDataNewTime(d2, irc)
    If irc = False Then
        ic = irc
      GoTo JUMP
    End If
    If d2 > d1 Then
'        n = DateDiff("h", d1, d2) + 1
'        If n > 100 Then                     '100�͓K���Ɍ��߂��l�A�v�����if���Ɉ���������Ȃ��悤�ɂ����B2002/08/07 in YOKOHAMA
'            ans = MsgBox("�ǉ��Ŏ擾���悤�Ƃ��Ă��鐅�ʃf�[�^�X�e�b�v���Q�S������" & vbCrLf & _
'                         "�Ԋu������܂��B��Ƃ��p�����܂����H" & vbCrLf & _
'                         "�V�K�̍^���v�Z�ł͂��߂邱�Ƃ����i�߂��܂��B" & vbCrLf & _
'                         "[�͂�]�ł��̃W���u�͏I�����܂��A[������]�Ōp�����܂��B", vbInformation + vbYesNo)
'            If ans = vbYes Then
'                Close
'                End
'            End If
'        End If
        d1 = DateAdd("n", 10, d1)
        ORA_LOG "���ʃf�[�^��荞�݊J�n " & d1 & " ���� " & d2 & "�܂�"
        ORA_P_WATER d1, d2, ic
        If Not ic Then
            ORA_LOG "���m���͐���V�X�e���f�[�^�x�[�X���u��萅�ʃf�[�^���擾���悤�Ƃ�������" & vbCrLf & _
                    "�G���[���������Ă��܂��B"
            GoTo JUMP
        Else
            ORA_LOG "���ʃf�[�^��荞�ݐ���I��"
            ORA_LOG "���ʃf�[�^�����������݊J�n " & d2
            nf = FreeFile
            Open App.Path & "\data\P_WATER.DAT" For Output As #nf
            Print #nf, Format(d2, "yyyy/mm/dd hh:nn")
'            Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
            Print #nf, Format(d1, "yyyy/mm/dd hh:nn")
'            d1 = CDate(buf)
            Close #nf
            ORA_LOG "���ʃf�[�^�����������ݏI��"
        End If
    End If
JUMP:
End Sub

'******************************************************************************
'�T�u���[�`���FCheck_ORA_OWARI_WATER()
'�����T�v�F
'���P�[�u�����ʃf�[�^���`�F�b�N����B
'******************************************************************************
Sub Check_ORA_OWARI_WATER(ic As Boolean)
    Dim nf  As Integer
    Dim n   As Long
    Dim d1  As Date
    Dim d2  As Date
    Dim d3  As Date
    Dim ans As Long
    Dim buf As String
    Dim irc As Boolean
    
    Exit Sub
    
    nf = FreeFile
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'Debug.Print " Freefile="; nf
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    Open App.Path & "\data\OWARI_WATER.DAT" For Input As #nf
    Line Input #nf, buf
    d1 = CDate(buf)
    Close #nf
'    d3 = Format(Now, "yyyy/mm/dd hh") & ":00"
'    n = DateDiff("h", d1, d3)
'    If n > 25 Then
'        d1 = DateAdd("h", -25, d3)          '�O��I�����Q�T���Ԃ��O�������̂Ŏ�荞�݊J�n������ύX
'    End If
    ORA_KANSOKU_JIKOKU_GET "OWARI_WATER", d2, irc
    If irc = False Then
        ic = irc
      GoTo JUMP
    End If
    If d2 > d1 Then
'        n = DateDiff("h", d1, d2) + 1
'        If n > 100 Then                     '100�͓K���Ɍ��߂��l�A�v�����if���Ɉ���������Ȃ��悤�ɂ����B2002/08/07 in YOKOHAMA
'            ans = MsgBox("�ǉ��Ŏ擾���悤�Ƃ��Ă��鐅�ʃf�[�^�X�e�b�v���Q�S������" & vbCrLf & _
'                         "�Ԋu������܂��B��Ƃ��p�����܂����H" & vbCrLf & _
'                         "�V�K�̍^���v�Z�ł͂��߂邱�Ƃ����i�߂��܂��B" & vbCrLf & _
'                         "[�͂�]�ł��̃W���u�͏I�����܂��A[������]�Ōp�����܂��B", vbInformation + vbYesNo)
'            If ans = vbYes Then
'                Close
'                End
'            End If
'        End If
        d1 = DateAdd("n", 10, d1)
        ORA_LOG "�����ʃf�[�^��荞�݊J�n " & d1 & " ���� " & d2 & "�܂�"
        ORA_OWARI_WATER d1, d2, ic
        If Not ic Then
            ORA_LOG "�I���N���f�[�^�x�[�X�������ʃf�[�^���擾���悤�Ƃ�������" & vbCrLf & _
                    "�G���[���������Ă��܂��B"
            GoTo JUMP
        Else
            ORA_LOG "�����ʃf�[�^��荞�ݐ���I��"
            ORA_LOG "�����ʃf�[�^�����������݊J�n " & d2
            nf = FreeFile
            Open App.Path & "\data\OWARI_WATER.DAT" For Output As #nf
            Print #nf, Format(d2, "yyyy/mm/dd hh:nn")
'            Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
            Print #nf, Format(d1, "yyyy/mm/dd hh:nn")
'            d1 = CDate(buf)
            Close #nf
            ORA_LOG "���ʃf�[�^�����������ݏI��"
        End If
    End If
JUMP:
End Sub

'******************************************************************************
'�T�u���[�`���FCheck_F_RADAR_Time()
'�����T�v�F
' FRICS�\���J�ʃf�[�^���`�F�b�N����B
'******************************************************************************
'Sub Check_F_RADAR_Time(ic As Boolean)
'    Dim nf     As Integer
'    Dim n      As Long
'    Dim d1     As Date
'    Dim d2     As Date
'    Dim d3     As Date
'    Dim ans    As Long
'    Dim buf    As String
'    Dim irc    As Boolean
'    Dim d1st   As String
'    Dim d2st   As String
'    nf = FreeFile
'    Open App.Path & "\data\F_RADAR.DAT" For Input As #nf
'    Line Input #nf, buf
'    d1 = CDate(buf)
'    Close #nf
'    d3 = Format(Now, "yyyy/mm/dd hh") & ":00"
'    n = DateDiff("h", d1, d3)
'    If n > 25 Then
'        d1 = DateAdd("h", -25, d3)          '�O��I�����Q�T���Ԃ��O�������̂Ŏ�荞�݊J�n������ύX
'    End If
'    ORA_KANSOKU_JIKOKU_GET "F_RADAR", d2, irc
'    If irc = False Then
'        ic = irc
'        GoTo JUMP
'    End If
'    d1st = Format(d1, "yyyy/mm/dd hh:nn")
'    d2st = Format(d2, "yyyy/mm/dd hh:nn")
'    If d2st > d1st Then
'        ORA_LOG "FRICS�\�����[�_�f�[�^��荞�݊J�n " & d2
'        '******************************************************
'        'Ver1.0.0 �C���J�n 2015/08/05 O.OKADA�y06-01�z
'        '��OracleDB.Timer1_Timer()�̏������C�����邱�ƁB�y06-01-01�z
'        '******************************************************
'        'ORA_F_RADAR d2, ic
'        Exit Sub
'        '******************************************************
'        'Ver1.0.0 �C���I�� 2015/08/05 O.OKADA�y06-01�z
'        '******************************************************
'        If Not ic Then
'            ORA_LOG "�I���N���f�[�^�x�[�X���FRICS�\�����[�_�f�[�^���擾���悤�Ƃ�������" & vbCrLf & _
'                    "�G���[���������Ă��܂��B"
'            GoTo JUMP
'        Else
'            ORA_LOG "FRICS�\�����[�_�f�[�^��荞�ݐ���I��"
'        End If
'    End If
'JUMP:
'End Sub

'******************************************************************************
'�T�u���[�`���FCheck_F_MESSYU_10MIN_1_Time()
'�����T�v�F
'�C�ے��\���莞�J�ʃf�[�^���`�F�b�N����B
'���P�O���\���i�P���ԕ��j�P�O�����U��
'******************************************************************************
'Sub Check_F_MESSYU_10MIN_1_Time(ic As Boolean)
'    Dim nf     As Integer
'    Dim n      As Long
'    Dim d1     As Date
'    Dim d2     As Date
'    Dim d3     As Date
'    Dim dw     As Date
'    Dim ans    As Long
'    Dim buf    As String
'    Dim irc    As Boolean
'    Dim da     As String
'    Dim db     As String
'    Dim d1st   As String
'    Dim d2st   As String
'    nf = FreeFile
'    Open App.Path & "\data\F_MESSYU_10MIN_1.DAT" For Input As #nf
'    Line Input #nf, buf
'    d1 = CDate(buf)
'    Close #nf
'    d3 = Format(Now, "yyyy/mm/dd hh") & ":00"
'    n = DateDiff("h", d1, d3)
'    If n > 25 Then
'        d1 = DateAdd("h", -25, d3)          '�O��I�����Q�T���Ԃ��O�������̂Ŏ�荞�݊J�n������ύX
'    End If
'    ORA_KANSOKU_JIKOKU_GET "F_MESSYU_10MIN_1", d2, irc
'    If irc = False Then
'        ic = irc
'        GoTo JUMP
'    End If
'    d1st = Format(d1, "yyyy/mm/dd hh:nn")
'    d2st = Format(d2, "yyyy/mm/dd hh:nn")
'    If d2st > d1st Then
'        n = DateDiff("h", d1, d2) + 1
'        If n > 100 Then                     '100�͓K���Ɍ��߂��l�A�v�����if���Ɉ���������Ȃ��悤�ɂ����B2002/08/07 in YOKOHAMA
'            ans = MsgBox("�ǉ��Ŏ擾���悤�Ƃ��Ă���C�ے����[�_�f�[�^�X�e�b�v���Q�S������" & vbCrLf & _
'                         "�Ԋu������܂��B��Ƃ��p�����܂����H" & vbCrLf & _
'                         "�V�K�̍^���v�Z�ł͂��߂邱�Ƃ����i�߂��܂��B" & vbCrLf & _
'                         "[�͂�]�ł��̃W���u�͏I�����܂��A[������]�Ōp�����܂��B", vbInformation + vbYesNo)
'            If ans = vbYes Then
'                Close
'                End
'            End If
'        End If
'        dw = DateAdd("n", 10, d1)
'        ORA_LOG "�C�ے��P�O���\���J�ʃf�[�^��荞�݊J�n " & dw & " ���� " & d2 & "�܂�"
'        da = Format(dw, "yyyy/mm/dd hh:nn")
'        db = Format(d2, "yyyy/mm/dd hh:nn")
'        Do Until da > db
'            '******************************************************
'            'Ver1.0.0 �C���J�n 2015/08/05�y03-01�z
'            'Timer1_Timer()�̏������C�����邱�ƁB�y03-01-01�z
'            '******************************************************
'            'ORA_F_MESSYU_10MIN_1 dw, ic
'            Exit Sub
'            '******************************************************
'            'Ver1.0.0 �C���I�� 2015/08/05 O.OKADA�y03-01�z
'            '******************************************************
'            If Not ic Then
'                ORA_LOG "�I���N���f�[�^�x�[�X���C�ے��P�O���\���J�ʃf�[�^���擾���悤�Ƃ�������" & vbCrLf & _
'                        "�G���[���������Ă��܂��Bdw=" & Format(dw, "yyyy/mm/dd hh:nn")
'            End If
'            dw = DateAdd("n", 10, dw)
'            da = Format(dw, "yyyy/mm/dd hh:nn")
'        Loop
'    End If
'JUMP:
'End Sub

'******************************************************************************
'�T�u���[�`���GCheck_F_MESSYU_10MIN_2_Time()
'�����T�v�F
'�C�ے��\�������J�ʃf�[�^���`�F�b�N����B
'�������\���i�U���ԕ��j�P�O�����P�W��
'******************************************************************************
'Sub Check_F_MESSYU_10MIN_2_Time(ic As Boolean)
'    Dim i   As Integer
'    Dim nf  As Integer
'    Dim n   As Long
'    Dim d1  As Date
'    Dim d2  As Date
'    Dim d3  As Date
'    Dim dw  As Date
'    Dim ans As Long
'    Dim buf As String
'    Dim irc As Boolean
'    nf = FreeFile
'    Open App.Path & "\data\F_MESSYU_10MIN_2.DAT" For Input As #nf
'    Line Input #nf, buf
'    d1 = CDate(buf)
'    Close #nf
'    d3 = Format(Now, "yyyy/mm/dd hh") & ":00"
'    n = DateDiff("h", d1, d3)
'    If n > 25 Then
'        d1 = DateAdd("h", -25, d3)          '�O��I�����Q�T���Ԃ��O�������̂Ŏ�荞�݊J�n������ύX
'    End If
'    ORA_KANSOKU_JIKOKU_GET "F_MESSYU_10MIN_2", d2, irc
'    If irc = False Then
'        ic = irc
'        GoTo JUMP
'    End If
'    dw = DateAdd("h", 1, d1)
'    If d2 > d1 Then
'        n = DateDiff("h", dw, d2) + 1
'        If n > 100 Then                     '100�͓K���Ɍ��߂��l�A�v�����if���Ɉ���������Ȃ��悤�ɂ����B2002/08/07 in YOKOHAMA
'            ans = MsgBox("�ǉ��Ŏ擾���悤�Ƃ��Ă���C�ے����[�_�f�[�^�X�e�b�v���Q�S������" & vbCrLf & _
'                         "�Ԋu������܂��B��Ƃ��p�����܂����H" & vbCrLf & _
'                         "�V�K�̍^���v�Z�ł͂��߂邱�Ƃ����i�߂��܂��B" & vbCrLf & _
'                         "[�͂�]�ł��̃W���u�͏I�����܂��A[������]�Ōp�����܂��B", vbInformation + vbYesNo)
'            If ans = vbYes Then
'                Close
'                End
'            End If
'        End If
'        ORA_LOG "�C�ے������\���J�ʃf�[�^��荞�݊J�n " & dw & " ���� " & d2 & "�܂�"
'        For i = 1 To n
'            '******************************************************
'            'Ver1.0.0 �C���J�n 2015/08/05 O.OKADA�y04-01�z
'            'OracleDB.Timer1_Timer()�̏������C�����邱�ƁB�y04-01-01�z
'            '******************************************************
'            'ORA_F_MESSYU_10MIN_2 dw, ic
'            Exit Sub
'            '******************************************************
'            'Ver1.0.0 �C���I�� 2015/08/05 O.OKADA�y04-01�z
'            '******************************************************
'            If Not ic Then
'                ORA_LOG "�I���N���f�[�^�x�[�X���C�ے������\���J�ʃf�[�^���擾���悤�Ƃ�������" & vbCrLf & _
'                        "�G���[���������Ă��܂��B"
'                GoTo JUMP
'            Else
'                ORA_LOG "�C�ے������\���J�ʃf�[�^��荞�ݐ���I��"
'                ORA_LOG "�C�ے������\���J�ʃf�[�^�����������݊J�n " & d2
'                nf = FreeFile
'                Open App.Path & "\data\F_MESSYU_10MIN_2.DAT" For Output As #nf
'                Print #nf, Format(d2, "yyyy/mm/dd hh:nn")
'                Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
'                Close #nf
'                ORA_LOG "�C�ے������\���J�ʃf�[�^�����������ݏI��"
'            End If
'            dw = DateAdd("h", 1, dw)
'        Next i
'    End If
'JUMP:
'End Sub

'******************************************************************************
'�T�u���[�`���FCheck_P_MESSYU_1HOUR_Time()
'�����T�v�F
'�C�ے����ѐ����J�ʃf�[�^���`�F�b�N����B
'���g�p���Ă��Ȃ��B
'******************************************************************************
'Sub Check_P_MESSYU_1HOUR_Time(ic As Boolean)
'    Dim nf  As Integer
'    Dim n   As Long
'    Dim d1  As Date
'    Dim d2  As Date
'    Dim d3  As Date
'    Dim ans As Long
'    Dim buf As String
'    Dim irc As Boolean
'    nf = FreeFile
'    Open App.Path & "\data\P_MESSYU_1HOUR.DAT" For Input As #nf
'    Line Input #nf, buf
'    d1 = CDate(buf)
'    Close #nf
'    d3 = Format(Now, "yyyy/mm/dd hh") & ":00"
'    n = DateDiff("h", d1, d3)
'    If n > 25 Then
'        d1 = DateAdd("h", -25, d3)          '�O��I�����Q�T���Ԃ��O�������̂Ŏ�荞�݊J�n������ύX
'    End If
'    ORA_KANSOKU_JIKOKU_GET "P_MESSYU_1HOUR", d2, irc
'    If irc = False Then
'        ic = irc
'        GoTo JUMP
'    End If
'    If d2 > d1 Then
'        n = DateDiff("h", d1, d2) + 1
'        If n > 100 Then                     '100�͓K���Ɍ��߂��l�A�v�����if���Ɉ���������Ȃ��悤�ɂ����B2002/08/07 in YOKOHAMA
'            ans = MsgBox("�ǉ��Ŏ擾���悤�Ƃ��Ă���C�ے����[�_�f�[�^�X�e�b�v���Q�S������" & vbCrLf & _
'                         "�Ԋu������܂��B��Ƃ��p�����܂����H" & vbCrLf & _
'                         "�V�K�̍^���v�Z�ł͂��߂邱�Ƃ����i�߂��܂��B" & vbCrLf & _
'                         "[�͂�]�ł��̃W���u�͏I�����܂��A[������]�Ōp�����܂��B", vbInformation + vbYesNo)
'            If ans = vbYes Then
'                Close
'                End
'            End If
'        End If
'        d1 = DateAdd("n", 10, d1)
'        ORA_LOG "�C�ے����щJ�ʃf�[�^��荞�݊J�n " & d1 & " ���� " & d2 & "�܂�"
'        '******************************************************
'        'Ver1.0.0 �C���J�n 2015/08/05 O.OKADA�y08-01�z
'        '���g�p���Ă��Ȃ��A���̃T�u���[�`�����Ăяo����Ă��炸�A�e���͈͂Ȃ��Ɣ��f����B
'        '******************************************************
'        'ORA_P_MESSYU_1Hour d1, d2, ic
'        Exit Sub
'        '******************************************************
'        'Ver1.0.0 �C���I�� 2015/08/05 O.OKADA�y08-01�z
'        '******************************************************
'        If Not ic Then
'            ORA_LOG "�I���N���f�[�^�x�[�X���C�ے����[�_���уf�[�^���擾���悤�Ƃ�������" & vbCrLf & _
'                    "�G���[���������Ă��܂��B"
'            GoTo JUMP
'        Else
'            ORA_LOG "�C�ے����[�_���уf�[�^��荞�ݐ���I��"
'            ORA_LOG "�C�ے����[�_���уf�[�^�����������݊J�n " & d2
'            nf = FreeFile
'            Open App.Path & "\data\P_MESSYU_1HOUR.DAT" For Output As #nf
'            Print #nf, Format(d2, "yyyy/mm/dd hh:nn")
'            Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
'            d1 = CDate(buf)
'            Close #nf
'            ORA_LOG "�C�ے����[�_���уf�[�^�����������ݏI��"
'        End If
'    End If
'JUMP:
'End Sub

'******************************************************************************
'�T�u���[�`���FCheck_P_MESSYU_10MIN_Time()
'�����T�v�F
'�C�ے����ђ莞�J�ʃf�[�^�`�F�b�N
'******************************************************************************
'Sub Check_P_MESSYU_10MIN_Time(ic As Boolean)
'    Dim nf  As Integer
'    Dim n   As Long
'    Dim d1  As Date
'    Dim d2  As Date
'    Dim d3  As Date
'    Dim ans As Long
'    Dim buf As String
'    Dim irc As Boolean
'    nf = FreeFile
'    Open App.Path & "\data\P_MESSYU_10MIN.DAT" For Input As #nf
'    Line Input #nf, buf
'    d1 = CDate(buf)
'    Close #nf
'    d3 = Format(Now, "yyyy/mm/dd hh") & ":00"
'    n = DateDiff("h", d1, d3)
'    If n > 25 Then
'        d1 = DateAdd("h", -25, d3)          '�O��I�����Q�T���Ԃ��O�������̂Ŏ�荞�݊J�n������ύX
'    End If
'    ORA_KANSOKU_JIKOKU_GET "P_MESSYU_10MIN", d2, irc
'    If irc = False Then
'        ic = irc
'        GoTo JUMP
'    End If
'    If d2 > d1 Then
'        n = DateDiff("h", d1, d2) + 1
'        If n > 100 Then                     '100�͓K���Ɍ��߂��l�A�v�����if���Ɉ���������Ȃ��悤�ɂ����B2002/08/07 in YOKOHAMA
'            ans = MsgBox("�ǉ��Ŏ擾���悤�Ƃ��Ă���C�ے����[�_�f�[�^�X�e�b�v���Q�S������" & vbCrLf & _
'                         "�Ԋu������܂��B��Ƃ��p�����܂����H" & vbCrLf & _
'                         "�V�K�̍^���v�Z�ł͂��߂邱�Ƃ����i�߂��܂��B" & vbCrLf & _
'                         "[�͂�]�ł��̃W���u�͏I�����܂��A[������]�Ōp�����܂��B", vbInformation + vbYesNo)
'            If ans = vbYes Then
'                Close
'                End
'            End If
'        End If
'        d1 = DateAdd("n", 10, d1)
'        ORA_LOG "�C�ے����щJ�ʃf�[�^��荞�݊J�n " & d1 & " ���� " & d2 & "�܂�"
'        '******************************************************
'        'Ver1.0.0 �C���J�n 2015/08/05 O.OKADA �y07-01�z
'        '��OracleDB.Check_P_MESSYU_10MIN_Time()���C�����邱�ƁB�y07-01-01�z
'        '******************************************************
'        ORA_P_MESSYU_10MIN d1, d2, ic
'        '******************************************************
'        'Ver1.0.0 �C���I�� 2015/08/05 O.OKADA�y07-01�z
'        '******************************************************
'        If Not ic Then
'            ORA_LOG "�I���N���f�[�^�x�[�X���C�ے����[�_���уf�[�^���擾���悤�Ƃ�������" & vbCrLf & _
'                    "�G���[���������Ă��܂��B"
'            ORA_LOG "�C�ے����[�_�[���уf�[�^�G���[�����������݊J�n " & Format(d2, "yyyy/mm/dd h:nn")
'            nf = FreeFile
'            Open App.Path & "\data\P_MESSYU_10MIN.DAT" For Output As #nf
'            Print #nf, Format(d2, "yyyy/mm/dd h:nn")
'            Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
'            Close #nf
'            ORA_LOG "�C�ے����[�_�[���уf�[�^�G���[�����������ݏI��"
'            GoTo JUMP
'        Else
'            ORA_LOG "�C�ے����[�_���уf�[�^��荞�ݐ���I��"
'        End If
'    End If
'JUMP:
'End Sub

'******************************************************************************
'�T�u���[�`���FCommand1_Click()
'�����T�v�F
'�f�[�^�x�[�X�����C���e�i���X����B
'******************************************************************************
Private Sub Command1_Click()
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'Me.Timer1.Enabled = False
    'CompactMDB
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
End Sub

'******************************************************************************
'�T�u���[�`���FCommand2_Click()
'�����T�v�F
'�g�p���Ă��Ȃ��B
'******************************************************************************
'Private Sub Command2_Click()
'    Timer1.Enabled = False
'    Load �\�񕶃e�X�g���M
'    �\�񕶃e�X�g���M.Show
'End Sub

'******************************************************************************
'�T�u���[�`���FCommand3_Click()
'�����T�v�F
'******************************************************************************
Private Sub Command3_Click()
    Timer1.Enabled = False
    ORA_DataBase_Close
    MsgBox "�I���܂���"
    Close
    End
End Sub

'******************************************************************************
'�T�u���[�`���FForm_Click()
'�����T�v�F
'******************************************************************************
Private Sub Form_Click()
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'Dim ic As Boolean
    'ORA_DataBase_Connection ic
    'Dim d1 As Date
    'Dim d2 As Date
    'Dim ic As Boolean
    'd1 = "2006/04/02 01:20"
    'd2 = "2006/04/02 01:20"
    'ORA_DataBase_Connection ic
    'ORA_P_WATER d1, d2, ic
    'ORA_DataBase_Close
    'Check_OWARI_PUMP ic                    '�����y�؃|���v�f�[�^
    'Check_ORA_OWARI_WATER ic               '���P�[�u�����ʃf�[�^
    'Dim Name    As String                  '�|���v����
    'Dim Code    As Long                    '�|���v���R�[�h
    'Dim sv      As Long                    'sv�ԍ�
    'Dim N_P     As Long                    '�|���v��
    'Dim np      As Long                    '�|���v��̒ʂ��ԍ�
    'Dim d1      As Date
    'Dim d2      As Date
    'Dim n       As Long
    'd1 = "2005/05/21 01:00"
    'd2 = "2005/05/21 02:00"
    'n = DateDiff("n", d1, d2) / 10 + 1     '10���f�[�^�̌�
    'ReDim Pump(17, n)  '17=�|���v�ꐔ  n=�����X�e�b�v��
    'Name = "�y���|���v��"
    'Code = 2605
    'sv = 1
    'N_P = 4
    'np = 12
    'Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
End Sub

'******************************************************************************
'�T�u���[�`���FForm_Load()
'�����T�v�F
'******************************************************************************
Private Sub Form_Load()
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'Me.Timer1.Enabled = True
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    If App.PrevInstance Then
        MsgBox "���̃v���O�����͂��łɋN������Ă��܂��B"
        End
    End If
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'Me.Left = (Screen.Width - Me.Width) * 0.5
    'Me.Top = (Screen.Height - Me.Height) * 0.3
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    Me.Left = Screen.Width - Me.Width
    Me.Top = Screen.Height - Me.Height
    Dim i    As Integer
    Dim ic   As Boolean
    Dim nf   As Integer
    Dim buf  As String
    Dim dw   As Date
    Dim d1   As Date
    Dim d2   As Date
    Dim Rtry As Long
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    '2002/05/20 11:28 Frics Data center �ɂďC��
    'nf = FreeFile
    'Open App.Path & "\DBpath.dat" For Input As #nf
    'Input #nf, MDB_Path
    'Close #nf
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    LOG_File = "\data\LOGFILE" & Format(Now, "yyyy-mm-dd-hh-nn") & ".DAT"
    LOG_N = FreeFile
    Open App.Path & LOG_File For Output As #LOG_N
    Mesh_2km_to_1km_data
    ���[�_�[�e�B�[�Z���ǂݍ���
    jobg = False                            '��荞�݉\���
    FRICS_CVT_DATA                          '2�����b�V���f�[�^��315����ɐU�蕪����f�[�^��ǂ�
    Bit_Intial
    M_Link_Read
    ic = True
    Rtry = 0
ret_MDB:
    MDB_Connection ic
    Pump_Inital
    MDB_�ŐV����
    If Not ic Then
        ORA_LOG " MDB Connetion ���g���C��"
        Rtry = Rtry + 1
        If Rtry > 10 Then
            MsgBox "���[�J���c�a�ɐڑ��ł��܂���ł����B" & vbCrLf & _
                   "�W���u���I�����܂��B"
            End
        End If
        Short_Break 10
        GoTo ret_MDB
    End If
    ic = True
    Rtry = 0
ret_Ora:
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'ORA_DataBase_Connection ic
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
    If Not ic Then
       ORA_LOG " OraDB Connetion ���g���C��"
        Rtry = Rtry + 1
        If Rtry > 20 Then
            MsgBox "���m���I���N���c�a�ɐڑ��ł��܂���ł����B" & vbCrLf & _
                   "�W���u���I�����܂��B"
            End
        End If
        Short_Break 20
        GoTo ret_Ora
    End If
    'Ver0.0.0 �C���J�n 1900/01/01 00:00
    'ORA_KANSOKU_JIKOKU buf, dw
    'd1 = "2002/05/23 12:10"
    'd2 = "2002/05/25 12:10"
    'ORA_P_EATER d1, d2, ic
    'Ver0.0.0 �C���I�� 1900/01/01 00:00
End Sub

'******************************************************************************
'�T�u���[�`���FTimer1_Timer()
'�Ɩ��T�v�F
'******************************************************************************
Private Sub Timer1_Timer()
    Dim sec   As Integer
    Dim ic    As Boolean
    Dim rc    As Boolean
    Dim nf    As Integer
    Dim n     As Long
    Dim buf   As String
    Dim d1    As Date
    Dim d2    As Date
    Dim ans   As Long
    Dim ret   As Long
    Dim Rrun  As Boolean
    DoEvents
    '******************************************************
    '���[�J��DB�����k����B
    '******************************************************
    If (Day(Now) Mod 10) = 0 And Hour(Now) = 0 And Minute(Now) = 0 And Second(Now) < 2 Then
        Me.Timer1.Enabled = False
        Pre_Compact rc, Rrun
        If rc Then
            CompactMDB
        End If
        If Rrun Then ret = Shell("D:\SHINKAWA\���[�_�[�^���\��\RSHINKAWA.EXE", 1)
        Me.Timer1.Enabled = True
    End If
    '******************************************************
    '���̑��̏���
    '******************************************************
    sec = Second(Now)
    If (sec = 0 Or sec = 30) And Not jobg Then
        Short_Break 1
        jobg = True
        Me.Command1.Enabled = False
        Me.Command2.Enabled = False
        Me.Command3.Enabled = False
        ORA_LOG "���m���͐���V�X�e���f�[�^�x�[�X���u�Ɛڑ��J�n"
        ORA_DataBase_Connection ic          '���m�����I���N���T�[�o�[�ƃZ�b�V�������J�n
        If ic Then
            ORA_LOG "���m���͐���V�X�e���f�[�^�x�[�X���u�Ɛڑ�����"
        Else
            ORA_LOG "���m���͐���V�X�e���f�[�^�x�[�X���u�Ɛڑ��ł��܂���ł����B"
            jobg = False
            Me.Command1.Enabled = True
            Me.Command2.Enabled = True
            Me.Command3.Enabled = True
            Exit Sub
        End If
        Me.Timer1.Enabled = False
        '**************************************************
        '�T�u���[�`�����R�[������B
        '**************************************************
        '**************************************************
        'Ver1.0.0 �C���J�n 2015/08/05 O.OKADA�y01-01-01�z�y03-01-01�z�y04-01-01�z�y06-01-01�z�y07-01-01�z�y09-01-01�z
        '**************************************************
        'Check_Araizeki_Time ic                  '�􉁃f�[�^
        'Check_ORA_OWARI_WATER ic                '���P�[�u�����ʃf�[�^
        'Check_Suii_Time ic                      '���ʃf�[�^
        'Check_P_MESSYU_10MIN_Time ic            '�C�ے��J�ʎ��уf�[�^
        'Check_P_RADAR_Time ic                   'FRICS���[�_�[����
        'Check_F_RADAR_Time ic                   'FRICS���[�_�[�\��
        
        'Check_Araizeki_Time ic                  '�􉁃f�[�^
'        Check_ORA_OWARI_WATER ic                '���P�[�u�����ʃf�[�^
        Check_Suii_Time ic                      '���ʃf�[�^
        'Check_P_MESSYU_10MIN_Time ic            '�C�ے��J�ʎ��уf�[�^
        'Check_P_RADAR_Time ic                   'FRICS���[�_�[����
        'Check_F_RADAR_Time ic                   'FRICS���[�_�[�\��
        '**************************************************
        'Ver1.0.0 �C���I�� 2015/08/05 O.OKADA�y01-01-01�z�y03-01-01�z�y04-01-01�z�y06-01-01�z�y07-01-01�z�y09-01-01�z
        '**************************************************
        
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        'Check_F_MESSYU_10MIN_2_Time ic         '�C�ے��������\���J�ʃf�[�^
        'Check_F_MESSYU_10MIN_1_Time ic         '�C�ے���10���\���J�ʃf�[�^
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        
'        Check_OWARI_PUMP ic                     '�����y�؃|���v�f�[�^
'        ORA_NEW_DATA_TIME
        
        'Ver0.0.0 �C���J�n 1900/01/01 00:00
        '�C�ے�1km���b�V���J�ʒǉ� 2007/05/02
        Check_1kmMesh_Time "VDXA70", ic         '�C�ے����у��[�_�J�ʃf�[�^
        Check_1kmMesh_Time "VCXB70", ic         '�C�ے��~�J�Z���ԃ��[�_�J�ʃf�[�^(1-3)
        Check_1kmMesh_Time "VCXB71", ic         '�C�ے��~�J�Z���ԃ��[�_�J�ʃf�[�^(4-5)
        Check_1kmMesh_Time "VCXB75", ic         '�C�ے��~�J�Z���ԃ��[�_�J�ʃf�[�^30(1-3)
        Check_1kmMesh_Time "VCXB76", ic         '�C�ے��~�J�Z���ԃ��[�_�J�ʃf�[�^30(4-5)
        Check_1kmMesh_Time "VDXB70", ic         '�C�ے��i�E�L���X�g�f�[�^
        'Ver0.0.0 �C���I�� 1900/01/01 00:00
        Me.Timer1.Enabled = True
JUMP:
        OracleDB.Label3 = "��荞�ݑҋ@��"
        OracleDB.Label3.Refresh
        jobg = False
        ORA_DataBase_Close
        ORA_LOG "���m���͐���V�X�e���f�[�^�x�[�X���u�Ɛڑ�����"
        Me.Command1.Enabled = True
        Me.Command2.Enabled = True
        Me.Command3.Enabled = True
    End If
End Sub

'******************************************************************************
'�T�u���[�`���FTimer2_Timer()
'�����T�v�F
'******************************************************************************
Private Sub Timer2_Timer()
    Label4 = Format(Now, "yyyy�Nmm��dd�� hh��nn��ss�b")
End Sub
