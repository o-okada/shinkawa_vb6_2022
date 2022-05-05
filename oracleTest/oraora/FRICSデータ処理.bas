Attribute VB_Name = "FRICSデータ処理"
Option Explicit
Option Base 1

Public Dim2_Num               As Long
Public Dim2_mesh_Name(10)     As String
Public Dim2_mesh_Number(10)   As Long
Public Dim2_To_315(10, 100)   As Cvt_Mesh

Type Cvt_Mesh
    Rn  As Long  '流域番号(1～315)
    Mn  As Long  '2次メッシュ上の順番号
End Type

Public Dim2_WHERE2            As String

'******************************************************************************
'サブルーチン：FRICS_CVT_DATA()
'処理概要：
'使用する2次メッシュ番号
'523607
'523606
'523770
'523677
'523676
'523760
'523667
'523666
'523656
'523646
'******************************************************************************
Sub FRICS_CVT_DATA()
    '******************************************************
    '変数セット処理
    '******************************************************
    ORA_LOG "IN   Sub FRICS_CVT_DATA"
    Dim nf   As Integer
    Dim i    As Long
    Dim j    As Long
    Dim m    As Long
    Dim n    As Long
    Dim buf  As String
    '******************************************************
    'ファイル入力処理
    '******************************************************
    nf = FreeFile
    Open App.Path & "\data\FRICS_RADAR.txt" For Input As #nf
    Line Input #nf, buf
    Dim2_Num = CLng(buf)
    For i = 1 To Dim2_Num
        Line Input #nf, buf
        Dim2_mesh_Name(i) = Mid(buf, 1, 6)
        Dim2_mesh_Number(i) = CLng(Mid(buf, 7, 4))
        For j = 1 To Dim2_mesh_Number(i)
            Line Input #nf, buf
            Dim2_To_315(i, j).Rn = CLng(Mid(buf, 1, 3))
            Dim2_To_315(i, j).Mn = CLng(Mid(buf, 4, 4))
        Next j
    Next i
    Close #nf
    '******************************************************
    '変数セット処理
    '******************************************************
    Dim2_WHERE2 = "((LATITUDE='" & Mid(Dim2_mesh_Name(1), 1, 2) & "' AND " & _
                 "LONGITUDE='" & Mid(Dim2_mesh_Name(1), 3, 2) & "' AND " & _
                      "CODE=" & Mid(Dim2_mesh_Name(1), 5, 2) & ") OR " & _
                 "(LATITUDE='" & Mid(Dim2_mesh_Name(2), 1, 2) & "' AND " & _
                 "LONGITUDE='" & Mid(Dim2_mesh_Name(2), 3, 2) & "' AND " & _
                      "CODE=" & Mid(Dim2_mesh_Name(2), 5, 2) & ") OR " & _
                 "(LATITUDE='" & Mid(Dim2_mesh_Name(3), 1, 2) & "' AND " & _
                 "LONGITUDE='" & Mid(Dim2_mesh_Name(3), 3, 2) & "' AND " & _
                      "CODE=" & Mid(Dim2_mesh_Name(3), 5, 2) & ") OR " & _
                 "(LATITUDE='" & Mid(Dim2_mesh_Name(4), 1, 2) & "' AND " & _
                 "LONGITUDE='" & Mid(Dim2_mesh_Name(4), 3, 2) & "' AND " & _
                      "CODE=" & Mid(Dim2_mesh_Name(4), 5, 2) & ") OR " & _
                 "(LATITUDE='" & Mid(Dim2_mesh_Name(5), 1, 2) & "' AND " & _
                 "LONGITUDE='" & Mid(Dim2_mesh_Name(5), 3, 2) & "' AND " & _
                      "CODE=" & Mid(Dim2_mesh_Name(5), 5, 2) & ") OR " & _
                 "(LATITUDE='" & Mid(Dim2_mesh_Name(6), 1, 2) & "' AND " & _
                 "LONGITUDE='" & Mid(Dim2_mesh_Name(6), 3, 2) & "' AND " & _
                      "CODE=" & Mid(Dim2_mesh_Name(6), 5, 2) & ") OR " & _
                 "(LATITUDE='" & Mid(Dim2_mesh_Name(7), 1, 2) & "' AND " & _
                 "LONGITUDE='" & Mid(Dim2_mesh_Name(7), 3, 2) & "' AND " & _
                      "CODE=" & Mid(Dim2_mesh_Name(7), 5, 2) & ") OR " & _
                 "(LATITUDE='" & Mid(Dim2_mesh_Name(8), 1, 2) & "' AND " & _
                 "LONGITUDE='" & Mid(Dim2_mesh_Name(8), 3, 2) & "' AND " & _
                      "CODE=" & Mid(Dim2_mesh_Name(8), 5, 2) & ") OR "
    Dim2_WHERE2 = Dim2_WHERE2 & _
                 "(LATITUDE='" & Mid(Dim2_mesh_Name(9), 1, 2) & "' AND " & _
                 "LONGITUDE='" & Mid(Dim2_mesh_Name(9), 3, 2) & "' AND " & _
                      "CODE=" & Mid(Dim2_mesh_Name(9), 5, 2) & ") OR " & _
                 "(LATITUDE='" & Mid(Dim2_mesh_Name(10), 1, 2) & "' AND " & _
                 "LONGITUDE='" & Mid(Dim2_mesh_Name(10), 3, 2) & "' AND " & _
                      "CODE=" & Mid(Dim2_mesh_Name(10), 5, 2) & ")) "
    '******************************************************
    '戻り値セット処理
    '******************************************************
    ORA_LOG "Sub FRICS_CVT_DATA Complete."
End Sub
