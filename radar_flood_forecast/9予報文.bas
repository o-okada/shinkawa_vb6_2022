Attribute VB_Name = "�\�񕶍쐬"
Option Explicit
Option Base 1
Public Pattan_Now    As Long       '���Y�p�^�[��
Public Message()     As Pattan     '����
Public Wng_Last_Time As Long       '�������ӁE�x���ԍ�
Public NPat          As Long       '�����̎�ސ�
Type Pattan
    Patn(16)   As Variant
       ' 1=�p�^�[��                2=�@�\�񕶂̎��
       ' 3=�A���o��                4=�B�啶
       ' 5=�C���ӁE�x�����        6=�I�^���\���̐��ʏ�
       ' 7=�K�^���\���̐��ʏ󋵂Q  8=�M�^���\���̐��ʊ댯�x���x��
       ' 9=�B�啶�\���p           16=�\�񕶎�ʔԍ��R�[�h
End Type
Public ICat          As Long       '�I�������p�^�[���ԍ�
Public Msgz          As Long       '�p����啶�ԍ�
Public Const CYUBN_1 = "�@�@����̏o���́A����3�N9���̑䕗17�E18���ɕC�G����K��" & _
                       "�ƌ����܂�܂��B"

Public Const CYUBN_2 = "�@�@����̏o���́A����3�N9���̑䕗17�E18��������K�͂�" & _
                       "�����܂�܂��B"

Public Const CYUBN_3 = "�@�@����̏o���́A����12�N9���̓��C���J�ɕC�G����K�͂�" & _
                       "�����܂�܂��B"

Public Add_Main_Message(20)    As String
Public �啶1                   As String
Public �啶2                   As String

Public ����                    As six_time
Type six_time
     hm3     As Single '3���ԑO����
     hm2     As Single '2���ԑO����
     hm1     As Single '1���ԑO����
     h       As Single '���ݎ�������
     hy1     As Single '1���Ԍ㐅��
     hy2     As Single '2���Ԍ㐅��
     hy3     As Single '3���Ԍ㐅��
End Type
Sub Cat_Num_Read()

    Dim nf   As Long

    On Error GoTo jump

    nf = FreeFile
    Open App.Path & "\data\Cat.dat" For Input As #nf

    Print #nf, ICat

    Close #nf

    Exit Sub

jump:
    ICat = 1
    On Error GoTo 0

End Sub
Sub Cat_Num_Write()

    Dim nf   As Long

    nf = FreeFile
    Open App.Path & "\data\Cat.dat" For Output As #nf

    Print #nf, Msgz

    Close #nf

End Sub
Sub Disp_Msg(i As Long)

    If i > NPat Then
        i = 1
    End If
    If i < 1 Then
        i = NPat
    End If

'    With �\�񕶑��M
'        .Label7 = Message(i).Patn(2)
'        .Label8 = Message(i).Patn(1)
'        .Text5.Text = Message(i).Patn(9)
'    End With

    Msgz = i

End Sub
Sub Pattan_Add_Lf(w As Variant, ww As Variant, x As Long)

    Dim i    As Long
    Dim L    As Long
    Dim c    As String
    Dim cc   As String
    Dim LF

    If x = 1 Then
        LF = vbCrLf
    Else
        LF = vbLf
    End If

    ww = ""
    cc = ""
    L = Len(w)
    For i = 1 To L
        c = Mid(w, i, 1)
        If c = "%" Or c = LF Then
            cc = cc & LF & "�@�@"
        Else
            cc = cc & c
        End If
    Next i

    ww = cc

End Sub
Sub �p�^�[�����W_Read()

    Dim i     As Long
    Dim j     As Long
    Dim nf    As Long
    Dim buf   As String
    Dim m     As String
    Dim w
    Dim ww
    Dim p     As Long

    LOG_Out "IN    �p�^�[�����W_Read"

    nf = FreeFile
    Open App.Path & "\Data\�p�^�[��MK.txt" For Input As #nf

'�����̐��𒲂ׂ�
    i = 0
    Do
        Line Input #nf, buf
        i = i + 1
    Loop Until EOF(nf)
    Close #nf
    NPat = i - 2
    ReDim Message(NPat)

    Open App.Path & "\Data\�p�^�[��MK.txt" For Input As #nf
    Line Input #nf, buf
    Line Input #nf, buf
    For i = 1 To NPat
        Line Input #nf, buf
        w = Split(buf, vbTab)
        ww = "�@�@�i�啶�j%�V��̐��{�s�����O���ʐ��ʊϑ����ł́A%" & w(4)
        Pattan_Add_Lf ww, w(4), 0
        Pattan_Add_Lf ww, w(9), 1
        p = CLng(w(1))
        For j = 1 To 16
'            If j = 4 Or j = 9 Then
                Message(p).Patn(j) = w(j)           '�\�񕶂ɂ��������
'            Else
'                Message(p).Patn(j) = Trim(w(j))     '�\�񕶂ɂ��������
'            End If
        Next j
        m = "�V��@" & Message(p).Patn(3)
        Message(p).Patn(3) = m
        ' 1=�p�^�[���ԍ�            2=�@�\�񕶂̎��
        ' 3=�A���o��                4=�B�啶
        ' 5=�C���ӁE�x�����        6=�I�^���\���̐��ʏ�
        ' 7=�K�^���\���̐��ʏ󋵂Q  8=�M�^���\���̐��ʊ댯�x���x��
        '16=�\�񕶎�ʔԍ��R�[�h
    Next i
    Close #nf

    LOG_Out "OUT   �p�^�[�����W_Read"

End Sub
'
'
' 1�@�@�V��̐��ʂ͂P�P���P�U���S�O�����݁A���̂Ƃ���ƂȂ��Ă��܂��B
' 2�@�@�����O���ʊϑ����m�V�쒬�厚�����n���n�ŁA�R�D�W�O��------�ȍ~�ɒǉ��i���x��2���߁j
' 3�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@-----�ȍ~�ɒǉ��i�P���ԂɂS�Ocm�̑����ŏ㏸���j
' 4�@�@�I�ǉ� + �P�P���P�X���S�O�����ɂ́A---�ȍ~�ɒǉ� �͂񗔊댯���ʂɒB�����
' 5�@�@�����܂�܂��B
' 6�@�@�����O���ʊϑ����m�V�쒬�厚�����n���n�ŁA�U�D�Q�O���@---�ȍ~�ɒǉ��i���ʊ댯�x���x���S�j
' 7�@�y�Q�l�z
' 8�@�@�����O���ʊϑ����m���{�s�V�쒬�厚�����n���n
' 9�@�@�󂯎������
'10�@�@���E�݂Ƃ��A�����앪��_�i�n�n�s�������j����C�i�{�{�s�H�H���j�܂�
'11�@�@���ʊ댯�x���x��
'12�@�@�����x���P�@���h�c�ҋ@���ʒ���   �F�Q�D�O���`�R�D�O��
'13�@�@�����x���Q�@�͂񗔒��Ӑ��ʒ���   �F�R�D�O���`�S�D�S��
'14�@�@�����x���R�@���f���ʒ��� �@�@�F�S�D�S���`�T�D�Q��
'15�@�@�����x���S�@�͂񗔊댯���ʒ���   �F�T�D�Q���`���D����
'16�@�@�����x���T�@�͂񗔂̔���
'17�@�y�₢���킹��z
'18�@�@���ʊ֌W�@�@�@���m���������ݎ�����  �ێ��Ǘ���   �d�b  �O�T�Q�|�X�U�P�|�S�S�Q�P
'
'
'�쐬����镶����
'
'�����V�쐅���O���ʊϑ����i���{�s�V�쒬�厚�����n���j�ł́A
'�����V�삪����ɑ������A�Q���Ԍ�ɂ́A�͂񗔊댯���ʂɓ��B���錩���݂ł��B�s���ɂ����Ĕ��ׂ��Ɣ��f�����ꍇ������܂��̂ŁAOO�sOO�n�悩��{�{�n��ł́A�s������̔����ɒ��ӂ��ĉ������B
'���y���ӁE�x�����z
'��������̏o���́A�����P�Q�N�X���̓��C���J�ɕC�G����K�͂ƌ����܂�܂��B
'�����܂��A�z���̋��ꂪ����܂��̂Ō��d�Ȍx�����K�v�ł��B
'���y�����E�\�z�z
'�����V��㗬��̗��敽�ωJ��
'�����P�P���P�O���S�O������P�P���P�U���S�O���܂ł̂U���Ԃ̌������P�O�O�~��
'�����P�P���P�U���S�O������P�P���P�X���S�O���܂ł̂R���Ԃ̗\�z���W�O�~��
'�����V��̐��ʂ͂P�P���P�U���S�O�����݁A���̂Ƃ���ƂȂ��Ă��܂��B
'���������O���ʊϑ����m�V�쒬�厚�����n���n�ŁA�R�D�W�O�����i���x��2���߁j
'�����������������������������������i�P���ԂɂS�Ocm�̑����ŏ㏸���j
'�����V��̐��ʂ́A�㏸�X���ɂ���A�P�P���P�X���S�O�����ɂ́A�͂񗔊댯���ʂɒB����ƌ����܂�܂��B
'�����V��̐��ʂͤ�㏸�X���ɂ���03��14��36�����ɂͤ�V��̐��ʂͤ�㏸�X���ɂ���
'���������O���ʊϑ����m�V�쒬�厚�����n���n�ŁA�U�D�Q�O�����i���ʊ댯�x���x���S�j
'���y�Q�l�z
'���������O���ʊϑ����m���{�s�V�쒬�厚�����n���n
'�����󂯎������
'�������E�݂Ƃ��A�����앪��_�i�n�n�s�������j����C�i�{�{�s�H�H���j�܂�
'���������x���P�����h�c�ҋ@���ʒ��߁������F�Q�D�Om�`�R�D�O��
'���������x���Q���͂񗔒��Ӑ��ʒ��߁������F�R�D�O���`�S�D�S��
'���������x���R�����f���ʒ��߁��������F�S�D�S���`�T�D�Q��
'���������x���S���͂񗔊댯���ʒ��߁������F�T�D�Q���`���D����
'���������x���T���͂񗔂̔���
'���y�₢���킹��z
'�����ʊ֌W���������m���������ݎ��������ێ��Ǘ��ہ��d�b���O�T�Q�|�X�U�P�|�S�S�Q�P
'
'
'
Sub �啶�쐬1()

    Dim m     As String
    Dim m1    As String  '�y�啶�z
    Dim m2    As String  '�y���ӁE�x�����z
    Dim m3    As String  '�y�����E�\���z
    Dim m4    As String
    Dim m5    As String
    Dim m6    As String
    Dim mw    As String
    Dim dw    As Date
    Dim h0    As Single
    Dim dh    As Single

    h0 = ����.h
    dh = (h0 - ����.hm2) * 100#

    �啶1 = ""
    �啶2 = ""

    m1 = Message(Pattan_Now).Patn(4)

    ���ӌx����� m2

    m3 = vbLf & "�@�@�i�����E�\�z�j"

    �啶1 = m1 & m2 & m3


    m3 = "�@�@�V��̐��ʂ�" & Format(jgd, "dd��hh��nn��")
    m3 = m3 & "���݁A���̂Ƃ���ƂȂ��Ă��܂��B"
    m3 = m3 & vbLf & "�@�@�����O���ʊϑ����m���{�s�n�ŁA"
    m3 = m3 & Format(h0, "#0.00") & "m"
    ���ʃ��x��_Check ����.h, mw
    m3 = m3 & mw
    ���ʕϓ�_Check mw, dh
    If InStr(mw, "��") = 0 Then
        m3 = m3 & vbLf & "�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�i�P���Ԃ�"
        m3 = m3 & Format(Abs(dh * 0.5), "##0") & "cm�̑�����"
    Else
        m3 = m3 & vbLf & "�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@"
    End If
    m3 = m3 & mw
    If Pattan_Now <> 4 Then
        If Pattan_Now <> 14 Then
            m3 = m3 & vbLf & "�@�@" & Message(Pattan_Now).Patn(11)
        End If
        If Pattan_Now = 1 Then
            dw = DateAdd("h", 2, jgd)
            mw = Format(dw, "dd��hh��nn�����ɂ́A")
        Else
            dw = DateAdd("h", 3, jgd)
            mw = Format(dw, "dd��hh��nn�����ɂ́A")
        End If
        If Pattan_Now <> 14 Then
            m3 = m3 & mw & vbLf & "�@�@" & Message(Pattan_Now).Patn(13)
            m3 = m3 & "�ƌ����܂�܂��B"
            m4 = vbLf & "�@�@����삪���ʊϑ����m���{�s�n�ŁA"
        End If
        If Pattan_Now = 1 Then
            mw = Format(����.hy2, "#0.0") & "0m"
            ���ʃ��x��_Check ����.hy2 + 0.04, m
        Else
            mw = Format(����.hy3 + 0.04, "#0.0") & "0m"
            ���ʃ��x��_Check ����.hy3 + 0.04, m
        End If
        m4 = m4 & mw & m
    Else
        m4 = m3
        m3 = ""
    End If
    If Pattan_Now = 14 Then
        m4 = ""
    End If

    m5 = vbLf & "�@�@�@�y�Q�l�z"
    m5 = m5 & vbLf & "�@�@�@�����O���ʊϑ����m���{�s�����n"
    m5 = m5 & vbLf & "�@�@�@�󂯎������"
    m5 = m5 & vbLf & "�@�@�@���E�݂Ƃ��A�����앪��_����C�܂�"
    m5 = m5 & vbLf & "�@�@�@���ʊ댯�x���x��"
    m5 = m5 & vbLf & "�@�@�@�����x���P�@���h�c�ҋ@���ʒ��߁@�@�F�Q�D�O���`�R�D�O��"
    m5 = m5 & vbLf & "�@�@�@�����x���Q�@�͂񗔒��Ӑ��ʒ��߁@�@�F�R�D�O���`�S�D�S��"
    m5 = m5 & vbLf & "�@�@�@�����x���R�@���f���ʒ��߁@�@�@�F�S�D�S���`�T�D�Q��"
    m5 = m5 & vbLf & "�@�@�@�����x���S�@�͂񗔊댯���ʒ��߁@�@�F�T�D�Q���`���D����"
    m5 = m5 & vbLf & "�@�@�@�����x���T�@�͂񗔂̔���"
    m5 = m5 & vbLf & "�@�@�@�k�₢���킹��l"
    m5 = m5 & vbLf & "�@�@���ʊ֌W�F���m���@�������ݎ������@�@�ێ��Ǘ��ہ@�d�b 052-961-4421"

    �啶2 = m3 & m4 & m5

Debug.Print "�@�@���o��(" & Message(Pattan_Now).Patn(3) & ")"
Debug.Print " "
Debug.Print �啶1
Debug.Print �啶2










End Sub
'
'
' 1�@�@�V��̐��ʂ͂P�P���P�U���S�O�����݁A���̂Ƃ���ƂȂ��Ă��܂��B
' 2�@�@�����O���ʊϑ����m�V�쒬�厚�����n���n�ŁA�R�D�W�O��------�ȍ~�ɒǉ��i���x��2���߁j
' 3�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@-----�ȍ~�ɒǉ��i�P���ԂɂS�Ocm�̑����ŏ㏸���j
' 4�@�@�I�ǉ� + �P�P���P�X���S�O�����ɂ́A---�ȍ~�ɒǉ� �͂񗔊댯���ʂɒB�����
' 5�@�@�����܂�܂��B
' 6�@�@�����O���ʊϑ����m�V�쒬�厚�����n���n�ŁA�U�D�Q�O���@---�ȍ~�ɒǉ��i���ʊ댯�x���x���S�j
' 7�@�y�Q�l�z
' 8�@�@�����O���ʊϑ����m���{�s�V�쒬�厚�����n���n
' 9�@�@�󂯎������
'10�@�@���E�݂Ƃ��A�����앪��_�i�n�n�s�������j����C�i�{�{�s�H�H���j�܂�
'11�@�@���ʊ댯�x���x��
'12�@�@�����x���P�@���h�c�ҋ@���ʒ���   �F�Q�D�O���`�R�D�O��
'13�@�@�����x���Q�@�͂񗔒��Ӑ��ʒ���   �F�R�D�O���`�S�D�S��
'14�@�@�����x���R�@���f���ʒ��� �@�@�F�S�D�S���`�T�D�Q��
'15�@�@�����x���S�@�͂񗔊댯���ʒ���   �F�T�D�Q���`���D����
'16�@�@�����x���T�@�͂񗔂̔���
'17�@�y�₢���킹��z
'18�@�@���ʊ֌W�@�@�@���m���������ݎ�����  �ێ��Ǘ���   �d�b  �O�T�Q�|�X�U�P�|�S�S�Q�P
'
'
'�쐬����镶����
'

'�@�@�͂񗔒��Ӑ��ʂɓ��B�A���ʂ͂���ɏ㏸���邨����
'�啶1
'�@�@�i�啶�j
'�@�@�V��̐��{�s�����O���ʐ��ʊϑ����ł́A
'�@�@�͂񗔒��Ӑ���(���x��2�j�ɒB���܂����B���ʂ͂���ɏ㏸���錩���݂ł��B
'�@�@����̍^���\��ɒ��ӂ��ĉ������B
'
'�@�@�i�����E�\�z�j
'�啶2
'�@�@�V��̐����O���ʐ��ʊϑ����k���{�s�l�̐���
'�@�@19��05��50���̌����@3.36m�i�}�㏸���j�i���ʊ댯�x���x���Q�j
'�@�@19��07��50���̗\�z�@4.50m�i���ʊ댯�x���x���R�j
'�@�@�y�Q�l�z
'�@�@�@�����O���ʐ��ʊϑ����k���{�s�����l
'�@�@�@�͂񗔊댯���ʁ@5.20m�@�@�@�@�@�@���f���ʁ@�@4.40m
'�@�@�@�͂񗔒��Ӑ��ʁi�x�����ʁj3.00m�@���h�c�ҋ@���ʁ@2.00m
'
'�@�@���ʊ댯�x���x��
'�@�@�@�����x���T�@�͂񗔂̔���
'�@�@�@�����x���S�@�͂񗔊댯���ʒ���
'�@�@�@�����x���R�@���f���ʒ���
'�@�@�@�����x���Q�@�͂񗔒��Ӑ��ʒ���
'�@�@�@�����x���P�@���h�c�ҋ@���ʒ���
'
'
'
Sub �啶�쐬2()

    Dim m     As String
    Dim m1    As String  '�y�啶�z
    Dim m2    As String  '�y���ӁE�x�����z
    Dim m3    As String  '�y�����E�\���z
    Dim m4    As String
    Dim m5    As String
    Dim m6    As String
    Dim mw    As String
    Dim dw    As Date
    Dim h0    As Single
    Dim dh    As Single

    h0 = ����.h
    dh = (h0 - ����.hm2) * 100#

    �啶1 = ""
    �啶2 = ""

    m1 = Message(Pattan_Now).Patn(4)

    ���ӌx����� m2

    m3 = vbLf & "�@�@�i�����E�\�z�j"

    �啶1 = m1 & m2 & m3


    m3 = "�@�@�V��̐����O���ʐ��ʊϑ����k���{�s�l�̐���" & vbLf
    m3 = m3 & "�@�@" & Format(jgd, "dd��hh��nn��") & "�̌����@" & Format(h0, "#0.00") & "m"
    ���ʕϓ�_Check mw, dh
    m3 = m3 & mw
    ���ʃ��x��_Check ����.h, mw
    m3 = m3 & mw
    If Pattan_Now <> 4 Then
        If Pattan_Now <> 14 Then
            m3 = m3 & vbLf & "�@�@"   ' & Message(Pattan_Now).Patn(11)
        End If
        If Pattan_Now = 1 Then
            dw = DateAdd("h", 2, jgd)
            mw = Format(dw, "dd��hh��nn���̗\�z�@")
        Else
            dw = DateAdd("h", 3, jgd)
            mw = Format(dw, "dd��hh��nn���̗\�z�@")
        End If
        If Pattan_Now <> 14 Then
            m3 = m3 & mw
'            m3 = m3 & "�ƌ����܂�܂��B"
'            m4 = vbLf & "�@�@����삪���ʊϑ����m���{�s�n�ŁA"
        End If
        If Pattan_Now = 1 Then
            mw = Format(����.hy2, "#0.0") & "0m"
            ���ʃ��x��_Check ����.hy2 + 0.04, m
        Else
            mw = Format(����.hy3 + 0.04, "#0.0") & "0m"
            ���ʃ��x��_Check ����.hy3 + 0.04, m
        End If
        m4 = m4 & mw & m
    Else
        m4 = m3
        m3 = ""
    End If
    If Pattan_Now = 14 Then
        m4 = ""
    End If

    m5 = vbLf & "�@�@�y�Q�l�z"
    m5 = m5 & vbLf & "�@�@�@�����O���ʐ��ʊϑ����k���{�s�����l"
    m5 = m5 & vbLf & "�@�@�@�͂񗔊댯���ʁ@5.20m�@�@�@�@�@�@���f���ʁ@�@4.40m"
    m5 = m5 & vbLf & "�@�@�@�͂񗔒��Ӑ��ʁi�x�����ʁj3.00m�@���h�c�ҋ@���ʁ@2.00m"
    m5 = m5 & vbLf & "�@�@�@"
    m5 = m5 & vbLf & "�@�@���ʊ댯�x���x��"
    m5 = m5 & vbLf & "�@�@�@�����x���T�@�͂񗔂̔���"
    m5 = m5 & vbLf & "�@�@�@�����x���S�@�͂񗔊댯���ʒ���"
    m5 = m5 & vbLf & "�@�@�@�����x���R�@���f���ʒ���"
    m5 = m5 & vbLf & "�@�@�@�����x���Q�@�͂񗔒��Ӑ��ʒ���"
    m5 = m5 & vbLf & "�@�@�@�����x���P�@���h�c�ҋ@���ʒ���"
    m5 = m5 & vbLf & " "
    m5 = m5 & vbLf & "�@�k�₢���킹��l"
    m5 = m5 & vbLf & "�@�@���ʊ֌W�F���m���@�������ݎ������@�@�ێ��Ǘ��ہ@�d�b 052-961-4421"


    �啶2 = m3 & m4 & m5

Debug.Print "�@�@���o��(" & Message(Pattan_Now).Patn(3) & ")"
Debug.Print " "
Debug.Print �啶1
Debug.Print �啶2










End Sub

Sub ���ʃ��x��_Check(h As Single, m As String)


    If h < 2# Then
        m = ""
        Exit Sub
    End If

    Select Case h
        Case Is < 3#
            m = "�i���ʊ댯�x���x���P�j"

        Case Is < 4.4
            m = "�i���ʊ댯�x���x���Q�j"

        Case Is < 5.2
            m = "�i���ʊ댯�x���x���R�j"

        Case Is < 100#
            m = "�i���ʊ댯�x���x���S�j"

    End Select

End Sub
Sub ���ʕϓ�_Check(hg As String, dh As Single)

    hg = ""

    Select Case dh
        Case Is > 30#
            hg = "�i�}�㏸���j"
        Case Is > 10#
            hg = "�i�㏸���j"
        Case Is > -10#
            hg = "�i���݂̐��ʂ͉��΂��j"
        Case Is > -110#
            hg = "�i���~���j"
    End Select

End Sub
'
'H0   ���ݎ������ʐ���
'H1   1���Ԍ�\������
'H2   2���Ԍ�\������
'H3   3���Ԍ�\������
'
'm--���ӌx�����
'
'
Sub ���ӌx�����(CYUBN As String)

    Dim Wng    As Long
    Dim HV     As Single
    Dim h0     As Single
    Dim H1     As Single
    Dim H2     As Single
    Dim H3     As Single

    h0 = ����.h
    H1 = ����.hy1
    H2 = ����.hy2
    H3 = ����.hy3

    If Pattan_Now < 5 Or Pattan_Now > 13 Then
        CYUBN = vbLf
        Exit Sub
    End If

'���ӕ�
    HV = H3 - H2
    Select Case HV
        Case Is < 0.5
            Wng = 1
        Case Is < 1#
            Wng = 2
        Case Is >= 1#
            Wng = 3
    End Select
    If Wng_Last_Time > Wng Then Wng = Wng_Last_Time
    Wng_Last_Time = Wng
    Select Case Wng
        Case 1
            CYUBN = vbLf & "�@�@�i���ӎ����j" & vbLf & CYUBN_1
        Case 2
            CYUBN = vbLf & "�@�@�i���ӎ����j" & vbLf & CYUBN_2
        Case 3
            CYUBN = vbLf & "�@�@�i���ӎ����j" & vbLf & CYUBN_3
    End Select
    If h0 >= 6.2 Or H3 >= 6.2 Then '�v���h��(T.P 6.2m)�𒴂���
        CYUBN = CYUBN & vbLf & "�@�@�܂��A�z���̋��ꂪ����܂��̂Ō��d�Ȍx�����K�v�ł��B"
    End If

End Sub
'
'2008/03/03���݈ȉ���19�f�[�^��ǂ�
'
' 1�@�@�V��̐��ʂ͂P�P���P�U���S�O�����݁A���̂Ƃ���ƂȂ��Ă��܂��B
' 2�@�@�����O���ʊϑ����m�V�쒬�厚�����n���n�ŁA�R�D�W�O��------�ȍ~�ɒǉ��i���x��2���߁j
' 3�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@-----�ȍ~�ɒǉ��i�P���ԂɂS�Ocm�̑����ŏ㏸���j
' 4�@�@�I�ǉ� + �P�P���P�X���S�O�����ɂ́A---�ȍ~�ɒǉ� �͂񗔊댯���ʂɒB�����
' 5�@�@�����܂�܂��B
' 6�@�@�����O���ʊϑ����m�V�쒬�厚�����n���n�ŁA�U�D�Q�O���@---�ȍ~�ɒǉ��i���ʊ댯�x���x���S�j
' 7�@�y�Q�l�z
' 8�@�@�����O���ʊϑ����m���{�s�V�쒬�厚�����n���n
' 9�@�@�󂯎������
'10�@�@���E�݂Ƃ��A�����앪��_�i�n�n�s�������j����C�i�{�{�s�H�H���j�܂�
'11�@�@���ʊ댯�x���x��
'12�@�@�����x���P�@���h�c�ҋ@���ʒ���   �F�Q�D�O���`�R�D�O��
'13�@�@�����x���Q�@�͂񗔒��Ӑ��ʒ���   �F�R�D�O���`�S�D�S��
'14�@�@�����x���R�@���f���ʒ���     �F�S�D�S���`�T�D�Q��
'15�@�@�����x���S�@�͂񗔊댯���ʒ���   �F�T�D�Q���`���D����
'16�@�@�����x���T�@�͂񗔂̔���
'17�@�y�₢���킹��z
'18�@�@���ʊ֌W�@�@�@���m���������ݎ�����  �ێ��Ǘ���   �d�b  �O�T�Q�|�X�U�P�|�S�S�Q�P
'
'
'�쐬����镶����
'
'
'
'
'
Sub �ǉ��啶_Read()

    Dim i      As Long
    Dim nf     As Long
    Dim buf    As String
    Dim f      As String

    f = App.Path & "\data\�ǉ��啶.txt"

    nf = FreeFile
    Open f For Input As #nf
    For i = 1 To 18
        Line Input #nf, Add_Main_Message(i)
    Next i
        
    Close #nf

End Sub
