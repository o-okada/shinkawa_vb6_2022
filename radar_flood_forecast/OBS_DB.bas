Attribute VB_Name = "OBS_DB"
Option Explicit
Global OBS_Con As New ADODB.Connection
Global OBS_Rst As New ADODB.Recordset
Global OBS_DB  As Boolean

Sub OBS_DB_Connection()

    Dim Con_str As String

    On Error GoTo ER1

    Con_str = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=D:\SHINKAWA\OracleTest\oraDB\Data\����.mdb"
'    Con_str = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source= " & App.Path & "\data\����.mdb"
    OBS_Con.ConnectionString = Con_str
    OBS_Con.Open

    Set OBS_Rst.ActiveConnection = OBS_Con
    OBS_DB = True
    On Error GoTo 0

'    �\���l�i�[

    Exit Sub
ER1:
    MsgBox "����.MDB�ɐڑ��ł��܂���B"
    OBS_DB = False
    On Error GoTo 0

End Sub
