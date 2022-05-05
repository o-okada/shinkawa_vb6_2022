Attribute VB_Name = "OracleConst"
Option Explicit
''''''''''''''''''''''''''''
' Oracle Objects for OLE global constant file.
' This file can be loaded into a code module.
''''''''''''''''''''''''''''

'Editmode property values
' These are intended to match similar constants in the
' Visual Basic file CONSTANT.TXT
Global Const ORADATA_EDITNONE = 0
Global Const ORADATA_EDITMODE = 1
Global Const ORADATA_EDITADD = 2

'Parameter Types
Global Const ORAPARM_INPUT = 1
Global Const ORAPARM_OUTPUT = 2
Global Const ORAPARM_BOTH = 3

'Parameter Status
Global Const ORAPSTAT_INPUT = &H1&
Global Const ORAPSTAT_OUTPUT = &H2&
Global Const ORAPSTAT_AUTOENABLE = &H4&
Global Const ORAPSTAT_ENABLE = &H8&

'CreateDynaset Method Options
Global Const ORADYN_DEFAULT = &H0&
Global Const ORADYN_NO_AUTOBIND = &H1&
Global Const ORADYN_NO_BLANKSTRIP = &H2&
Global Const ORADYN_READONLY = &H4&
Global Const ORADYN_NOCACHE = &H8&
Global Const ORADYN_ORAMODE = &H10&
Global Const ORADYN_DBDEFAULT = &H20&
Global Const ORADYN_NO_MOVEFIRST = &H40&
Global Const ORADYN_DIRTY_WRITE = &H80&

'CreateSql Method options
Global Const ORASQL_DEFAULT = &H0&
Global Const ORASQL_NO_AUTOBIND = &H1&
Global Const ORASQL_FAILEXEC = &H2&

'OpenDatabase Method Options
Global Const ORADB_DEFAULT = &H0&
Global Const ORADB_ORAMODE = &H1&
Global Const ORADB_NOWAIT = &H2&
Global Const ORADB_DBDEFAULT = &H4&
Global Const ORADB_DEFERRED = &H8&
Global Const ORADB_NONBLK = &H8&

'Parameter Types (ServerType)
Global Const ORATYPE_VARCHAR2 = 1
Global Const ORATYPE_NUMBER = 2
Global Const ORATYPE_SINT = 3
Global Const ORATYPE_FLOAT = 4
Global Const ORATYPE_STRING = 5
Global Const ORATYPE_VARCHAR = 9
Global Const ORATYPE_DATE = 12
Global Const ORATYPE_UINT = 68
Global Const ORATYPE_CHAR = 96
Global Const ORATYPE_CHARZ = 97
Global Const ORATYPE_CURSOR = 102

'OIP errors returned as part of the OLE Automation error.
Global Const OERROR_ADVISEULINK = 4096  ' Invalid advisory connection
Global Const OERROR_POSITION = 4098 ' Invalid database position
Global Const OERROR_NOFIELDNAME = 4099  ' Field 'field-name' not found
Global Const OERROR_TRANSIP = 4101  ' Transaction already in process
Global Const OERROR_TRANSNIPC = 4104    ' Commit detected with no active transaction
Global Const OERROR_TRANSNIPR = 4105    ' Rollback detected with no active transaction
Global Const OERROR_NODSET = 4106   ' No such set attached to connection
Global Const OERROR_INVROWNUM = 4108    ' Invalid row reference
Global Const OERROR_TEMPFILE = 4109 ' Error creating temporary file
Global Const OERROR_DUPSESSION = 4110   ' Duplicate session name
Global Const OERROR_NOSESSION = 4111    ' Session not found during detach
Global Const OERROR_NOOBJECTN = 4112    ' No such object named 'object-name'
Global Const OERROR_DUPCONN = 4113  ' Duplicate connection name
Global Const OERROR_NOCONN = 4114   ' No such connection during detach
Global Const OERROR_BFINDEX = 4115  ' Invalid field index
Global Const OERROR_CURNREADY = 4116    ' Cursor not ready for I-O
Global Const OERROR_NOUPDATES = 4117    ' Not an updatable set
Global Const OERROR_NOTEDITING = 4118   ' Attempt to update without edit or add operation
Global Const OERROR_DATACHANGE = 4119   ' Data has been modified
Global Const OERROR_NOBUFMEM = 4120 ' No memory for data transfer buffers
Global Const OERROR_INVBKMRK = 4121 ' Invalid bookmark
Global Const OERROR_BNDVNOEN = 4122 ' Bind variable not fully enabled
Global Const OERROR_DUPPARAM = 4123 ' Duplicate parameter name
Global Const OERROR_INVARGVAL = 4124    ' Invalid argument value
Global Const OERROR_INVFLDTYPE = 4125   ' Invalid field type
Global Const OERROR_TRANSFORUP = 4127   ' For Update detected with no active transaction
Global Const OERROR_NOTUPFORUP = 4128   ' For Update detected but not updatable set

Global Const OERROR_TRANSLOCK = 4129    ' Commit/Rollback with SELECT FOR UPDATE in progress
Global Const OERROR_CACHEPARM = 4130    ' Invalid cache parameter
Global Const OERROR_FLDRQROWID = 4131   ' Field processing requires ROWID

'Booleans-These are defined as reserved
' words in Visual Basic
'Global Const True = -1
'Global Const False = 0

'Constants for Update Source
Global Const OFIELD_THISFIELD = 0
Global Const OFIELD_SYSDATE = 1
Global Const OFIELD_TIMESTAMP = 2
Global Const OFIELD_DATESTAMP = 3
Global Const OFIELD_FUNCTION = 4
Global Const OFIELD_OTHERFIELD = 5
Global Const OFIELD_STRLITERAL = 6

Global Const OFIELD_FOR_NONE = 0
Global Const OFIELD_FOR_INSERT = 10
Global Const OFIELD_FOR_UPDATE = 100
Global Const OFIELD_FOR_ALL = 110


