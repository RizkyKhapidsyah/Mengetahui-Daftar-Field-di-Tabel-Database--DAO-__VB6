VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mengetahui Daftar Field di Tabel Database (DAO)"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   1680
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function FieldNames(dbPath As String, _
TableName As String) As Collection
'Input:
'dbPath: Path lengkap file database MS Access
'TableName: Nama tabel di dalam database
Dim oCol As Collection
Dim db As DAO.Database
Dim oTD As DAO.TableDef
Dim lCount As Long, lCtr As Long
Dim f As DAO.Field
On Error GoTo errorhandler
Set db = Workspaces(0).OpenDatabase(dbPath)
Set oTD = db.TableDefs(TableName)
Set oCol = New Collection

With oTD
    lCount = .Fields.Count
      For lCtr = 0 To lCount - 1
        oCol.Add .Fields(lCtr).Name
        List1.AddItem .Fields(lCtr).Name
    Next
End With
    MsgBox FieldNames
    db.Close
    Set FieldNames = oCol
Exit Function
errorhandler:
    On Error Resume Next
    If Not db Is Nothing Then db.Close
    Set FieldNames = Nothing
    Exit Function
End Function

Private Sub Command1_Click()
   Call FieldNames(App.Path & "\mahasiswa.mdb", _
                   "Mahasiswa")
End Sub


