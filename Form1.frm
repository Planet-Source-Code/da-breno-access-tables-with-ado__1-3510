VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADO Connect to Access Databases"
   ClientHeight    =   8595
   ClientLeft      =   1815
   ClientTop       =   1380
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   8220
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   15769
            Text            =   "ADO Demo Program to copy Tables between Access Databases"
            TextSave        =   "ADO Demo Program to copy Tables between Access Databases"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "3:02 PM"
            Object.Tag             =   "Time Field"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "9/14/99"
            Object.Tag             =   "date"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11040
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "ADO"
      Height          =   7455
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   10455
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   600
         TabIndex        =   15
         Top             =   6360
         Width           =   3855
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   600
         TabIndex        =   14
         Text            =   "Enter the name of the Table to be copied to"
         Top             =   5880
         Width           =   3255
      End
      Begin VB.CommandButton Command7 
         Caption         =   "<-"
         Height          =   495
         Left            =   3120
         TabIndex        =   13
         Top             =   3000
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Caption         =   "->"
         Height          =   495
         Left            =   3120
         TabIndex        =   12
         Top             =   2280
         Width           =   495
      End
      Begin VB.ListBox List2 
         Height          =   2010
         ItemData        =   "Form1.frx":0000
         Left            =   3960
         List            =   "Form1.frx":0002
         TabIndex        =   11
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   600
         TabIndex        =   10
         Top             =   5040
         Width           =   3855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "..."
         Height          =   375
         Left            =   4440
         TabIndex        =   9
         Top             =   5040
         Width           =   375
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Text            =   "Select the database the table is to be copied to"
         Top             =   4560
         Width           =   3855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         Height          =   495
         Left            =   7440
         TabIndex        =   7
         Top             =   1920
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   375
         Left            =   4440
         TabIndex        =   6
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   600
         TabIndex        =   4
         Text            =   "Select the Database to copy from"
         Top             =   480
         Width           =   3855
      End
      Begin VB.ListBox List1 
         Height          =   2010
         ItemData        =   "Form1.frx":0004
         Left            =   600
         List            =   "Form1.frx":0006
         TabIndex        =   3
         Top             =   2160
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Copy Tables"
         Height          =   495
         Left            =   7440
         TabIndex        =   2
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   165
         Left            =   600
         TabIndex        =   1
         Text            =   "Select a Table from the selected database"
         Top             =   1680
         Width           =   3375
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

Dim cnn1 As ADODB.Connection
Dim cmdQuery As ADODB.Command
Dim strCnn As String
Dim Rs1 As ADODB.Recordset
Dim prm As ADODB.Parameter
   
Dim First, UserPath As String

'error handler
On Error GoTo ErrorHandler

sTable = List2.List(0)  ' the table we are going to copy

' ADO connection string that MS uses to to setup a link between
' the program and a JET database
' the string would be different for non JET DB's such as Oracle.
strCnn = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source="
strCnn = strCnn & sTmp

' create and open a connection to the databse
Set cnn1 = New ADODB.Connection
cnn1.Open strCnn

' declare working variables we'll need for executing SQL statements
Set cmdQuery = New ADODB.Command
Set Rs1 = New ADODB.Recordset

'active connection
Set cmdQuery.ActiveConnection = cnn1
  
' build the query up
SelectString = "SELECT * INTO"
FromString = "FROM"
FromSource = sTmp
NewTable = Text7.Text

'builds up query that will copy tables and data
SQLtext = SelectString & Space(1) & "[" & Destination & "]" _
            & "." & NewTable & Space(1) & FromString & Space(1) _
            & "[" & FromSource & "]" & "." & sTable


' assigns the SQL statement to the command object
cmdQuery.CommandText = SQLtext

' runs the SQL statement
Set Rs1 = cmdQuery.Execute()

If Err.Number = 0 Then
    MsgBox "The copy of Tables is complete. So There!!", vbOKOnly, "ADO Copy System Message"
End If


'close the connection to the database
cnn1.Close
Exit Sub



ErrorHandler:   ' Error-handling routine.
   Select Case Err.Number   ' Evaluate error number.
       
       Case Else
         
      Msg = "Unexpected error #" & Str(Err.Number)
      Msg = Msg & " occurred: " & Err.Description
      ' Display message box with Stop sign icon and
      ' OK button.
      MsgBox Msg, vbCritical
      
   End Select
   Resume Next  ' Resume execution at same line
            ' that caused the error.


End Sub

Private Sub Command3_Click()
End
   
End Sub

Private Sub Command4_Click()

' the database to copy from

' sets up a common dialog box that will only show Access .mdb files
  CommonDialog1.Filter = "DB Files" & _
  "(*.mdb)|*.mdb"
  
CommonDialog1.ShowOpen


  If Err = 32755 Then         ' if the user has cancelled
    Exit Sub
  Else
    sTmp = CommonDialog1.FileName    'get the filename selected by the user
    Text3.Text = sTmp
    
    Set dbList = Workspaces(0).OpenDatabase(sTmp)
    List1.Clear
    For Each tdFrom In dbList.TableDefs
        '  screen out all tables beginning with MSys as they are not needed
       If Mid(tdFrom.Name, 1, 4) <> "MSys" Then
       
            List1.AddItem tdFrom.Name   ' add the name of the table to the List
       
       End If
    Next
  End If


End Sub

Private Sub Command5_Click()

' the database to copy to

' sets up a common dialog box that will only show Access .mdb files
CommonDialog1.Filter = "DB Files" & _
  "(*.mdb)|*.mdb"
  
CommonDialog1.ShowOpen

  If Err = 32755 Then    ' if the user has cancelled
    Exit Sub
  Else
    Destination = CommonDialog1.FileName  'get the filename selected by the user
    Text5.Text = Destination   ' and display it
    
   End If
End Sub

Private Sub Command6_Click()

' writes the selected Table name from List1 to List2
List2.AddItem List1.Text

End Sub

Private Sub Command7_Click()

'takes out the Table name entry in List2
List2.RemoveItem (0)

End Sub



Private Sub List1_DblClick()
' writes the selected Table name from List1 to List2
List2.AddItem List1.Text

End Sub

