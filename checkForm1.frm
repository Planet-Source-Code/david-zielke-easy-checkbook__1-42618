VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Checking"
   ClientHeight    =   7200
   ClientLeft      =   765
   ClientTop       =   840
   ClientWidth     =   9615
   Icon            =   "checkForm1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9615
   Begin VB.CommandButton Command10 
      Caption         =   "&Find"
      Height          =   375
      Left            =   8160
      TabIndex        =   29
      Top             =   1200
      Width           =   972
   End
   Begin VB.TextBox Text1 
      DataField       =   "PrimaryKey"
      DataSource      =   "Data1"
      Height          =   372
      Left            =   120
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Calcu&tator"
      Height          =   375
      Left            =   8160
      TabIndex        =   27
      Top             =   720
      Width           =   972
   End
   Begin VB.CheckBox Check1 
      DataField       =   "Cashed"
      DataSource      =   "Data1"
      Height          =   252
      Left            =   8280
      TabIndex        =   15
      Top             =   2160
      Width           =   252
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "checkForm1.frx":000C
      Height          =   4332
      Left            =   720
      OleObjectBlob   =   "checkForm1.frx":0046
      TabIndex        =   16
      Top             =   2520
      Width           =   8292
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Update"
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   1200
      Width           =   1092
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   720
      Width           =   1092
   End
   Begin VB.CommandButton Command6 
      Caption         =   "De&lete"
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   720
      Width           =   1092
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   324
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6840
      Width           =   3615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   1200
      Width           =   1092
   End
   Begin VB.CommandButton Command4 
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   6600
      TabIndex        =   8
      Top             =   1200
      Width           =   1092
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Debit"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   1092
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      DataField       =   "Check Number"
      DataSource      =   "Data1"
      Height          =   252
      Left            =   720
      TabIndex        =   9
      Top             =   2160
      Width           =   612
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393216
      ClipMode        =   1
      Enabled         =   0   'False
      MaxLength       =   4
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Submit"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1200
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Credit"
      Height          =   372
      Left            =   2640
      TabIndex        =   2
      Top             =   720
      Width           =   1092
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      DataField       =   "EntryDate"
      DataSource      =   "Data1"
      Height          =   252
      Left            =   1320
      TabIndex        =   10
      Top             =   2160
      Width           =   852
      _ExtentX        =   1508
      _ExtentY        =   450
      _Version        =   393216
      ClipMode        =   1
      Enabled         =   0   'False
      MaxLength       =   8
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox3 
      DataField       =   "Description"
      DataSource      =   "Data1"
      Height          =   252
      Left            =   2160
      TabIndex        =   11
      Top             =   2160
      Width           =   3012
      _ExtentX        =   5318
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   40
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox4 
      DataField       =   "Debit"
      DataSource      =   "Data1"
      Height          =   252
      Left            =   5160
      TabIndex        =   12
      Top             =   2160
      Width           =   972
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox5 
      DataField       =   "Credit"
      DataSource      =   "Data1"
      Height          =   252
      Left            =   6120
      TabIndex        =   13
      Top             =   2160
      Width           =   972
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox6 
      DataField       =   "Balance"
      DataSource      =   "Data1"
      Height          =   252
      Left            =   7080
      TabIndex        =   14
      Top             =   2160
      Width           =   972
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   8
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   720
      TabIndex        =   24
      Top             =   480
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   4560
      TabIndex        =   25
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Cleared"
      Height          =   192
      Left            =   8280
      TabIndex        =   26
      Top             =   1920
      Width           =   576
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   " "
      Height          =   72
      Left            =   7320
      TabIndex        =   0
      Top             =   6840
      Width           =   372
   End
   Begin VB.Line Line6 
      X1              =   2760
      X2              =   360
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line5 
      X1              =   6600
      X2              =   9240
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line4 
      X1              =   9240
      X2              =   9240
      Y1              =   7080
      Y2              =   240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   6480
      X2              =   9240
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   2640
      X2              =   360
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   360
      X2              =   360
      Y1              =   240
      Y2              =   7080
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "$ Checkbook $"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   552
      Left            =   2760
      TabIndex        =   23
      Top             =   0
      Width           =   3696
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Balance"
      Height          =   192
      Left            =   7080
      TabIndex        =   22
      Top             =   1920
      Width           =   588
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Credit"
      Height          =   192
      Left            =   6120
      TabIndex        =   21
      Top             =   1920
      Width           =   408
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Debit"
      Height          =   192
      Left            =   5160
      TabIndex        =   20
      Top             =   1920
      Width           =   372
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   192
      Left            =   2160
      TabIndex        =   19
      Top             =   1920
      Width           =   792
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Check"
      Height          =   195
      Left            =   720
      TabIndex        =   18
      Top             =   1920
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      Height          =   192
      Left            =   1320
      TabIndex        =   17
      Top             =   1920
      Width           =   348
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oldtotal As Currency
Dim newtotal As Currency
Dim newnumber As Currency
Dim oldnumber As Currency
Dim finaltotal As Currency
Dim difference As Currency
Dim dbName As String
Dim strSQL As String
Dim MySearch As String
Dim Ask As Boolean
Private Sub Command1_Click() 'add(Credit)
    On Error Resume Next 'for empty table
    Command1.Enabled = False
    Command3.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False
    Command8.Enabled = False
    
    MaskEdBox1.Enabled = True
    MaskEdBox2.Enabled = True
    MaskEdBox3.Enabled = True
    MaskEdBox5.Enabled = True
    
    Data1.Recordset.AddNew
    
    MaskEdBox2.Text = Date
End Sub

Private Sub Command10_Click()
    If Ask = False Then
        Data1.Recordset.MoveFirst
        MySearch = InputBox("What should I look for?", "Search")
    End If
    Ask = True
    Command10.Caption = "&Find Next"
    With Data1.Recordset
        .FindNext "'Check Number' LIKE " & "'*" & MySearch & "*'"
        If .NoMatch = True Then
            .FindNext "EntryDate LIKE " & "'*" & MySearch & "*'"
            If .NoMatch = True Then
                .FindNext "Description LIKE " & "'*" & MySearch & "*'"
                If .NoMatch = True Then
                    MsgBox "I can't find any more."
                    Command10.Caption = "&Find"
                    Ask = False
                End If
            End If
        End If
    End With
End Sub

Private Sub Command2_Click() 'submit
    'MUST BE HERE. Text1 is hidden & holds the pkey.
    FindPKey = Text1.Text
    'MsgBox FindPKey
    
    If MaskEdBox5.Enabled = True Then 'credit
        If MaskEdBox5.Text = "" Then
            Data1.Recordset.CancelUpdate
            MsgBox "What are you submitting? Pay attention!"
            GoTo getout
        End If
        Data1.Recordset.MoveFirst
    ElseIf MaskEdBox4.Enabled = True Then 'debit
        If MaskEdBox4.Text = "" Then
            Data1.Recordset.CancelUpdate
            MsgBox "What are you submitting? Pay attention!"
            GoTo getout
        End If
        Data1.Recordset.MoveFirst
    Else
        MsgBox "Don't do that!"
        GoTo getout
    End If

    MySubmit 'guess?

getout:
    MaskEdBox1.Enabled = False
    MaskEdBox2.Enabled = False
    MaskEdBox3.Enabled = False
    MaskEdBox4.Enabled = False
    MaskEdBox5.Enabled = False
    Command1.Enabled = True
    Command3.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command7.Enabled = True
    Command8.Enabled = True
    'need the following two
    Data1.Refresh   ' refresh data control
    Data1.Recordset.MoveLast
    
    Label8.Caption = "Balance = $" & MaskEdBox6.Text & " "
End Sub

Private Sub Command3_Click() 'subtract(debit)
    On Error Resume Next 'for empty table
    Command1.Enabled = False
    Command3.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False
    Command8.Enabled = False
    
    MaskEdBox1.Enabled = True
    MaskEdBox2.Enabled = True
    MaskEdBox3.Enabled = True
    MaskEdBox4.Enabled = True
    
    Data1.Recordset.AddNew
    
    MaskEdBox2.Text = Date
End Sub

Private Sub Command4_Click() 'cancel
    If MaskEdBox4.Enabled = True Or MaskEdBox5.Enabled = True Then
        On Error Resume Next 'for calcelling when you switch records when in edit mode.
        Data1.Recordset.CancelUpdate
        MaskEdBox1.Enabled = False
        MaskEdBox2.Enabled = False
        MaskEdBox3.Enabled = False
        MaskEdBox4.Enabled = False
        MaskEdBox5.Enabled = False
    Else
        On Error Resume Next
        Data1.Recordset.CancelUpdate
        MsgBox "There's nothing to cancel! Get outa here!"
    End If
    Data1.Recordset.MoveLast
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command7.Enabled = True
    Command8.Enabled = True
End Sub

Private Sub Command5_Click() 'quit
    End
End Sub

Private Sub Command6_Click() 'delete
    'MUST BE HERE. Text1 is hidden & holds the pkey.
    FindPKey = Text1.Text
    
    prompt$ = "Do you really want to delete the highlighted record?" & Chr(13) + Chr(10) & _
    "the record will be gone forever?"
    reply = MsgBox(prompt$, vbYesNo, "delete record")
    If reply = vbYes Then
        
        MyDelete
        
        Data1.Refresh   ' refresh data control
        Data1.Recordset.MoveLast
        Label8.Caption = "Balance = $" & MaskEdBox6.Text & " "
            
    Else
        Exit Sub
    End If
    Data1.Recordset.Requery
    Data1.Recordset.MoveLast
    Label8.Caption = "Balance = $" & MaskEdBox6.Text
End Sub

Private Sub Command7_Click() 'edit
    prompt$ = "When you're done editing, you MUST enter the changes by clicking the <UPDATE> button"
    reply = MsgBox(prompt$, , "edit record")
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False
    If MaskEdBox4.Text <> "" Then 'debit
        book = Data1.Recordset.Bookmark
        oldnumber = Val(MaskEdBox4.Text)
        Data1.Recordset.Edit
        MaskEdBox1.Enabled = True
        MaskEdBox2.Enabled = True
        MaskEdBox3.Enabled = True
        MaskEdBox4.Enabled = True
    Else
        book = Data1.Recordset.Bookmark
        Data1.Recordset.Edit
        oldnumber = Val(MaskEdBox5.Text)
        MaskEdBox1.Enabled = True 'credit
        MaskEdBox2.Enabled = True
        MaskEdBox3.Enabled = True
        MaskEdBox5.Enabled = True
    End If
End Sub

Private Sub Command8_Click() 'update
    'MUST BE HERE. Text1 is hidden & holds the pkey.
    FindPKey = Text1.Text
    
    If MaskEdBox3.Enabled = False Then
        MsgBox "You must edit a record before you can update it."
        GoTo imout
    End If
    
    MyUpdate
    
    Data1.Refresh   ' refresh data control
    Data1.Recordset.MoveLast
    Label8.Caption = "Balance = $" & MaskEdBox6.Text & " "
    
imout:
    MaskEdBox1.Enabled = False
    MaskEdBox2.Enabled = False
    MaskEdBox3.Enabled = False
    MaskEdBox4.Enabled = False
    MaskEdBox5.Enabled = False
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command7.Enabled = True
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    Shell ("c:\windows\calc.exe")
End Sub

Private Sub Data1_Reposition()
    Data1.Caption = "Record " & (Data1.Recordset.AbsolutePosition + 1) _
    & " of " & (Data1.Recordset.RecordCount) & " total records."
End Sub

Private Sub Form_Activate()
    Ask = False
    
    dbName = App.Path & "\occasions.mdb"
    strSQL = "SELECT * FROM CheckingTable1 ORDER BY EntryDate, Description, PrimaryKey"
    
    Data1.DatabaseName = dbName
    Data1.RecordSource = strSQL  ' load grid-bound data control
    
    Data1.Refresh   ' refresh data control
    
    recordend = Data1.Recordset.RecordCount
    
    If Data1.Recordset.BOF = False Then
        Data1.Recordset.MoveLast
        Label8.Caption = "Balance = $" & MaskEdBox6.Text & " "
    End If

    On Error Resume Next 'for empty db
    Data1.Recordset.MoveFirst
    Old6 = MaskEdBox6.Text
    
Done:
    Data1.Recordset.MoveLast
End Sub

Private Sub MaskEdBox1_GotFocus()
    MaskEdBox1.PromptChar = "_"
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
    MaskEdBox1.Mask = "####"
End Sub

Private Sub MaskEdBox1_LostFocus()
    If MaskEdBox1.Text = "" Then
        MaskEdBox1.Text = " "
    Else
        MaskEdBox1.Mask = ""
        MaskEdBox1.PromptChar = " "
    End If
End Sub

Private Sub MaskEdBox2_GotFocus()
    MaskEdBox2.PromptChar = "_"
End Sub

Private Sub MaskEdBox2_KeyPress(KeyAscii As Integer)
    TextLen = Trim(MaskEdBox2.Text)
    If Val(TextLen) < 1 Then
        MaskEdBox2.Mask = "##/##/##"
    End If
End Sub

Private Sub MaskEdBox2_LostFocus()
    If MaskEdBox2.Text = "" Then
        MsgBox "You must enter a date"
        MaskEdBox2.SetFocus
    Else
        MaskEdBox2.Mask = ""
        MaskEdBox2.PromptChar = " "
        MaskEdBox2.Text = Mid(MaskEdBox2.Text, 1, Len(MaskEdBox2.Text) - 2) + "0" _
        + Right(Trim(MaskEdBox2.Text), 1)
    End If
End Sub
