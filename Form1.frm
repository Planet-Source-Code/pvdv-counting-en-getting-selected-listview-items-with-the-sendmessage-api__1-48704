VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   8220
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy selected items to textbox"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   3120
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Count the selected items"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   4335
   End
   Begin MSComctlLib.ListView lvwDetails 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5106
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
    
    MsgBox CountSelectedItemsInListview(lvwDetails) & " item(s) are selected"
End Sub

Private Sub Command2_Click()
Dim itemx As ListItem
Dim myCol As Collection

    Set myCol = GetSelectedItemsFromListview(lvwDetails)
    Text1.Text = ""
    
    For Each itemx In myCol
        Text1.Text = Text1.Text & itemx.Text & vbCrLf
    Next itemx
End Sub

Private Sub Form_Load()
Dim i
    
    'Add some items in listview
    
    For i = 1 To 100
        lvwDetails.ListItems.Add , , i
    Next i
    
    
End Sub

