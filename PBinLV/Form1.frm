VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "PB in LV"
   ClientHeight    =   1995
   ClientLeft      =   2925
   ClientTop       =   2700
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   133
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   Begin VB.CommandButton Command2 
      Caption         =   "Remove Item"
      Height          =   330
      Left            =   1380
      TabIndex        =   3
      Top             =   60
      Width           =   1260
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   5280
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   106
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   4740
      Top             =   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Item"
      Height          =   330
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1260
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1455
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Downloading"
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Progress"
         Object.Width           =   3705
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   1984
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intProgress As Integer

Private Sub Command1_Click()
    Me.ListView1.ListItems.Add , , "Item" & Me.ListView1.ListItems.Count + 1
    Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
    If Me.ListView1.ListItems.Count >= 1 Then
        Me.ListView1.ListItems.Remove Me.ListView1.SelectedItem.Index
    End If
End Sub

Private Sub Form_Load()
    Me.Show
    Me.Refresh
    InitPBinLV ListView1
End Sub

Private Sub Form_Resize()
    ListView1.Width = Me.ScaleWidth - 8
    ListView1.Height = Me.ScaleHeight - 36
End Sub

Private Sub Timer1_Timer()
    Dim intI As Integer
    
    For intI = 1 To ListView1.ListItems.Count
        SetProgress intI, GetProgress(intI) + 1
    Next
End Sub
