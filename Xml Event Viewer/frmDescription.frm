VERSION 5.00
Begin VB.Form frmDescription 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Event Detail"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   330
      Left            =   3825
      TabIndex        =   4
      Top             =   4455
      Width           =   825
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Previous"
      Height          =   330
      Left            =   1260
      TabIndex        =   3
      Top             =   4455
      Width           =   825
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   330
      Left            =   2160
      TabIndex        =   2
      Top             =   4455
      Width           =   825
   End
   Begin VB.TextBox txtDescription 
      BackColor       =   &H80000000&
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   2565
      Width           =   4650
   End
   Begin VB.Label lblComputerField 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1170
      TabIndex        =   20
      Top             =   1935
      Width           =   3480
   End
   Begin VB.Label lblComputer 
      Caption         =   "Computer:"
      Height          =   195
      Left            =   0
      TabIndex        =   19
      Top             =   1935
      Width           =   1095
   End
   Begin VB.Label lblUserField 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1170
      TabIndex        =   18
      Top             =   1665
      Width           =   3480
   End
   Begin VB.Label lblUser 
      Caption         =   "User:"
      Height          =   195
      Left            =   0
      TabIndex        =   17
      Top             =   1665
      Width           =   1095
   End
   Begin VB.Label lblTypeField 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1170
      TabIndex        =   16
      Top             =   1395
      Width           =   3480
   End
   Begin VB.Label lblType 
      Caption         =   "Type:"
      Height          =   195
      Left            =   0
      TabIndex        =   15
      Top             =   1395
      Width           =   1095
   End
   Begin VB.Label lblEventField 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1170
      TabIndex        =   14
      Top             =   1125
      Width           =   3480
   End
   Begin VB.Label lblEvent 
      Caption         =   "Event ID:"
      Height          =   195
      Left            =   0
      TabIndex        =   13
      Top             =   1125
      Width           =   1095
   End
   Begin VB.Label lblCategoryField 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1170
      TabIndex        =   12
      Top             =   855
      Width           =   3480
   End
   Begin VB.Label lblCategory 
      Caption         =   "Category:"
      Height          =   195
      Left            =   0
      TabIndex        =   11
      Top             =   855
      Width           =   1095
   End
   Begin VB.Label lblSourceField 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1170
      TabIndex        =   10
      Top             =   585
      Width           =   3480
   End
   Begin VB.Label lblSource 
      Caption         =   "Source:"
      Height          =   195
      Left            =   0
      TabIndex        =   9
      Top             =   585
      Width           =   1095
   End
   Begin VB.Label lblTimeField 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1170
      TabIndex        =   8
      Top             =   315
      Width           =   3480
   End
   Begin VB.Label lblTime 
      Caption         =   "Time:"
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   315
      Width           =   1095
   End
   Begin VB.Label lblDateField 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1170
      TabIndex        =   6
      Top             =   45
      Width           =   3480
   End
   Begin VB.Label lblDate 
      Caption         =   "Date:"
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   45
      Width           =   1095
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description"
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   2295
      Width           =   1185
   End
End
Attribute VB_Name = "frmDescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    Dim itmx As ListItem
    
    On Error Resume Next
    Set itmx = frmXmlEventViewer.lvXMLLog.ListItems(frmXmlEventViewer.lvXMLLog.SelectedItem.Index + 1)
    
    If Err.Number <> 0 Then
        Me.lblEventField.Caption = ""
        Me.lblDateField.Caption = ""
        Me.lblTimeField.Caption = ""
        Me.lblSourceField.Caption = ""
        Me.lblCategoryField.Caption = ""
        Me.lblTypeField.Caption = ""
        Me.lblUserField.Caption = ""
        Me.lblComputerField.Caption = ""
        Me.txtDescription.Text = ""
        Me.txtDescription.Text = "End of Event Log"
        Exit Sub
    End If
    
   
       Me.lblEventField.Caption = itmx.Text
       Me.lblDateField.Caption = itmx.SubItems(1)
       Me.lblTimeField.Caption = itmx.SubItems(2)
       Me.lblSourceField.Caption = itmx.SubItems(3)
       Me.lblCategoryField.Caption = itmx.SubItems(4)
       
       Me.lblTypeField.Caption = itmx.SubItems(5)
       Me.lblUserField.Caption = itmx.SubItems(6)
       Me.lblComputerField.Caption = itmx.SubItems(7)
       
       Me.txtDescription.Text = itmx.SubItems(8)
   
    
    itmx.Selected = True
    
End Sub

Private Sub cmdPrevious_Click()
    Dim itmx As ListItem
    
    On Error Resume Next
    Set itmx = frmXmlEventViewer.lvXMLLog.ListItems(frmXmlEventViewer.lvXMLLog.SelectedItem.Index - 1)
    
    If Err.Number <> 0 Then
        Me.lblEventField.Caption = ""
        Me.lblDateField.Caption = ""
        Me.lblTimeField.Caption = ""
        Me.lblSourceField.Caption = ""
        Me.lblCategoryField.Caption = ""
        Me.lblTypeField.Caption = ""
        Me.lblUserField.Caption = ""
        Me.lblComputerField.Caption = ""
        Me.txtDescription.Text = ""
        Me.txtDescription.Text = "Start of Event Log"
        Exit Sub
    End If
    
   With Me
       .lblEventField.Caption = itmx.Text
       .lblDateField.Caption = itmx.SubItems(1)
       .lblTimeField.Caption = itmx.SubItems(2)
       .lblSourceField.Caption = itmx.SubItems(3)
       .lblCategoryField.Caption = itmx.SubItems(4)
       
       .lblTypeField.Caption = itmx.SubItems(5)
       .lblUserField.Caption = itmx.SubItems(6)
       .lblComputerField.Caption = itmx.SubItems(7)
       
       .txtDescription.Text = itmx.SubItems(8)
    End With
    
    itmx.Selected = True
End Sub

Private Sub Form_Load()
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, 1
End Sub
