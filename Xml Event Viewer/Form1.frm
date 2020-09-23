VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmXmlEventViewer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "XML Event Log Viewer"
   ClientHeight    =   9495
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   540
      Top             =   9225
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "auditwarning"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":031A
            Key             =   "auditfailure"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0634
            Key             =   "warning"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0A86
            Key             =   "information"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0ED8
            Key             =   "error"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":132A
            Key             =   "sucess"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   -45
      Top             =   9270
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Event Log"
      Height          =   420
      Left            =   9045
      TabIndex        =   2
      Top             =   9000
      Width           =   1680
   End
   Begin MSComctlLib.ListView lvXMLLog 
      Height          =   8835
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   15584
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "EventID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Source"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Category"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Event"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "User"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Computer"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Description"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton cmdWriteEvent 
      Caption         =   "WriteEvent"
      Height          =   510
      Left            =   9585
      TabIndex        =   0
      Top             =   0
      Width           =   1185
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuCSVExport 
         Caption         =   "Export To CSV"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmXmlEventViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xml As New DOMDocument
Public root As IXMLDOMNode
Public strLog As String


Private Sub Command1_Click()

Me.CommonDialog1.Filter = "*.xml|*.xml"

Me.CommonDialog1.ShowOpen
Me.lvXMLLog.ListItems.Clear
ReadXMLEventLog Me.CommonDialog1.FileName

End Sub


Private Sub cmdWriteEvent_Click()
    Dim X As New clsXMLEventLog
    X.WriteXMLEventLog "C:\MyTemp.xml", "ME", "1", EVENTLOG_SUCCESS, "HELLO"
End Sub


Public Sub ReadXMLEventLog(ByVal strPath As String)

Dim bolLoaded As Boolean
Dim readroot As IXMLDOMNode
Dim xml As New DOMDocument
Dim i As Long
Dim itmx As ListItem
Dim j As Long

bolLoaded = xml.Load(strPath)

Set readroot = xml.documentElement

Dim ndNext As IXMLDOMNode
Dim currNode As IXMLDOMNode
Dim nd As IXMLDOMNodeList

Set nd = xml.getElementsByTagName("event")

For i = 0 To (nd.length - 1)

        Set currNode = nd.nextNode

        On Error Resume Next
        
        Set itmx = Me.lvXMLLog.ListItems.Add(, , currNode.Attributes(0).Text)
        
        Select Case UCase(currNode.childNodes(4).Text)
            Case "SUCCESS"
                itmx.SmallIcon = "sucess"
            Case "ERROR"
                itmx.SmallIcon = "error"
            Case "WARNING"
                itmx.SmallIcon = "warning"
            Case "INFORMATION"
                itmx.SmallIcon = "information"
            Case "AUDIT_SUCCESS"
                itmx.SmallIcon = "auditwarning"
            Case "AUDIT_FAILURE"
                itmx.SmallIcon = "auditfailure"
        End Select
        
        
        If Err.Number <> 0 Then
            Exit For
        End If
        
        itmx.SubItems(1) = currNode.childNodes(0).Text
        itmx.SubItems(2) = currNode.childNodes(1).Text
        itmx.SubItems(3) = currNode.childNodes(2).Text
        itmx.SubItems(4) = currNode.childNodes(3).Text
        itmx.SubItems(5) = currNode.childNodes(4).Text
        itmx.SubItems(6) = currNode.childNodes(5).Text
        itmx.SubItems(7) = currNode.childNodes(6).Text
        itmx.SubItems(8) = currNode.childNodes(7).Text
        Set currNode = nd.nextNode

      
    
Next



End Sub

Private Sub lvXMLLog_ItemClick(ByVal Item As MSComctlLib.ListItem)

Dim itmx As ListItem

Set itmx = Me.lvXMLLog.SelectedItem

    
    With frmDescription
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


frmDescription.Show


End Sub

Private Sub mnuCSVExport_Click()

Dim fso As New FileSystemObject


Dim itmx As ListItem
Dim strBuild As String

Dim strFileName As String


Me.CommonDialog1.Filter = "*.csv|*.csv"
Me.CommonDialog1.DialogTitle = "Export Event Log to CSV"
Me.CommonDialog1.ShowSave

strFileName = Me.CommonDialog1.FileName

If strFileName = "" Then
    MsgBox "Please Provide a file name"
    Exit Sub
End If

strBuild = Chr(34) & "Event_ID" & Chr(34) & Chr(44)
strBuild = strBuild & Chr(34) & "Date" & Chr(34) & Chr(44)
strBuild = strBuild & Chr(34) & "Time" & Chr(34) & Chr(44)
strBuild = strBuild & Chr(34) & "Source" & Chr(34) & Chr(44)
strBuild = strBuild & Chr(34) & "Category" & Chr(34) & Chr(44)
strBuild = strBuild & Chr(34) & "Event" & Chr(34) & Chr(44)
strBuild = strBuild & Chr(34) & "User" & Chr(34) & Chr(44)
strBuild = strBuild & Chr(34) & "Computer" & Chr(34) & Chr(44)
strBuild = strBuild & Chr(34) & "Description" & Chr(34) & vbCrLf


fso.OpenTextFile(strFileName, ForAppending, True).WriteLine strBuild

For i = 1 To Me.lvXMLLog.ListItems.Count

   Set itmx = Me.lvXMLLog.ListItems(i)

    strBuild = Chr(34) & itmx.Text & Chr(34) & Chr(44)
    strBuild = strBuild & Chr(34) & itmx.SubItems(1) & Chr(34) & Chr(44)
    strBuild = strBuild & Chr(34) & itmx.SubItems(2) & Chr(34) & Chr(44)
    strBuild = strBuild & Chr(34) & itmx.SubItems(3) & Chr(34) & Chr(44)
    strBuild = strBuild & Chr(34) & itmx.SubItems(4) & Chr(34) & Chr(44)
    strBuild = strBuild & Chr(34) & itmx.SubItems(5) & Chr(34) & Chr(44)
    strBuild = strBuild & Chr(34) & itmx.SubItems(6) & Chr(34) & Chr(44)
    strBuild = strBuild & Chr(34) & itmx.SubItems(7) & Chr(34) & Chr(44)
    strBuild = strBuild & Chr(34) & itmx.SubItems(8) & Chr(34) & vbCrLf
   
    fso.OpenTextFile(strFileName, ForAppending, True).WriteLine strBuild
   
    strBuild = ""
    

Next

MsgBox "Export Complete", vbInformation, "Export"
Set fso = Nothing

End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub
