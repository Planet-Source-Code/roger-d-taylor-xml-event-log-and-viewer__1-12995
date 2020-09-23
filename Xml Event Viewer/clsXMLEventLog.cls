VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXMLEventLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public xml As New DOMDocument
Public root As IXMLDOMNode

Public Enum NTEventType
    EVENTLOG_SUCCESS = 0
    EVENTLOG_ERROR_TYPE = 1
    EVENTLOG_WARNING_TYPE = 2
    EVENTLOG_INFORMATION_TYPE = 4
    EVENTLOG_AUDIT_SUCCESS = 8
    EVENTLOG_AUDIT_FAILURE = 10
End Enum

Public Function WriteXMLEventLog( _
                strLogLocation As String, _
                strSource As String, _
                strCategory As String, _
                lngEventType As NTEventType, _
                strDescription As String) As Long
                
Dim bLoaded As Boolean   ' Return value stating the success of the load
Dim objXMLcdata As IXMLDOMElement

xml.async = False
xml.validateOnParse = False
xml.preserveWhiteSpace = False
xml.resolveExternals = False


If Dir(strLogLocation) <> "" Then
    'Load an existing XML document
    bLoaded = xml.Load(strLogLocation)
    If bLoaded = True Then
        Set root = xml.documentElement
        Debug.Print CStr(CLng(root.lastChild.Attributes.Item(0).Text) + 1)
    End If
Else
    xml.appendChild xml.createProcessingInstruction("xml", "version='1.0'")
    Set root = xml.createNode("element", "xmleventlog", "")
    xml.appendChild root
End If


'Adding New event node to the root
Set newnode = xml.createNode(1, "event", "")
    On Error Resume Next
    newnode.setAttribute "ID", CStr(CLng(root.lastChild.Attributes.Item(0).Text) + 1)
    If Err.Number <> 0 Then
        newnode.setAttribute "ID", "1"
    End If
    
    root.appendChild newnode
        
        'Adding a child node to the parent node
        Set newnode = xml.createNode("element", "date", "")
            newnode.Text = CStr(Format(Now(), "DD-MM-YYYY"))
            On Error Resume Next
            xml.documentElement.childNodes(xml.documentElement.childNodes.length - 1).appendChild newnode
            
            If Err.Number <> 0 Then
                xml.documentElement.childNodes(0).appendChild newnode
            End If
        Set newnode = Nothing
    
        'Adding a child node to the parent node
        Set newnode = xml.createNode("element", "time", "")
            newnode.Text = CStr(Format(Now(), "HH:MM:SS"))
            On Error Resume Next
            xml.documentElement.childNodes(xml.documentElement.childNodes.length - 1).appendChild newnode
            
            If Err.Number <> 0 Then
                xml.documentElement.childNodes(0).appendChild newnode
            End If
        Set newnode = Nothing
    
        'Adding a child node to the parent node
        Set newnode = xml.createNode("element", "source", "")
            newnode.Text = strSource
            On Error Resume Next
            xml.documentElement.childNodes(xml.documentElement.childNodes.length - 1).appendChild newnode
            
            If Err.Number <> 0 Then
                xml.documentElement.childNodes(0).appendChild newnode
            End If
        Set newnode = Nothing

        'Adding a child node to the parent node
        Set newnode = xml.createNode("element", "category", "")
            newnode.Text = IIf(Trim("" & strCategory) = "", "NONE", Trim("" & strCategory))
            On Error Resume Next
            xml.documentElement.childNodes(xml.documentElement.childNodes.length - 1).appendChild newnode
            
            If Err.Number <> 0 Then
                xml.documentElement.childNodes(0).appendChild newnode
            End If
        Set newnode = Nothing
    
        'Adding a child node to the parent node
        Set newnode = xml.createNode("element", "event", "")
            
            Select Case lngEventType
                Case 0
                    newnode.Text = "Success"
                Case 1
                    newnode.Text = "Error"
                Case 2
                    newnode.Text = "Warning"
                Case 4
                    newnode.Text = "Information"
                Case 8
                    newnode.Text = "Autdit_Success"
                Case 10
                    newnode.Text = "Audit_Failure"
            End Select
            
            On Error Resume Next
            xml.documentElement.childNodes(xml.documentElement.childNodes.length - 1).appendChild newnode
            
            If Err.Number <> 0 Then
                xml.documentElement.childNodes(0).appendChild newnode
            End If
        Set newnode = Nothing
        
        
                'Adding a child node to the parent node
        Set newnode = xml.createNode("element", "user", "")
            newnode.Text = Trim("" & GetUser())
            On Error Resume Next
            xml.documentElement.childNodes(xml.documentElement.childNodes.length - 1).appendChild newnode
            
            If Err.Number <> 0 Then
                xml.documentElement.childNodes(0).appendChild newnode
            End If
        Set newnode = Nothing
        
        Dim osys As New clsSystemFunctions
        
       'Adding a child node to the parent node
        Set newnode = xml.createNode("element", "computer", "")
            newnode.Text = Trim("" & osys.ComputerName())
            On Error Resume Next
            xml.documentElement.childNodes(xml.documentElement.childNodes.length - 1).appendChild newnode
            
            If Err.Number <> 0 Then
                xml.documentElement.childNodes(0).appendChild newnode
            End If
        Set newnode = Nothing
    
    
                'Adding a child node to the parent node
        Set newnode = xml.createNode("element", "description", "")
            newnode.Text = strDescription
            On Error Resume Next
            xml.documentElement.childNodes(xml.documentElement.childNodes.length - 1).appendChild newnode
            
            If Err.Number <> 0 Then
                xml.documentElement.childNodes(0).appendChild newnode
            End If
                
        Set newnode = Nothing
    
Set newnode = Nothing
Set root = Nothing
xml.save strLogLocation
End Function

