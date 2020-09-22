VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmvbpTree 
   Caption         =   "Project Tree"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5895
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   10398
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   540
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBPTree.frx":0000
            Key             =   "form1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBPTree.frx":0452
            Key             =   "form2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBPTree.frx":08A4
            Key             =   "vbbas"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBPTree.frx":0E36
            Key             =   "form3"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmvbpTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mNode As Node

Dim rs As Recordset
Private MenuStack() As String
Dim stkCount As Integer
Const L = "    |"
Const txtCon = "|____"
Const B = "     "
Dim FSO As New FileSystemObject
Dim ts As TextStream
Dim f As File

Private Sub cleanupText() '
'this is a waste of time and does not work
ts.Close
Dim txtlast As String
Dim txtPrev As String
Dim BlenL As Integer
Dim BlenP As Integer
Set ts = FSO.OpenTextFile(App.Path & "\" & frmPM.Tag & ".txt", ForReading, False)
Dim txtary() As String
Dim i As Integer
Dim lvl As Integer
Do While Not ts.AtEndOfStream
    i = i + 1
    ReDim Preserve txtary(i)
    txtary(i) = ts.ReadLine

Loop
ts.Close
Set ts = FSO.OpenTextFile(App.Path & "\" & frmPM.Tag & ".txt", ForWriting, True)
For i = UBound(txtary) To 1 Step -1
    txtlast = txtary(i)
    BlenL = Len(GetDlmData(txtary(i), "|", 1))
    If i - 2 > 1 Then
        txtPrev = txtary(i - 2)
        BlenP = Len(GetDlmData(txtary(i), "|", 1))
        If BlenL > BlenP And i <> UBound(txtary) Then
            txtlast = Mid(txtlast, 1, BlenL - 4) & "|"
        End If
        
        
    Else
        Exit For
    End If
    

Next


End Sub





Private Sub Form_Load()
Dim CallingID As Long
Dim CalledID As Long
Dim LastCallingID As Long
Dim LastCalledID As Long
Dim RootKey As String
Dim PrevRootKey As String

Dim DupItem As Integer
Dim Key As String
ReDim MenuStack(0)
stkCount = 0

Set FSO = New FileSystemObject
Set ts = FSO.OpenTextFile(App.Path & "\" & frmPM.Tag & ".txt", ForWriting, True)
On Error GoTo errh

Set rs = Currentdb.OpenRecordset("Select * from tblformcalls order by ID")

CallingID = rs.Fields("CallingFormID")
CalledID = rs.Fields("CalledFormID")

'Set first node
Me.TreeView1.ImageList = Me.ImageList1
Me.TreeView1.Nodes.add , , GetFormName(CallingID), GetFormName(CallingID), GetFormIcon(CallingID)
RootKey = GetFormName(CallingID)
PushStack (RootKey)
WriteText ts, RootKey, UBound(MenuStack)
Me.TreeView1.Nodes.add RootKey, tvwChild, GetFormName(CalledID), GetFormName(CalledID), GetFormIcon(CalledID)
PrevRootKey = RootKey
RootKey = GetFormName(CalledID)
LastCalledID = CalledID
LastCallingID = CallingID
PushStack (RootKey)
WriteText ts, RootKey, UBound(MenuStack)
rs.MoveNext
With rs
    Do While Not .EOF
        
        CallingID = rs.Fields("CallingFormID")
        CalledID = rs.Fields("CalledFormID")
        
        If FindStackItem(GetDlmData(GetFormName(CallingID), "-", 1)) Then
            RootKey = GetFormName(CallingID)
        Else
            PushStack (GetFormName(CallingID))
            RootKey = GetFormName(CallingID)
        End If
        'ShowStack
        If CallingID = LastCalledID Then
            
            
            WriteText ts, GetFormName(CalledID), UBound(MenuStack)
            
                
                Me.TreeView1.Nodes.add RootKey, tvwChild, TestNode(CalledID), GetFormName(CalledID), GetFormIcon(CalledID)
  
        ElseIf CallingID = LastCallingID Then
            Me.TreeView1.Nodes.add RootKey, tvwChild, TestNode(CalledID), GetFormName(CalledID), GetFormIcon(CalledID)
            WriteText ts, GetFormName(CalledID), UBound(MenuStack)
        ElseIf CallingID <> LastCalledID Then
            Me.TreeView1.Nodes.add RootKey, tvwChild, TestNode(CalledID), GetFormName(CalledID), GetFormIcon(CalledID)
            WriteText ts, GetFormName(CalledID), UBound(MenuStack)
        End If
        LastCallingID = CallingID
        LastCalledID = CalledID
        
        .MoveNext
    Loop
End With
Exit Sub
errh:

Resume Next
End Sub

Private Function GetFormName(pformid As Long) As String
Dim rs As Recordset
Set rs = Currentdb.OpenRecordset("Select formName from tblforms where ID=" & pformid)
GetFormName = rs.Fields(0)
End Function

Private Function GetFormIcon(pformid As Long) As String
Dim rs As Recordset
Set rs = Currentdb.OpenRecordset("Select type from tblforms where ID=" & pformid)
GetFormIcon = rs.Fields(0)
End Function


Private Sub Form_Resize()
Me.TreeView1.Width = Me.Width - 100
Me.TreeView1.Height = Me.Height - 100
End Sub

Private Sub PushStack(pKey As String)
stkCount = stkCount + 1
ReDim Preserve MenuStack(stkCount)
MenuStack(stkCount) = pKey
End Sub

Private Function FindStackItem(pKey As String) As Boolean
Dim i As Integer
For i = UBound(MenuStack) To 1 Step -1
    If MenuStack(i) = pKey Then
        FindStackItem = True
        ReDim Preserve MenuStack(i)
        stkCount = UBound(MenuStack)
        Exit For
    End If
Next

End Function

Private Sub WriteText(pTS As TextStream, pText As String, pLevel As Integer)
Dim Outline As String
Dim i As Integer

For i = 1 To pLevel
    Outline = Outline & B
Next
'Outline = Outline & L
pTS.WriteLine Outline & "|"
Outline = Outline & txtCon & pText
pTS.WriteLine Outline
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

ts.Close
End Sub

Private Sub ShowStack()
Dim i As Integer
For i = UBound(MenuStack) To 1 Step -1
    Debug.Print i & " " & MenuStack(i)
Next

End Sub

Private Function TestNode(pCalledID As Long) As String
Dim tN As String
On Error GoTo errh
tN = Me.TreeView1.Nodes.Item(GetFormName(pCalledID)).Text
If tN = "" Then
                TestNode = GetFormName(pCalledID)
            Else
                
                
                TestNode = GetFormName(pCalledID) & "-" & Format(Now, "HHMMSS")
            End If
Exit Function
errh:
Resume Next
End Function
