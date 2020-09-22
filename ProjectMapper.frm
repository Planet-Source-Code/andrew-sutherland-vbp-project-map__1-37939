VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPM 
   Caption         =   "Projectmapper"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   7680
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View Tree"
      Height          =   375
      Left            =   5700
      TabIndex        =   3
      Top             =   360
      Width           =   1155
   End
   Begin VB.TextBox Text2 
      Height          =   6255
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   1260
      Width           =   8715
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open VBP"
      Height          =   375
      Left            =   4380
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   300
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objFSO As New FileSystemObject
Dim ts As TextStream
Dim ffile As File
Dim recCount As Long
Dim TotRec As Long

Private Sub cmdOpen_Click()

Me.CommonDialog1.FileName = "*.vbp"
Me.CommonDialog1.Action = 1
Me.Text1 = Me.CommonDialog1.FileName
If IsBlank(Me.Text1) Or Me.Text1 = "*.vbp" Then
Me.Command1.Enabled = False

Exit Sub
Else
Me.Command1.Enabled = True
End If
Me.Text2 = ""
Set ffile = objFSO.GetFile(Me.Text1)
Set ts = ffile.OpenAsTextStream(ForReading)
Do While Not ts.AtEndOfStream
Me.Text2 = Me.Text2 & ts.ReadLine & Chr(13) & Chr(10)
Me.Tag = GetPathFileName(Me.CommonDialog1.FileName, 1)

Loop
Screen.MousePointer = vbHourglass
ts.Close
Process
Me.ProgressBar1.Value = 0
SearchCode
Me.ProgressBar1.Value = 0
Screen.MousePointer = vbNormal

End Sub

Private Sub Process()
Dim inline As String
Dim rs As Recordset
Set rs = Currentdb.OpenRecordset("tblforms")
Set ts = ffile.OpenAsTextStream(ForReading)
Currentdb.Execute ("Delete * from tblForms")
Do While Not ts.AtEndOfStream
inline = ts.ReadLine
If InStr(inline, "Form=") <> 0 And GetDlmData(inline, "=", 1) = "Form" Then
    rs.AddNew
    rs.Fields("FormName") = GetDlmData(inline, "=", 2)
    rs.Update
End If
If InStr(inline, "Module=") <> 0 And GetDlmData(inline, "=", 1) = "Module" Then
    rs.AddNew
    rs.Fields("FormName") = Trim(GetDlmData(inline, ";", 2))
    rs.Update
End If
Loop
ts.Close
rs.Close
Set rs = Currentdb.OpenRecordset("tblforms")
TotRec = rs.RecordCount
FindStartup


End Sub

Private Sub FindStartup()
Set ts = ffile.OpenAsTextStream(ForReading)
Dim rs As Recordset
Dim SUF As String
Dim CheckBas As Boolean
Set rs = Currentdb.OpenRecordset("select * from tblForms order by formname")
Dim inline As String

Do While Not ts.AtEndOfStream
    inline = ts.ReadLine
    If InStr(inline, "Startup=") <> 0 Then
        SUF = GetDlmData(inline, "=", 2)
    End If

    DoEvents
Loop
SUF = Trim(Replace(SUF, """", ""))

FindModuleVBNames (SUF)
FindFormVBNames (SUF)

End Sub

Private Sub FindModuleVBNames(pSUF As String)
Dim rs As Recordset
Dim inline As String
Set rs = Currentdb.OpenRecordset("select * from tblForms where right(formname,3) ='bas'")
Do While Not rs.EOF
    rs.Edit
    Set ffile = objFSO.GetFile(rs.Fields("formname"))
    Set ts = ffile.OpenAsTextStream(ForReading)
    Do While Not ts.AtEndOfStream
        inline = ts.ReadLine
        If InStr(inline, "Attribute VB_Name =") <> 0 Then
            rs.Fields("formVBname") = Trim(Replace(GetDlmData(inline, "=", 2), """", ""))
        End If
        If InStr(UCase(inline), UCase(pSUF)) <> 0 Then
            rs.Fields("Startform") = True
        End If
        
        
        rs.Fields("Type") = "vbbas"
        
        rs.Fields("sourcecode") = rs.Fields("SourceCode") & inline & Chr(13) & Chr(10)
        DoEvents
    Loop
    rs.Update
    rs.MoveNext
    ts.Close
Loop
End Sub


Private Sub FindFormVBNames(pSUF As String)
Dim rs As Recordset
Dim inline As String
Set rs = Currentdb.OpenRecordset("select * from tblForms where right(formname,3) ='frm'")
recCount = 0
Do While Not rs.EOF
    rs.Edit
    Set ffile = objFSO.GetFile(rs.Fields("formname"))
    Set ts = ffile.OpenAsTextStream(ForReading)
    Do While Not ts.AtEndOfStream
        inline = ts.ReadLine
        If InStr(inline, "Attribute VB_Name =") <> 0 And InStr(inline, "Attribute VB_Name =") = 1 Then
            
            rs.Fields("formVBname") = Trim(Replace(GetDlmData(inline, "=", 2), """", ""))
           
        End If
         If InStr(UCase(inline), UCase(pSUF)) <> 0 Then
            rs.Fields("Startform") = True
        End If
        If InStr(UCase(inline), UCase("MDIChild")) Then
            rs.Fields("Type") = "form2"
        End If
        
        If InStr(UCase(inline), UCase("MDIForm")) Then
            rs.Fields("Type") = "form3"
        End If
        
        rs.Fields("sourcecode") = rs.Fields("SourceCode") & inline & Chr(13) & Chr(10)
        
        DoEvents
    Loop
    If IsBlank(rs.Fields("type")) Then
            rs.Fields("type") = "form1"
    End If
    rs.Update
    rs.MoveNext
    recCount = recCount + 1
    Me.ProgressBar1.Value = IIf(Int((recCount / TotRec) * 100) > 100, 100, Int((recCount / TotRec) * 100))
    ts.Close
Loop
End Sub

Private Sub SearchCode()
Dim rs As Recordset
Dim lastform As Long
Set rs = Currentdb.OpenRecordset("Select * from tblForms where startform =true")
Currentdb.Execute ("Delete * from tblFormcalls")

lastform = SearchForms(rs.Fields("id"))




End Sub

Private Function LookUpForm(pFormName As String) As Long
Dim rs As Recordset
Set rs = Currentdb.OpenRecordset("Select * from tblforms where formvbname='" & FixText(pFormName) & "'")
If rs.RecordCount <> 0 Then
    LookUpForm = rs.Fields("ID")
Else
    LookUpForm = 0
End If

End Function

Private Sub Command1_Click()
frmvbpTree.Show
End Sub

Private Function SearchForms(pformid As Long) As Long
Dim rs As Recordset
Set rs = Currentdb.OpenRecordset("Select * from tblForms where id =" & pformid)
Dim rsCf As Recordset
Set rsCf = Currentdb.OpenRecordset("tblformcalls")
Dim CodeStr As String
Dim cLen As Long
Dim cPos As Long
Dim cNewPos As Long
Dim i As Long
Dim tstr As String
Dim X As Integer
Dim formStr As String
Dim CallingFormID As Long
Dim CalledFormID As Long
Dim UsesWith As Boolean

DoEvents
recCount = recCount + 1
Me.ProgressBar1.Value = IIf(Int((recCount / TotRec) * 100) > 100, 100, Int((recCount / TotRec) * 100))
CallingFormID = rs.Fields("ID")
CodeStr = UCase(rs.Fields("SourceCode"))
cLen = Len(CodeStr)
i = 1
Do While i < cLen
i = InStr(i, CodeStr, UCase(".show"))
cNewPos = i + 5
If i = 0 Then
    Exit Do
Else
   ' Stop
    cPos = i - 2000
    tstr = Mid(CodeStr, cPos, 2000)
   ' Debug.Print tstr
    X = Len(tstr)
    If Mid(tstr, X, 1) = " " Or Mid(tstr, X, 1) = Chr(10) Then
            UsesWith = True
        Else
            UsesWith = False
        End If
    For X = Len(tstr) To 1 Step -1
       
        If Not UsesWith Then
            If Mid(tstr, X, 1) = " " Or Mid(tstr, X, 1) = Chr(10) Then Exit For
            formStr = Mid(tstr, X, 1) & formStr
        Else
            formStr = Mid(tstr, X, 1) & formStr
            If InStr(formStr, UCase("With ")) Then
                formStr = Right(GetDlmData(formStr, Chr(13), 1), Len(GetDlmData(formStr, Chr(13), 1)) - 5)
            Exit For
            End If
            
        End If
        
    Next
    CalledFormID = LookUpForm(formStr)
    If CalledFormID <> 0 Then
        rsCf.AddNew
        rsCf.Fields("callingformID") = CallingFormID
        rsCf.Fields("CalledformID") = CalledFormID
        If Not CallDuplicates(CallingFormID, CalledFormID) Then
            rsCf.Update
            'recurse here
            SearchForms (CalledFormID)
        End If
    End If
    
End If
formStr = ""
i = cNewPos
Loop
rs.Close
rsCf.Close


End Function

Private Function findCalledFormID(pformid As Long) As Long
Dim rs As Recordset
Set rs = Currentdb.OpenRecordset("Select * from tblformcalls where callingformid =" & pformid)
If rs.RecordCount = 0 Then
    findCalledFormID = pformid
Else
    findCalledFormID = 0
End If
rs.Close
End Function

Private Function CallDuplicates(pCallingID As Long, pCalledID As Long) As Boolean
Dim rs As Recordset
Set rs = Currentdb.OpenRecordset("Select * from tblFormCalls where Callingformid=" & pCallingID _
    & "and  CalledformID =" & pCalledID)
If rs.RecordCount <> 0 Then
    CallDuplicates = True
Else
    CallDuplicates = False
End If
rs.Close
End Function

