Attribute VB_Name = "bMain"
Option Explicit



Global DatabaseName As String
Global Currentdb As Database





Public Function IsBlank(rvar As Variant) As Boolean
' Purpose:   Test a string for null or zero length
' Arguments: String to check
' Returns:   True/False, True=value is Null or zero-length
' Example:   IsBlank(strWork)

  On Error GoTo IsBlank_Err
  Const CstrProc As String = "IsBlank"
  
  If IsNull(rvar) Then
    IsBlank = True
    GoTo IsBlank_Exit
  End If
  If Len(rvar) = 0 Then
    IsBlank = True
    GoTo IsBlank_Exit
  End If
  IsBlank = False

IsBlank_Exit:
  On Error Resume Next
  Exit Function

IsBlank_Err:
  'Call ErrMsgStd(mcstrMod & "." & CstrProc, Err.Number, Err.Description, True)
  Resume IsBlank_Exit

End Function
  

Sub main()

    DatabaseName = App.Path & "\projectmapper.mdb"
    Set Currentdb = OpenDatabase(DatabaseName)
    
    frmPM.Show
    
End Sub



Public Function GetDlmData(inline As String, Dlm As String, FieldNum As Integer) As String
'returns string from delimited data base on pos param sent, if no data sends back ""
'updated 3-26-99 to work with strings such as |45|4|5||5|
'where position 1 would return "", pos2 = 45, p3 =4, p4=5, p6=""
'if a fieldnum is entered that is beyond the number of fields "" is returned

Dim spos As Integer
Dim epos As Integer
Dim i As Integer
spos = 1
For i = 1 To FieldNum
epos = InStr(spos, inline, Dlm)
If epos = 0 And i = FieldNum Then
    GetDlmData = Mid(inline, spos + Len(Dlm) - 1, Len(inline) - (spos - 1))
    Exit Function
ElseIf epos = 0 Then
    GetDlmData = ""
    Exit Function
End If
If epos = 1 And i = FieldNum Then
    GetDlmData = ""
Else
    GetDlmData = Mid(inline, spos, (epos) - spos)
    spos = epos + 1
End If
If Len(GetDlmData) <> 0 Then
spos = epos + 1
End If
Next


End Function
Function LastOccurrence(strIn As String, strFind As String) As Integer
  ' Comments  : returns the last position of a string in a string
  ' Parameters: strIn - string to search in
  '             strFind - string to search for
  ' Returns   : Position or zero
  '
  Dim intPos As Integer
  Dim intWordCount As Integer

  intWordCount = 1
  intPos = InStr(strIn, strFind)
  
  Do While intPos > 0
    
    intPos = InStr(intPos + 1, strIn, strFind)
    
    If intPos > 0 Then
      LastOccurrence = intPos
    End If

  Loop

  
End Function



Public Function GetPathFileName(FullPath As String, GetType As Integer) As String
'Date: 7/7/98
'Modified: 7/7/98
'Author: Andy Sutherland
'Purpose: Strip apart a passed full file name into its path or file name components
'Paramaters: The full pathnameof the file, Type of Get 2 for path
'1 for file name only
'good for use with common dialog where you get the full path and file and want to use
'the path for other things later

Dim pos As Integer
Dim FileName As String

If GetType = 1 Then
    For pos = Len(FullPath) To 1 Step -1
        If Mid(FullPath, pos, 1) = "\" Then
            FileName = Mid(FullPath, pos + 1, Len(FullPath) - pos)
            If InStr(FileName, ".") Then
                FileName = Mid(FileName, 1, InStr(FileName, ".") - 1)
            End If
            Exit For
        End If
    Next
ElseIf GetType = 2 Then
    For pos = Len(FullPath) To 1 Step -1
        If Mid(FullPath, pos, 1) = "\" Then
            FileName = Mid(FullPath, 1, pos - 1)
            Exit For
        End If
        
    Next
End If

GetPathFileName = FileName
End Function


Public Function FixText(sText As String) As String
FixText = ""
If sText <> "" Then
   FixText = Replace(sText, "'", "''")
End If
End Function







