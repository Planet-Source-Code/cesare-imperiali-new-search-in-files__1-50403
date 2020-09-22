VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSearchInFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Raise events to give feedback
Event FileSearched(ByVal FileInfo As String)
Event FileError(ByVal ErrDescr As String)
Event FileJobPercent(ByVal FileRatio As Integer)

Event FileFound(ByVal FilePathName As String)

Event FolderSearched(ByVal FolderInfo As String)
Event FolderError(ByVal ErrDescr As String)
Event FolderJobPercent(ByVal FolderRatio As Integer)


Private mMsgTitle As String
Private mbStop As Boolean
Const MaxLen As Long = 102400 'hundred times 1024 = 100Kb
Public Property Get bStop() As Boolean
   bStop = mbStop
End Property

Public Property Let bStop(ByVal StopSearch As Boolean)
   mbStop = StopSearch
End Property

Public Sub SearchInFiles( _
          ByVal theFolder As Scripting.Folder _
        , ByRef sFilter() As String _
        , ByVal SerachedText As String _
        , ByVal MatchCase As Integer _
        , ByVal bRecurse As Boolean _
        )

    'there are folders - like system ones,
    'where you cannot even count the files in...
    On Error Resume Next
    
    Dim Counter As Integer
    Dim FoundPos As Integer
    Dim strExtension As String
    Dim FileContent As String
    Dim tmpPath As String
    Dim tmpFile As Scripting.File
    
    tmpPath = theFolder.Path
   
    If Right(tmpPath, 1) <> "\" Then
        tmpPath = tmpPath & "\"
    End If
Dim lProgress As Long
Dim lCountFile As Long
Err.Clear
lProgress = theFolder.Files.Count

RaiseEvent FolderJobPercent(0)

If Err.Number <> 0 Then 'System Volume Information is a non searchable folder
    RaiseEvent FolderError("Error in searching " & theFolder.Name & Space(1) & Err.Description)
    Err.Clear
    Exit Sub 'do not try to search inside and to list subfolders...there are none.
Else
   RaiseEvent FolderSearched("Files in " & tmpPath & "= " & lProgress)
End If

    
    
    For Each tmpFile In theFolder.Files
        DoEvents
        If mbStop Then Exit Sub
        'new file: start with 0
        RaiseEvent FileJobPercent(0)
        'get extension
        
        strExtension = tmpFile.Name
        FoundPos = InStrRev(strExtension, ".")
        If FoundPos > 0 Then
            strExtension = Right(strExtension, Len(strExtension) - FoundPos)
        End If
        
        'check if match with filters
        For Counter = 0 To UBound(sFilter)
            If mbStop Then Exit Sub
            
            If LCase(strExtension) = sFilter(Counter) Or sFilter(Counter) = "*" Then
                'this file is to be added
                'search inside it
                'scan chunked...
                If ScanChunked(tmpFile.Path, SerachedText, MatchCase) Then
                      Exit For
                End If
            End If
        Next
        lCountFile = lCountFile + 1
        RaiseEvent FolderJobPercent(CInt((lCountFile / lProgress) * 100))
        
    Next
    'do you need to look in subfolders?
    Dim tmpFolder As Scripting.Folder
    If bRecurse Then
        For Each tmpFolder In theFolder.SubFolders
            If mbStop Then Exit Sub
            Call SearchInFiles(tmpFolder, sFilter, SerachedText, MatchCase, bRecurse)
        Next
    End If

End Sub

Private Function ScanChunked( _
          ByVal fileName As String _
        , ByVal SearchString As String _
        , ByVal MatchCase As Integer _
        ) As Boolean
    Dim intFree As Integer
    intFree = FreeFile 'a free file number
    
    Dim bytePos As Long 'bytes to get from file
    Dim sContent As String 'content of file to get
    
    Dim Remind As Long 'to get remind part of file
    Dim sPrevBytes As String 'to add to content last read bytes so that if chunk split
                             'the searched word, you can still get it
    
    
    
    'this may take a lot.
    'let user see a progress working
    Dim lProgress As Long
    Dim lFileSize As Long
    Dim lIter As Long
    ' ensure that the file exists
    'skip system files!
    If Len(Dir$(fileName, vbArchive Or vbHidden Or vbNormal Or vbReadOnly)) = 0 Then
        RaiseEvent FileError("File " & fileName & " skipped as it is a system file!")
        
        Exit Function
        'Err.Raise 53  ' File not found
    End If
   
    Open fileName For Binary Access Read As #intFree
    
    lFileSize = LOF(intFree)
    RaiseEvent FileSearched(fileName & " size: " & Format$(CLng(lFileSize / 1024), "Standard") & " Kb.")
    'mFileInfo.Text = fileName & " size: " & Format$(CLng(lFileSize / 1024), "Standard") & " Kb."
    
    Remind = lFileSize Mod MaxLen
    lIter = CLng(lFileSize / MaxLen)
    If lIter > 0 Then
        For bytePos = 1 To lFileSize Step MaxLen
            
            DoEvents 'lets user stop process if needed
             If mbStop Then
                Close #intFree
                Exit Function
             End If
            
             'read data
             sContent = ReadData(intFree, bytePos, MaxLen)
            
           'check for match, including previous chars (if any) enough
           'to make the searched string
            If CheckInstr(sPrevBytes & sContent, SearchString, MatchCase, fileName) Then
                ScanChunked = True
                Close #intFree
                RaiseEvent FileJobPercent(100) 'ended with this one
                Exit Function
            End If
            
            'see if ending previus reading plus actual ones can give
            'you what you're looking for
            sPrevBytes = getPrevString(sContent, SearchString)
            RaiseEvent FileJobPercent(CInt(bytePos / lFileSize * 100))
        Next bytePos
    End If
    If Remind > 0 Then
        'get last bytes
        For bytePos = lFileSize - Remind + 1 To lFileSize Step Remind
            'should be one single step...
             If mbStop Then
                Close #intFree
                Exit Function
            End If
            
            'read remind
            sContent = ReadData(intFree, bytePos, Remind)
            
            If CheckInstr(sPrevBytes & sContent, SearchString, MatchCase, fileName) Then
                ScanChunked = True
                Close #intFree
                RaiseEvent FileJobPercent(100)
                Exit Function
             End If
             sPrevBytes = getPrevString(sContent, SearchString)
             RaiseEvent FileJobPercent(CInt(bytePos / lFileSize * 100))   'should be 100
        Next
    End If
    Close #intFree
End Function

Private Function CheckInstr(ByVal sContent As String, ByVal SearchString As String, ByVal MatchCase As Integer, ByVal sFileName As String) As Boolean
    If InStr(1, sContent, SearchString, MatchCase) Then
        RaiseEvent FileFound(sFileName)
        CheckInstr = True
    End If
End Function

Function ReadData(intFree As Integer, bytePos As Long, byteLength As Long) As String
    Dim sBuffer As String
    sBuffer = Space(byteLength)
    Get #intFree, bytePos, sBuffer
    ReadData = sBuffer
End Function

Private Function getPrevString( _
          ByVal sContent As String _
        , ByVal SearchString As String _
        ) As String
    
    If Len(sContent) > Len(SearchString) Then
        getPrevString = Right$(sContent, Len(SearchString))
    ElseIf Len(sContent) > 0 Then
        getPrevString = sContent
    End If

End Function

Public Function writeToFile( _
          ByVal sFileName As String _
        , sContent As String _
        ) As Boolean
    Dim intFree As Integer
    intFree = FreeFile
    If Dir$(sFileName) <> "" Then
      Select Case MsgBox("File " & sFileName & " already exists. Yes: append, No:OverWrite, Cancel: skip saving.", vbYesNoCancel, mMsgTitle)
         Case vbYes
            Open sFileName For Append As #intFree
         Case vbNo
            Open sFileName For Output As #intFree
         Case Else
            Exit Function
      End Select
   Else
      Open sFileName For Output As #intFree
         
   End If
   
   Print #intFree, sContent
   Close #intFree
   writeToFile = True

End Function

Public Function AddSlash(ByVal sPath As String) As String
    sPath = Trim(sPath)
    If Len(sPath) > 0 Then
        If Right$(sPath, 1) <> "/" Then
            If Right$(sPath, 1) <> "\" Then
                sPath = sPath & "\"
            End If
        End If
        AddSlash = sPath
    End If
End Function

Public Property Get MsgTitle() As String
    MsgTitle = mMsgTitle
End Property

Public Property Let MsgTitle(ByVal TitleMsg As String)
   mMsgTitle = TitleMsg
End Property
