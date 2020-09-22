VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmSearchInFiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search In files --CesareI"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   Icon            =   "FrmSearchInFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdExportMsg 
      Caption         =   "Export Msg"
      Height          =   330
      Left            =   8100
      TabIndex        =   21
      Top             =   1125
      Width           =   1080
   End
   Begin MSComctlLib.ProgressBar PBarFolder 
      Height          =   240
      Left            =   0
      TabIndex        =   19
      Top             =   6525
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar PBarFile 
      Height          =   240
      Left            =   0
      TabIndex        =   18
      Top             =   6825
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar SBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   17
      Top             =   6765
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Key             =   "pBar"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   12144
            Key             =   "LblPath"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   7650
      Top             =   1050
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdChoose 
      Caption         =   "..."
      Height          =   255
      Left            =   9000
      TabIndex        =   15
      Top             =   2460
      Width           =   315
   End
   Begin VB.TextBox TxtExport 
      Height          =   285
      Left            =   4620
      TabIndex        =   14
      Top             =   2460
      Width           =   4335
   End
   Begin VB.CommandButton CmdClearMsg 
      Caption         =   "Clear Msg"
      Height          =   330
      Left            =   5850
      TabIndex        =   13
      Top             =   1125
      Width           =   1080
   End
   Begin VB.CommandButton CmdExport 
      Caption         =   "Export List"
      Height          =   330
      Left            =   8100
      TabIndex        =   12
      Top             =   720
      Width           =   1080
   End
   Begin VB.CheckBox ChkCase 
      Caption         =   "Match Case"
      Height          =   315
      Left            =   3750
      TabIndex        =   11
      Top             =   750
      Width           =   1965
   End
   Begin VB.TextBox TxtMsg 
      BackColor       =   &H80000018&
      ForeColor       =   &H80000017&
      Height          =   840
      Left            =   3780
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   1575
      Width           =   5550
   End
   Begin VB.CommandButton CmdStop 
      Caption         =   "Stop"
      Height          =   330
      Left            =   6975
      TabIndex        =   9
      Top             =   720
      Width           =   1080
   End
   Begin VB.ListBox LstResults 
      Height          =   3765
      Left            =   0
      TabIndex        =   8
      Top             =   2760
      Width           =   9360
   End
   Begin VB.CommandButton CmdSearch 
      Caption         =   "Search"
      Height          =   330
      Left            =   5850
      TabIndex        =   7
      Top             =   720
      Width           =   1080
   End
   Begin VB.TextBox TxtExtensions 
      Height          =   285
      Left            =   4725
      TabIndex        =   6
      Text            =   "cls,frm"
      Top             =   375
      Width           =   4620
   End
   Begin VB.CheckBox ChkRecurse 
      Caption         =   "Recurse subfolders"
      Height          =   315
      Left            =   3750
      TabIndex        =   4
      Top             =   1125
      Width           =   1965
   End
   Begin VB.TextBox TxtSearch 
      Height          =   315
      Left            =   4725
      TabIndex        =   2
      Top             =   0
      Width           =   4620
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   0
      TabIndex        =   1
      Top             =   375
      Width           =   3690
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3690
   End
   Begin VB.Label LblFolder 
      AutoSize        =   -1  'True
      Caption         =   "000 % files in folder scanned:"
      Height          =   195
      Left            =   7200
      TabIndex        =   20
      Top             =   6525
      Width           =   2130
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Export to"
      Height          =   195
      Left            =   3840
      TabIndex        =   16
      Top             =   2520
      Width           =   630
   End
   Begin VB.Label LblFileTypes 
      AutoSize        =   -1  'True
      Caption         =   "Files types"
      Height          =   195
      Left            =   3750
      TabIndex        =   5
      Top             =   375
      Width           =   735
   End
   Begin VB.Label LblSearchFor 
      AutoSize        =   -1  'True
      Caption         =   "Search for"
      Height          =   195
      Left            =   3750
      TabIndex        =   3
      Top             =   75
      Width           =   735
   End
End
Attribute VB_Name = "FrmSearchInFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hi all.
'This will search inside files for a string.
'Useful if you get some duty like the following:
'
'"Find all docs or modules (cls or bas) or txt files
'where you could read inside a specific word;
'they should be around three of the over 2000 you might
'have on your Pc...Find the string by opening
'the docs, txt or vbp one by one, and write down a list
'of file and their location..."
'
'Well, no thanks. I built this, instead, and runned
'searching in "cls" and "bas" and "txt" and "doc" of
'main folder where files laied. Much, much better...
'
'It can search in any kind of file, but if you try
'with compiled ones (exes, dll) or zipped ones, it
'will most likely not find what you're looking for...
'
'This code is built a bit in a hurry, but I tested
'severely and (at least with Vb6 - Vb5 users will have to
'substitute split function - and on win2k ) it is working
'fine. So you may find you are able to make it a bit cleaner.
'If so, do it and post on Planet source code under your name:
'if I step into yours, I will rate you ;-)
'
'The ScanChunked function is based on a function I saw
'herearound, which I rated, but had a mistake:
'stepped for x=1 to lenOfFile, step 32000
'and did not check for remainder or if file was
'shorter.
'This code is also a fix for that one.
'Hope you enjoy,
'Cesare Imperiali

Option Explicit

Private WithEvents SearchInFiles As ClsSearchInFile
Attribute SearchInFiles.VB_VarHelpID = -1

Const TheTitle As String = "Search in Files utility"

Private Sub CmdChoose_Click()
   Dim sFileN As String
   sFileN = ChooseFile
   If Len(sFileN) > 0 Then
      TxtExport.Text = sFileN
   End If
End Sub

Private Sub CmdClearMsg_Click()
    TxtMsg.Text = ""
End Sub

Private Sub CmdExport_Click()
    
    If LstResults.ListCount = 0 Then
       MsgBox "No list to export!", vbOKOnly + vbInformation, TheTitle
       Exit Sub
    End If
    
    If Len(TxtExport.Text) > 0 Then
       
       Dim sList As String
       Dim lngCounter As Long
       
       sList = String(20, "*") & vbCrLf
       sList = sList & Now() & vbCrLf
       sList = sList & "Files (type: """ & TxtExtensions.Text & """) found containing """ & TxtSearch.Text & """" & vbCrLf
       For lngCounter = 0 To LstResults.ListCount - 1
          sList = sList & LstResults.List(lngCounter) & vbCrLf
       Next
       sList = sList & String(20, "*") & vbCrLf
       
       If SearchInFiles.writeToFile(TxtExport.Text, sList) Then
            MsgBox "List saved in " & TxtExport.Text, vbOKOnly, TheTitle
       End If
    Else
       MsgBox "choose export file", vbOKOnly + vbInformation, TheTitle
    End If
End Sub

Private Sub CmdExportMsg_Click()
   Dim sFileN As String
   
   If Len(TxtMsg.Text) = 0 Then
        MsgBox "No messages to export", vbOKOnly + vbInformation, TheTitle
        Exit Sub
   End If
   
   sFileN = ChooseFile
   If Len(sFileN) > 0 Then
      'try
      If SearchInFiles.writeToFile(sFileN, TxtMsg.Text) Then
            MsgBox "Messages saved in " & sFileN, vbOKOnly, TheTitle
      End If
   End If
End Sub

Private Sub Form_Load()
    
    Set SearchInFiles = New ClsSearchInFile
    SearchInFiles.MsgTitle = TheTitle
    TxtExport.Text = SearchInFiles.AddSlash(App.Path) & "aLstSrch.txt"
    
    PBarFile.Visible = False
    PBarFolder.Visible = False
    ChkRecurse.Value = vbChecked
    
    PBarFile.Move SBar.Panels("pBar").Left, SBar.Top + (SBar.Height - PBarFile.Height) / 2, SBar.Panels("pBar").Width
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SearchInFiles.bStop = True
End Sub

Private Sub CmdSearch_Click()
    Dim Fso As Scripting.FileSystemObject
   
    Dim aFolder As Scripting.Folder
    Dim sFilter() As String
    Dim sPath As String
    
    Dim MatchCase As Integer
    sPath = SearchInFiles.AddSlash(Dir1.Path)
    
    LstResults.Clear
    TxtMsg.Text = "Search for """ & TxtSearch.Text & """" & vbCrLf
    If Len(sPath) > 0 Then
        CmdSearch.Enabled = False
        Set Fso = New Scripting.FileSystemObject
            If Fso.FolderExists(sPath) Then
                Set aFolder = Fso.GetFolder(sPath)

                'remove spaces and convert all to lowercase and split string in array
                sFilter = Split(LCase(Replace(TxtExtensions.Text, " ", "")), ",")
                Screen.MousePointer = vbArrowHourglass
                   
                    SearchInFiles.bStop = False
                    
                    setPBar PBarFile, 0, True
                    setPBar PBarFolder, 0, True
                    LblFolder.Caption = "000% of files in folder scanned"
                    LblFolder.Visible = True
                    If ChkCase.Value = vbChecked Then
                        MatchCase = vbBinaryCompare
                    Else
                        MatchCase = vbTextCompare
                    End If
                    
                    Call SearchInFiles.SearchInFiles(aFolder _
                                                   , sFilter _
                                                   , TxtSearch.Text _
                                                   , MatchCase _
                                                   , CBool(ChkRecurse.Value = vbChecked))
                    LblFolder.Visible = False
                    setPBar PBarFile, 0, False
                    setPBar PBarFolder, 0, False
                    SBar.Panels("LblPath").Text = ""
                Screen.MousePointer = vbDefault
                Set aFolder = Nothing
            End If
            
        Set Fso = Nothing
        CmdSearch.Enabled = True
    End If
End Sub

Private Sub CmdStop_Click()
    SearchInFiles.bStop = True
End Sub

Private Sub Drive1_Change()
    On Error GoTo errHandler
    Dir1.Path = Drive1.Drive
    Exit Sub
errHandler:
    Drive1.Drive = Left$(Dir1.Path, 1) & ":"
End Sub





Private Function ChooseFile() As String
   With CDialog
      .Filter = "Text (*.txt)|*.txt|All (*.*)|*.*"
      .Flags = cdlOFNHideReadOnly Or cdlOFNNoChangeDir
      .ShowOpen
       ChooseFile = CDialog.fileName
   End With
End Function

Private Sub setPBar( _
          ByVal PBar As ProgressBar _
        , iValue As Integer _
        , bVisible As Boolean)
    
    With PBar
        .Value = iValue
        .Visible = bVisible
    End With
    
End Sub

'events from class
Private Sub SearchInFiles_FileFound(ByVal FilePathName As String)
   'file containing searched val
   LstResults.AddItem FilePathName
   LstResults.Refresh
End Sub

Private Sub SearchInFiles_FileJobPercent(ByVal FileRatio As Integer)
   'percent of single file analized
   PBarFile.Value = FileRatio
End Sub


Private Sub SearchInFiles_FileSearched(ByVal FileInfo As String)
   'infos about file now scanned (path and size)
   SBar.Panels("LblPath").Text = FileInfo
End Sub

Private Sub SearchInFiles_FileError(ByVal ErrDescr As String)
   'info on errors accessing file for reading
   TxtMsg.Text = TxtMsg.Text & ErrDescr & vbCrLf
   TxtMsg.SelStart = Len(TxtMsg.Text)
End Sub

Private Sub SearchInFiles_FolderError(ByVal ErrDescr As String)
   'info on error processing a folder
    TxtMsg.Text = TxtMsg.Text & ErrDescr & vbCrLf
    TxtMsg.SelStart = Len(TxtMsg.Text)
End Sub

Private Sub SearchInFiles_FolderJobPercent(ByVal FolderRatio As Integer)
   'percent of files in folder analized
   PBarFolder.Value = FolderRatio
   LblFolder.Caption = Format$(FolderRatio, "000") & "% of files in folder scanned"
End Sub

Private Sub SearchInFiles_FolderSearched(ByVal FolderInfo As String)
   'info about folder now searched
   TxtMsg.Text = TxtMsg.Text & FolderInfo & vbCrLf
   TxtMsg.SelStart = Len(TxtMsg.Text)
End Sub
