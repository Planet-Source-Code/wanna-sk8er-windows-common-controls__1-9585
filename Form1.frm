VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Slider Slider1 
      Height          =   480
      Left            =   60
      TabIndex        =   18
      Top             =   600
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load C:\"
      Height          =   285
      Left            =   60
      TabIndex        =   17
      Top             =   3930
      Width           =   2895
   End
   Begin MSComctlLib.TreeView TV 
      Height          =   2625
      Left            =   60
      TabIndex        =   16
      Top             =   1260
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   4630
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.PictureBox Picture1 
      Height          =   1065
      Index           =   4
      Left            =   4470
      ScaleHeight     =   1005
      ScaleWidth      =   2355
      TabIndex        =   11
      Top             =   3030
      Width           =   2415
      Begin VB.Label Label4 
         Caption         =   "Tab 4"
         Height          =   435
         Left            =   330
         TabIndex        =   15
         Top             =   330
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1275
      Index           =   3
      Left            =   4260
      ScaleHeight     =   1215
      ScaleWidth      =   2325
      TabIndex        =   10
      Top             =   2790
      Width           =   2385
      Begin VB.Label Label3 
         Caption         =   "Tab 3"
         Height          =   675
         Left            =   360
         TabIndex        =   14
         Top             =   300
         Width           =   1365
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1395
      Index           =   2
      Left            =   4050
      ScaleHeight     =   1335
      ScaleWidth      =   2355
      TabIndex        =   9
      Top             =   2550
      Width           =   2415
      Begin VB.Label Label2 
         Caption         =   "Tab 2"
         Height          =   675
         Left            =   450
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Index           =   1
      Left            =   3750
      ScaleHeight     =   1515
      ScaleWidth      =   2445
      TabIndex        =   8
      Top             =   2250
      Width           =   2505
      Begin VB.Label Label1 
         Caption         =   "Tab 1"
         Height          =   885
         Left            =   300
         TabIndex        =   12
         Top             =   240
         Width           =   1785
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2415
      Left            =   3090
      TabIndex        =   7
      Top             =   1800
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   4260
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tab1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tab2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tab3"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tab4"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   4245
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9657
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   176
            MinWidth        =   176
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   176
            MinWidth        =   176
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   60
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   741
      ButtonWidth     =   609
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "blah1"
            Object.ToolTipText     =   "Button1"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "blah2"
            Object.ToolTipText     =   "Button2"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "blah3"
            Object.ToolTipText     =   "Button3"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "blah4"
            Object.ToolTipText     =   "Button4"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "blah5"
            Object.ToolTipText     =   "Button5"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "blah6"
            Object.ToolTipText     =   "Button6"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "blah7"
            Object.ToolTipText     =   "Button7"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "blah8"
            Object.ToolTipText     =   "Button8"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "blah9"
            Object.ToolTipText     =   "Button9"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "ProgressBar"
      Height          =   1245
      Left            =   3090
      TabIndex        =   0
      Top             =   450
      Width           =   4065
      Begin VB.CommandButton Command1 
         Caption         =   "&Start"
         Height          =   375
         Left            =   150
         TabIndex        =   3
         Top             =   360
         Width           =   1035
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   3240
         Top             =   240
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Slow"
         Height          =   285
         Left            =   1470
         TabIndex        =   2
         Top             =   390
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Fast"
         Height          =   255
         Left            =   2280
         TabIndex        =   1
         Top             =   390
         Width           =   855
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   165
         Left            =   120
         TabIndex        =   4
         Top             =   870
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   291
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Numtabs = 4
Dim X As Integer

Const MAX_PATH = 260
Const FILE_ATTRIBUTE_DIRECTORY = &H10

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Type DirInfo
    DirName     As String
End Type

Sub FindDirs(D$, T As TreeView)
    Dim nx As Node, C$
    C$ = CurDir$
    
    ChDir D$
    
    If Len(Dir$("*.*", vbDirectory)) Then
        On Local Error Resume Next
        ChDir ".."
        ChDir ".."
        Set nx = T.Nodes.Add(CurDir$, 4, C$, LastPath$(C$))
        If Err Then
            Set nx = T.Nodes.Add(, , C$, C$)
        End If
        ChDir C$
        ChDir D$
        'Set nx = T.Nodes.Add(C$, 4, , D$)
    Else
        Set nx = T.Nodes.Add(C$, 4, , D$)
    End If
    'T.Nodes(T.Nodes.Count).EnsureVisible
    
    DoEvents
    
    Dim N As Integer, Srch$, i As Integer, NewD$
    
    Srch$ = "*.*"
    ReDim Dees(1 To 10) As DirInfo
    Call LoadDirs(Dees(), N, Srch$)
    If N = 0 Then
        ChDir ".."
        Exit Sub
    End If

    
    For i = 1 To N
        NewD$ = RTrim$(Dees(i).DirName)
        Call FindDirs(NewD$, T)
    Next
    
    ChDir ".."
End Sub

Function LastPath$(P$)
    Dim i
    For i = Len(P$) To 1 Step -1
        If Mid$(P$, i, 1) = "\" Then
            LastPath$ = Mid$(P$, i + 1)
            Exit For
        End If
    Next
End Function

Private Sub LoadDirs(D() As DirInfo, N As Integer, Srch$)
    Dim a$, Max As Integer, i As Integer, k As Integer, W32 As WIN32_FIND_DATA, fHandle As Long, lResult As Long
    Max = UBound(D)
    N = 0
    
    fHandle = FindFirstFile(Srch$, W32)

    If fHandle Then
        Do
            a$ = Left$(W32.cFileName, InStr(W32.cFileName, Chr$(0)) - 1)
            If a$ <> "." And a$ <> ".." And ((W32.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) > 0) Then
                N = N + 1
                If Max < N Then
                    Max = Max + 10
                    ReDim Preserve D(1 To Max) As DirInfo
                End If
                D(N).DirName = a$
            End If
            DoEvents
            lResult = FindNextFile(fHandle, W32)
        Loop While lResult
        lResult = FindClose(fHandle)
    End If

    For i = 1 To N - 1
        For k = i + 1 To N
            If D(i).DirName > D(k).DirName Then
                a$ = D(k).DirName
                D(k).DirName = D(i).DirName
                D(i).DirName = a$
            End If
        Next
    Next
End Sub

Private Sub Command1_Click()
If Option1.Value = True Then
Timer1.Interval = 100
Timer1.Enabled = True
End If
If Option2.Value = True Then
Timer1.Interval = 25
Timer1.Enabled = True
End If
End Sub

Private Sub Command2_Click()
  'TreeView
      Static done
    If done Then Exit Sub
    done = True
    ChDrive "c:\"
    ChDir "c:\"
    Call FindDirs("c:\", TV)
End Sub

Private Sub Form_Load()

'ProgressBar
Option1.Value = True

'TabStrip
  For X = 1 To Numtabs
    
    With Picture1(X)
      .BorderStyle = 0
      .Left = TabStrip1.ClientLeft
      .Top = TabStrip1.ClientTop
      .Width = TabStrip1.ClientWidth
      .Height = TabStrip1.ClientHeight
      .Visible = False
    End With
    
  Next X
  
  Picture1(1).Visible = True

End Sub

Private Sub mnuAbout_Click()
frmAbout.Show 1, Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub TabStrip1_Click()
    
  Static PrevTab As Integer
  PrevTab = Switch(PrevTab = 0, 1, PrevTab >= 1 And PrevTab <= Numtabs, PrevTab)
  Picture1(PrevTab).Visible = False
  Picture1(TabStrip1.SelectedItem.Index).Visible = True
  Picture1(TabStrip1.SelectedItem.Index).Refresh
  PrevTab = TabStrip1.SelectedItem.Index
End Sub

Private Sub Timer1_Timer()
If ProgressBar1.Value > 95 Then
ProgressBar1.Value = 100
Timer1.Enabled = False
Command1.Enabled = False
Else
ProgressBar1.Value = Val(ProgressBar1.Value) + 3
End If
If ProgressBar1.Value = 100 Then
    ProgressBar1.Value = 0
    Command1.Enabled = True
Else
End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'Setting up the toolbar
Select Case Button.Key
Case "blah1"
StatusBar1.Panels(1).Text = "Button 1"
MsgBox "Button 1"
Case "blah2"
StatusBar1.Panels(1).Text = "Button 2"
MsgBox "Button 2"
Case "blah3"
StatusBar1.Panels(1).Text = "Button 3"
MsgBox "Button 3"
Case "blah4"
StatusBar1.Panels(1).Text = "Button 4"
MsgBox "Button 4"
Case "blah5"
StatusBar1.Panels(1).Text = "Button 5"
MsgBox "Button 5"
Case "blah6"
StatusBar1.Panels(1).Text = "Button 6"
MsgBox "Button 6"
Case "blah7"
StatusBar1.Panels(1).Text = "Button 7"
MsgBox "Button 7"
Case "blah8"
StatusBar1.Panels(1).Text = "Button 8"
MsgBox "Button 8"
Case "blah9"
StatusBar1.Panels(1).Text = "Button 9"
MsgBox "Button 9"
Case Else
End Select
End Sub
