VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmStableQuick 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Stable Quick Sort v2.3"
   ClientHeight    =   6870
   ClientLeft      =   2835
   ClientTop       =   1680
   ClientWidth     =   9570
   ClipControls    =   0   'False
   Icon            =   "StableQuick.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   9570
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   960
      Top             =   270
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtDisplay 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4680
      Index           =   0
      Left            =   270
      MaxLength       =   65500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   990
      Width           =   4510
   End
   Begin VB.Frame fraTitle 
      BackColor       =   &H80000005&
      Height          =   780
      Left            =   250
      TabIndex        =   6
      Top             =   60
      Width           =   9105
      Begin VB.Image imgState 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   195
         Left            =   8640
         Picture         =   "StableQuick.frx":0442
         ToolTipText     =   "Maximize"
         Top             =   315
         Width           =   195
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Stable Quick Sort v2.3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   30
         TabIndex        =   7
         Top             =   225
         Width           =   9000
      End
   End
   Begin VB.Frame fraBorder 
      BackColor       =   &H80000005&
      Height          =   960
      Left            =   225
      TabIndex        =   5
      Top             =   5760
      Width           =   9150
      Begin VB.CheckBox chkDesc 
         BackColor       =   &H80000005&
         Caption         =   "Descending Order"
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Top             =   285
         Width           =   1785
      End
      Begin VB.CheckBox chkCaps 
         BackColor       =   &H80000005&
         Caption         =   "List Capitals First"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   285
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CommandButton cmdOpen 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Open file to sort"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3855
         TabIndex        =   0
         Top             =   300
         Width           =   1635
      End
      Begin VB.CommandButton cmdSort 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sort list"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5640
         TabIndex        =   1
         Top             =   300
         Width           =   1635
      End
      Begin VB.CommandButton cmdQuit 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Quit program"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7425
         TabIndex        =   2
         Top             =   300
         Width           =   1485
      End
      Begin VB.Label lblTime 
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   630
         TabIndex        =   10
         Top             =   570
         Width           =   2895
      End
   End
   Begin VB.TextBox txtDisplay 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4680
      Index           =   1
      Left            =   4850
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   990
      Width           =   4510
   End
   Begin VB.Image imgMax 
      Height          =   165
      Left            =   9750
      Picture         =   "StableQuick.frx":0550
      Top             =   1890
      Width           =   165
   End
   Begin VB.Image imgNorm 
      Height          =   165
      Left            =   9720
      Picture         =   "StableQuick.frx":065E
      Top             =   2490
      Width           =   165
   End
End
Attribute VB_Name = "frmStableQuick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function PerfCount Lib "kernel32" Alias "QueryPerformanceCounter" (lpPerformanceCount As Currency) As Long
Private Declare Function PerfFreq Lib "kernel32" Alias "QueryPerformanceFrequency" (lpFrequency As Currency) As Long
Private mCurFreq As Currency

Private sA() As String

Private Sub Form_Load()
    ' Set the default Common Dialog properties
    With dlgCommonDialog
        .CancelError = True
        .InitDir = App.Path
        .Filter = "Text files (*.txt)|*.txt"
        .DefaultExt = "txt"
        ' If the user does not include an extension then '.txt' will
        ' be appended to the filename.
        .FilterIndex = 1 ' text files
        ' If you set more than one file type to the Filter property
        ' you can set the FilterIndex property (indexed from 1) to
        ' specify the default filter in the dialogs 'Files of type'
        ' combo box and determine the initial file types displayed
        ' in the file select listbox.
    End With
End Sub

Private Sub cmdOpen_Click()
    Dim sFileName As String
    Dim sTemp As String
    With dlgCommonDialog
        .DialogTitle = "Open File To Sort..."
        .FileName = vbNullString
        .Flags = 4096 + 4 ' File must exist, no read-only checkbox
        On Error GoTo dlgCancelHandler
        .ShowOpen
        sFileName = .FileName
    End With
    Screen.MousePointer = vbHourglass
    ' Clear the sorted display
    txtDisplay(1).Text = vbNullString
    ' Call the OpenFile function
    sTemp = OpenFile(sFileName)
    ' Load the file into the array
    sA = Split(sTemp, vbCrLf)
    ' Display the file in the text box
    txtDisplay(0).Text = sTemp
    ' Display the filename in the titlebar
    Me.Caption = " Stable QuickSort v2.3 - " & sFileName
    Screen.MousePointer = vbDefault
dlgCancelHandler:
End Sub

Private Sub cmdSort_Click()
    Screen.MousePointer = vbHourglass
    If txtDisplay(0).Text <> vbNullString Then
        txtDisplay(1).Text = PrettySort
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Function OpenFile(sFileSpec As String) As String
    ' Handle errors if they occur
    On Error GoTo GetFileError
    Dim iFile As Integer
    iFile = FreeFile
    ' Open in binary mode, let others read but not write
    Open sFileSpec For Binary Access Read Lock Write As #iFile
    ' Allocate the length first
    OpenFile = Space$(LOF(iFile))
    ' Get the file in one chunk
    Get #iFile, , OpenFile
GetFileError:
    Close #iFile ' Close the file
End Function

Private Function PrettySort() As String
    ' Handle errors if they occur
    On Error GoTo GetFileError
    Dim curElapse As Currency
    Dim r1 As Single
    Dim lb As Long, ub As Long
    Dim bCapsFirst As Boolean
    bCapsFirst = (chkCaps.Value = 1)
    If (chkDesc.Value = 0) Then
        SortOrder = Ascending
    Else '(chkDesc.Value = 1)
        SortOrder = Descending
    End If
    lb = LBound(sA)
    ub = UBound(sA)
    ' Display the array bounds in the titlebar
    Me.Caption = " Stable QuickSort v2.3 - " & lb & " to " & ub
    '+++++++++++++++++++++++
    curElapse = ProfileStart
    strPrettySort sA, lb, ub, bCapsFirst
    r1 = CSng(ProfileStop(curElapse))
    '+++++++++++++++++++++++
    PrettySort = Join(sA, vbCrLf)
    lblTime = "Pretty Sorting took " & Format$(r1, "##0.0000") & " seconds"
GetFileError:
End Function

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Exit Sub
    ElseIf Me.WindowState = vbMaximized Then
        Set imgState.Picture = imgNorm.Picture
    Else
        Set imgState.Picture = imgMax.Picture
    End If
    fraTitle.Left = (Me.Width - fraTitle.Width) \ 2
    txtDisplay(0).Width = (Me.Width - 600) \ 2
    txtDisplay(1).Width = txtDisplay(0).Width
    txtDisplay(1).Left = (Me.Width \ 2) + 20
    fraBorder.Left = (Me.Width - fraBorder.Width) \ 2
    fraBorder.Top = Me.Height - 1485
    txtDisplay(0).Height = ((fraBorder.Top - txtDisplay(1).Top) - 80)
    txtDisplay(1).Height = txtDisplay(0).Height
End Sub

Private Sub imgState_Click()
    If Me.WindowState = vbMaximized Then
        Me.WindowState = vbNormal
        imgState.ToolTipText = "Maximize"
    Else
        Me.WindowState = vbMaximized
        imgState.ToolTipText = "Normal Size"
    End If
End Sub

Private Sub chkCaps_Click()
    If txtDisplay(1).Text <> vbNullString Then
        cmdSort_Click
    End If
End Sub

Private Sub chkDesc_Click()
    If txtDisplay(1).Text <> vbNullString Then
        cmdSort_Click
    End If
End Sub

Private Function ProfileStart() As Currency
    If mCurFreq = 0 Then PerfFreq mCurFreq
    If (mCurFreq) Then PerfCount ProfileStart
End Function

Private Function ProfileStop(ByVal curStart As Currency) As Currency
    If (mCurFreq) Then
        Dim curStop As Currency
        PerfCount curStop
        ProfileStop = (curStop - curStart) / mCurFreq ' cpu tick accurate
        curStop = 0
    End If
End Function

