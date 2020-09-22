VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                               Auto Wallpaper"
   ClientHeight    =   1335
   ClientLeft      =   2625
   ClientTop       =   1485
   ClientWidth     =   4230
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4230
   WindowState     =   1  'Minimized
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   3975
      Begin VB.Timer Timer1 
         Interval        =   60000
         Left            =   3240
         Top             =   360
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Hide"
         Height          =   375
         Left            =   1200
         TabIndex        =   19
         Top             =   3840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   3015
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   18
         Text            =   "Form1.frx":08CA
         Top             =   1200
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Info"
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         ToolTipText     =   "Information"
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Set Directory"
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         ToolTipText     =   "To Select The Directory Of Background Files (*.bmp)"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   1695
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Always Run Auto"
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         ToolTipText     =   "To Enable Change Desktop Wallpaper Automatically While Windows Starts."
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Never Run"
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         ToolTipText     =   "To Disable Run While Windows Starts."
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Caption         =   "EXIT"
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         ToolTipText     =   "Exit To Windows System"
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NB: Selected Directory Must Be A  *.Bmp Directory."
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   3960
         Width           =   3630
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "CopyRight By General Corporation                      Bangladesh"
         Height          =   495
         Left            =   840
         TabIndex        =   16
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "CopyRight By General Corporation"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   720
         TabIndex        =   15
         Top             =   4200
         Width           =   2415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1"
         Height          =   195
         Left            =   1560
         TabIndex        =   14
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auto Wallpaper"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1080
         TabIndex        =   13
         Top             =   120
         Width           =   1860
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Walpaper"
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Text            =   "0"
      Top             =   1800
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Hidden          =   -1  'True
      Left            =   6000
      Pattern         =   "*.bmp"
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open"
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   960
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
 Const HKEY_LOCAL_MACHINE = &H80000002
 Const SPIF_UPDATEINIFILE = &H1
Const SPI_SETDESKWALLPAPER = 20
Const SPIF_SENDWININICHANGE = &H2
Public Sub SaveString(hKey As Long, StrPath As String, StrValue As String, StrData As String)
   Dim KeyH&
    r = RegCreateKey(hKey, StrPath, KeyH&)
    r = RegSetValueEx(KeyH&, StrValue, 0, 1, ByVal StrData, Len(StrData))
    r = RegCloseKey(KeyH&)
End Sub





Private Sub Command1_Click()
Dim FileName As String
    Dim Y As Long
 On Error Resume Next

Kill ("c:\windows\aspin.txt")
Text1.Text = Dir1.Path
FileNumber = FreeFile
FileName = "c:\windows\aspin.txt"
Open FileName For Append As #FileNumber
Print #FileNumber, Text1.Text
Close #FileNumber
Timer1.Enabled = False
File1.Path = Dir1.Path
 FileName = (Dir1.Path & "\" & File1.List(0))
Y = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, FileName, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)


 End Sub

Private Sub Command2_Click()
MsgBox File1.List(Text2.Text)
End Sub

Private Sub Command3_Click()
On Error Resume Next
If Command3.Value = 1 Then
SaveSetting App.Title, App.Title, "RunWithSystem", 1
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "AutoWalpaper", "AutoWal"
Else
SaveSetting App.Title, App.Title, "RunWithSystem", 0
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "AutoWalpaper", "C:\Windows\AutoWal.exe"
End If
End Sub

Private Sub Command4_Click()
On Error Resume Next
If Command4.Value = 1 Then
SaveSetting App.Title, App.Title, "RunWithSystem", 1
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "AutoWalpaper", "AutoWal"
Else
SaveSetting App.Title, App.Title, "RunWithSystem", 0
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "AutoWalpaper", "0"
End If
End Sub

Private Sub Command5_Click()
 Dim FileName As String
    Dim Y As Long
    FileName = (Dir1.Path & "\" & File1.List(Text2.Text))
   Y = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, FileName, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
    End Sub

Private Sub Command6_Click()
Dim ans As Variant
ans = MsgBox("Do You Really Want To Quit?..", vbYesNo, "Quit")
If ans = vbNo Then
Load Me
Else
End
End If
End Sub

Private Sub Command7_Click()
On Error Resume Next
Text3.Visible = True
Command8.Visible = True
Timer1.Enabled = False
End Sub

Private Sub Command8_Click()
Text3.Visible = False
Command8.Visible = False
End Sub


Private Sub Dir1_Change()
Command1.Enabled = True
End Sub

Private Sub Drive1_Change()
On Error GoTo err
Dir1.Path = Drive1.Drive
Exit Sub
err:
MsgBox "Device Is Not Available.", 16, "Device Error!!"
End Sub

Private Sub Form_Load()
Dim count As Variant
On Error Resume Next
Command1.Enabled = False
FileNumber = FreeFile
FileName = "c:\windows\num.txt"
Open "c:\windows\num.txt" For Input As #FileNumber
Text2.Text = Input(LOF(1), 1)
Close #1

count = (Len(Text2.Text) - 3)
Text2.Text = Left(Text2.Text, count)





FileNumber = FreeFile
FileName = "c:\windows\aspin.txt"
Open "c:\windows\aspin.txt" For Input As #FileNumber
Text1.Text = Input(LOF(1), 1)
Close #1
count = (Len(Text1.Text) - 2)
Text1.Text = Left(Text1.Text, count)
Dir1.Path = Text1.Text
File1.Path = Dir1.Path
Command1.Enabled = False



count = (Len(Text2.Text) - 3)
Text2.Text = Left(Text2.Text, count)
If Text2.Text >= File1.ListCount Then
Text2.Text = 0
End If

If File1.ListCount <= 0 Then
MsgBox "No *.BMP Directory Selected. Select A *.Bmp Filegroup.", 16, "No BMP File Found."
Form1.WindowState = 0
Form1.Height = 4850
Frame1.Height = 4445
End If





Kill ("c:\windows\num.txt")
FileNumber = FreeFile
FileName = "c:\windows\num.txt"
Open FileName For Append As #FileNumber
Print #FileNumber, Text2.Text + 1
Close #FileNumber


FileCopy (App.Path & "\Autowal.exe"), ("c:\windows\Autowal.exe")
Command3_Click
Command5_Click

End Sub




Private Sub frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Height = 4850
Frame1.Height = 4445
End Sub


Private Sub Timer1_Timer()
End
End Sub

