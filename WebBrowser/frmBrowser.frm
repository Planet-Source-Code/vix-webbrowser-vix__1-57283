VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBrowser 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   7200
   ClientLeft      =   3060
   ClientTop       =   3345
   ClientWidth     =   10380
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmBrowser.frx":0442
   ScaleHeight     =   7200
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox google 
      Height          =   315
      Left            =   8520
      TabIndex        =   8
      Text            =   "Google"
      Top             =   480
      Width           =   1755
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stop"
      Height          =   255
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Refresh"
      Height          =   255
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdForward 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Forward"
      Height          =   255
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back"
      Height          =   255
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox footer 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      Picture         =   "frmBrowser.frx":4157
      ScaleHeight     =   615
      ScaleWidth      =   1395
      TabIndex        =   2
      Top             =   6480
      Width           =   1395
   End
   Begin VB.ComboBox cboAddress 
      Height          =   315
      Left            =   1670
      TabIndex        =   1
      Top             =   480
      Width           =   3795
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   1335
      Left            =   1665
      TabIndex        =   0
      Top             =   840
      Width           =   5520
      ExtentX         =   9737
      ExtentY         =   2355
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   8040
      Top             =   360
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   1665
      TabIndex        =   3
      Top             =   6960
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7680
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/-----------------------------------------------------------------------------------------------------------\'
'|---=|                                     Author: Vix aka Nutz                                        |=---|'
'|---=| Commants: I made this just to show a more graphical browser, feel free to use it for what ever. |=---|'
'|---=|                                Site: http://Trackers-Alliance.com                               |=---|'
'\-----------------------------------------------------------------------------------------------------------/'




'/-------------------------------------------------------------------------\'
'|---=| Allows us to move the form by clicking on the form background |=---|'
'\-------------------------------------------------------------------------/'
Option Explicit

Public StartingAddress As String
Dim mbDontNavigateNow As Boolean


Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Const WM_NCLBUTTONDOWN = &HA1
    Private Const HTCAPTION = 2


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
'/-------------------------------------------------------------------------\'
'|------------------------------=| End |=----------------------------------|'
'\-------------------------------------------------------------------------/'



'/----------------------------------------------------------------------------------\'
'|---=| Soon as the browser changes/starts to load the progressBar changes too |=---|'
'\----------------------------------------------------------------------------------/'
Private Sub brwWebBrowser_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    On Error Resume Next
    ProgressBar1.Max = ProgressMax
    ProgressBar1.Value = Progress
    ProgressBar1.Refresh
End Sub
'/----------------------------------------------------------------------------------\'
'|---------------------------------=| End |=----------------------------------------|'
'\----------------------------------------------------------------------------------/'




'/---------------------------------------------------------\'
'|---=| These are just very easy web browser commands |=---|'
'\---------------------------------------------------------/'

'-----=| Back |=-----'
Private Sub cmdBack_Click()
On Error Resume Next
brwWebBrowser.GoBack
End Sub

'-----=| Forward |=-----'
Private Sub cmdForward_Click()
On Error Resume Next
brwWebBrowser.GoForward
End Sub

'-----=| Refresh |=-----'
Private Sub cmdRefresh_Click()
On Error Resume Next
brwWebBrowser.Refresh
End Sub

'-----=| Stop |=-----'
Private Sub cmdStop_Click()
On Error Resume Next
brwWebBrowser.Stop
End Sub
'/---------------------------------------------------------\'
'|--------------------------=| End |=----------------------|'
'\---------------------------------------------------------/'




'/-----------------------------------------------\'
'|---=| GO to starting address on form load |=---|'
'\-----------------------------------------------/'
Private Sub Form_Load()
    On Error Resume Next
    brwWebBrowser.Navigate ("http://google.com")

End Sub
'/-----------------------------------------------\'
'|-------------------=| End |=-------------------|'
'\-----------------------------------------------/'



'/-----------------------------------------\'
'|---=| When the page is fully loaded |=---|'
'\-----------------------------------------/'
Private Sub brwWebBrowser_DocumentComplete(ByVal pDisp As Object, URL As Variant)
'--- ..This grabs the page's name and location, once the page has loaded then it is displayed in the form caption .. ---'
frmBrowser.Caption = brwWebBrowser.LocationName & " <-> " & brwWebBrowser.LocationURL & " - Vix"
'--- .. This shows the loaded page's location in the address bar.. ---'
cboAddress = brwWebBrowser.LocationURL
End Sub
'/-----------------------------------------\'
'|----------------End----------------------|'
'\-----------------------------------------/'



Private Sub footer_Click()
Unload Me
End Sub

Private Sub brwWebBrowser_DownloadComplete()
    On Error Resume Next
    Me.Caption = brwWebBrowser.LocationName
End Sub

'/------------------------------------------------------\'
'|---=| Takes you to the address in the adress bar |=---|'
'\------------------------------------------------------/'
Private Sub cboAddress_Click()
    If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
    brwWebBrowser.Navigate cboAddress.Text
End Sub
Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub
'/------------------------------------------------------\'
'|---------------------=| End |=------------------------|'
'\------------------------------------------------------/'




'/------------------------------------------------------\'
'|---=| This is just getting things neat on resize |=---|'
'\------------------------------------------------------/'
Private Sub Form_Resize()
    On Error Resume Next
    cboAddress.Width = Me.ScaleWidth - 3700
    brwWebBrowser.Width = Me.ScaleWidth - 1800
    brwWebBrowser.Height = Me.ScaleHeight - 1050
    footer.Left = 0
    footer.Top = Me.ScaleHeight - footer.Height - 0
    ProgressBar1.Width = Me.ScaleWidth - 1820
    ProgressBar1.Top = Me.ScaleHeight - ProgressBar1.Height - 50
    google.Left = Me.ScaleWidth - google.Width - 140
    Line1.X2 = Me.Width
End Sub
'/------------------------------------------------------\'
'|----------------------=| End |=-----------------------|'
'\------------------------------------------------------/'




'/---------------------------------------------------\'
'|---=| This is the fire fox like google search |=---|'
'\---------------------------------------------------/'
Private Sub google_click()
    
    If mbDontNavigateNow Then Exit Sub
    '-- ..Start the timer.. --'
    timTimer.Enabled = True
    '--- ..This just uses google's URL string with our text as key words.. ---'
    brwWebBrowser.Navigate "http://www.google.com/search?hl=en&lr=&q=" & google.Text
    '-- ..After the click the text is added to the drop down box.. ---'
    google.AddItem google.Text
End Sub

Private Sub google_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    '--- ..When you press Enter/Return inside the box it runs 'google_click'.. ---'
    If KeyAscii = vbKeyReturn Then
        google_click
    End If
End Sub
'/---------------------------------------------------\'
'|----------------------=| End |=---------------------|'
'\---------------------------------------------------/'




'/------------------------------------------------------------\'
'|---=| This is that timer you've been hearing all about |=---|'
'\------------------------------------------------------------/'
'--- ..This checks to see if the browser is busy and if it is sets the form caption to 'Working...'..---'
Private Sub timTimer_Timer()
    If brwWebBrowser.Busy = False Then
        '--- ..If not turn off.. ---'
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser.LocationName & " <-> " & brwWebBrowser.LocationURL & " - Vix"
    Else
        Me.Caption = "Working..."
    End If
End Sub
'/-----------------------------------------------------------\'
'|------------------------=| End |=--------------------------|'
'\-----------------------------------------------------------/'
