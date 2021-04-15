VERSION 5.00
Object = "{DF2BBE39-40A8-433B-A279-073F48DA94B6}#1.0#0"; "axvlc.dll"
Begin VB.Form Form1 
   Caption         =   "Cam Handle"
   ClientHeight    =   10155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12330
   LinkTopic       =   "Form1"
   ScaleHeight     =   10155
   ScaleWidth      =   12330
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   4575
      Left            =   240
      TabIndex        =   4
      Top             =   5160
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Text            =   "http://192.168.1.3:4747/mjpegfeed?320x240"
      Top             =   4560
      Width           =   4215
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "Capture"
      Height          =   495
      Left            =   10920
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "On"
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   4560
      Width           =   855
   End
   Begin AXVLCCtl.VLCPlugin2 cam1 
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5415
      AutoLoop        =   0   'False
      AutoPlay        =   -1  'True
      Toolbar         =   0   'False
      ExtentWidth     =   9551
      ExtentHeight    =   7223
      MRL             =   ""
      Object.Visible         =   -1  'True
      Volume          =   0
      StartTime       =   0
      BaseURL         =   ""
      BackColor       =   0
      FullscreenEnabled=   0   'False
      Branding        =   -1  'True
   End
   Begin VB.Image img2 
      BorderStyle     =   1  'Fixed Single
      Height          =   4575
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   6615
   End
   Begin VB.Image img1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4095
      Left            =   6000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmd1_Click()
If cmd1.Caption = "On" Then
    cmd1.Caption = "Off"
    cam1.playlist.Add (Text1.Text)
    cam1.playlist.play
Else
    cmd1.Caption = "On"
    cam1.playlist.stop
End If

End Sub

Private Sub cmd2_Click()
Dim resp As Variant
Dim ls_prev_path As String
Dim ls_temp_path As String
Dim pic As IPictureDisp


ls_prev_path = CurDir
ls_temp_path = App.Path + IIf(Right(App.Path, 1) = "\", "", "\")
ls_temp_path = ls_temp_path + "tmp_snapshot\"

If Dir(ls_temp_path, vbDirectory) <> "" Then
    If Dir(ls_temp_path + "*.bmp") <> "" Then Kill ls_temp_path + "*.bmp"
Else
    MkDir ls_temp_path
End If

ChDir ls_temp_path

'capture
cam1.playlist.pause
cam1.video.takeSnapshot
cam1.playlist.play

DoEvents

If Dir(ls_temp_path + "*.*") <> "" Then
    ls_pic_file = ls_temp_path + Dir(ls_temp_path + "*.*")
End If

Filename = ls_pic_file
fileJpg = App.Path & "\gbr\" & Format(Now, "YYYYMMdd HHmmss") & ".jpg"
            
 'convert bmp to jpg
dib = FreeImage_LoadEx(Filename)
If (dib) Then
   Call FreeImage_SaveEx(dib, fileJpg)
   Call FreeImage_Unload(dib)
End If

'delete file bmp
Kill Filename
        
'tampil hasil capture
If Dir$(fileJpg) <> "" Then
    img1.Picture = LoadPicture(fileJpg)
End If
            
ChDir ls_prev_path
File1.Refresh
End Sub

Private Sub File1_Click()
    img2.Picture = LoadPicture(App.Path & "\gbr\" & File1.Filename)
End Sub

Private Sub Form_Load()
File1 = App.Path & "\gbr"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If cmd1.Caption = "On" Then
cam1.playlist.stop
End If
End Sub
