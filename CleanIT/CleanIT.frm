VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clean IT Out Ver 1.1"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "CleanIT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit Program Now"
      Height          =   375
      Left            =   2055
      TabIndex        =   9
      Top             =   4335
      Width           =   1830
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Clean Selected Items"
      Height          =   375
      Left            =   150
      TabIndex        =   8
      Top             =   4335
      Width           =   1830
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4125
      Left            =   90
      TabIndex        =   0
      Top             =   105
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   7276
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Clean General"
      TabPicture(0)   =   "CleanIT.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line1(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblInfo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkDocuments"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkRunMenu"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkUrls"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkTemp"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkRecBin"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Clean Applications"
      TabPicture(1)   =   "CleanIT.frx":069C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "chkWzip1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "chkWzip2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "chkVb"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "NetScape1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "netscape2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "chkWordPad"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "chkPbrush"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "chkMediaPlayer"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "About Clean IT"
      TabPicture(2)   =   "CleanIT.frx":0A2E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Image1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label4"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label5"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Image2"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Line2(0)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Line2(1)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Text1"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   990
         Left            =   -74895
         MultiLine       =   -1  'True
         TabIndex        =   19
         Text            =   "CleanIT.frx":0DC0
         Top             =   1290
         Width           =   4680
      End
      Begin VB.CheckBox chkMediaPlayer 
         Caption         =   "Clean Items From Media Player Recent File Menu"
         Height          =   255
         Left            =   -74640
         TabIndex        =   18
         Top             =   2925
         Width           =   4095
      End
      Begin VB.CheckBox chkPbrush 
         Caption         =   "Clean Items Form Paint Brush Recent File Menu"
         Height          =   255
         Left            =   -74640
         TabIndex        =   17
         Top             =   2670
         Width           =   3705
      End
      Begin VB.CheckBox chkWordPad 
         Caption         =   "Clean Items FromWord Pad Recent File Menu"
         Height          =   255
         Left            =   -74640
         TabIndex        =   16
         Top             =   2415
         Width           =   3705
      End
      Begin VB.CheckBox netscape2 
         Caption         =   "Clean Items From Netscape Cache Folder"
         Height          =   255
         Left            =   -74640
         TabIndex        =   15
         Top             =   2160
         Width           =   3705
      End
      Begin VB.CheckBox NetScape1 
         Caption         =   "Clean Netscape Internet Typed Web Address"
         Height          =   255
         Left            =   -74640
         TabIndex        =   14
         Top             =   1875
         Width           =   3705
      End
      Begin VB.CheckBox chkVb 
         Caption         =   "Clean Visual Basic Recent Opened files"
         Height          =   255
         Left            =   -74640
         TabIndex        =   13
         Top             =   1605
         Width           =   3435
      End
      Begin VB.CheckBox chkWzip2 
         Caption         =   "Clean Items From WinZip Extract To Menu"
         Height          =   255
         Left            =   -74640
         TabIndex        =   12
         Top             =   1350
         Width           =   3435
      End
      Begin VB.CheckBox chkWzip1 
         Caption         =   "Clean Items From WinZip File Menu"
         Height          =   255
         Left            =   -74640
         TabIndex        =   11
         Top             =   1080
         Width           =   3045
      End
      Begin VB.CheckBox chkRecBin 
         Caption         =   "Empty Trash From Recycle Bin"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   2220
         Width           =   2610
      End
      Begin VB.CheckBox chkTemp 
         Caption         =   "Clean Items From Temp Folder"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   1935
         Width           =   2610
      End
      Begin VB.CheckBox chkUrls 
         Caption         =   "Clean Internet Typed Web Address"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   1650
         Width           =   3030
      End
      Begin VB.CheckBox chkRunMenu 
         Caption         =   "Clean Recsent Items From Run Menu"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   1365
         Width           =   3075
      End
      Begin VB.CheckBox chkDocuments 
         Caption         =   "Clean Recent Doucements List"
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Width           =   2610
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   -74730
         X2              =   -70470
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         Index           =   0
         X1              =   -74745
         X2              =   -70485
         Y1              =   2325
         Y2              =   2325
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   -73860
         Picture         =   "CleanIT.frx":0EAC
         Top             =   585
         Width           =   2805
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "website dreamvb.s5.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -73410
         TabIndex        =   22
         Top             =   3195
         Width           =   1755
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Email dreamvb@yahoo.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -73485
         TabIndex        =   21
         Top             =   2865
         Width           =   1980
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Writen and designed by Ben Jones"
         Height          =   195
         Left            =   -73665
         TabIndex        =   20
         Top             =   2550
         Width           =   2475
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   -74730
         Picture         =   "CleanIT.frx":A6C4
         Top             =   615
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clean Applications Recent Lists"
         Height          =   195
         Left            =   -74790
         TabIndex        =   10
         Top             =   720
         Width           =   2235
      End
      Begin VB.Label lblInfo 
         ForeColor       =   &H000000FF&
         Height          =   1005
         Left            =   255
         TabIndex        =   7
         Top             =   2775
         Width           =   4590
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   255
         X2              =   4875
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   255
         X2              =   4875
         Y1              =   2625
         Y2              =   2625
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Clean Recent lists."
         Height          =   195
         Left            =   210
         TabIndex        =   1
         Top             =   690
         Width           =   1320
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Check1_Click()
    lblInfo.Caption = ""
    lblInfo.Caption = "This option will remove any item, files or folders from the windows Recycle Bin"
    
End Sub

Private Sub chkDocuments_Click()
    lblInfo.Caption = ""
    lblInfo.Caption = "This option will clean all items form the Recent Doucements menu you may have opened"
End Sub

Private Sub chkRunMenu_Click()
    lblInfo.Caption = ""
    lblInfo.Caption = "This option will remove any program name you have typed in the run menu"
    
End Sub

Private Sub chkTemp_Click()
    lblInfo.Caption = ""
    lblInfo.Caption = "This option removes all files form the TEMP folder that someone install programs leave behind"
    
End Sub

Private Sub chkUrls_Click()
    lblInfo.Caption = ""
    lblInfo.Caption = "This option will remove all web URL'S you have typed in Internet Explorers address bar"
    End Sub

Private Sub Command1_Click()
    Ans = MsgBox("Do you wnat to clean your system now", _
    vbYesNo)
    If Ans = vbNo Then
        Exit Sub
    Else
        Me.MousePointer = 11
        If chkDocuments.Value = 1 Then
            CleanIT.CleanDocs
        End If
        
        If chkRunMenu Then
            CleanIT.CleanRunMnu
        End If
        
        If chkUrls Then
            CleanIT.CleanIEUrls
        End If
        
        If chkTemp Then
            CleanIT.CleanTemp
        End If
        
        If chkRecBin Then
            CleanIT.EmptyBin Me.hwnd
        End If
        
        If chkWzip1 Then
            CleanIT.CleanWZipFile
        End If
        
        If chkWzip2 Then
            CleanIT.CleanWZipExtract
        End If
        
        If chkVb Then
            CleanIT.VBFileMenu
        End If
        
        If NetScape1 Then
            CleanIT.CleanNetScapeUrls
        End If
        
        If netscape2 Then
            CleanIT.CleanNetScapeCache
        End If
        
        If chkWordPad Then
            CleanIT.CleanWordPadFileLst
        End If
    
        If chkPbrush Then
            CleanIT.CleanPaintBrushFileLst
        End If
        
        If chkMediaPlayer Then
            CleanIT.CleanMediaPlayer
        End If
        Me.MousePointer = 0
        MsgBox "Your system has now been cleaned", vbInformation, "Done...."
        End If
        
   
End Sub

Private Sub Command2_Click()
Dim Ans
    Ans = MsgBox("Do you want to exit now", _
    vbYesNo)
    If Ans = vbnow Then
        Exit Sub
    Else
        Unload Form1: End
    End If
    
End Sub

