VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AMA Computer University - Student Ebook(Expert System)"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   7800
   ScaleWidth      =   11295
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox tempWord 
      Height          =   375
      Left            =   3480
      TabIndex        =   30
      Top             =   24000
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":94A4
   End
   Begin VB.Timer tmrEntrance 
      Interval        =   10
      Left            =   9240
      Top             =   5400
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   8280
      Picture         =   "frmMain.frx":9526
      ScaleHeight     =   1695
      ScaleWidth      =   2895
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   2895
      Begin VB.PictureBox picCancel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         Picture         =   "frmMain.frx":A256
         ScaleHeight     =   255
         ScaleWidth      =   735
         TabIndex        =   5
         Top             =   1200
         Width           =   735
      End
      Begin VB.PictureBox picSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         Picture         =   "frmMain.frx":A7AC
         ScaleHeight     =   255
         ScaleWidth      =   735
         TabIndex        =   4
         Top             =   1200
         Width           =   735
      End
      Begin RichTextLib.RichTextBox rtf 
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393217
         MultiLine       =   0   'False
         TextRTF         =   $"frmMain.frx":AD40
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblTitles 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "What would you like to do?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   2220
      End
   End
   Begin VB.PictureBox picResult 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   8280
      Picture         =   "frmMain.frx":ADC7
      ScaleHeight     =   2655
      ScaleWidth      =   2895
      TabIndex        =   7
      Top             =   2040
      Visible         =   0   'False
      Width           =   2895
      Begin VB.PictureBox picCancel2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1920
         Picture         =   "frmMain.frx":BF4A
         ScaleHeight     =   255
         ScaleWidth      =   735
         TabIndex        =   10
         Top             =   2160
         Width           =   735
      End
      Begin VB.ListBox lstTitles 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         ItemData        =   "frmMain.frx":C4A0
         Left            =   240
         List            =   "frmMain.frx":C4A2
         TabIndex        =   9
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lblTitles 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "What would you like to do?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   2220
      End
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   9840
      OLEDropMode     =   1  'Manual
      Picture         =   "frmMain.frx":C4A4
      ScaleHeight     =   1575
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   6000
      Width           =   1215
   End
   Begin VB.PictureBox picAnimate 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   7800
      Picture         =   "frmMain.frx":DB99
      ScaleHeight     =   1815
      ScaleWidth      =   105
      TabIndex        =   12
      Top             =   4920
      Visible         =   0   'False
      Width           =   105
      Begin VB.Label lblSubEvents 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   7
         Left            =   3000
         TabIndex        =   20
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label lblSubEvents 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   6
         Left            =   3000
         TabIndex        =   19
         Top             =   960
         Width           =   540
      End
      Begin VB.Label lblSubEvents 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   5
         Left            =   3000
         TabIndex        =   18
         Top             =   600
         Width           =   540
      End
      Begin VB.Label lblSubEvents 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   4
         Left            =   3000
         TabIndex        =   17
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblSubEvents 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label lblSubEvents 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   540
      End
      Begin VB.Label lblSubEvents 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   540
      End
      Begin VB.Label lblSubEvents 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.PictureBox picEvents 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Index           =   5
      Left            =   5520
      Picture         =   "frmMain.frx":1012D
      ScaleHeight     =   1815
      ScaleWidth      =   2055
      TabIndex        =   29
      Top             =   4920
      Width           =   2055
   End
   Begin VB.PictureBox picEvents 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Index           =   4
      Left            =   3240
      Picture         =   "frmMain.frx":12AB1
      ScaleHeight     =   1815
      ScaleWidth      =   2055
      TabIndex        =   28
      Top             =   4920
      Width           =   2055
   End
   Begin VB.PictureBox picEvents 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Index           =   3
      Left            =   960
      Picture         =   "frmMain.frx":1573B
      ScaleHeight     =   1815
      ScaleWidth      =   2055
      TabIndex        =   27
      Top             =   4920
      Width           =   2055
   End
   Begin VB.PictureBox picEvents 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Index           =   2
      Left            =   5520
      Picture         =   "frmMain.frx":17D89
      ScaleHeight     =   1815
      ScaleWidth      =   2055
      TabIndex        =   26
      Top             =   2880
      Width           =   2055
   End
   Begin VB.PictureBox picEvents 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Index           =   1
      Left            =   3240
      Picture         =   "frmMain.frx":1AA3E
      ScaleHeight     =   1815
      ScaleWidth      =   2055
      TabIndex        =   25
      Top             =   2880
      Width           =   2055
   End
   Begin VB.PictureBox picEvents 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Index           =   0
      Left            =   960
      Picture         =   "frmMain.frx":1D204
      ScaleHeight     =   1815
      ScaleWidth      =   2055
      TabIndex        =   24
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Timer tmrAction 
      Interval        =   2500
      Left            =   9240
      Top             =   6600
   End
   Begin VB.Timer tmrSlide 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   10200
      Top             =   5400
   End
   Begin VB.Timer tmrAnimate 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   9720
      Top             =   5400
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   0
      Picture         =   "frmMain.frx":1F8FD
      ScaleHeight     =   1935
      ScaleWidth      =   11295
      TabIndex        =   11
      Top             =   0
      Width           =   11295
   End
   Begin RichTextLib.RichTextBox tempRTF 
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   10005
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"frmMain.frx":26F68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picFooter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   0
      ScaleHeight     =   180
      ScaleWidth      =   8895
      TabIndex        =   21
      Top             =   1800
      Width           =   8895
      Begin VB.PictureBox picBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   7440
         Picture         =   "frmMain.frx":26FDF
         ScaleHeight     =   375
         ScaleWidth      =   1215
         TabIndex        =   23
         Top             =   5520
         Width           =   1215
      End
      Begin SHDocVwCtl.WebBrowser Web 
         Height          =   5055
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   8415
         ExtentX         =   14843
         ExtentY         =   8916
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.Label lblGreet 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   405
      Left            =   9000
      TabIndex        =   31
      Top             =   5400
      Width           =   120
   End
   Begin ComctlLib.ImageList IList 
      Left            =   9120
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   81
      ImageHeight     =   105
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   22
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":27755
            Key             =   "sleep1"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2DBBB
            Key             =   "admission2"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":39EC9
            Key             =   "smile1"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4032F
            Key             =   "smile2"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":46795
            Key             =   "enrolment2"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":52AA3
            Key             =   "general2"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5EDB1
            Key             =   "policy1"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":6B0BF
            Key             =   "guidelines2"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":773CD
            Key             =   "policy2"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":836DB
            Key             =   "scholarship2"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":8F9E9
            Key             =   "sleep2"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":95E4F
            Key             =   "admission1"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":A215D
            Key             =   "enrolment1"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":AE46B
            Key             =   "general1"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":BA779
            Key             =   "guidelines1"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":C6A87
            Key             =   "scholarship1"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":D2D95
            Key             =   "up1"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":D91FB
            Key             =   "default2"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":DF661
            Key             =   "default"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":E5AC7
            Key             =   "up2"
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":EBF2D
            Key             =   "write1"
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":F2393
            Key             =   "write2"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StartX, StartY

Dim Titles(40) As String
Dim CommonWords(127) As String
Dim OAspeech(10) As String
Dim OACounter As Integer

Dim EventIndex As Integer
Dim subWindow As Integer

Dim counter As Integer

Private Sub Form_Click()
    picAnimate.Visible = False
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    
    Source.Move X - StartX, Y - StartY
    
    If picInfo.Visible = True Then
        
        picInfo.Left = picChar.Left - (picInfo.Width / 2)
        picInfo.Top = picChar.Top - (picInfo.Height + 50)
        lblGreet.Top = picChar.Top
        lblGreet.Left = (picChar.Left - lblGreet.Width) + picChar.Width + 50
        
    End If
    
    If picResult.Visible = True Then
        
        picResult.Left = picChar.Left - (picResult.Width / 2)
        picResult.Top = picChar.Top - (picResult.Height + 50)
        lblGreet.Top = picChar.Top
        lblGreet.Left = (picChar.Left - lblGreet.Width) + picChar.Width + 50
        
    End If
    
End Sub

Private Sub Form_Load()
    
    Dim count As Integer
    
    For count = 0 To 5
        picEvents(count).Left = 0 - picEvents(count).Width
    Next
    
    Call Initialize
    picFooter.Height = 1
    OACounter = 0
    
    lblGreet.Top = picChar.Top
    lblGreet.Left = (picChar.Left - lblGreet.Width) + picChar.Width + 50
        
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
            
            picEvents(0).Picture = IList.ListImages("admission1").Picture

            picEvents(1).Picture = IList.ListImages("enrolment1").Picture

            picEvents(2).Picture = IList.ListImages("scholarship1").Picture

            picEvents(3).Picture = IList.ListImages("general1").Picture

            picEvents(4).Picture = IList.ListImages("policy1").Picture

            picEvents(5).Picture = IList.ListImages("guidelines1").Picture
            
End Sub

Private Sub lblSubEvents_Click(Index As Integer)
    
    Dim count As Integer
    
        For count = 0 To 5
            picEvents(count).Visible = False
        Next
        
        tmrAction.Enabled = True
        picAnimate.Visible = False
        tmrSlide.Enabled = True

        Select Case lblSubEvents(Index).Caption
        
                Case "Incoming freshmen students"
                    Web.Navigate2 App.Path & "\html\Incoming freshmen students.htm"
                Case "Transferees"
                    Web.Navigate2 App.Path & "\html\Transferees.htm"
                Case "Foreign students"
                    Web.Navigate2 App.Path & "\html\Foreign students.htm"
                Case "Enrolment procedure"
                    Web.Navigate2 App.Path & "\html\Enrolment procedure.htm"
                Case "Student load"
                    Web.Navigate2 App.Path & "\html\Student load.htm"
                Case "Pre-requisites of subjects"
                    Web.Navigate2 App.Path & "\html\Pre-requisites of subjects.htm"
                Case "Adding/Dropping of subjects"
                    Web.Navigate2 App.Path & "\html\Adding_Dropping of subjects.htm"
                Case "Dropping of subjects before midterm"
                    Web.Navigate2 App.Path & "\html\Dropping of subjects before midterm.htm"
                Case "Withdrawal of enrolment"
                    Web.Navigate2 App.Path & "\html\Withdrawal of enrolment.htm"
                Case "Shifting of course"
                    Web.Navigate2 App.Path & "\html\Shifting of course.htm"
                Case "Discontinuance of studies / Leave of absense"
                    Web.Navigate2 App.Path & "\html\Discontinuance of studies.htm"
                Case "Transfer credential"
                    Web.Navigate2 App.Path & "\html\Transfer credential.htm"
                Case "Scholarship grants"
                    Web.Navigate2 App.Path & "\html\Scholarship grants.htm"
                Case "Honors"
                    Web.Navigate2 App.Path & "\html\Honors.htm"
                Case "Graduation awards"
                    Web.Navigate2 App.Path & "\html\Graduation awards.htm"
                Case "Other awards"
                    Web.Navigate2 App.Path & "\html\Other awards.htm"
                Case "Attendance"
                    Web.Navigate2 App.Path & "\html\Attendance.htm"
                Case "Grading system"
                    Web.Navigate2 App.Path & "\html\Grading system.htm"
                Case "Grade complaints"
                    Web.Navigate2 App.Path & "\html\Grade complaints.htm"
                Case "Examination guidelines"
                    Web.Navigate2 App.Path & "\html\Examination guidelines.htm"
                Case "Application of ID card, ID validation"
                    Web.Navigate2 App.Path & "\html\Application for ID Card.htm"
                Case "Graduation"
                    Web.Navigate2 App.Path & "\html\Graduation.htm"
                Case "Standard of conduct"
                    Web.Navigate2 App.Path & "\html\Standard of conduct.htm"
                Case "Student organizations"
                    Web.Navigate2 App.Path & "\html\Student organizations.htm"
                Case "Student activities"
                    Web.Navigate2 App.Path & "\html\Student activities.htm"
                Case "Rules governing scholastic delinquency"
                    Web.Navigate2 App.Path & "\html\Rules governing scholastic delinquency.htm"
                Case "Administration"
                    Web.Navigate2 App.Path & "\html\Administration.htm"
                Case "Guidance and counseling"
                    Web.Navigate2 App.Path & "\html\Guidance and counseling.htm"
                Case "Health services"
                    Web.Navigate2 App.Path & "\html\Health services.htm"
                Case "Placement of services"
                    Web.Navigate2 App.Path & "\html\Placement Services.htm"
                Case "Computer center"
                    Web.Navigate2 App.Path & "\html\Computer center.htm"
                Case "Other facilities"
                    Web.Navigate2 App.Path & "\html\Other facilities.htm"
                Case "Code of conduct"
                    Web.Navigate2 App.Path & "\html\Code of conduct.htm"
                Case "Role of OSA"
                    Web.Navigate2 App.Path & "\html\Role of OSA.htm"
                Case "Student disciplinary tribunal"
                    Web.Navigate2 App.Path & "\html\Student disciplinary tribunal.htm"
                Case "Offenses"
                    Web.Navigate2 App.Path & "\html\Offenses.htm"
                Case "Sanctions"
                    Web.Navigate2 App.Path & "\html\Sanctions.htm"
                Case "Jurisdiction and venue"
                    Web.Navigate2 App.Path & "\html\Jurisdiction and venue.htm"
                
        End Select
        
End Sub

Private Sub lblSubEvents_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSubEvents(Index).ForeColor = &HFFFF80
    lblSubEvents(Index).MousePointer = 14
End Sub

Private Sub lstTitles_DblClick()
    
    Dim count As Integer
    
        For count = 0 To 5
            picEvents(count).Visible = False
        Next
            
        picAnimate.Visible = False
        tmrSlide.Enabled = True

        Select Case lstTitles.Text
        
                Case "Incoming freshmen students"
                    Web.Navigate2 App.Path & "\html\Incoming freshmen students.htm"
                Case "Transferees"
                    Web.Navigate2 App.Path & "\html\Transferees.htm"
                Case "Foreign students"
                    Web.Navigate2 App.Path & "\html\Foreign students.htm"
                Case "Enrolment procedure"
                    Web.Navigate2 App.Path & "\html\Enrolment procedure.htm"
                Case "Student load"
                    Web.Navigate2 App.Path & "\html\Student load.htm"
                Case "Pre-requisites of subjects"
                    Web.Navigate2 App.Path & "\html\Pre-requisites of subjects.htm"
                Case "Adding/Dropping of subjects"
                    Web.Navigate2 App.Path & "\html\Adding_Dropping of subjects.htm"
                Case "Dropping of subjects before midterm"
                    Web.Navigate2 App.Path & "\html\Dropping of subjects before midterm.htm"
                Case "Withdrawal of enrolment"
                    Web.Navigate2 App.Path & "\html\Withdrawal of enrolment.htm"
                Case "Shifting of course"
                    Web.Navigate2 App.Path & "\html\Shifting of course.htm"
                Case "Discontinuance of studies / Leave of absense"
                    Web.Navigate2 App.Path & "\html\Discontinuance of studies.htm"
                Case "Transfer credential"
                    Web.Navigate2 App.Path & "\html\Transfer credential.htm"
                Case "Scholarship grants"
                    Web.Navigate2 App.Path & "\html\Scholarship grants.htm"
                Case "Honors"
                    Web.Navigate2 App.Path & "\html\Honors.htm"
                Case "Graduation awards"
                    Web.Navigate2 App.Path & "\html\Graduation awards.htm"
                Case "Other awards"
                    Web.Navigate2 App.Path & "\html\Other awards.htm"
                Case "Attendance"
                    Web.Navigate2 App.Path & "\html\Attendance.htm"
                Case "Grading system"
                    Web.Navigate2 App.Path & "\html\Grading system.htm"
                Case "Grade complaints"
                    Web.Navigate2 App.Path & "\html\Grade complaints.htm"
                Case "Examination guidelines"
                    Web.Navigate2 App.Path & "\html\Examination guidelines.htm"
                Case "Application of ID card, ID validation"
                    Web.Navigate2 App.Path & "\html\Application for ID Card.htm"
                Case "Graduation"
                    Web.Navigate2 App.Path & "\html\Graduation.htm"
                Case "Standard of conduct"
                    Web.Navigate2 App.Path & "\html\Standard of conduct.htm"
                Case "Student organizations"
                    Web.Navigate2 App.Path & "\html\Student organizations.htm"
                Case "Student activities"
                    Web.Navigate2 App.Path & "\html\Student activities.htm"
                Case "Rules governing scholastic delinquency"
                    Web.Navigate2 App.Path & "\html\Rules governing scholastic delinquency.htm"
                Case "Administration"
                    Web.Navigate2 App.Path & "\html\Administration.htm"
                Case "Guidance and counseling"
                    Web.Navigate2 App.Path & "\html\Guidance and counseling.htm"
                Case "Health services"
                    Web.Navigate2 App.Path & "\html\Health services.htm"
                Case "Placement of services"
                    Web.Navigate2 App.Path & "\html\Placement Services.htm"
                Case "Computer center"
                    Web.Navigate2 App.Path & "\html\Computer center.htm"
                Case "Other facilities"
                    Web.Navigate2 App.Path & "\html\Other facilities.htm"
                Case "Code of conduct"
                    Web.Navigate2 App.Path & "\html\Code of conduct.htm"
                Case "Role of OSA"
                    Web.Navigate2 App.Path & "\html\Role of OSA.htm"
                Case "Student disciplinary tribunal"
                    Web.Navigate2 App.Path & "\html\Student disciplinary tribunal.htm"
                Case "Offenses"
                    Web.Navigate2 App.Path & "\html\Offenses.htm"
                Case "Sanctions"
                    Web.Navigate2 App.Path & "\html\Sanctions.htm"
                Case "Jurisdiction and venue"
                    Web.Navigate2 App.Path & "\html\Jurisdiction and venue.htm"
                    
                Case "AMA Mission and Vision"
                    Web.Navigate2 App.Path & "\html\AMA Mission and Vission.htm"
                Case "Philosophy of Education"
                    Web.Navigate2 App.Path & "\html\Philosophy.htm"
                Case "History"
                    Web.Navigate2 App.Path & "\html\History.htm"
        End Select
    
End Sub

Private Sub picAnimate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim count As Integer
    
        For count = 0 To 7
            lblSubEvents(count).ForeColor = vbWhite
        Next
    
End Sub

Private Sub picBack_Click()
        
        Dim count As Integer
        
        For count = 0 To 5
            picEvents(count).Visible = True
            picEvents(count).Left = 0 - picEvents(count).Width
        Next
            
        picFooter.Height = 10
        tmrSlide.Enabled = False
        tmrEntrance.Enabled = True
        
End Sub

Private Sub picCancel_Click()
    
    picInfo.Visible = False
    picChar.Picture = IList.ListImages("default").Picture
    counter = 0
    tmrAction.Enabled = True
    
End Sub

Private Sub picCancel2_Click()
    picResult.Visible = False
    picChar.Picture = IList.ListImages("default").Picture
    counter = 0
End Sub

Private Sub picChar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    StartX = X
    StartY = Y
    picChar.Drag vbBeginDrag
    
    picAnimate.Visible = False
    counter = 0
    picChar.Picture = IList.ListImages("default").Picture
    
    If picInfo.Visible = True Or picResult.Visible = True Then
        Exit Sub
    Else
        picInfo.Visible = True
        picInfo.Left = picChar.Left - (picInfo.Width / 2)
        picInfo.Top = picChar.Top - (picInfo.Height + 50)
        lblGreet.Top = picChar.Top
        lblGreet.Left = (picChar.Left - lblGreet.Width) + picChar.Width + 50
        
        rtf.Text = "Type a word here..."
        rtf.SelStart = 0
        rtf.SelLength = Len(rtf.Text)
        rtf.SetFocus
    End If
    
End Sub

Private Sub picChar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    picChar.Drag vbEndDrag
    
End Sub

Private Sub Initialize()

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
            OAspeech(0) = "Welcome!"
            OAspeech(1) = "Click me!"
            OAspeech(2) = "Learned som'tin"
            OAspeech(3) = "Are you amazed?"
            OAspeech(4) = "That's a fact!"
            
            OAspeech(5) = "Want more!?"
            OAspeech(6) = "Ask me! or else!"
            OAspeech(7) = "Im exausted"
            OAspeech(8) = "I like dis program!"
            OAspeech(9) = "Enjoying?"
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Titles(0) = "Incoming freshmen students"
                Titles(1) = "Transferees"
                Titles(2) = "Foreign students"

                Titles(3) = "Enrolment procedure"
                Titles(4) = "Student load"
                Titles(5) = "Pre-requisites of subjects"
                Titles(6) = "Adding/Dropping of subjects"
                Titles(7) = "Dropping of subjects before midterm"
                Titles(8) = "Withdrawal of enrolment"
                Titles(9) = "Shifting of course"
                Titles(10) = "Discontinuance of studies / Leave of absense"
                Titles(11) = "Transfer credential"

                Titles(12) = "Scholarship grants"
                Titles(13) = "Honors"
                Titles(14) = "Graduation awards"
                Titles(15) = "Other awards"

                Titles(16) = "Attendance"
                Titles(17) = "Grading system"
                Titles(18) = "Grade complaints"
                Titles(19) = "Examination guidelines"
                Titles(20) = "Application of ID card, ID validation"
                Titles(21) = "Graduation"
                Titles(22) = "Standard of conduct"
                Titles(23) = "Student organizations"
                Titles(24) = "Student activities"
            
                Titles(25) = "Rules governing scholastic delinquency"
                Titles(26) = "Administration"
                Titles(27) = "Guidance and counseling"
                Titles(28) = "Health services"
                Titles(29) = "Placement of services"
                Titles(30) = "Computer center"
                Titles(31) = "Other facilities"
                
                Titles(32) = "Code of conduct"
                Titles(33) = "Role of OSA"
                Titles(34) = "Student disciplinary tribunal"
                Titles(35) = "Offenses"
                Titles(36) = "Sanctions"
                Titles(37) = "Jurisdiction and venue"
                
                Titles(38) = "AMA Mission and Vision"
                Titles(39) = "Philosophy of Education"
                Titles(40) = "History"
                
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                
            CommonWords(0) = "THE"
            CommonWords(1) = "OF"
            CommonWords(2) = "TO"
            CommonWords(3) = "AND"
            CommonWords(4) = "A"
            CommonWords(5) = "IN"
            CommonWords(6) = "IS"
            CommonWords(7) = "IT"
            CommonWords(8) = "YOU"
            CommonWords(9) = "THAT"
            
            CommonWords(10) = "HE"
            CommonWords(11) = "WAS"
            CommonWords(12) = "FOR"
            CommonWords(13) = "ON"
            CommonWords(14) = "ARE"
            CommonWords(15) = "WITH"
            CommonWords(16) = "AS"
            CommonWords(17) = "I"
            CommonWords(18) = "HIS"
            CommonWords(19) = "THEY"
            
            CommonWords(20) = "BE"
            CommonWords(21) = "AT"
            CommonWords(22) = "ONE"
            CommonWords(23) = "HAVE"
            CommonWords(24) = "THIS"
            CommonWords(25) = "FROM"
            CommonWords(26) = "OR"
            CommonWords(27) = "HAD"
            CommonWords(28) = "BY"
            CommonWords(29) = "HOT"
            
            CommonWords(30) = "BUT"
            CommonWords(31) = "SOME"
            CommonWords(32) = "WHAT"
            CommonWords(33) = "THERE"
            CommonWords(34) = "WE"
            CommonWords(35) = "CAN"
            CommonWords(36) = "OUT"
            CommonWords(37) = "OTHER"
            CommonWords(38) = "WERE"
            CommonWords(39) = "ALL"
            
            CommonWords(40) = "YOUR"
            CommonWords(41) = "WHEN"
            CommonWords(42) = "UP"
            CommonWords(43) = "USE"
            CommonWords(44) = "WORD"
            CommonWords(45) = "HOW"
            CommonWords(46) = "SAID"
            CommonWords(47) = "AN"
            CommonWords(48) = "EACH"
            CommonWords(49) = "SHE"
            
            CommonWords(50) = "WHICH"
            CommonWords(51) = "DO"
            CommonWords(52) = "THEIR"
            CommonWords(53) = "TIME"
            CommonWords(54) = "IF"
            CommonWords(55) = "WILL"
            CommonWords(56) = "WAY"
            CommonWords(57) = "ABOUT"
            CommonWords(58) = "MANY"
            CommonWords(59) = "THEN"
            
            CommonWords(60) = "THEM"
            CommonWords(61) = "WOULD"
            CommonWords(62) = "WRITE"
            CommonWords(63) = "LIKE"
            CommonWords(64) = "SO"
            CommonWords(65) = "THESE"
            CommonWords(66) = "HER"
            CommonWords(67) = "LONG"
            CommonWords(68) = "MAKE"
            CommonWords(69) = "THING"
            
            CommonWords(70) = "SEE"
            CommonWords(71) = "HIM"
            CommonWords(72) = "TWO"
            CommonWords(73) = "HAS"
            CommonWords(74) = "LOOK"
            CommonWords(75) = "MORE"
            CommonWords(76) = "DAY"
            CommonWords(77) = "COULD"
            CommonWords(78) = "GO"
            CommonWords(79) = "COME"
            
            CommonWords(80) = "DID"
            CommonWords(81) = "MY"
            CommonWords(82) = "SOUND"
            CommonWords(83) = "NO"
            CommonWords(84) = "MOST"
            CommonWords(85) = "NUMBER"
            CommonWords(86) = "WHO"
            CommonWords(87) = "OVER"
            CommonWords(88) = "KNOW"
            CommonWords(89) = "WATER"
            
            CommonWords(90) = "THAN"
            CommonWords(91) = "CALL"
            CommonWords(92) = "FIRST"
            CommonWords(93) = "PEOPLE"
            CommonWords(94) = "MAY"
            CommonWords(95) = "DOWN"
            CommonWords(96) = "SIDE"
            CommonWords(97) = "BEEN"
            CommonWords(98) = "NOW"
            CommonWords(99) = "FIND"
            
            CommonWords(100) = "ANY"
            CommonWords(101) = "NEW"
            CommonWords(102) = "WORK"
            CommonWords(103) = "PART"
            CommonWords(104) = "TAKE"
            CommonWords(105) = "GET"
            CommonWords(106) = "PLACE"
            CommonWords(107) = "MADE"
            CommonWords(108) = "LIVE"
            CommonWords(109) = "WHERE"
            
            CommonWords(110) = "AFTER"
            CommonWords(111) = "BACK"
            CommonWords(112) = "LITTLE"
            CommonWords(113) = "ONLY"
            CommonWords(114) = "ROUND"
            CommonWords(115) = "MAN"
            CommonWords(116) = "YEAR"
            CommonWords(117) = "CAME"
            CommonWords(118) = "SHOW"
            CommonWords(119) = "EVERY"
            
            CommonWords(120) = "GOOD"
            CommonWords(121) = "ME"
            CommonWords(122) = "GIVE"
            CommonWords(123) = "OUR"
            CommonWords(124) = "UNDER"
            
            CommonWords(125) = "."
            CommonWords(126) = "!"
            CommonWords(127) = "?"

End Sub

Private Sub picEvents_Click(Index As Integer)
    
    tmrAnimate.Enabled = True
    tmrAction.Enabled = True
    picAnimate.Visible = True
    
    picInfo.Visible = False
    picResult.Visible = False
    
    picAnimate.Top = picEvents(Index).Top
    picAnimate.Left = picEvents(Index).Left
    picAnimate.Width = 1
    
    Dim count As Integer
            
            For count = 0 To 3
                lblSubEvents(count).Left = 240
            Next
            
            For count = 4 To 7
                lblSubEvents(count).Left = 3000
            Next
            
    Select Case Index
        
        Case 0
        
            For count = 0 To 2
                lblSubEvents(count).Caption = Titles(count)
            Next
            
            For count = 3 To 7
                lblSubEvents(count).Caption = ""
            Next
            
            subWindow = 4000
            
        Case 1
            
            For count = 0 To 7
                lblSubEvents(count).Caption = Titles(count + 3)
            Next
            
            subWindow = 7000
            
        Case 2
        
            For count = 0 To 3
                lblSubEvents(count).Caption = Titles(count + 12)
            Next
            
            For count = 4 To 7
                lblSubEvents(count).Caption = ""
            Next
            
            subWindow = 4000
            
        Case 3
        
            For count = 0 To 7
                lblSubEvents(count).Caption = Titles(count + 16)
            Next
            
            subWindow = 6000
            
        Case 4
            
            For count = 4 To 7
                lblSubEvents(count).Left = 4000
            Next
            
            For count = 0 To 6
                lblSubEvents(count).Caption = Titles(count + 25)
            Next
            
            lblSubEvents(7).Caption = ""
            
            subWindow = 7000
            
        Case 5
        
            For count = 0 To 5
                lblSubEvents(count).Caption = Titles(count + 32)
            Next
            
            For count = 6 To 7
                lblSubEvents(count).Caption = ""
            Next
            
            subWindow = 5100
            
    End Select
    
End Sub

Private Sub picEvents_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Select Case Index
        
        Case 0
            picEvents(0).Picture = IList.ListImages("admission2").Picture
        Case 1
            picEvents(1).Picture = IList.ListImages("enrolment2").Picture
        Case 2
            picEvents(2).Picture = IList.ListImages("scholarship2").Picture
        Case 3
            picEvents(3).Picture = IList.ListImages("general2").Picture
        Case 4
            picEvents(4).Picture = IList.ListImages("policy2").Picture
        Case 5
            picEvents(5).Picture = IList.ListImages("guidelines2").Picture
            
    End Select
    
End Sub

Private Sub picFooter_DragDrop(Source As Control, X As Single, Y As Single)
    Call Form_DragDrop(Source, X, Y)
End Sub

Private Sub picHeader_DragDrop(Source As Control, X As Single, Y As Single)
    Call Form_DragDrop(Source, X, Y)
End Sub

Private Sub picSearch_Click()
    
    Dim count As Integer
    Dim count2 As Integer
    Dim count3 As Integer
    
    Dim pos As Integer
    Dim res As Integer
    Dim res2 As Integer
    Dim res3 As Integer
    Dim foundDATA As Boolean
    Dim optionNum As Integer
    
    counter = 0
    foundDATA = False
    lstTitles.Clear
    pos = 0
    
    tmrAction.Enabled = True
    
    tempWord.Text = rtf.Text
    
    'Delete Common Words (ascending)
    For count = 0 To 127
        
        If count >= 125 Then
            res2 = tempWord.Find(CommonWords(count), pos)
        Else
            res2 = tempWord.Find(CommonWords(count), pos, , optionNum + 2)
        End If
        
        If res2 = -1 Then
        Else
            tempWord.SelText = ""
            pos = tempWord.SelStart + tempWord.SelLength
            tempWord.SetFocus
            count = 0
            pos = 0
        End If
        
    Next
    
    'If search is exact
    For count = 0 To 40
            
            tempRTF.Text = Titles(count)
            res = tempRTF.Find(tempWord.Text, 0)
            
            If res = -1 Then
            Else
                lstTitles.AddItem Titles(count)
                foundDATA = True
                
                picInfo.Visible = False
                picResult.Visible = True
                picResult.Left = picChar.Left - (picResult.Width / 2)
                picResult.Top = picChar.Top - (picResult.Height + 100)
                lblGreet.Top = picChar.Top
                lblGreet.Left = (picChar.Left - lblGreet.Width) + picChar.Width + 50
                
            End If
    
    Next
    
    If foundDATA = True Then
        Exit Sub
    End If
    
    pos = 0
    'Delete spaces
    For count = 0 To Len(tempWord)
        
        res3 = tempWord.Find(" ", pos)
        
        If res3 = -1 Then
        Else
            tempWord.SelText = ""
            pos = tempWord.SelStart + tempWord.SelLength
            tempWord.SetFocus
        End If
        
        If Left$(tempWord.Text, 1) <> " " Then
            
            For count3 = 0 To Len(tempWord)
                
                res3 = tempWord.Find("  ", pos)
                
                If res3 = -1 Then
                Else
                    tempWord.SelText = " "
                    pos = tempWord.SelStart + tempWord.SelLength
                    tempWord.SetFocus
                End If
            Next
            
                'Find similar topics
                For count2 = 0 To 40
                
                    tempRTF.Text = Titles(count2)
                    res = tempRTF.Find(tempWord.Text, 0)
                    
                    If res = -1 Then
                    Else
                        lstTitles.AddItem Titles(count2)
                        foundDATA = True
                        
                        picInfo.Visible = False
                        picResult.Visible = True
                        picResult.Left = picChar.Left - (picResult.Width / 2)
                        picResult.Top = picChar.Top - (picResult.Height + 100)
                        lblGreet.Top = picChar.Top
                        lblGreet.Left = (picChar.Left - lblGreet.Width) + picChar.Width + 50
                        
                    End If

                Next
            
            Exit For
        End If
        
    Next
        
        If foundDATA = True Then Exit Sub
        
        'Find similar topics
        For count = 0 To 40
        
            tempRTF.Text = Titles(count)
            res = tempRTF.Find(Left$(tempWord.Text, 4), 0)
            
            If res = -1 Then
            Else
                lstTitles.AddItem Titles(count)
                foundDATA = True
                
                picInfo.Visible = False
                picResult.Visible = True
                picResult.Left = picChar.Left - (picResult.Width / 2)
                picResult.Top = picChar.Top - (picResult.Height + 100)
                lblGreet.Top = picChar.Top
                lblGreet.Left = (picChar.Left - lblGreet.Width) + picChar.Width + 50
                
            End If
            
        Next
    
    If foundDATA = False Then
        MsgBox "No word is associated with the input!", vbInformation, "AMA Computer University"
        rtf.SelStart = 0
        rtf.SelLength = Len(rtf.Text)
        rtf.SetFocus
    End If
    
End Sub

Private Sub rtf_Change()
        
        If rtf.Text = "" Or rtf.Text = "Type a word here..." Then
            tmrAction.Enabled = True
        Else
            tmrAction.Enabled = False
            
            If picChar.Picture = IList.ListImages("write1").Picture Then
                picChar.Picture = IList.ListImages("write2").Picture
            Else
                picChar.Picture = IList.ListImages("write1").Picture
            End If
            
        End If
        
End Sub

Private Sub rtf_Click()
    rtf.Text = ""
End Sub

Private Sub tmrAction_Timer()
    
    Dim count As Integer
    Dim delay As Integer
    
    delay = 30
    
    If picInfo.Visible = True Then
        If picChar.Picture = IList.ListImages("up1").Picture Then
            picChar.Picture = IList.ListImages("up2").Picture
            tmrAction.Interval = 500
        Else
            picChar.Picture = IList.ListImages("up1").Picture
            tmrAction.Interval = 500
        End If
        
        Exit Sub
        
    End If
    
    If picResult.Visible = True Then
        If picChar.Picture = IList.ListImages("smile1").Picture Then
            picChar.Picture = IList.ListImages("smile2").Picture
            tmrAction.Interval = 1500
        Else
            picChar.Picture = IList.ListImages("smile1").Picture
            tmrAction.Interval = 1500
        End If
        
        Exit Sub
        
    End If
    
    
    If picInfo.Visible = False And picResult.Visible = False And counter < delay Then
        
        If picChar.Picture = IList.ListImages("default").Picture Then
            picChar.Picture = IList.ListImages("default2").Picture
            tmrAction.Interval = 100
        Else
            picChar.Picture = IList.ListImages("default").Picture
            tmrAction.Interval = 2500
            
            If counter < delay Then
                counter = counter + 1
            End If
            
        End If
        
        Exit Sub
        
    ElseIf picInfo.Visible = False And picResult.Visible = False And counter >= delay Then
        
        If picChar.Picture = IList.ListImages("sleep1").Picture Then
            picChar.Picture = IList.ListImages("sleep2").Picture
            tmrAction.Interval = 2000
        Else
            picChar.Picture = IList.ListImages("sleep1").Picture
            tmrAction.Interval = 2000
        End If
        
        Exit Sub
        
    End If
    
End Sub

Private Sub tmrAnimate_Timer()
    
    If picAnimate.Width > subWindow Then
        tmrAnimate.Enabled = False
    Else
        picAnimate.Width = picAnimate.Width + 50
    End If

End Sub

Private Sub tmrEntrance_Timer()
            
    Dim delay As Integer
    Dim space As Integer
    
    delay = 200
    space = 300
    
            If picEvents(2).Left < 5500 Then
                picEvents(2).Left = picEvents(2).Left + delay
            Else
                If picEvents(1).Left < picEvents(2).Left - picEvents(2).Width - space Then
                    picEvents(1).Left = picEvents(1).Left + delay
                Else
                    If picEvents(0).Left < picEvents(1).Left - picEvents(1).Width - space Then
                        picEvents(0).Left = picEvents(0).Left + delay
                        
                        lblGreet.Caption = OAspeech(OACounter)
                        lblGreet.Left = (picChar.Left - lblGreet.Width) + picChar.Width + 50
                        lblGreet.Top = lblGreet.Top - 25
                        
                    Else
                        
                        If picEvents(5).Left < 5500 Then
                            picEvents(5).Left = picEvents(5).Left + delay
                        Else
                            If picEvents(4).Left < picEvents(5).Left - picEvents(5).Width - space Then
                                picEvents(4).Left = picEvents(4).Left + delay
                            Else
                                If picEvents(3).Left < picEvents(4).Left - picEvents(4).Width - space Then
                                    
                                    picEvents(3).Left = picEvents(3).Left + delay
                                    lblGreet.Top = lblGreet.Top + 25
                                    
                                Else
                                    lblGreet.Caption = ""
                                    
                                    If OACounter < 10 Then
                                        OACounter = OACounter + 1
                                    Else
                                        OACounter = 0
                                    End If
                                    
                                    tmrEntrance.Enabled = False
                                    lblGreet.Top = picChar.Top
                                End If
                            End If
                        End If
                        
                        
                    End If
                End If
            End If
            
End Sub

Private Sub tmrSlide_Timer()

    If picFooter.Height < Me.Height Then
        picFooter.Height = picFooter.Height + 50
    Else
        tmrSlide.Enabled = False
    End If
    
End Sub
