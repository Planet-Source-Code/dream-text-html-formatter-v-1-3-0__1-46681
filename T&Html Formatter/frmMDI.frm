VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Text & Html Formatter"
   ClientHeight    =   6630
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10590
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6000
      Left            =   2640
      ScaleHeight     =   6000
      ScaleWidth      =   10845
      TabIndex        =   11
      Top             =   360
      Width           =   10845
      Begin TabDlg.SSTab Tab1 
         Height          =   6015
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   10610
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         Tab             =   2
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Text Editor"
         TabPicture(0)   =   "frmMDI.frx":1CCA
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "RTB"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Html Code"
         TabPicture(1)   =   "frmMDI.frx":1CE6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Text1"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Html Preview"
         TabPicture(2)   =   "frmMDI.frx":1D02
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "WB1"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Html Source Grabber"
         TabPicture(3)   =   "frmMDI.frx":1D1E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Picture3"
         Tab(3).Control(1)=   "txtGrab"
         Tab(3).Control(2)=   "txtAddress"
         Tab(3).Control(3)=   "Inet1"
         Tab(3).Control(4)=   "Label1"
         Tab(3).ControlCount=   5
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   -68400
            ScaleHeight     =   255
            ScaleWidth      =   1095
            TabIndex        =   19
            Top             =   360
            Visible         =   0   'False
            Width           =   1095
            Begin VB.CommandButton cmdGrab 
               Caption         =   "Html Grab"
               Height          =   255
               Left            =   0
               TabIndex        =   20
               Top             =   0
               Width           =   1095
            End
         End
         Begin RichTextLib.RichTextBox txtGrab 
            Height          =   5265
            Left            =   -75000
            TabIndex        =   17
            Top             =   690
            Width           =   7800
            _ExtentX        =   13758
            _ExtentY        =   9287
            _Version        =   393217
            ScrollBars      =   2
            TextRTF         =   $"frmMDI.frx":1D3A
         End
         Begin VB.TextBox txtAddress 
            Height          =   285
            Left            =   -74160
            TabIndex        =   15
            Top             =   360
            Width           =   5655
         End
         Begin VB.TextBox Text1 
            Height          =   5535
            Left            =   -74985
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Text            =   "frmMDI.frx":1E3B
            Top             =   330
            Visible         =   0   'False
            Width           =   7800
         End
         Begin SHDocVwCtl.WebBrowser WB1 
            Height          =   5535
            Left            =   15
            TabIndex        =   14
            Top             =   330
            Visible         =   0   'False
            Width           =   7800
            ExtentX         =   13758
            ExtentY         =   9763
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
            Location        =   "http:///"
         End
         Begin InetCtlsObjects.Inet Inet1 
            Left            =   -69120
            Top             =   180
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
         End
         Begin RichTextLib.RichTextBox RTB 
            Height          =   5625
            Left            =   -74985
            TabIndex        =   18
            Top             =   330
            Visible         =   0   'False
            Width           =   7800
            _ExtentX        =   13758
            _ExtentY        =   9922
            _Version        =   393217
            ScrollBars      =   2
            TextRTF         =   $"frmMDI.frx":1E41
         End
         Begin VB.Label Label1 
            Caption         =   "Address:"
            Height          =   255
            Left            =   -74880
            TabIndex        =   16
            Top             =   360
            Width           =   735
         End
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   333
      Left            =   6720
      Top             =   600
   End
   Begin VB.PictureBox picMain 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   6000
      Left            =   0
      ScaleHeight     =   6000
      ScaleWidth      =   2640
      TabIndex        =   1
      Tag             =   "3585"
      Top             =   360
      Width           =   2640
      Begin VB.FileListBox File1 
         Height          =   1260
         Left            =   720
         TabIndex        =   10
         Top             =   3600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.PictureBox PicBrowse 
         BorderStyle     =   0  'None
         Height          =   1320
         Left            =   120
         ScaleHeight     =   88
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   153
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmMDI.frx":1F42
         Left            =   120
         List            =   "frmMDI.frx":1F52
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   2280
         Width           =   2415
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   120
         Left            =   480
         MousePointer    =   7  'Size N S
         Picture         =   "frmMDI.frx":1FDE
         ScaleHeight     =   120
         ScaleWidth      =   1575
         TabIndex        =   4
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton cmdSizer 
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5985
         Index           =   0
         Left            =   0
         MousePointer    =   9  'Size W E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Size Code Listing"
         Top             =   0
         Width           =   90
      End
      Begin VB.CommandButton cmdSizer 
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5985
         Index           =   1
         Left            =   2520
         MousePointer    =   9  'Size W E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Resize Directory And File's Window"
         Top             =   0
         Width           =   90
      End
      Begin MSComctlLib.ListView lvCode 
         CausesValidation=   0   'False
         Height          =   3375
         Left            =   120
         TabIndex        =   6
         Top             =   2640
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   5953
         View            =   2
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         _Version        =   393217
         Icons           =   "imgListTvw"
         SmallIcons      =   "imgListTvw"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDragMode     =   1
         NumItems        =   0
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlToolbarIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Desktop"
               Object.ToolTipText     =   "DeskTop"
               ImageIndex      =   16
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "UpOne"
               Object.ToolTipText     =   "Move up one directory"
               ImageIndex      =   17
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Key             =   "Tags"
               Object.ToolTipText     =   "View Html Tags"
               ImageIndex      =   18
            EndProperty
         EndProperty
      End
      Begin VB.Image Imagedown 
         Height          =   120
         Left            =   600
         Picture         =   "frmMDI.frx":2A02
         Top             =   2040
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Image Imageup 
         Height          =   120
         Left            =   840
         Picture         =   "frmMDI.frx":3426
         Top             =   2040
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   120
         Picture         =   "frmMDI.frx":3E4A
         Top             =   2040
         Width           =   2340
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   120
         Picture         =   "frmMDI.frx":5BCE
         Top             =   360
         Width           =   2340
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   6840
      Top             =   5640
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   3240
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":7952
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":7A64
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":7B76
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":7C88
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":7D9A
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":7EAC
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":7FBE
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":80D0
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":81E2
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":82F4
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":8406
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":8518
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":862A
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":873C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":8AD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":8E74
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":928C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":9554
            Key             =   "Tag"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   6360
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3598
            MinWidth        =   3598
            Text            =   "Status:"
            TextSave        =   "Status:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8334
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   2011
            MinWidth        =   2011
            TextSave        =   "6:37 PM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1588
            MinWidth        =   1588
            TextSave        =   "7/6/03"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgListTvw 
      Left            =   3240
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   18
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":992C
            Key             =   "FolderClosed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":9BF0
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":9F44
            Key             =   "Bas"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":A35C
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":A658
            Key             =   "Gif"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":AA70
            Key             =   "Bmp"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":AE44
            Key             =   "Jpg"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":B218
            Key             =   "Ctl"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":B5F0
            Key             =   "Exe"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":B950
            Key             =   "Html"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":BD64
            Key             =   "Txt"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":C178
            Key             =   "Zip"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":C58C
            Key             =   "Dll"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":C8E0
            Key             =   "Ini"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":CC34
            Key             =   "Log"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New File"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open File"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save Text/Html"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print Text/Html"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut Text"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy Text"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste Text"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "HtmlGrab"
            Object.ToolTipText     =   "Html Source Grabber"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "strip"
            Object.ToolTipText     =   "Strip HTML From Text"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu khk 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu hkh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu dg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReset 
         Caption         =   "&Reset"
      End
      Begin VB.Menu dshh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit (Esc)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "E&dit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu hgggggg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearAll 
         Caption         =   "Clear &All"
      End
      Begin VB.Menu mnuAll 
         Caption         =   "Select &All"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      WindowList      =   -1  'True
      Begin VB.Menu mnuSettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu dhfhf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTextShow 
         Caption         =   "Show &Editor"
      End
      Begin VB.Menu mnuHTML 
         Caption         =   "Show &Preview Window"
      End
      Begin VB.Menu mnuBrowser 
         Caption         =   "Show &Web Browser"
      End
      Begin VB.Menu lllll 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStatBar 
         Caption         =   "StatusBar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuToolBar 
         Caption         =   "ToolBar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuBold 
         Caption         =   "Bold (Selected text)"
      End
      Begin VB.Menu mnuBreak 
         Caption         =   "Line Break"
      End
      Begin VB.Menu mnuSpc 
         Caption         =   "&Space"
      End
      Begin VB.Menu jjgj 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFontColr 
         Caption         =   "Font Colo&r (Selected text)"
         Begin VB.Menu mnuFCBlue 
            Caption         =   "&Blue"
         End
         Begin VB.Menu mnuFCGreen 
            Caption         =   "&Green"
         End
         Begin VB.Menu mnuFCRed 
            Caption         =   "&Red"
         End
         Begin VB.Menu mnuFCOrange 
            Caption         =   "&Orange"
         End
         Begin VB.Menu mnuFCPrpl 
            Caption         =   "&Purple"
         End
         Begin VB.Menu mnuFCYlw 
            Caption         =   "&Yellow"
         End
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Font &Size (Selected text)"
         Begin VB.Menu mnuFont1 
            Caption         =   "&Size 1"
         End
         Begin VB.Menu mnuFont2 
            Caption         =   "&Size 2"
         End
         Begin VB.Menu mnuFont3 
            Caption         =   "&Size 3"
         End
         Begin VB.Menu mnuFont4 
            Caption         =   "&Size 4"
         End
         Begin VB.Menu mnuFont5 
            Caption         =   "&Size 5"
         End
         Begin VB.Menu mnuFont6 
            Caption         =   "&Size 6"
         End
      End
      Begin VB.Menu yoooooy 
         Caption         =   "-"
      End
      Begin VB.Menu mnuT 
         Caption         =   "Table"
         Begin VB.Menu mnuTblStrt 
            Caption         =   "&Start"
         End
         Begin VB.Menu mnuTblEnd 
            Caption         =   "En&d"
         End
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "Tools"
      Begin VB.Menu mnuDbMargin 
         Caption         =   "&Double Space Margin"
      End
      Begin VB.Menu mnuDouble 
         Caption         =   "&Double Space Text"
      End
      Begin VB.Menu mnuIndent 
         Caption         =   "&Indent Margin (2 spaces)"
      End
      Begin VB.Menu gg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHtmGrab 
         Caption         =   "Html &Source Grabber"
      End
      Begin VB.Menu mnuStrTexta 
         Caption         =   "&Strip Html Source"
         Begin VB.Menu mnuStrText 
            Caption         =   "This Will Strip Html Tags From The Html Source Grabber Window, Proceed?"
         End
      End
      Begin VB.Menu dggg 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuFormatFunc 
      Caption         =   "&Format"
      Begin VB.Menu mnuFormat 
         Caption         =   "&Format Text/Code "
      End
      Begin VB.Menu fdfd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFrameWrap 
         Caption         =   "&Frame Wrap On Preview"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuHtmlPage 
         Caption         =   "&Set As Html Page"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuExample 
         Caption         =   "&Example"
      End
      Begin VB.Menu fiiuy 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInstruct 
         Caption         =   "Text To Html &Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu fdhfh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCtact 
         Caption         =   "&Contact"
         Begin VB.Menu mnuRegister 
            Caption         =   "&Register"
         End
         Begin VB.Menu dfds 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRBug 
            Caption         =   "Report &Bug"
         End
      End
      Begin VB.Menu hhhh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowAbout 
         Caption         =   "A&bout Text To Html"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************* _
 Copyright © 2002-2003 Dream-Domain.net _
 ********************************* _
  NOTE: THIS HEADER MUST STAY INTACT.
' Terms of Agreement: _
  By using this code, you agree to the following terms... _
  1) You may use this code in your own programs (and may compile it into _
  a program and distribute it in compiled format for languages that allow _
  it) freely and with no charge. _
  2) You MAY NOT redistribute this code (for example to a web site) without _
  written permission from the original author. Failure to do so is a _
  violation of copyright laws. _
  3) You may link to this code from another website, but ONLY if it is not _
  wrapped in a frame. _
  4) You will abide by any additional copyright restrictions which the _
  author may have placed in the code or code's description. _
 **********************************
' Text & HTML Formatter v1.3.0 _
  ----------------- _
  Example By Dream _
  Date: 5th July 2003 3:27:28 PM _
  Email:  baddest_attitude@hotmail.com _
 ********************************** _
  Additional Terms of Agreement: _
  You MAY NOT Sell This Code _
  You MAY NOT Sell Any Program Containing This Code _
  You use this code knowing I hold no responsibilities for any results _
  occuring from the use and/or misuse of this code _
  If you make any improvements it would be nice if you would send me a copy. _
 ********************************** _
  TITLE HERE _
 **********************************

'Thanks to mrk* for his additions/modifications wherever you see this ..
'* Didnt know if he wanted his full name here or not.
'05.2003 mrk Change/Add
Private iPanelDrag As Integer

Private Const meMinHeight As Single = 5595
Private Const meMinWidth As Single = 10795

Private counter As Integer              ' Variable for the No. of Opened Child forms
Private ListClick As Boolean
Private openeded As Boolean
Private ResizeGate As Boolean
Dim bdirty As Boolean

Private Sub BoxInCursor(bSet As Boolean)
'======================================================================
' This routine forces mouse in a set rectangle when resizing objects
'======================================================================
If bSet Then
    Dim client As RECT
    Dim upperleft As PointAPI
    'Get information about our wndow
    GetClientRect hwnd, client
    upperleft.X = client.Left
    upperleft.Y = client.Top
    'Convert window coördinates to screen coördinates
    ClientToScreen hwnd, upperleft
    'move our rectangle
    OffsetRect client, upperleft.X, upperleft.Y
    InflateRect client, -2, -2
    Select Case iPanelDrag
    Case 1 ''''''
    
    Case 2 ' right side of main listing is being resized
        client.Left = client.Left + 120
        client.Right = client.Right - 320
    Case 3  ' updown  sidebar being resized
        client.Top = client.Top + 108
        client.Bottom = client.Bottom - 105
    End Select
    'limit the cursor movement
    ClipCursor client
Else
    ClipCursor ByVal 0&
End If
End Sub



Private Sub cmdGrab_Click()
 On Error GoTo Grab_Err
   Screen.MousePointer = vbHourglass
   txtGrab.Text = Inet1.OpenURL(txtAddress.Text)
   
Err.Clear
Grab_Err:
   Screen.MousePointer = vbDefault
   If Err <> 0 Then
      MsgBox "The Url Could Not Be Downloaded," & vbCrLf & "An Error Occured: " & Err.Number & "   " & Err.Description, vbExclamation, "HTML Grabber"
      Exit Sub
   End If
   If txtGrab.Text = vbNullString Then MsgBox "Please Check You Inserted A Valid Url", vbInformation, "Html Source Grabber"
End Sub

Private Sub cmdSizer_MouseDown(Index As Integer, Button As Integer, shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  If Button <> vbRightButton Then
     iPanelDrag = Index + 1
     SetCapture cmdSizer(Index).hwnd
     BoxInCursor True
  End If
 
  If Index = 1 Then ResizeGate = True
End Sub

Private Sub cmdSizer_MouseMove(Index As Integer, Button As Integer, shift As Integer, X As Single, Y As Single)
On Error Resume Next
Select Case iPanelDrag
Case 1
    picMain.Width = picMain.Width - X
    cmdSizer(1).Left = picMain.Width - cmdSizer(1).Width
Case 2
    picMain.Width = picMain.Width + X
    cmdSizer(1).Left = picMain.Width - cmdSizer(1).Width
    Picture1.Left = picMain.Width / 2 - Picture1.Width / 2 + 10
    Combo1.Width = Combo1.Width + X
    resize_Tab
Case 3
   '''
End Select
End Sub

Private Sub cmdSizer_MouseUp(Index As Integer, Button As Integer, shift As Integer, X As Single, Y As Single)
BoxInCursor False
ReleaseCapture
If iPanelDrag Then
    If iPanelDrag = 2 Then lvCode.SetFocus
    If iPanelDrag = 4 Then Me.SetFocus
    If iPanelDrag < 3 Then Me.SetFocus   'DoGradient picMain, 1
    iPanelDrag = 0
End If
ResizeGate = False
End Sub

Private Sub Combo1_Click()
    File1.Path = m_CurrentDirectory
    LockWindowUpdate 0&
    ReList_Files
End Sub

Private Sub lvCode_DblClick()
'============================================================
' Allow a double click on the listview to view selected file
'============================================================
     Dim Char As String
     Dim Construct As String
10    On Error GoTo ErrTrap

20    If lvCode.ListItems.Count = 0 Then Exit Sub     ' of course if no code in the listview then abort
30    LockedWindow = True          'Tell RTB change to cease while loading file
40    Screen.MousePointer = 11
     'This gives us the extension but when loading the same somehow it drops the
     'filename so we check it here then use the file.path & "\" & lvcode.selecteditem.text
50    Char = Mid(lvCode.SelectedItem.Key, _
             Len(lvCode.SelectedItem.Key) - 3)
     'check if it is a file extension or exit the sub.
60    If InStr(1, Char, ".") = False Then Exit Sub

      Construct = File1.Path & "\" & lvCode.SelectedItem.Text
     'Check filetype for .frm .bas & .cls
     'If VBDoc then tell ToFileLoad what kinda form!
70    SB.Panels(1).Text = "Status: Loading File"

80    Select Case Char
        'You can customize the filetypes and the load procedure in the ToFile.bas
        'to load just about any filetype you wish (text type files)
         Case ".frm", ".bas", ".cls":    RTB.Text = ToFileLoad(Construct, _
             VBDoc) 'if a VBDoc then filter i
90       Case Else:  RTB.Text = ToFileLoad(Construct)   'or ToFileLoad(.FileName, Default)
100   End Select

     'color in keywords
110   ColorIn RTB
120   Timer1.Enabled = True
130   LockedWindow = False
140   Call PreviewHtml(False)
150   Screen.MousePointer = 1
160   SB.Panels(1).Text = "Status: File Loaded: "
170   SB.Panels(2).Text = lvCode.SelectedItem.Key & "    "

180   On Error GoTo 0
190   Exit Sub
ErrTrap:
200      If ErrorLogAndStop(Err.Number, Err.Description, Erl, Err.Source, _
                                     "{lvCode_DblClick}", _
                                     "{Private Sub}", _
                                     "{frmMDI}", _
                                     "{frmMDI}") Then
210      End If
220     Resume Next ' Resume 'Exit Sub
End Sub

Private Sub lvCode_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
'============================================================
' option gives the user a right click function within the code listing
'============================================================
10    On Error GoTo lvCode_MouseUp_General_ErrTrap

20    If lvCode.ListItems.Count = 0 Then Exit Sub         ' No code in the listing, no right click

30    If Button = 2 Then
40       MsgBox "right click here"
50    End If

60    On Error GoTo 0
70    Exit Sub
lvCode_MouseUp_General_ErrTrap:
80       If ErrorLogAndStop(Err.Number, Err.Description, Erl, Err.Source, _
                                   "{lvCode_MouseUp}", _
                                   "{Private Sub}", _
                                   "{frmMDI}", _
                                   "{frmMDI}") Then
90      End If
100     Resume Next ' Resume 'Exit Sub
End Sub

Private Sub MDIForm_Load()

'        To Do ==>  API for desktop path
 
10       On Error GoTo LogError
20       LockWindowUpdate PicBrowse.hwnd
30       Combo1.ListIndex = 1
40       lvCode.Height = picMain.ScaleHeight - lvCode.Top - 60
    '    tvProject.Height = picFilters.Height - 200
50       cmdSizer(0).Enabled = False: cmdSizer(1).Enabled = True
60       Screen.MousePointer = vbHourglass
          
         OpenGate = True
75       LoadKeyWords
      '  ChangePath App.Path
         Screen.MousePointer = vbDefault
80       On Error GoTo 0
         Exit Sub
LogError:
100      If ErrorLogAndStop(Err.Number, Err.Description, Erl, Err.Source, _
                                     "{MDIForm_Load}", _
                                     "{Private Sub}", _
                                     "{frmMDI}", _
                                     "{frmMDI}") Then
110     End If
120     Resume Next ' Resume 'Exit Sub
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormCode Then
     'Space for MsgBoxTitle----------,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
     Cancel = CBool(MsgBox("Do you wish to exit Text To Html Formatter ?", _
              vbYesNo Or vbInformation Or vbApplicationModal Or vbMsgBoxSetForeground Or vbSystemModal, _
              Me.Caption) = vbNo)
  End If
End Sub '05.2003 mrk Change/Add

Private Sub mdiform_resize()
    resize_Tab
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)                    'Unload application
  CloseUp
  Timer1.Enabled = False
  ToFileKill ToAppPath & "preview.htm"                        'Delete the preview html
  End
End Sub

Private Sub mnuAll_Click()                                   'Select all text
  RTB.SetFocus
  RTB.SelStart = 0
  RTB.SelLength = Len(RTB.Text)
End Sub

Private Sub mnuBold_Click()
    RTB.SelBold = True
End Sub

Private Sub mnuBreak_Click()
  
   Insert vbNewLine
End Sub

Private Sub mnuBrowser_Click()
Tab1.Visible = True
WB1.Visible = True
End Sub

Private Sub mnuClearAll_Click()
    RTB.Text = vbNullString                             'Clear RTB TextBox
End Sub

Private Sub mnuCopy_Click()                                   'Copy selected text
    Clipboard.Clear
    Clipboard.SetText RTB.SelText
End Sub

Private Sub mnuCut_Click()                                    'Cut selected text
    Clipboard.Clear
    Clipboard.SetText RTB.SelText
    RTB.SelText = vbNullString
End Sub

Private Sub mnuDbMargin_Click()                               'Double Space the Margin
    DoubleUpMargin False
End Sub

Private Sub mnuDelete_Click()                                 'Delete selected text
    RTB.SelText = vbNullString
End Sub

Private Sub mnuDouble_Click()
    RTB.Text = Replace(RTB.Text, Chr$(32), Chr$(32) & Chr$(32))
End Sub

Public Sub mnuExample_Click()                                   'Example Text
    LockedWindow = False
    ExampleLoad
    Tab1.TabIndex = 0
    FormatText
    Call PreviewHtml(False)
End Sub

Private Sub mnuExit_Click()                                   'Unload form
    Unload Me
End Sub

Private Sub mnuFCBlue_Click()
    RTB.SelColor = vbBlue
End Sub

Private Sub mnuFCGreen_Click()
    RTB.SelColor = vbGreen
End Sub

Private Sub mnuFCOrange_Click()
  RTB.SelColor = 211
End Sub

Private Sub mnuFCPrpl_Click()
  RTB.SelColor = 214
End Sub

Private Sub mnuFCRed_Click()
  RTB.SelColor = vbRed
End Sub

Private Sub mnuFCYlw_Click()
  RTB.SelColor = vbYellow
End Sub

Private Sub mnuFont1_Click()
 RTB.SelFontSize = 8
End Sub

Private Sub mnuFont2_Click()
 RTB.SelFontSize = 10
End Sub

Private Sub mnuFont3_Click()
 RTB.SelFontSize = 12
End Sub

Private Sub mnuFont4_Click()
 RTB.SelFontSize = 15
End Sub

Private Sub mnuFont5_Click()
 RTB.SelFontSize = 20
End Sub

Private Sub mnuFont6_Click()
 RTB.SelFontSize = 26
End Sub

Private Sub mnuFormat_Click()
  FormatText
  Call PreviewHtml(True)
End Sub

Private Sub mnuFrameWrap_Click()
  mnuFrameWrap.Checked = Not mnuFrameWrap.Checked
End Sub

Private Sub mnuHtmGrab_Click()
    txtGrab.Visible = True
    Tab1.Visible = True
End Sub

Private Sub mnuHTML_Click()
    Tab1.Visible = True
    Text1.Visible = True
End Sub

Private Sub mnuHtmlPage_Click()
     mnuHtmlPage.Checked = Not mnuHtmlPage.Checked
End Sub

Private Sub mnuIndent_Click()
     DoubleUpMargin True
End Sub
  
Private Sub mnuInstruct_Click()
''
End Sub

Private Sub mnuLoad_Click()                                  'Load file
    FileLoader   'ToFile bas
End Sub

Private Sub mnuNew_Click()                                  'Load New File
    If Reset = True Then
       SB.Panels(1).Text = "Status: New File"
       SB.Panels(2).Text = "FileName: None         "
       Tab1.Visible = True
       RTB.Visible = True
       Tab1.TabIndex = 0
    End If
End Sub

Private Sub mnuPaste_Click()                               'Paste text
  On Error Resume Next
  Dim strData As String
  strData = Clipboard.GetText(vbCFText)
  RTB.SelText = strData
End Sub

Private Sub mnuPrint_Click()
   Call PrintOut
End Sub

Private Sub mnuRBug_Click()
  OpenLink frmMDI, "Mailto:dream@dream-domain.net?Subject=TandH Bug Report"
End Sub

Private Sub mnuRegister_Click()
  OpenLink frmMDI, "Mailto:dream@dream-domain.net?Subject=TandH Register"
End Sub

Private Sub mnuReset_Click()                               'Clear Program
   Reset
End Sub

Private Sub mnuSave_Click()                            'Save Text To File
   Call FileSaver
End Sub
Private Sub mnuText_Click()
  Show
End Sub

Private Sub mnuSettings_Click()
'frmSettings.Show
End Sub

Private Sub mnuShowAbout_Click()                       'Display About Window
  ToDlgAbout Me, _
             "    " & App.ProductName, _
             "    Example By " & App.CompanyName & ", Version " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & "    " & App.LegalCopyright, _
             Me.Icon
End Sub '05.2003 mrk Change/Add

Private Sub mnuSpc_Click()
Insert "&nbsp;"
End Sub


Private Sub mnuStatBar_Click()                         'Hide/Show Status Bar
  mnuStatBar.Checked = Not mnuStatBar.Checked
  Select Case mnuStatBar.Checked
    Case 1
      SB.Visible = True
    Case 0
      SB.Visible = False
    Case Else
     'Blah Blah @ Microsoft
  End Select
End Sub

Private Sub mnuStrText_Click()
    Dim s As String
    If txtGrab.Text = vbNullString Then MsgBox "No html code present in the Source Grabber window", vbInformation, "No Html Found": Exit Sub
    Screen.MousePointer = vbHourglass
    s = txtGrab.Text
    txtGrab.Text = StripTags(s, "")
    Screen.MousePointer = vbDefault
End Sub

Public Sub mnuTblEnd_Click()
  Insert "</td></tr></table>"
End Sub

Private Sub mnuTblStrt_Click()
  TableStart
End Sub

Private Sub mnuTextShow_Click()
    RTB.Visible = True
    Tab1.Visible = True
    Tab1.TabIndex = 0
    
    Tab1.TabIndex = 0
End Sub

Private Sub mnuToolBar_Click()                         'Hide/Show Tool Bar
  mnuToolBar.Checked = Not mnuToolBar.Checked
  Select Case mnuToolBar.Checked
    Case 1
      TB.Visible = True                                'Show toolbar and resize
    Case 0
      TB.Visible = False                               'Hide toolbar and resize
    Case Else
     'Blah Blah @ Microsoft
  End Select
End Sub

Public Sub PathChange()
  On Error GoTo ErrControl   'if path not dir (user clicked file) then catch error
    'Recieve path change from the Treeview
    If openeded = False Then Exit Sub
    File1.Path = m_CurrentDirectory
    LockWindowUpdate 0&
    ReList_Files
ErrControl:
On Error GoTo 0
End Sub

Private Sub PicBrowse_Resize()
    'resize the Treeview as needed
    On Error Resume Next
    SizeTV 0, 0, PicBrowse.ScaleWidth, PicBrowse.ScaleHeight
End Sub

Private Sub picMain_Resize()
'============================================================
'  Whenever the form resizes, adjust the code listing to fit the window & repaint the window
'============================================================
  On Error Resume Next
  If iPanelDrag = 0 Then
     cmdSizer(0).Height = picMain.Height
     cmdSizer(1).Height = picMain.Height
     cmdSizer(2).Height = picMain.Height
     cmdSizer(3).Height = picMain.Height
  End If
  lvCode.Height = picMain.ScaleHeight - lvCode.Top ' Adjust the code listing size
  lvCode.Width = picMain.Width - (cmdSizer(0) + 60) * 2 - 120
  PicBrowse.Width = picMain.Width - (cmdSizer(0) + 60) * 2 - 120
End Sub

Private Sub picture1_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button <> vbRightButton Then
    iPanelDrag = 3
    SetCapture Picture1.hwnd
    BoxInCursor True
    Picture1.Picture = Imagedown.Picture
End If
End Sub
 
Private Sub picture1_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
 If iPanelDrag = 3 Then
    Picture1.Top = Picture1.Top + Y
    Combo1.Top = Picture1.Top + 360
    PicBrowse.Height = Picture1.Top - 630
    Image2.Top = Picture1.Top + 120
    lvCode.Top = Combo1.Top + 320
    lvCode.Height = frmMDI.Height - PicBrowse.Height - Picture1.Height - Combo1.Height - 2190  ' - 500
 End If
End Sub

Private Sub picture1_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
  BoxInCursor False
  ReleaseCapture
  iPanelDrag = 0
  Picture1.Picture = Imageup.Picture
End Sub

Public Sub ProduceHtml()                                      'Produce a blank html
     Dim intFreeFile As Integer
10    On Error GoTo LogError
20    If ToFileIsExist(ToAppPath & "blank.htm") = False Then       'If doesnt exist then..
30      intFreeFile = FreeFile
40      Open ToAppPath & "blank.htm" For Output As #intFreeFile    'Create it!
50           Print " "  ' vbNullString
60      Close #intFreeFile
70    End If
80    On Error GoTo 0
90    Exit Sub
LogError:
100      If ErrorLogAndStop(Err.Number, Err.Description, Erl, Err.Source, _
                                      "{ProduceHtml}", _
                                      "{Public Sub}", _
                                      "{frmMDI}", _
                                      "{frmMDI}") Then
110     End If
120     Resume Next ' Resume 'Exit Sub
End Sub


Private Sub resize_Tab()
 Dim TabW As Integer
 Dim TabH As Integer

 Picture2.Width = frmMDI.Width - picMain.Width - 120

 Tab1.Width = Picture2.Width
 Tab1.Height = Picture2.Height
 TabW = Tab1.Width
 TabH = Tab1.Height
 RTB.Height = TabH - 340
 RTB.Width = TabW - 80
 Text1.Height = TabH - 360
 Text1.Width = TabW - 80
 WB1.Height = TabH - 360
 WB1.Width = TabW - 80
 txtGrab.Height = TabH - 730
 txtGrab.Width = TabW - 80
 txtAddress.Width = TabW - 2200
 Picture3.Left = TabW - 1300

End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button) 'Toolbar click
  Select Case Button.Key                            ' Read them for what they say
    Case "New": mnuNew_Click
    Case "Open": mnuLoad_Click
    Case "Save": mnuSave_Click
    Case "Print": PrintOut
    Case "Cut": mnuCut_Click
    Case "Copy": mnuCopy_Click
    Case "Paste": mnuPaste_Click
    Case "HtmlGrab": mnuHtmGrab_Click
    Case "strip": frmMDI.PopupMenu frmMDI.mnuStrTexta
    Case Else: 'Blah Blah @ Microsoft
  End Select
End Sub

Private Sub Timer1_Timer()                               'Scroll Timer
    With SB.Panels(2)
     .Text = ToStrScrollToLeft(.Text)                    'Call textscroll function
    End With
End Sub

Private Sub Timer2_Timer()
  'Allow change now
   openeded = True
  'Change Dir
   ChangePath App.Path
   PicBrowse.Visible = True
  'Kill timer after first run
   Timer2 = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Key                            ' Read them for what they say
      Case "Desktop"
            ChangePath "C:\Windows\Desktop"
            SB.Panels(2).Text = "C:\Windows\Desktop         "
      Case "UpOne"
            Dim arr() As String
            Dim Char As String
            Dim char2 As String
            arr() = Split(m_CurrentDirectory, "\")
            Char = arr(UBound(arr))
            char2 = Replace(m_CurrentDirectory, "\" & Char, vbNullString)
            If Len(char2) <= 2 Then char2 = char2 & "\"
            ChangePath char2
            SB.Panels(2).Text = char2 & "        "
      Case "Tags" ': 'LoadTags
   End Select
End Sub

Private Sub tvProject_NodeClick(ByVal Node As MSComctlLib.Node)
    Insert CStr(Node)
End Sub

Private Function IsControlKey(ByVal KeyCode As Long) As Boolean  'For RTB KeyDown
 'check if the key is a control key
  Select Case KeyCode
    Case vbKeyLeft, vbKeyRight, vbKeyHome, _
         vbKeyEnd, vbKeyPageUp, vbKeyPageDown, _
         vbKeyShift, vbKeyControl
         IsControlKey = True
    Case Else
        IsControlKey = False
  End Select
End Function

Private Sub rtb_mouseup(Button As Integer, shift As Integer, X As Single, Y As Single)
   If Button = 2 Then frmMDI.PopupMenu frmMDI.mnuEdit
End Sub

Private Sub RTB_change()                  'RTB Text changed so lets re color it
   If LockedWindow = True Then Exit Sub
   Dim position As Long
   position = RTB.SelStart
   ColorIn RTB, position
End Sub

Private Sub RTB_KeyPress(KeyAscii As Integer)
   Dim lCursor As Long
   Dim lSelectLen As Long
   Dim lStart As Long
   Dim lFinish As Long
   Dim sText As String
   On Error Resume Next                      'Here's the on the fly coloring
   Screen.MousePointer = vbDefault
   
'   If KeyAscii = vbKeyC And shift = 2 Then Exit Sub  'check for Ctrl+C
'   If KeyAscii = vbKeyV And shift = 2 Then  'check for text being pasted into the box
'      Screen.MousePointer = vbHourglass
'      DoClipBoardPaste RTB
'      KeyAscii = 0
'      Screen.MousePointer = vbNormal
'      Exit Sub
'   End If
  
  'if the cursor is moving to a different
  'line then process the orginal line
   If KeyAscii = 13 Or _
      KeyAscii = vbKeyUp Or _
      KeyAscii = vbKeyDown Then
       LockedWindow = True
       
      If bdirty Or KeyAscii = 13 Then 'only color this line if it's been changed
         LockWindowUpdate RTB.hwnd  'lock the window to cancel out flickering
         lCursor = RTB.SelStart      'store the current cursor pos
         lSelectLen = RTB.SelLength  'and current selection if there is any
         If lCursor <> 0 Then        'get the line start and end
            lStart = InStrRev(RTB.Text, vbCrLf, RTB.SelStart) + 2
            If lStart = 2 Then lStart = 1
          Else
            lStart = 1
         End If
         lFinish = InStr(lCursor + 1, RTB.Text, vbCrLf)
         If lFinish = 0 Then lFinish = Len(RTB.Text)
            basColor.sText = RTB.Text            'do the coloring
            DoColor RTB, lStart, lFinish
           'if ENTER was pressed, we should color the next line
           'as well, so that if a line is broken by the ENTER
           'the new line and the old line are colored properly
            If KeyAscii = 13 Then
               lStart = lCursor + 1
               lFinish = InStr(lStart, RTB.Text, vbCrLf)
               If lFinish = 0 Then lFinish = Len(RTB.Text)
               If lStart - 1 <> lFinish Then  'only color if another line exists
                  RTB.SelStart = lStart - 1
                  RTB.SelLength = lFinish - lStart
                  RTB.SelColor = vbBlack
                  DoColor RTB, lStart, lFinish
               End If
            End If
            RTB.SelStart = lCursor      'reset the properties
            RTB.SelLength = lSelectLen
            RTB.SelColor = vbBlack
            bdirty = False              'reset the flag and release the window
            LockWindowUpdate 0&
        End If
     ElseIf Not IsControlKey(KeyAscii) Then
       'a different key was pressed - and
       'this will alter the line so it
       'needs recoloring when we move off it
        If Not bdirty Then
           LockWindowUpdate RTB.hwnd
           lStart = InStrRev(RTB.Text, vbCrLf, RTB.SelStart + 1) + 1 'get the line start & end
           lFinish = InStr(RTB.SelStart + 1, RTB.Text, vbCrLf)
           If lFinish = 0 Then lFinish = Len(RTB.Text)
           lCursor = RTB.SelStart     'color the line (remembering the cursor position)
           lSelectLen = RTB.SelLength
           RTB.SelStart = lStart
           RTB.SelLength = lFinish - lStart
           RTB.SelColor = vbBlack
           RTB.SelStart = lCursor
           RTB.SelLength = lSelectLen
           bdirty = True
           LockWindowUpdate 0&
       End If
       LockedWindow = False
    End If
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
       cmdGrab_Click
    End If

End Sub
 ' Dreams Tools (3:27:28 PM  7/6/03) 48 + 848 lines <D.E.I.>
