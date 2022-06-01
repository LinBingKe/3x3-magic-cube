<<<<<<< HEAD
VERSION 5.00
Object = "{E1208DE3-A783-11D0-9161-00A024D24992}#1.0#0"; "MILApplication.ocx"
Object = "{6D9F7F71-9658-11D0-BDB5-00608CC9F9FB}#1.0#0"; "MILSystem.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{03985961-6B33-11D0-AB4A-00608CC9CA57}#1.0#0"; "MilBuffer.ocx"
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   Caption         =   "Form"
   ClientHeight    =   11715
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   18585
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   14.25
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11715
   ScaleWidth      =   18585
   StartUpPosition =   3  '系統預設值
   Begin MILAPPLICATIONLib.Application Application1 
      Height          =   480
      Left            =   14880
      TabIndex        =   43
      Top             =   8880
      Visible         =   0   'False
      Width           =   480
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      AutomaticAllocation=   -1  'True
   End
   Begin MILSYSTEMLib.System System1 
      Height          =   480
      Left            =   15600
      TabIndex        =   44
      Top             =   8880
      Visible         =   0   'False
      Width           =   480
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      AutomaticAllocation=   -1  'True
      SystemType      =   "VGA"
      ProcessingSystem=   1699376
      ProcessingSystemName=   "[Default]"
   End
   Begin MILBUFFERLib.Buffer Buffer1 
      Height          =   480
      Left            =   16320
      TabIndex        =   46
      Top             =   8880
      Visible         =   0   'False
      Width           =   480
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      AutomaticAllocation=   -1  'True
      OwnerSystem     =   "System1"
      SizeX           =   640
      SizeY           =   480
      NumberOfBands   =   3
      AbsoluteValue   =   252
      Saturation      =   252
      ChildRegionEndX =   639
      ChildRegionEndY =   479
      ChildRegionCenterX=   319
      ChildRegionCenterY=   239
      ChildRegionSizeX=   640
      ChildRegionSizeY=   480
      ChildRegionMode =   1
      CanDisplay      =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   16920
      Top             =   8880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Caption         =   "測試"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12360
      TabIndex        =   45
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "分析樣本所在位置"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9960
      TabIndex        =   42
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   3
      Left            =   8880
      TabIndex        =   41
      Top             =   1320
      Width           =   350
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   3
      Left            =   8160
      TabIndex        =   40
      Top             =   2040
      Width           =   350
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   3
      Left            =   7440
      TabIndex        =   39
      Top             =   2760
      Width           =   350
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   4
      Left            =   8880
      TabIndex        =   38
      Top             =   1800
      Width           =   350
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   4
      Left            =   8160
      TabIndex        =   37
      Top             =   2520
      Width           =   350
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   4
      Left            =   7440
      TabIndex        =   36
      Top             =   3240
      Width           =   350
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   4
      Left            =   8040
      TabIndex        =   35
      Top             =   1080
      Width           =   850
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   5
      Left            =   7320
      TabIndex        =   34
      Top             =   1800
      Width           =   850
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   6
      Left            =   6600
      TabIndex        =   33
      Top             =   2400
      Width           =   850
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   2
      Left            =   1440
      TabIndex        =   32
      Top             =   1440
      Width           =   350
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   2
      Left            =   2160
      TabIndex        =   31
      Top             =   2040
      Width           =   350
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   2
      Left            =   2880
      TabIndex        =   30
      Top             =   2640
      Width           =   350
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   1440
      TabIndex        =   29
      Top             =   1920
      Width           =   350
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   2160
      TabIndex        =   28
      Top             =   2520
      Width           =   350
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   2880
      TabIndex        =   27
      Top             =   3120
      Width           =   350
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   4
      Left            =   4440
      TabIndex        =   26
      Top             =   3600
      Width           =   350
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   4
      Left            =   5160
      TabIndex        =   25
      Top             =   3600
      Width           =   350
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   4
      Left            =   5880
      TabIndex        =   24
      Top             =   3600
      Width           =   350
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   6
      Left            =   6360
      TabIndex        =   23
      Top             =   5640
      Width           =   850
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   4
      Left            =   6360
      TabIndex        =   22
      Top             =   4200
      Width           =   850
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   5
      Left            =   6360
      TabIndex        =   21
      Top             =   4920
      Width           =   850
   End
   Begin VB.CommandButton Command5 
      Caption         =   "暫停"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   20
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "開始"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   19
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   17400
      Top             =   8160
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   3
      Left            =   5880
      TabIndex        =   18
      Top             =   3120
      Width           =   350
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   2
      Left            =   5880
      TabIndex        =   17
      Top             =   6720
      Width           =   350
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   5880
      TabIndex        =   16
      Top             =   6240
      Width           =   350
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   3
      Left            =   5160
      TabIndex        =   15
      Top             =   3120
      Width           =   350
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   2
      Left            =   5160
      TabIndex        =   14
      Top             =   6720
      Width           =   350
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   5160
      TabIndex        =   13
      Top             =   6240
      Width           =   350
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   3
      Left            =   3360
      TabIndex        =   12
      Top             =   2400
      Width           =   850
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   2
      Left            =   2520
      TabIndex        =   11
      Top             =   1800
      Width           =   850
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   1800
      TabIndex        =   10
      Top             =   1080
      Width           =   850
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   3
      Left            =   4440
      TabIndex        =   9
      Top             =   3120
      Width           =   350
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   2
      Left            =   4440
      TabIndex        =   8
      Top             =   6720
      Width           =   350
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   4440
      TabIndex        =   7
      Top             =   6240
      Width           =   350
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   2
      Left            =   3360
      TabIndex        =   6
      Top             =   4920
      Width           =   850
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   3360
      TabIndex        =   5
      Top             =   4200
      Width           =   850
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   3
      Left            =   3360
      TabIndex        =   4
      Top             =   5640
      Width           =   850
   End
   Begin VB.CommandButton Command4 
      Caption         =   "關閉檔案"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   3
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "下一步"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   2
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "記錄旋轉指令"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9960
      TabIndex        =   1
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "判斷開始"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9960
      TabIndex        =   0
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   2
      Left            =   240
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   1
      Left            =   240
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H80000000&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   0
      Left            =   240
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   1
      Left            =   2760
      Shape           =   1  '正方形
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   4
      Left            =   2040
      Shape           =   1  '正方形
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   7
      Left            =   1320
      Shape           =   1  '正方形
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   2
      Left            =   2760
      Shape           =   1  '正方形
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   5
      Left            =   2040
      Shape           =   1  '正方形
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   8
      Left            =   1320
      Shape           =   1  '正方形
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   3
      Left            =   2760
      Shape           =   1  '正方形
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   6
      Left            =   2040
      Shape           =   1  '正方形
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   9
      Left            =   1320
      Shape           =   1  '正方形
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   3
      Left            =   5760
      Shape           =   1  '正方形
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   2
      Left            =   5040
      Shape           =   1  '正方形
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   1
      Left            =   4320
      Shape           =   1  '正方形
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   6
      Left            =   5760
      Shape           =   1  '正方形
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   5
      Left            =   5040
      Shape           =   1  '正方形
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   4
      Left            =   4320
      Shape           =   1  '正方形
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   9
      Left            =   5760
      Shape           =   1  '正方形
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   8
      Left            =   5040
      Shape           =   1  '正方形
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   7
      Left            =   4320
      Shape           =   1  '正方形
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   0
      Left            =   240
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   1
      Left            =   4320
      Shape           =   1  '正方形
      Top             =   7320
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   2
      Left            =   5040
      Shape           =   1  '正方形
      Top             =   7320
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   3
      Left            =   5760
      Shape           =   1  '正方形
      Top             =   7320
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   4
      Left            =   4320
      Shape           =   1  '正方形
      Top             =   8040
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   5
      Left            =   5040
      Shape           =   1  '正方形
      Top             =   8040
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   6
      Left            =   5760
      Shape           =   1  '正方形
      Top             =   8040
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   7
      Left            =   4320
      Shape           =   1  '正方形
      Top             =   8760
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   8
      Left            =   5040
      Shape           =   1  '正方形
      Top             =   8760
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   9
      Left            =   5760
      Shape           =   1  '正方形
      Top             =   8760
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   0
      Left            =   240
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   1
      Left            =   4320
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   2
      Left            =   5040
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   3
      Left            =   5760
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   4
      Left            =   4320
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   5
      Left            =   5040
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   6
      Left            =   5760
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   7
      Left            =   4320
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   8
      Left            =   5040
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   9
      Left            =   5760
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   0
      Left            =   240
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   1
      Left            =   7440
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   2
      Left            =   7440
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   3
      Left            =   7440
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   4
      Left            =   8160
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   5
      Left            =   8160
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   6
      Left            =   8160
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   7
      Left            =   8880
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   8
      Left            =   8880
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   9
      Left            =   8880
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   0
      Left            =   240
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H000080FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   0
      Left            =   240
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   0
      Left            =   240
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim box(1 To 3, 1 To 3, 1 To 3, 1 To 3) As Integer
Dim box2(1 To 3, 1 To 3, 1 To 3, 1 To 3) As Integer
Dim showBox(1 To 3, 1 To 3, 1 To 3, 1 To 3) As Integer
Dim Third0_1(1 To 3, 1 To 3) As Integer
Dim Third0_2(1 To 3, 1 To 3) As Integer
Dim Third1_1(1 To 3, 1 To 3) As Integer
Dim Third1_2(1 To 3, 1 To 3) As Integer
Dim Third1_3(1 To 3, 1 To 3) As Integer
Dim Third2_1(1 To 3, 1 To 3) As Integer
Dim Third2_2(1 To 3, 1 To 3) As Integer
Dim Third2_3(1 To 3, 1 To 3) As Integer
Dim Third2_4(1 To 3, 1 To 3) As Integer
Dim Third2_5(1 To 3, 1 To 3) As Integer
Dim Third2_6(1 To 3, 1 To 3) As Integer
Dim Third2_7(1 To 3, 1 To 3) As Integer


Dim R, L, G, U, E, W As Integer



Dim term As Integer '選擇判斷
Dim showterm As Integer '旋轉指示方塊顏色變換判斷

Private Sub Command1_Click() '讓程式判斷該怎麼轉
Dim i, j, k, n As Integer
Call colorSET

'先歸零
For i = 1 To 3
    For j = 1 To 3
        For k = 1 To 3
            For n = 1 To 3
            
                box2(i, j, k, n) = 0
                box(i, j, k, n) = 0
            
            Next n
        Next k
    Next j
Next i



'開啟文件輸入顏色到box

Call OpenColorfile(2, 3, 0, 4, 1) '前面

Call OpenColorfile(4, 1, 4, 4, 1) '後面

Call OpenColorfile(1, 0, 1, 4, 2) '左面

Call OpenColorfile(3, 4, 3, 4, 2) '右面

Call OpenColorfile(5, 0, 0, 3, 3) '上面

Call OpenColorfile(6, 4, 0, 1, 3) '下面



'-----------------------初始狀況
For k = 1 To 3
    For i = 1 To 3
        For j = 1 To 3
            
                
                Call printcolor(i, j, k)
          
        Next j
    Next i
Next k

Call Openfile(1)
Call showmethod2
'----------------------------------------------開始第一層判斷
    Debug.Print "第一層邊塊測試"
    Debug.Print "判斷121"
    Call firstEdgeBlock(1, 2, 1, E, 0, L)
    
    Debug.Print "判斷211"
    Call firstEdgeBlock(2, 1, 1, 0, U, L)
    
    Debug.Print "判斷231"
    Call firstEdgeBlock(2, 3, 1, 0, G, L)
    
    Debug.Print "判斷321"
    Call firstEdgeBlock(3, 2, 1, R, 0, L)
    
    Debug.Print "第一層角塊測試"
    Debug.Print "判斷111"
    Call firstCornerBlock(1, 1, 1, E, U, L)
    Call text
    Debug.Print "判斷131"
    Call firstCornerBlock(1, 3, 1, E, G, L)
    
    
    'Call text
    Debug.Print "判斷311"
    
    Call firstCornerBlock(3, 1, 1, R, U, L)
    
    Debug.Print "判斷331"
    Call firstCornerBlock(3, 3, 1, R, G, L)
    
'--------------------第一層復原以後的狀況
    
    
    For k = 1 To 3
    For i = 1 To 3
        For j = 1 To 3
            
                
                Call printcolor(i, j, k)
          
        Next j
    Next i
Next k

    


'----------------------------------------第二層判斷
    Debug.Print "第二層測試"
    Debug.Print "判斷112"
    Call secondmethod(1, 1, 2, E, U, 0)
 
    Debug.Print "判斷132"
    Call secondmethod(1, 3, 2, E, G, 0)
    
    Debug.Print "判斷312"
    Call secondmethod(3, 1, 2, R, U, 0)
    
    Debug.Print "判斷332"
    Call secondmethod(3, 3, 2, R, G, 0)
    
'--------------------第二層復原以後的狀況
For k = 1 To 3
    For i = 1 To 3
        For j = 1 To 3
            
                
                Call printcolor(i, j, k)
          
        Next j
    Next i
Next k

'-----------------------------------------第三層判斷
    Debug.Print "第三層測試"
    Call thirdmethod
    
 '--------------------第三層復原以後的狀況
For k = 1 To 3
    For i = 1 To 3
        For j = 1 To 3
            
                
                'Call printcolor(i, j, k)
          
        Next j
    Next i
Next k
   
Call Openfile(3)
'測試
'Call choose1(0, 0, 3) '旋轉的方法:第三層z軸方向旋轉




End Sub
Private Sub printcolor(i, j, k) '查看方塊顏色
Dim a As Integer

For a = 1 To 3
    If box(i, j, k, a) = 0 Then
    Debug.Print i, j, k, a, "0"
    End If
    
    If box(i, j, k, a) = 1 Then
    Debug.Print i, j, k, a, "紅"
    End If
    
    If box(i, j, k, a) = 2 Then
    Debug.Print i, j, k, a, "黑"
    End If
    
    If box(i, j, k, a) = 3 Then
    Debug.Print i, j, k, a, "綠"
    End If
    
    If box(i, j, k, a) = 4 Then
    Debug.Print i, j, k, a, "藍"
    End If
    
    If box(i, j, k, a) = 5 Then
    Debug.Print i, j, k, a, "黃"
    End If
    
    If box(i, j, k, a) = 6 Then
    Debug.Print i, j, k, a, "橘"
    End If
    
Next a
Debug.Print ""




End Sub
Private Sub OpenColorfile(NUMBER As Integer, a As Integer, b As Integer, c As Integer, d As Integer)
    '判斷開啟的文件
    If NUMBER = 1 Then
        Open "blue.txt" For Input As #1
    ElseIf NUMBER = 2 Then
        Open "red.txt" For Input As #1
    ElseIf NUMBER = 3 Then
        Open "green.txt" For Input As #1
    ElseIf NUMBER = 4 Then
        Open "orange.txt" For Input As #1
    ElseIf NUMBER = 5 Then
        Open "yellow.txt" For Input As #1
    ElseIf NUMBER = 6 Then
        Open "black.txt" For Input As #1
    Else
    End If
    
    Call ColorInput(NUMBER, a, b, c, d)
    
    Close #1
    
End Sub
Private Sub ColorInput(NUMBER As Integer, a, b, c, d)
Dim i As Integer
Dim j As Integer
Dim Color As Integer

    If NUMBER = 1 Then '方塊左面
        For i = 1 To 3
            For j = 1 To 3
                Input #1, Color '讀取一條紀錄
                box(a + i, b, c - j, d) = Color
            Next j
        Next i
        
    ElseIf NUMBER = 3 Then '方塊右面
        For i = 1 To 3
            For j = 1 To 3
                Input #1, Color '讀取一條紀錄
                box(a - i, b, c - j, d) = Color
            Next j
        Next i
    
    ElseIf NUMBER = 2 Then '方塊前面
        For i = 1 To 3
            For j = 1 To 3
                Input #1, Color '讀取一條紀錄
                box(a, b + i, c - j, d) = Color
            Next j
        Next i
        
    
      
    ElseIf NUMBER = 4 Then '方塊後面
        For i = 1 To 3
            For j = 1 To 3
                Input #1, Color '讀取一條紀錄
                box(a, b - i, c - j, d) = Color
            Next j
        Next i
        
    ElseIf NUMBER = 5 Then '方塊上面
        For i = 1 To 3
            For j = 1 To 3
                Input #1, Color '讀取一條紀錄
                box(a + j, b + i, c, d) = Color
            Next j
        Next i
       
    ElseIf NUMBER = 6 Then '方塊下面
        For i = 1 To 3
            For j = 1 To 3
                Input #1, Color '讀取一條紀錄
                box(a - j, b + i, c, d) = Color
            Next j
        Next i
        
    Else
    End If
    
End Sub
Private Sub Openfile(X As Integer) '開啟讀寫文件
    If X = 1 Then
        Open "123.txt" For Output As #1
    End If
    
    If X = 2 Then
        Open "123.txt" For Input As #1
    End If
    
    If X = 3 Then
        Close #1
    End If
End Sub
Private Sub colorSET() '設定顏色參數
R = 1 '紅色
L = 2 '黑色
G = 3 '綠色
U = 4 '藍色
W = 5 '黃色
E = 6 '橘色
End Sub
Private Sub choose(fox() As Integer, a, b, c) '建立選擇旋轉的方向和層數的方法
    If b = 0 And c = 0 Then
        Call X(a, fox) 'x軸旋轉
    End If

    If a = 0 And c = 0 Then
        Call Y(b, fox) 'y軸旋轉
    End If
    
    If a = 0 And b = 0 Then
        Call z(c, fox) 'z軸旋轉
    End If
End Sub
Private Sub choose1(fox() As Integer, a, b, c, show) '建立旋轉正方向的選擇
    Dim d As Integer
    d = 1
    
    Call choose(fox, a, b, c)
    If a <> 0 Then
        Debug.Print "X軸", a, "正轉"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    ElseIf b <> 0 Then
        Debug.Print "Y軸", b, "正轉"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    Else
        Debug.Print "Z軸", c, "正轉"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    End If
End Sub
Private Sub choose2(fox() As Integer, a, b, c, show) '建立旋轉反方向的選擇
    Dim d As Integer
    d = 2
    
    Call choose(fox, a, b, c)
    Call choose(fox, a, b, c)
    Call choose(fox, a, b, c)
    
    If a <> 0 Then
        Debug.Print "X軸", a, "反轉"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    ElseIf b <> 0 Then
        Debug.Print "Y軸", b, "反轉"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    Else
        Debug.Print "Z軸", c, "反轉"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    End If
End Sub
Private Sub X(a, fox() As Integer) '沿著x軸旋轉的方法
Dim i As Integer
Dim j As Integer

For i = 1 To 3
    For j = 1 To 3
        Call rotate(i2, j2, i, j)
        box2(a, i2, j2, 1) = fox(a, i, j, 1)
        box2(a, i2, j2, 2) = fox(a, i, j, 3)
        box2(a, i2, j2, 3) = fox(a, i, j, 2)
                
    Next j
Next i

'在扔回box1存取改變
For i = 1 To 3
    For j = 1 To 3
        fox(a, i, j, 1) = box2(a, i, j, 1)
        fox(a, i, j, 2) = box2(a, i, j, 2)
        fox(a, i, j, 3) = box2(a, i, j, 3)
                
    Next j
Next i
End Sub
Private Sub Y(b, fox() As Integer) '沿著y軸旋轉的方法
Dim i As Integer
Dim j As Integer

For i = 1 To 3
    For j = 1 To 3
        Call rotate(i2, j2, i, j)
        box2(j2, b, i2, 1) = fox(j, b, i, 3)
        box2(j2, b, i2, 2) = fox(j, b, i, 2)
        box2(j2, b, i2, 3) = fox(j, b, i, 1)
    Next j
Next i

'在扔回box1存取改變
For i = 1 To 3
    For j = 1 To 3
        fox(j, b, i, 1) = box2(j, b, i, 1)
        fox(j, b, i, 2) = box2(j, b, i, 2)
        fox(j, b, i, 3) = box2(j, b, i, 3)
                
    Next j
Next i
End Sub
Private Sub z(c, fox() As Integer) '沿著z軸旋轉的方法
Dim i As Integer
Dim j As Integer

'先用box2儲存改變
For i = 1 To 3
    For j = 1 To 3
        Call rotate(i2, j2, i, j) '前兩個是後來的座標
                                  '後兩個是原本的座標
        box2(i2, j2, c, 1) = fox(i, j, c, 2)
        box2(i2, j2, c, 2) = fox(i, j, c, 1)
        box2(i2, j2, c, 3) = fox(i, j, c, 3)
                
    Next j
Next i

For i = 1 To 3
    For j = 1 To 3
        fox(i, j, c, 1) = box2(i, j, c, 1)
        fox(i, j, c, 2) = box2(i, j, c, 2)
        fox(i, j, c, 3) = box2(i, j, c, 3)
                
    Next j
Next i
''

'測試
'For i = 1 To 3
'    For j = 1 To 3
'
'        Debug.Print i, j, "3  ", box2(i, j, 3, 1)
'        Debug.Print i, j, "3  ", box2(i, j, 3, 2)
'        Debug.Print i, j, "3  ", box2(i, j, 3, 3)
'
'    Next j
'Next i
End Sub
Private Sub rotate(a2, b2, a1, b1) '建立旋轉座標轉換的公式
For i = 1 To 3
    For j = 1 To 3
        a2 = b1
        b2 = 4 - a1
    Next j
Next i

End Sub

Private Sub firstEdgeBlock(X, Y, z, cx, cy, cz) '偵測第一層邊塊
    term = 0 '條件判斷變數先初始化
Debug.Print "判斷邊塊是否在正確位置"
    Call firstmethod(X, Y, z, cx, cy, cz)
    
    If term <> 1 And term <> 5 Then
Debug.Print "判斷邊塊是否在第二層"
        Call firstmethod3(X, Y, z, 1, 1, 2, cx, cy, cz)
        Call firstmethod3(X, Y, z, 1, 3, 2, cx, cy, cz)
        Call firstmethod3(X, Y, z, 3, 1, 2, cx, cy, cz)
        Call firstmethod3(X, Y, z, 3, 3, 2, cx, cy, cz)
    
    End If
    
    If term <> 1 And term <> 3 Then
Debug.Print "Z軸正確 判斷是否有在第一層"
        Call firstmethod1(X, Y, z, 1, 2, 1, cx, cy, cz)
        Call firstmethod1(X, Y, z, 2, 1, 1, cx, cy, cz)
        Call firstmethod1(X, Y, z, 2, 3, 1, cx, cy, cz)
        Call firstmethod1(X, Y, z, 3, 2, 1, cx, cy, cz)
Debug.Print "Z軸正確 判斷是否有在第三層"
        Call firstmethod1(X, Y, z, 1, 2, 3, cx, cy, cz)
        Call firstmethod1(X, Y, z, 2, 1, 3, cx, cy, cz)
        Call firstmethod1(X, Y, z, 2, 3, 3, cx, cy, cz)
        Call firstmethod1(X, Y, z, 3, 2, 3, cx, cy, cz)
    End If
    
    
    If term <> 1 And term <> 3 Then
Debug.Print "Z軸顛倒 判斷是否有在第一層"
        Call firstmethod2(X, Y, z, 1, 2, 1, cx, cy, cz)
        Call firstmethod2(X, Y, z, 2, 1, 1, cx, cy, cz)
        Call firstmethod2(X, Y, z, 2, 3, 1, cx, cy, cz)
        Call firstmethod2(X, Y, z, 3, 2, 1, cx, cy, cz)
Debug.Print "Z軸顛倒 判斷是否有在第三層"
        Call firstmethod2(X, Y, z, 1, 2, 3, cx, cy, cz)
        Call firstmethod2(X, Y, z, 2, 1, 3, cx, cy, cz)
        Call firstmethod2(X, Y, z, 2, 3, 3, cx, cy, cz)
        Call firstmethod2(X, Y, z, 3, 2, 3, cx, cy, cz)
    End If
  
End Sub

Private Sub firstCornerBlock(X, Y, z, cx, cy, cz) '偵測第一層角塊
    term = 0 '條件初始化
Debug.Print "角塊是否在正確位置"
    Call firstmethod4(X, Y, z, cx, cy, cz)
    
    If term <> 1 And term <> 3 Then
Debug.Print "角塊z軸是否黑色並進行處理"
        Call firstmethod5(X, Y, z, 1, 1, 1, cx, cy, cz)
        Call firstmethod5(X, Y, z, 1, 3, 1, cx, cy, cz)
        Call firstmethod5(X, Y, z, 3, 1, 1, cx, cy, cz)
        Call firstmethod5(X, Y, z, 3, 3, 1, cx, cy, cz)
        
        Call firstmethod5(X, Y, z, 1, 1, 3, cx, cy, cz)
        Call firstmethod5(X, Y, z, 1, 3, 3, cx, cy, cz)
        Call firstmethod5(X, Y, z, 3, 1, 3, cx, cy, cz)
        Call firstmethod5(X, Y, z, 3, 3, 3, cx, cy, cz)
        
    End If
    
    If term <> 1 And term <> 4 Then
Debug.Print "角塊是否z軸不為黑色並進行處理"
        Call firstmethod6(X, Y, z, 1, 1, 1, cx, cy, cz)
        Call firstmethod6(X, Y, z, 1, 3, 1, cx, cy, cz)
        Call firstmethod6(X, Y, z, 3, 1, 1, cx, cy, cz)
        Call firstmethod6(X, Y, z, 3, 3, 1, cx, cy, cz)
        
        Call firstmethod6(X, Y, z, 1, 1, 3, cx, cy, cz)
        Call firstmethod6(X, Y, z, 1, 3, 3, cx, cy, cz)
        Call firstmethod6(X, Y, z, 3, 1, 3, cx, cy, cz)
        Call firstmethod6(X, Y, z, 3, 3, 3, cx, cy, cz)
    End If

End Sub
Private Sub firstmethod(X, Y, z, cx, cy, cz)
    '判斷是否在原位
    Debug.Print "判斷是否在原位"
    
    If box(X, Y, z, 1) = cx And box(X, Y, z, 2) = cy And box(X, Y, z, 3) = cz Then
        term = 1
        Debug.Print "有在原位"
    End If
End Sub
Private Sub firstmethod1(x1, y1, z1, x2, y2, z2, cx, cy, cz)
    '判斷在第一層或第三層
    '判斷條件z軸方向為必為黑色
    If term <> 3 Then
        If box(x1, y1, z1, 1) <> 0 Then
        
            If cx = box(x2, y2, z2, 1) And cz = box(x2, y2, z2, 3) Then
                term = 2
            End If
    
            If cx = box(x2, y2, z2, 2) And cz = box(x2, y2, z2, 3) Then
                term = 2
            End If
        End If
        
        If box(x1, y1, z1, 2) <> 0 Then
        
            If cy = box(x2, y2, z2, 1) And cz = box(x2, y2, z2, 3) Then
                term = 2
            End If
    
            If cy = box(x2, y2, z2, 2) And cz = box(x2, y2, z2, 3) Then
                term = 2
            End If
        End If
    End If
    
    If term = 2 Then
        Call firstmethod1_1(x1, y1, z1, x2, y2, z2)
    End If

End Sub
Private Sub firstmethod1_1(x1, y1, z1, x2, y2, z2)
Debug.Print "進入旋轉處理firstmethod1_1"
        term = 3
        '轉到第三層處理
        If x2 <> 2 And z2 <> 3 Then
            Call secondchoose1(box, x2, 0, 0, 1)
            Call secondchoose1(box, x2, 0, 0, 1)
        End If

        If y2 <> 2 And z2 <> 3 Then
            Call secondchoose1(box, 0, y2, 0, 1)
            Call secondchoose1(box, 0, y2, 0, 1)
        End If
            
        '在第三層旋轉
        Call firstmethod2_2(x1, x2, y1, y2)
            
        '轉回第一層
        If x1 <> 2 Then
            Call secondchoose1(box, x1, 0, 0, 1)
            Call secondchoose1(box, x1, 0, 0, 1)
        End If
        
        If y1 <> 2 Then
            Call secondchoose1(box, 0, y1, 0, 1)
            Call secondchoose1(box, 0, y1, 0, 1)
        End If
End Sub
Private Sub firstmethod2(x1, y1, z1, x2, y2, z2, cx, cy, cz)
    '判斷在第一層或第三層
    '但是z軸顏色顛倒
    If term <> 3 Then
        If box(x1, y1, z1, 1) <> 0 Then
        
            If cx = box(x2, y2, z2, 3) And cz = box(x2, y2, z2, 1) Then
                term = 2
            End If
            
            If cx = box(x2, y2, z2, 3) And cz = box(x2, y2, z2, 2) Then
                term = 2
            End If
        End If
        
        If box(x1, y1, z1, 2) <> 0 Then
        
            If cy = box(x2, y2, z2, 3) And cz = box(x2, y2, z2, 1) Then
                term = 2
            End If
 
            If cy = box(x2, y2, z2, 3) And cz = box(x2, y2, z2, 2) Then
                term = 2
            End If
        End If
        
    End If
    
    If term = 2 Then
        Call firstmethod2_1(x1, y1, z1, x2, y2, z2)
    End If
    
End Sub
Private Sub firstmethod2_1(x1, y1, z1, x2, y2, z2)
Debug.Print "進入旋轉處理firstmethod2_1"
        term = 3
        '轉到第三層處理抓取位置座標
        If x2 <> 2 And z2 <> 3 Then
            Call secondchoose1(box, x2, 0, 0, 1)
            Call secondchoose1(box, x2, 0, 0, 1)
        End If

        If y2 <> 2 And z2 <> 3 Then
            Call secondchoose1(box, 0, y2, 0, 1)
            Call secondchoose1(box, 0, y2, 0, 1)
        End If
            
        '在第三層旋轉
        Call firstmethod2_2(x1, x2, y1, y2)
            
        '轉回第一層正確位置座標
        If x1 <> 2 Then
            
            Call secondchoose2(box, x1, 0, 0, 1)
            Call choose2(box, 0, 0, 2, 1)
            Call secondchoose1(box, x1, 0, 0, 1)
            Call choose1(box, 0, 0, 2, 1)
        End If
        
        If y1 <> 2 Then
            Call secondchoose2(box, 0, y1, 0, 1)
            Call choose2(box, 0, 0, 2, 1)
            Call secondchoose1(box, 0, y1, 0, 1)
            Call choose1(box, 0, 0, 2, 1)
        End If
        
End Sub
Private Sub firstmethod2_2(x1, x2, y1, y2) '副程式----------由firstmethod2_1-呼叫
'目的將在第三層的邊塊轉到第一層的位置上的xy軸   以利於置入第二層中
'1為正確位置 2為抓取位置
    If x1 = x2 And y1 = y2 Then
        
    ElseIf x1 = x2 Or y1 = y2 Then
        Call choose1(box, 0, 0, 3, 1)
        Call choose1(box, 0, 0, 3, 1)
    ElseIf x1 <> x2 And y1 <> y2 Then '
        If x2 = 1 Then
            If y1 < y2 Then
                Call choose2(box, 0, 0, 3, 1)
            Else
                Call choose1(box, 0, 0, 3, 1)
            End If
        End If
            
        If x2 = 3 Then
            If y1 < y2 Then
                Call choose1(box, 0, 0, 3, 1)
            Else
                Call choose2(box, 0, 0, 3, 1)
            End If
        End If
            
        If y2 = 1 Then
            If x1 < x2 Then
                Call choose1(box, 0, 0, 3, 1)
            Else
                Call choose2(box, 0, 0, 3, 1)
            End If
        End If
        
        If y2 = 3 Then
            If x1 < x2 Then
                Call choose2(box, 0, 0, 3, 1)
            Else
                Call choose1(box, 0, 0, 3, 1)
            End If
        End If
    Else
        Debug.Print "firstmethod2_2出錯"
            
        
    End If
End Sub
Private Sub firstmethod3(x1, y1, z1, x2, y2, z2, cx, cy, cz)
    '判斷邊塊在第二層
    '座標1為正確位置 座標2為抓取位置
    If (box(x2, y2, z2, 1) = cx And box(x2, y2, z2, 2) = cz) Or (box(x2, y2, z2, 1) = cz And box(x2, y2, z2, 2) = cx) Then
        term = 2
    End If
    
    If (box(x2, y2, z2, 1) = cy And box(x2, y2, z2, 2) = cz) Or (box(x2, y2, z2, 1) = cz And box(x2, y2, z2, 2) = cy) Then
        term = 2
    End If
        
    If term = 2 Then
        Call firstmethod3_1(x1, y1, z1, x2, y2, z2, cx, cy, cz)
    End If

End Sub
Private Sub firstmethod3_1(x1, y1, z1, x2, y2, z2, cx, cy, cz)
'將在第二層的邊塊轉到第三層
Debug.Print "進入旋轉處理firstmethod3_1"
    term = 5
    If x2 = 1 Then
    
        If y2 = 1 Then
            Call secondchoose2(box, 1, 0, 0, 1)
            Call choose2(box, 0, 0, 3, 1)
            Call secondchoose1(box, 1, 0, 0, 1)
        ElseIf y2 = 3 Then
            Call secondchoose1(box, 1, 0, 0, 1)
            Call choose2(box, 0, 0, 3, 1)
            Call secondchoose2(box, 1, 0, 0, 1)
        Else
Debug.Print "firstmethod3-1 X2=1 Y2錯誤"
        End If
        
    ElseIf x2 = 3 Then
    
        If y2 = 1 Then
            Call secondchoose1(box, 3, 0, 0, 1)
            Call choose2(box, 0, 0, 3, 1)
            Call secondchoose2(box, 3, 0, 0, 1)
        ElseIf y2 = 3 Then
            Call secondchoose2(box, 3, 0, 0, 1)
            Call choose2(box, 0, 0, 3, 1)
            Call secondchoose1(box, 3, 0, 0, 1)
        Else
Debug.Print "firstmethod3-1 X2=1 Y2 錯誤"
        End If
        
    Else
Debug.Print "firstmethod3-1 X2錯誤"
    End If
    
End Sub
Private Sub firstmethod4(x1, y1, z1, cx, cy, cz) '角塊
    '是否在正確的位置上
    If box(x1, y1, z1, 1) = cx And box(x1, y1, z1, 2) = cy And box(x1, y1, z1, 3) = cz Then
        term = 1
    End If
End Sub
Private Sub firstmethod5(x1, y1, z1, x2, y2, z2, cx, cy, cz) '角塊
'判斷條件Z軸為黑色
    If box(x2, y2, z2, 1) = cx And box(x2, y2, z2, 2) = cy And box(x2, y2, z2, 3) = cz Then
        term = 2
    End If
    If box(x2, y2, z2, 1) = cy And box(x2, y2, z2, 2) = cx And box(x2, y2, z2, 3) = cz Then
        term = 2
    End If
    
    If term = 2 Then
        Call firstmethod5_1(x1, y1, z1, x2, y2, z2, cx, cy, cz)
    End If
    
End Sub
Private Sub firstmethod5_1(x1, y1, z1, x2, y2, z2, cx, cy, cz) '角塊
'條件Z軸為黑色
'處理:邊塊z軸不為黑色
Debug.Print "進入旋轉處理firstmethod5_1"
    term = 3
    If z2 = 1 Then
        If x2 = y2 Then
            Call secondchoose1(box, 0, y2, 0, 1)
            Call choose1(box, 0, 0, 3, 1)
            Call secondchoose2(box, 0, y2, 0, 1)
        
        ElseIf x2 <> y2 Then
            Call secondchoose2(box, 0, y2, 0, 1)
            Call choose2(box, 0, 0, 3, 1)
            Call secondchoose1(box, 0, y2, 0, 1)
            
        Else
Debug.Print "firstmethod5_1 z2=1錯誤"

        End If
    End If
    
    If z2 = 3 Then
        If x2 = 1 And y2 = 1 Then
            Call choose1(box, 0, 0, 3, 1)
            Call choose1(box, 0, 0, 3, 1)
            
        ElseIf x2 = 1 And y2 = 3 Then
            Call choose1(box, 0, 0, 3, 1)
        
        ElseIf x2 = 2 And y2 = 1 Then
            Call choose2(box, 0, 0, 3, 1)
        Else
        
        End If
        
        Call secondchoose1(box, 0, 3, 0, 1)
        Call choose2(box, 0, 0, 3, 1)
        Call secondchoose2(box, 0, 3, 0, 1)
            
        Else
Debug.Print "firstmethod5_1 z2=3錯誤"

        
        
    End If
    

End Sub
Private Sub firstmethod6(x1, y1, z1, x2, y2, z2, cx, cy, cz) '角塊
'判斷條件Z軸不為黑色
    If box(x2, y2, z2, 1) = cz And box(x2, y2, z2, 2) = cy And box(x2, y2, z2, 3) = cx Then
        term = 2
    End If
    If box(x2, y2, z2, 1) = cy And box(x2, y2, z2, 2) = cz And box(x2, y2, z2, 3) = cx Then
        term = 2
    End If
    If box(x2, y2, z2, 1) = cx And box(x2, y2, z2, 2) = cz And box(x2, y2, z2, 3) = cy Then
        term = 2
    End If
    If box(x2, y2, z2, 1) = cz And box(x2, y2, z2, 2) = cx And box(x2, y2, z2, 3) = cy Then
        term = 2
    End If
    
    If term = 2 Then
        Call FIRSTMETHOD6_1(x1, y1, z1, x2, y2, z2, cx, cy, cz)
    End If
End Sub
Private Sub FIRSTMETHOD6_1(x1, y1, z1, x2, y2, z2, cx, cy, cz) '角塊
'如果在第一層先轉到第三層
Debug.Print "進入旋轉處理firstmethod6_1"
    term = 4
    If z2 = 1 Then
Debug.Print "如果在第一層先轉到第三層"
        If x2 = y2 Then
            If box(x2, y2, z2, 1) = cz Then
                Call secondchoose2(box, x2, 0, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose1(box, x2, 0, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
            
            
            ElseIf box(x2, y2, z2, 2) = cz Then
                Call secondchoose1(box, 0, y2, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose2(box, 0, y2, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
            End If
            
        Else
            If box(x2, y2, z2, 1) = cz Then
                Call secondchoose1(box, x2, 0, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose2(box, x2, 0, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
            
            
            ElseIf box(x2, y2, z2, 2) = cz Then
                Call secondchoose2(box, 0, y2, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose1(box, 0, y2, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
            End If
        End If
            
    End If
    
'在第三層旋轉到適當位置
Debug.Print "在第三層旋轉到適當位置"
    If x1 <> x2 And y1 <> y2 Then
        Call choose1(box, 0, 0, 3, 1)
        Call choose1(box, 0, 0, 3, 1)
    ElseIf x1 = x2 And y1 = y2 Then
    
    ElseIf x1 <> x2 Or y1 <> y2 Then
        If x1 = x2 And x1 = 1 Then
        
            If y1 < y2 Then
                Call choose2(box, 0, 0, 3, 1)
            Else
                Call choose1(box, 0, 0, 3, 1)
            End If
        ElseIf x1 = x2 And x1 = 3 Then
        
            If y1 < y2 Then
                Call choose1(box, 0, 0, 3, 1)
            Else
                Call choose2(box, 0, 0, 3, 1)
            End If
        ElseIf y1 = y2 And y1 = 1 Then
        
            If x1 < x2 Then
                Call choose1(box, 0, 0, 3, 1)
            Else
                Call choose2(box, 0, 0, 3, 1)
            End If
        ElseIf y1 = y2 And y1 = 3 Then
        
            If x1 < x2 Then
                Call choose2(box, 0, 0, 3, 1)
            Else
                Call choose1(box, 0, 0, 3, 1)
            End If
        End If
    Else
Debug.Print "firstmethod6_1 cx cy錯誤"
    End If
'將邊塊轉到正確的位置
    If x1 = y1 Then
        If box(x1, y1, 3, 2) = cz Then
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, x1, 0, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, x1, 0, 0, 1)
            
            Else
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, 0, y1, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, 0, y1, 0, 1)
            End If
    Else
        
        If box(x1, y1, 3, 2) = cz Then
            Call choose1(box, 0, 0, 3, 1)
            Call secondchoose1(box, x1, 0, 0, 1)
            Call choose2(box, 0, 0, 3, 1)
            Call secondchoose2(box, x1, 0, 0, 1)
            
            Else
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, 0, y1, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, 0, y1, 0, 1)
            End If
    End If
    
            
            
        
End Sub


Private Sub secondmethod(X, Y, z, cx, cy, cz) '副程式----------由主程式呼叫
'第二層單一邊塊位置判斷

    term = 0 '條件判斷變數先初始化
    
    '判斷是否在原位
    Debug.Print "判斷是否在原位"
    
    If box(X, Y, z, 1) = cx And box(X, Y, z, 2) = cy Then
        term = 1
        Debug.Print "有在原位"
    End If
    
    
    '判斷是否在第二層的某個位置
    Debug.Print "判斷是否在第二層"
    
    If term <> 1 Then
        Debug.Print "判斷在第二層的哪裡"
        Call secondmethod1_1(1, 1, 2, cx, cy, cz)
            Debug.Print "透過secondmethod1-1判斷112"
        Call secondmethod1_1(1, 3, 2, cx, cy, cz)
            Debug.Print "透過secondmethod1-1判斷132"
        Call secondmethod1_1(3, 1, 2, cx, cy, cz)
            Debug.Print "透過secondmethod1-1判斷312"
        Call secondmethod1_1(3, 3, 2, cx, cy, cz)
            Debug.Print "透過secondmethod1-1判斷332"
        
    End If
    
    '判斷是否在第三層的某個位置
    Debug.Print "判斷是否在第三層"
    
    If term <> 1 Then
        Debug.Print "判斷在第三層的哪裡"
            Debug.Print "透過secondmethod1-2判斷123"
        Call secondmethod1_2(1, 2, 3, cx, cy, cz)
            Debug.Print "透過secondmethod1-2判斷213"
        Call secondmethod1_2(2, 1, 3, cx, cy, cz)
            Debug.Print "透過secondmethod1-2判斷233"
        Call secondmethod1_2(2, 3, 3, cx, cy, cz)
            Debug.Print "透過secondmethod1-2判斷323"
        Call secondmethod1_2(3, 2, 3, cx, cy, cz)
            
        '置入第二層
            Debug.Print "透過secondmethod1-3將邊塊置入第二層"
        Call secondmethod1_3(X, Y, z)
             
        
    
    End If




End Sub
Private Sub secondmethod1_1(X, Y, z, cx, cy, cz) '副程式---------由secondmethod系列-呼叫
    '判斷是否在第二層的某個位置上
    Debug.Print "副程式判斷邊塊在第二層的哪裡"
    
    Dim termsecond1_1 As Integer
    If (box(X, Y, z, 1) = cx And box(X, Y, z, 2) = cy) Or (box(X, Y, z, 1) = cy And box(X, Y, z, 2) = cx) Then
        term = 2
        termsecond1_1 = 2
    End If
    
    '如果抓到位置再對其進行處理
    '處理方式為將邊塊移至第三層
    If termsecond1_1 = 2 Then
        Debug.Print "第二層移至第三層的旋轉步驟"
        If X = Y Then
            Call secondchoose1(box, 0, Y, 0, 1)
            Call choose2(box, 0, 0, 3, 1)
            Call secondchoose2(box, 0, Y, 0, 1)
            Call choose2(box, 0, 0, 3, 1)
            Call secondchoose2(box, X, 0, 0, 1)
            Call choose1(box, 0, 0, 3, 1)
            Call secondchoose1(box, X, 0, 0, 1)
        
        Else
            Call secondchoose1(box, X, 0, 0, 1)
            Call choose2(box, 0, 0, 3, 1)
            Call secondchoose2(box, X, 0, 0, 1)
            Call choose2(box, 0, 0, 3, 1)
            Call secondchoose2(box, 0, Y, 0, 1)
            Call choose1(box, 0, 0, 3, 1)
            Call secondchoose1(box, 0, Y, 0, 1)
  
        End If
    End If
    
    
End Sub
Private Sub secondmethod1_2(X, Y, z, cx, cy, cz) '副程式----------由secondmethod系列-呼叫
    '判斷是否在第三層的某個位置上
    Debug.Print "副程式判斷邊塊在第三層的哪裡"
    
    Dim termsecond1_2 As Integer
    If (box(X, Y, z, 1) = cx And box(X, Y, z, 3) = cy) Or (box(X, Y, z, 1) = cy And box(X, Y, z, 3) = cx) Then
        termsecond1_2 = 2
    End If
    
    If (box(X, Y, z, 2) = cx And box(X, Y, z, 3) = cy) Or (box(X, Y, z, 2) = cy And box(X, Y, z, 3) = cx) Then
        termsecond1_2 = 2
    End If
    
    
    Debug.Print "判斷變數termsecond1_2", termsecond1_2
    
    '如果抓到位置(在第三層)再對其進行處理
    If termsecond1_2 = 2 Then
        Debug.Print "進入第三層的適當位置旋轉步驟"
        
        '先讓在第三層的邊塊旋轉到第三層的適當位置
        If box(X, Y, z, 3) = box(1, 2, 2, 1) Then
            Call secondmethod1_2_1(X, 1, Y, 2)
            term = 3 '先轉x軸
            
        ElseIf box(X, Y, z, 3) = box(2, 1, 2, 2) Then
            Call secondmethod1_2_1(X, 2, Y, 1)
            term = 4 '先轉y軸
            
        ElseIf box(X, Y, z, 3) = box(2, 3, 2, 2) Then
            Call secondmethod1_2_1(X, 2, Y, 3)
            term = 4 '先轉y軸
            
        ElseIf box(X, Y, z, 3) = box(3, 2, 2, 1) Then
            Call secondmethod1_2_1(X, 3, Y, 2)
            term = 3 '先轉x軸
            
        Else
            Debug.Print "secondmethod1_2出錯"
        End If
       
    End If
End Sub
Private Sub secondmethod1_2_1(x1, x2, y1, y2) '副程式----------由secondmethod1_2-呼叫
'目的將在第三層的邊塊轉到正確的位置上
'以利於置入第二層中
    If x1 = x2 And y1 = y2 Then
        Call choose1(box, 0, 0, 3, 1)
        Call choose1(box, 0, 0, 3, 1)
    ElseIf x1 = x2 Or y1 = y2 Then
        
    ElseIf x1 <> x2 And y1 <> y2 Then
        If x1 = 1 Then
            If y1 < y2 Then
                Call choose2(box, 0, 0, 3, 1)
            Else
                Call choose1(box, 0, 0, 3, 1)
            End If
        End If
        
        If x1 = 3 Then
            If y1 < y2 Then
                Call choose1(box, 0, 0, 3, 1)
            Else
                Call choose2(box, 0, 0, 3, 1)
            End If
        End If
        
        If y1 = 1 Then
            If x1 < x2 Then
                Call choose1(box, 0, 0, 3, 1)
            Else
                Call choose2(box, 0, 0, 3, 1)
            End If
        End If
        
        If y1 = 3 Then
            If x1 < x2 Then
                Call choose2(box, 0, 0, 3, 1)
            Else
                Call choose1(box, 0, 0, 3, 1)
            End If
        End If
        
    Else
        Debug.Print "secondmethod1_2_1出錯"
        
    End If
End Sub
Private Sub secondmethod1_3(X, Y, z)
'再讓邊塊置入第二層
        If X = Y Then
            If term = 3 Then    '先轉X軸
                Call secondchoose2(box, X, 0, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, X, 0, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, 0, Y, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, 0, Y, 0, 1)
            
            End If
            
            If term = 4 Then    '先轉Y軸
                Call secondchoose1(box, 0, Y, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, 0, Y, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, X, 0, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, X, 0, 0, 1)
                
            End If
            
        Else
            If term = 3 Then    '先轉X軸
                Call secondchoose1(box, X, 0, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, X, 0, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, 0, Y, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, 0, Y, 0, 1)
            
            End If
            
            If term = 4 Then    '先轉Y軸
                Call secondchoose2(box, 0, Y, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, 0, Y, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, X, 0, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, X, 0, 0, 1)
                
            End If
        
        End If
End Sub
Private Sub secondchoose(fox() As Integer, a, b, c) '副程式由secondmethod-系列呼叫
'建立側面常規順時針轉動
    If a = 1 Or b = 1 Then
        Call choose(fox, a, b, c)
        Call choose(fox, a, b, c)
        Call choose(fox, a, b, c)
        
    ElseIf a = 3 Or b = 3 Then
        Call choose(fox, a, b, c)
          
    Else
        Debug.Print "secondchoose出現錯誤"
    End If
    
End Sub
Private Sub secondchoose1(fox() As Integer, a, b, c, show) '副程式由secondmethod-系列呼叫
'建立側面常規逆時針轉動
    Dim d As Integer
    d = 3
    Call secondchoose(fox, a, b, c)
    
    If a <> 0 Then
        Debug.Print "X軸", a, "側面順時針轉"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
    End If
    
    If b <> 0 Then
        Debug.Print "Y軸", b, "側面順時針轉"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
    End If
End Sub
Private Sub secondchoose2(fox() As Integer, a, b, c, show) '副程式由secondmethod-系列呼叫
'建立側面常規逆時針轉動
    Dim d As Integer
    d = 4
    Call secondchoose(fox, a, b, c)
    Call secondchoose(fox, a, b, c)
    Call secondchoose(fox, a, b, c)
    
    If a <> 0 Then
        
        Debug.Print "X軸", a, "側面逆時針轉"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
    End If
    
    If b <> 0 Then
        Debug.Print "Y軸", b, "側面逆時針轉"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    End If
End Sub


Private Sub thirdSET() '設定遮罩
    '先初始化
    For i = 1 To 3
        For j = 1 To 3
            Third0_1(i, j) = 0
            Third0_2(i, j) = 0
            Third1_1(i, j) = 0
            Third1_2(i, j) = 0
            Third1_3(i, j) = 0
            Third2_1(i, j) = 0
            Third2_2(i, j) = 0
            Third2_3(i, j) = 0
            Third2_4(i, j) = 0
            Third2_5(i, j) = 0
            Third2_6(i, j) = 0
            Third2_7(i, j) = 0
        Next j
    Next i
    
    '邊塊Z軸復原
        'Third1_1直角
    Third1_1(1, 2) = W
    Third1_1(2, 1) = W
    Third1_1(2, 2) = W
    
        'Third1_2直線
    Third1_2(1, 2) = W
    Third1_2(2, 2) = W
    Third1_2(3, 2) = W
    
        'Third1_3中心點
    Third1_3(2, 2) = W
    
    
    '頂面Z軸復原
        'Third2_1 c1和c2
    Third2_1(1, 1) = W
    Third2_1(1, 2) = W
    Third2_1(2, 1) = W
    Third2_1(2, 2) = W
    Third2_1(2, 3) = W
    Third2_1(3, 2) = W
    
        'Third2_2 c3和c4
    Third2_2(1, 2) = W
    Third2_2(2, 1) = W
    Third2_2(2, 2) = W
    Third2_2(2, 3) = W
    Third2_2(3, 1) = W
    Third2_2(3, 2) = W
    Third2_2(3, 3) = W
    
        'Third2_3 c5
    Third2_3(1, 2) = W
    Third2_3(1, 3) = W
    Third2_3(2, 1) = W
    Third2_3(2, 2) = W
    Third2_3(2, 3) = W
    Third2_3(3, 1) = W
    Third2_3(3, 2) = W
    
        'Third2_4 c6和c7
    Third2_4(1, 2) = W
    Third2_4(2, 1) = W
    Third2_4(2, 2) = W
    Third2_4(2, 3) = W
    Third2_4(3, 2) = W

End Sub
Private Sub thirdchoose1() '將通用模組Third0_1旋轉
Dim i As Integer
Dim j As Integer
    For i = 1 To 3
        For j = 1 To 3
            Call thirdchoose2(i2, j2, i, j)
            Third0_2(i2, j2) = Third0_1(i, j)
            
        Next j
    Next i
    
    For i = 1 To 3
        For j = 1 To 3
            Third0_1(i, j) = Third0_2(i, j)
            
        Next j
    Next i
    
End Sub

Private Sub thirdchoose2(i2, j2, i1, j1) '由Thirdchoose1呼叫
'進行座標轉換程序
    i2 = j1
    j2 = 4 - i1
End Sub
Private Sub thirdCopy(X() As Integer, Y() As Integer)
Dim i As Integer
Dim j As Integer
    For i = 1 To 3
        For j = 1 To 3
            X(i, j) = Y(i, j)
        Next j
    Next i
End Sub

Private Sub thirdmethod()
    Call colorSET
    Call thirdSET
Debug.Print "呼叫thirdmethod1_1"
    Call thirdmethod1_1
Debug.Print "呼叫thirdmethod2_1"
    Call thirdmethod2_1
Debug.Print "第三層角塊復原"
    Call thirdmethod3_1
Debug.Print "第三層邊塊復原"
    Call thirdmethod4_1
    
    
End Sub
Private Sub thirdmethod1_1() '將第三層邊塊z軸恢復
    term = 0
    
    Call thirdmethod1_3(Third2_4)
    
    If term <> 1 Then
        Call thirdmethod1_3(Third1_1)
        If term = 1 Then
            Call thirdmethod1_4
        End If
        
    End If
    
    If term <> 1 Then
        Call thirdmethod1_3(Third1_2)
        If term = 1 Then
            Call thirdmethod1_4
            Call thirdmethod1_4
        End If
    End If
    
    If term <> 1 Then
        Call thirdmethod1_3(Third1_3)
        If term = 1 Then
            Call thirdmethod1_4
            Call choose1(box, 0, 0, 3, 1)
            Call thirdmethod1_4
            Call thirdmethod1_4
        End If
    End If
    
End Sub
Private Sub thirdmethod2_1() '將第三層z軸面復原
    term = 0
    
    If term <> 1 Then
        Call thirdmethod2_3(Third2_2) 'c3 c4
        If term = 1 Then
            If box(1, 1, 3, 1) = W And box(1, 3, 3, 1) = W Then
                Call thirdmethod2_4
            
            ElseIf box(1, 1, 3, 2) = W And box(1, 3, 3, 2) = W Then
                Call choose2(box, 0, 0, 3, 1)
                Call thirdmethod2_4
            End If
            
        End If
    End If
    
    If term <> 1 Then
        Call thirdmethod2_3(Third2_3) 'c5
        If term = 1 Then
            If box(1, 1, 3, 1) = W And box(3, 3, 3, 2) = W Then
                Call thirdmethod2_4
            
            
            ElseIf box(1, 1, 3, 2) = W And box(3, 3, 3, 1) = W Then
                 Call choose1(box, 0, 0, 3, 1)
                 Call choose1(box, 0, 0, 3, 1)
                 Call thirdmethod2_4
            End If
            
        End If
    End If
    
    If term <> 1 Then
    
        Call thirdmethod2_3(Third2_4) 'c6 c7
        If term = 1 Then
            If box(1, 1, 3, 1) = W And box(1, 3, 3, 1) = W And box(3, 1, 3, 1) = W And box(3, 3, 3, 1) = W Then 'c7
                Call choose1(box, 0, 0, 3, 1)
                Call thirdmethod2_4
            
            ElseIf box(1, 1, 3, 2) = W And box(3, 3, 3, 2) = W And box(1, 1, 3, 2) = W And box(3, 3, 3, 2) = W Then 'c7
                
                Call thirdmethod2_4
                
            ElseIf box(1, 1, 3, 2) = W And box(1, 3, 3, 2) = W Then 'c6
                Call thirdmethod2_1
            
            ElseIf box(1, 1, 3, 2) <> W And box(1, 3, 3, 2) = W Then 'c6
                Call choose1(box, 0, 0, 3, 1)
                Call thirdmethod2_4
            
            ElseIf box(1, 1, 3, 2) <> W And box(1, 3, 3, 2) <> W Then 'c6
                Call choose1(box, 0, 0, 3, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call thirdmethod2_4
            
            
            ElseIf box(1, 1, 3, 2) = W And box(1, 3, 3, 2) <> W Then
                Call choose2(box, 0, 0, 3, 1)
                Call thirdmethod2_4
            
            End If

        End If
    End If
    
    term = 0
        Call thirdmethod2_3(Third2_1) 'c1 c2
        If term = 1 Then
            If box(1, 3, 3, 1) = W And box(3, 1, 3, 1) = W Then
                Call thirdmethod2_4
            End If
            
            If box(1, 3, 3, 2) = W And box(3, 1, 3, 2) = W Then
                Call thirdmethod2_5
            End If
                        
        End If
        
   
    
End Sub
Private Sub thirdmethod3_1() '將第三層角塊歸位
    term = 0
    If box(1, 1, 3, 1) = box(1, 3, 3, 1) And box(3, 1, 3, 1) = box(3, 3, 3, 1) Then
        term = 1
    End If
    
    If term <> 1 Then
        If box(1, 1, 3, 1) = box(1, 3, 3, 1) Then '上
            Call choose2(box, 0, 0, 3, 1)
            Call thirdmethod3_2
            
            Debug.Print "角塊1抓到"
        ElseIf box(1, 3, 3, 2) = box(3, 3, 3, 2) Then '右
            Call choose2(box, 0, 0, 3, 1)
            Call choose2(box, 0, 0, 3, 1)
            Call thirdmethod3_2
            
             Debug.Print "角塊2抓到"
        ElseIf box(3, 1, 3, 1) = box(3, 3, 3, 1) Then '下
            Call choose1(box, 0, 0, 3, 1)
            Call thirdmethod3_2
            
             Debug.Print "角塊3抓到"
        ElseIf box(1, 1, 3, 2) = box(3, 1, 3, 2) Then '左
            Call thirdmethod3_2
        Else
            Call thirdmethod3_2
            Call choose2(box, 0, 0, 3, 1)
            Call thirdmethod3_2
        End If
    End If

End Sub
Private Sub thirdmethod4_1() '將第三層的邊塊歸位
    term = 1
    
    Call thirdmethod4_4
    
    If term = 0 Then
        Call thirdmethod4_2
    
        If term = 2 Then
            Call thirdmethod4_2
        End If
    End If
    
    Call thirdmethod4_5
    
       
End Sub
Private Sub thirdmethod1_2(Thirdterm, X() As Integer)
'由THIRDMETHOD1_3呼叫
'與遮罩比較
Dim i As Integer
Dim j As Integer
    For i = 1 To 3
        For j = 1 To 3
            If X(i, j) = W And Thirdterm <> 2 Then
                If box(i, j, 3, 3) = W Then
                    Thirdterm = 1
                Else
                    Thirdterm = 2
                End If
            End If
        Next j
    Next i
End Sub
Private Sub thirdmethod1_3(X() As Integer)
'由THIRDMETHOD1_1呼叫
'呼叫此方法尋找適合的遮罩
Dim i As Integer
Dim i2 As Integer
Dim j As Integer
Dim k As Integer

 
        For i = 0 To 4
Debug.Print "運行第", i, "次"

            Call thirdmethod1_2(Thirdterm, X) '與遮罩比較
            If Thirdterm = 1 Then
Debug.Print "抓取成功"
                     
            End If
            
            '測試
           ' For j = 1 To 3
           '     For k = 1 To 3
           '         Debug.Print j, k, x(j, k)
           '     Next k
           ' Next j
            Debug.Print ""
            
            
            '旋轉
            Call thirdCopy(Third0_1, X)
            Call thirdchoose1
            Call thirdCopy(X, Third0_1)
            
            If Thirdterm = 1 Then term = 1
            If Thirdterm = 1 Then Exit For
            Thirdterm = 0
        Next i
        
        Debug.Print "測試I", i
        While (i <> 0 And i < 4)
            i = i - 1
            Call choose2(box, 0, 0, 3, 1)
        Wend
        
End Sub
Private Sub thirdmethod1_4() '第三層邊塊復原z面公式
    Call secondchoose2(box, 0, 3, 0, 1)
    Call choose2(box, 0, 0, 3, 1)
    Call secondchoose2(box, 3, 0, 0, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
    Call secondchoose1(box, 0, 3, 0, 1)
End Sub
Private Sub thirdmethod2_2(Thirdterm, X() As Integer)
'由THIRDMETHOD1_3呼叫
'與遮罩比較
Dim i As Integer
Dim j As Integer
    For i = 1 To 3
        For j = 1 To 3
            If (X(i, j) = W Or box(i, j, 3, 3) = W) And Thirdterm <> 2 Then
                If X(i, j) = box(i, j, 3, 3) Then
                    Thirdterm = 1
                Else
                    Thirdterm = 2
                End If
                
                
                
            End If
        Next j
    Next i
End Sub
Private Sub thirdmethod2_3(X() As Integer)
'由THIRDMETHOD1_1呼叫
'呼叫此方法尋找適合的遮罩
Dim i As Integer
Dim i2 As Integer
Dim j As Integer
Dim k As Integer

 
        For i = 0 To 4
Debug.Print "運行第", i, "次"

            Call thirdmethod2_2(Thirdterm, X) '與遮罩比較
            If Thirdterm = 1 Then
Debug.Print "抓取成功"
                     
            End If
            
            '測試
           ' For j = 1 To 3
           '     For k = 1 To 3
           '         Debug.Print j, k, x(j, k)
           '     Next k
           ' Next j
            Debug.Print ""
            
            
            '旋轉
            Call thirdCopy(Third0_1, X)
            Call thirdchoose1
            Call thirdCopy(X, Third0_1)
            
            If Thirdterm = 1 Then term = 1
            If Thirdterm = 1 Then Exit For
            
            Thirdterm = 0
        Next i
        
        Debug.Print "測試I", i
        While (i <> 0 And i < 4)
            i = i - 1
            Call choose2(box, 0, 0, 3, 1)
        Wend
        
End Sub
Private Sub thirdmethod2_4() 'c1第三層復原z面公式
    Call secondchoose2(box, 0, 3, 0, 1)
    Call choose2(box, 0, 0, 3, 1)
    Call secondchoose1(box, 0, 3, 0, 1)
    Call choose2(box, 0, 0, 3, 1)
    Call secondchoose2(box, 0, 3, 0, 1)
    Call choose2(box, 0, 0, 3, 1)
    Call choose2(box, 0, 0, 3, 1)
    Call secondchoose1(box, 0, 3, 0, 1)

End Sub
Private Sub thirdmethod2_5() 'c2第三層復原z面公式
    Call secondchoose1(box, 3, 0, 0, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call secondchoose2(box, 3, 0, 0, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call secondchoose2(box, 3, 0, 0, 1)
End Sub
Private Sub thirdmethod3_2() '第三層換角公式
    Call secondchoose1(box, 3, 0, 0, 1)
    Call secondchoose2(box, 0, 3, 0, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
    Call secondchoose2(box, 0, 1, 0, 1)
    Call secondchoose2(box, 0, 1, 0, 1)
    Call secondchoose2(box, 3, 0, 0, 1)
    Call secondchoose1(box, 0, 3, 0, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
    Call secondchoose1(box, 0, 1, 0, 1)
    Call secondchoose1(box, 0, 1, 0, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
    Call secondchoose1(box, 3, 0, 0, 1)

End Sub
Private Sub thirdmethod4_2()
    If box(1, 1, 3, 1) = box(1, 2, 3, 1) And box(1, 3, 3, 1) = box(1, 2, 3, 1) Then '上
            Call thirdmethod4_3
        
        ElseIf box(1, 3, 3, 2) = box(2, 3, 3, 2) And box(3, 3, 3, 2) = box(2, 3, 3, 2) Then '右
            Call choose2(box, 0, 0, 3, 1)
            Call thirdmethod4_3
        
        ElseIf box(3, 1, 3, 1) = box(3, 2, 3, 1) And box(3, 3, 3, 1) = box(3, 2, 3, 1) Then '下
            Call choose1(box, 0, 0, 3, 1)
            Call choose1(box, 0, 0, 3, 1)
            Call thirdmethod4_3
    
        ElseIf box(1, 1, 3, 2) = box(2, 1, 3, 2) And box(3, 1, 3, 2) = box(2, 1, 3, 2) Then '左
            Call choose1(box, 0, 0, 3, 1)
            Call thirdmethod4_3
            
        Else
            term = 2
            Call thirdmethod4_3_1
            
        End If
End Sub
Private Sub thirdmethod4_3()
    If box(2, 1, 3, 2) = box(3, 1, 3, 1) Then '逆時針觸發
        Call thirdmethod4_3_1
    End If
    
    If box(2, 3, 3, 2) = box(3, 1, 3, 1) Then '順時針觸發
        Call thirdmethod4_3_2
    End If
    
End Sub
Private Sub thirdmethod4_3_1() '第三層逆時針三角換邊
    Call secondchoose1(box, 3, 0, 0, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
    Call choose2(box, 0, 0, 3, 1)
    Call choose1(box, 0, 2, 0, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call choose2(box, 0, 2, 0, 1)
    Call choose2(box, 0, 0, 3, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
End Sub
Private Sub thirdmethod4_3_2() '第三層順時針三角換邊
    Call secondchoose1(box, 3, 0, 0, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call choose1(box, 0, 2, 0, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call choose2(box, 0, 2, 0, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
End Sub

Private Sub thirdmethod4_4()
    If box(1, 1, 3, 1) <> box(1, 2, 3, 1) Then '上
        term = 0
    End If
        
    If box(1, 3, 3, 2) <> box(2, 3, 3, 2) Then '右
        term = 0
            
    End If
        
    If box(3, 1, 3, 1) <> box(3, 2, 3, 1) Then '下
        term = 0
            
    End If
        
    If box(1, 1, 3, 2) <> box(2, 1, 3, 2) Then  '左
        term = 0
            
    End If
End Sub
Private Sub thirdmethod4_5()
    While box(1, 2, 3, 1) <> box(1, 2, 2, 1)
        
        Call choose1(box, 0, 0, 3, 1)
    Wend
    
    
      
End Sub
Private Sub Command2_Click()
    
    Timer1.Enabled = True
    
End Sub
Private Sub cmdShow_Click()
'測試用
    
    Call Openfile(2) '開啟檔案
    
    Call showmethod3
    
    
    
End Sub
Private Sub showmethod1()
    Dim a, b, c, d As Integer
    
    Input #1, a, b, c, d '讀取一條紀錄
    Call showmethod4(a, b, c, d)
        
    If d = 1 Then
        Call choose1(showBox, a, b, c, 0)
    End If
    
    If d = 2 Then
        Call choose2(showBox, a, b, c, 0)
    End If
    
    If d = 3 Then
        Call secondchoose1(showBox, a, b, c, 0)
    End If
    
    If d = 4 Then
        Call secondchoose2(showBox, a, b, c, 0)
    End If
    
    Call showmethod3
    
End Sub
Private Sub showmethod2() '紀錄初始狀態
    Dim i, j, k As Integer
    For k = 1 To 3
        For i = 1 To 3
            For j = 1 To 3
                showBox(i, j, k, 1) = box(i, j, k, 1)
                showBox(i, j, k, 2) = box(i, j, k, 2)
                showBox(i, j, k, 3) = box(i, j, k, 3)
            Next j
        Next i
    Next k
End Sub
Private Sub showmethod3() '介面顏色顯示

    Call showmethod3_1
End Sub
Private Sub showmethod3_1()
    
    'X前面
    Call showmethod3_2(showBox(3, 1, 3, 1), Shape1(1))
    Call showmethod3_2(showBox(3, 2, 3, 1), Shape1(2))
    Call showmethod3_2(showBox(3, 3, 3, 1), Shape1(3))
    Call showmethod3_2(showBox(3, 1, 2, 1), Shape1(4))
    Call showmethod3_2(showBox(3, 2, 2, 1), Shape1(5))
    Call showmethod3_2(showBox(3, 3, 2, 1), Shape1(6))
    Call showmethod3_2(showBox(3, 1, 1, 1), Shape1(7))
    Call showmethod3_2(showBox(3, 2, 1, 1), Shape1(8))
    Call showmethod3_2(showBox(3, 3, 1, 1), Shape1(9))
    'Y右面
    Call showmethod3_2(showBox(3, 3, 3, 2), Shape3(1))
    Call showmethod3_2(showBox(2, 3, 3, 2), Shape3(2))
    Call showmethod3_2(showBox(1, 3, 3, 2), Shape3(3))
    Call showmethod3_2(showBox(3, 3, 2, 2), Shape3(4))
    Call showmethod3_2(showBox(2, 3, 2, 2), Shape3(5))
    Call showmethod3_2(showBox(1, 3, 2, 2), Shape3(6))
    Call showmethod3_2(showBox(3, 3, 1, 2), Shape3(7))
    Call showmethod3_2(showBox(2, 3, 1, 2), Shape3(8))
    Call showmethod3_2(showBox(1, 3, 1, 2), Shape3(9))
    'X後面
    Call showmethod3_2(showBox(1, 1, 3, 1), Shape6(1))
    Call showmethod3_2(showBox(1, 2, 3, 1), Shape6(2))
    Call showmethod3_2(showBox(1, 3, 3, 1), Shape6(3))
    Call showmethod3_2(showBox(1, 1, 2, 1), Shape6(4))
    Call showmethod3_2(showBox(1, 2, 2, 1), Shape6(5))
    Call showmethod3_2(showBox(1, 3, 2, 1), Shape6(6))
    Call showmethod3_2(showBox(1, 1, 1, 1), Shape6(7))
    Call showmethod3_2(showBox(1, 2, 1, 1), Shape6(8))
    Call showmethod3_2(showBox(1, 3, 1, 1), Shape6(9))
    'Y左面
    Call showmethod3_2(showBox(3, 1, 3, 2), Shape4(1))
    Call showmethod3_2(showBox(2, 1, 3, 2), Shape4(2))
    Call showmethod3_2(showBox(1, 1, 3, 2), Shape4(3))
    Call showmethod3_2(showBox(3, 1, 2, 2), Shape4(4))
    Call showmethod3_2(showBox(2, 1, 2, 2), Shape4(5))
    Call showmethod3_2(showBox(1, 1, 2, 2), Shape4(6))
    Call showmethod3_2(showBox(3, 1, 1, 2), Shape4(7))
    Call showmethod3_2(showBox(2, 1, 1, 2), Shape4(8))
    Call showmethod3_2(showBox(1, 1, 1, 2), Shape4(9))
    'Z面
    Call showmethod3_2(showBox(1, 1, 3, 3), Shape2(1))
    Call showmethod3_2(showBox(1, 2, 3, 3), Shape2(2))
    Call showmethod3_2(showBox(1, 3, 3, 3), Shape2(3))
    Call showmethod3_2(showBox(2, 1, 3, 3), Shape2(4))
    Call showmethod3_2(showBox(2, 2, 3, 3), Shape2(5))
    Call showmethod3_2(showBox(2, 3, 3, 3), Shape2(6))
    Call showmethod3_2(showBox(3, 1, 3, 3), Shape2(7))
    Call showmethod3_2(showBox(3, 2, 3, 3), Shape2(8))
    Call showmethod3_2(showBox(3, 3, 3, 3), Shape2(9))
    

End Sub
Private Sub showmethod3_2(Color, Colorshape)
    If Color = 1 Then
        Colorshape.FillColor = Shape1(0).FillColor
    End If
    If Color = 2 Then
        Colorshape.FillColor = Shape2(0).FillColor
    End If
    If Color = 3 Then
        Colorshape.FillColor = Shape3(0).FillColor
    End If
    If Color = 4 Then
        Colorshape.FillColor = Shape4(0).FillColor
    End If
    If Color = 5 Then
        Colorshape.FillColor = Shape5(0).FillColor
    End If
    If Color = 6 Then
        Colorshape.FillColor = Shape6(0).FillColor
    End If
    
    
End Sub
Private Sub showmethod4(a, b, c, d) '轉動箭頭指示介面顯示

    If showterm = 0 Then showterm = 1
    Call showmethod4_1
    Call showmethod4_2(a, b, c, d)
    
    If showterm = 1 Then
        showterm = 2
    Else
        showterm = 1
    End If

End Sub
Private Sub showmethod4_1() '清空指示欄位
    Dim i As Integer
    
    For i = 1 To 3
        Text1(i).text = ""
        Text1(i + 3).text = ""
        Text3(i).text = ""
        Text3(i + 3).text = ""
        Call showmethod4_3_1(Text1(i))
        Call showmethod4_3_1(Text1(i + 3))
        Call showmethod4_3_1(Text3(i))
        Call showmethod4_3_1(Text3(i + 3))
    Next i
    
    For i = 1 To 4
        Text2(i).text = ""
        Text4(i).text = ""
        Text5(i).text = ""
        Text6(i).text = ""
        Text7(i).text = ""
        Text8(i).text = ""
        Call showmethod4_3_1(Text2(i))
        Call showmethod4_3_1(Text4(i))
        Call showmethod4_3_1(Text5(i))
        Call showmethod4_3_1(Text6(i))
        Call showmethod4_3_1(Text7(i))
        Call showmethod4_3_1(Text8(i))
        
    Next i
    
End Sub
Private Sub showmethod4_2(a, b, c, d) '選擇顯示正在被旋轉的軸
    If a <> 0 Then
        Call showmethod4_2_1(a, d)
    End If
    
    If b <> 0 Then
        Call showmethod4_2_2(b, d)
    End If
    
    If c <> 0 Then
        Call showmethod4_2_3(c, d)
    End If

    
End Sub
Private Sub showmethod4_2_1(a, d) 'x軸旋轉顯示介面
    If a = 1 Then
        If d = 1 Or d = 4 Then
            Text1(1).text = "--->"
            Text1(4).text = "--->"
        End If
        
        If d = 2 Or d = 3 Then
            Text1(1).text = "<---"
            Text1(4).text = "<---"
        End If
            
        Call showmethod4_3(Text1(1))
        Call showmethod4_3(Text1(4))
    End If
    
    If a = 2 Then
        If d = 1 Then
            Text1(2).text = "--->"
            Text1(5).text = "--->"
        End If
        If d = 2 Then
            Text1(2).text = "<---"
            Text1(5).text = "<---"
        End If
        
        Call showmethod4_3(Text1(2))
        Call showmethod4_3(Text1(5))
    End If
    
    If a = 3 Then
        If d = 1 Or d = 3 Then
            Text1(3).text = "--->"
            Text1(6).text = "--->"
        End If
        
        If d = 2 Or d = 4 Then
            Text1(3).text = "<---"
            Text1(6).text = "<---"
        End If
        
        Call showmethod4_3(Text1(3))
        Call showmethod4_3(Text1(6))
    End If
    
End Sub
Private Sub showmethod4_2_2(b, d) 'y軸旋轉顯示介面
     If b = 1 Then
        If d = 1 Or d = 4 Then
            Text2(1).text = "^"
            Text2(2).text = "|"
            Text2(3).text = "^"
            Text2(4).text = "|"
        End If
        
        If d = 2 Or d = 3 Then
            Text2(1).text = "|"
            Text2(2).text = "v"
            Text2(3).text = "|"
            Text2(4).text = "v"
        End If
        
        Call showmethod4_3(Text2(1))
        Call showmethod4_3(Text2(2))
        Call showmethod4_3(Text2(3))
        Call showmethod4_3(Text2(4))
    End If
    
    If b = 2 Then
        If d = 1 Then
            Text4(1).text = "^"
            Text4(2).text = "|"
            Text4(3).text = "^"
            Text4(4).text = "|"
           
        End If
        
        If d = 2 Then
            Text4(1).text = "|"
            Text4(2).text = "v"
            Text4(3).text = "|"
            Text4(4).text = "v"
        End If
        
        Call showmethod4_3(Text4(1))
        Call showmethod4_3(Text4(2))
        Call showmethod4_3(Text4(3))
        Call showmethod4_3(Text4(4))
    End If
    
    If b = 3 Then
        If d = 1 Or d = 3 Then
            Text5(1).text = "^"
            Text5(2).text = "|"
            Text5(3).text = "^"
            Text5(4).text = "|"
        End If
        
        If d = 2 Or d = 4 Then
            Text5(1).text = "|"
            Text5(2).text = "v"
            Text5(3).text = "|"
            Text5(4).text = "v"
        End If
        
        Call showmethod4_3(Text5(1))
        Call showmethod4_3(Text5(2))
        Call showmethod4_3(Text5(3))
        Call showmethod4_3(Text5(4))
    End If
            
End Sub
Private Sub showmethod4_2_3(c, d) 'z軸旋轉顯示介面
    If c = 1 Then
        If d = 1 Then
            Text3(1).text = "--->"
            Text3(4).text = "--->"
            Text6(1).text = "|"
            Text6(2).text = "^"
            Text6(3).text = "|"
            Text6(4).text = "v"
            
        End If
        
        If d = 2 Then
            Text3(1).text = "<---"
            Text3(4).text = "<---"
            Text6(1).text = "v"
            Text6(2).text = "|"
            Text6(3).text = "^"
            Text6(4).text = "|"
        End If
        
        Call showmethod4_3(Text3(1))
        Call showmethod4_3(Text3(4))
        Call showmethod4_3(Text6(1))
        Call showmethod4_3(Text6(2))
        Call showmethod4_3(Text6(3))
        Call showmethod4_3(Text6(4))
    End If
    
    If c = 2 Then
        If d = 1 Then
            Text3(2).text = "--->"
            Text3(5).text = "--->"
            Text7(1).text = "|"
            Text7(2).text = "^"
            Text7(3).text = "|"
            Text7(4).text = "v"
            
        End If
        
        If d = 2 Then
            Text3(2).text = "<---"
            Text3(5).text = "<---"
            Text7(1).text = "v"
            Text7(2).text = "|"
            Text7(3).text = "^"
            Text7(4).text = "|"
        End If
        
        Call showmethod4_3(Text3(2))
        Call showmethod4_3(Text3(5))
        Call showmethod4_3(Text7(1))
        Call showmethod4_3(Text7(2))
        Call showmethod4_3(Text7(3))
        Call showmethod4_3(Text7(4))
    End If
    
    If c = 3 Then
        If d = 1 Then
            Text3(3).text = "--->"
            Text3(6).text = "--->"
            Text8(1).text = "|"
            Text8(2).text = "^"
            Text8(3).text = "|"
            Text8(4).text = "v"
            
        End If
        
        If d = 2 Then
            Text3(3).text = "<---"
            Text3(6).text = "<---"
            Text8(1).text = "v"
            Text8(2).text = "|"
            Text8(3).text = "^"
            Text8(4).text = "|"
        End If
        
        Call showmethod4_3(Text3(3))
        Call showmethod4_3(Text3(6))
        Call showmethod4_3(Text8(1))
        Call showmethod4_3(Text8(2))
        Call showmethod4_3(Text8(3))
        Call showmethod4_3(Text8(4))
    End If
      
End Sub
Private Sub showmethod4_3(block) '旋轉指示箭頭顏色轉換
    If showterm = 1 Then
        block.BackColor = Shape7(1).FillColor
    End If
    
    If showterm = 2 Then
        block.BackColor = Shape7(2).FillColor
    End If
End Sub
Private Sub showmethod4_3_1(block)
    block.BackColor = Shape7(0).FillColor
End Sub

Private Sub text()
Dim i, j, k As Integer
For k = 1 To 3
    For i = 1 To 3
        For j = 1 To 3
            
                
                Call printcolor(i, j, k)
          
        Next j
    Next i
Next k

   

End Sub

Private Sub Command3_Click()
    Call showmethod1

End Sub

Private Sub Command4_Click()
    Call Openfile(3)
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Command5_Click()
'Call showmethod4_1
    Timer1.Enabled = False
End Sub



Private Sub Command6_Click()
    CommonDialog1.DialogTitle = "load an image"
    CommonDialog1.Filter = "*|*.*|mim|*.mim|tif|*.tif|"
    CommonDialog1.ShowOpen
    Buffer1.Load CommonDialog1.FileName, True
End Sub

Private Sub Command7_Click()
Dim i, j, k, n As Integer
Call colorSET




'先歸零
For i = 1 To 3
    For j = 1 To 3
        For k = 1 To 3
            For n = 1 To 3
            
                box2(i, j, k, n) = 0
                box(i, j, k, n) = 0
            
            Next n
        Next k
    Next j
Next i



'開啟文件輸入顏色到box

Call OpenColorfile(2, 3, 0, 4, 1) '前面

Call OpenColorfile(4, 1, 4, 4, 1) '後面

Call OpenColorfile(1, 0, 1, 4, 2) '左面

Call OpenColorfile(3, 4, 3, 4, 2) '右面

Call OpenColorfile(5, 0, 0, 3, 3) '上面

Call OpenColorfile(6, 4, 0, 1, 3) '下面


'-----------------------初始狀況
For k = 1 To 3
    For i = 1 To 3
        For j = 1 To 3
            
                
                Call printcolor(i, j, k)
          
        Next j
    Next i
Next k
End Sub

Private Sub Timer1_Timer()
    Call showmethod1
End Sub
=======
VERSION 5.00
Object = "{E1208DE3-A783-11D0-9161-00A024D24992}#1.0#0"; "MILApplication.ocx"
Object = "{6D9F7F71-9658-11D0-BDB5-00608CC9F9FB}#1.0#0"; "MILSystem.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{03985961-6B33-11D0-AB4A-00608CC9CA57}#1.0#0"; "MilBuffer.ocx"
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   Caption         =   "Form"
   ClientHeight    =   11715
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   18585
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   14.25
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11715
   ScaleWidth      =   18585
   StartUpPosition =   3  '系統預設值
   Begin MILAPPLICATIONLib.Application Application1 
      Height          =   480
      Left            =   14880
      TabIndex        =   43
      Top             =   8880
      Visible         =   0   'False
      Width           =   480
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      AutomaticAllocation=   -1  'True
   End
   Begin MILSYSTEMLib.System System1 
      Height          =   480
      Left            =   15600
      TabIndex        =   44
      Top             =   8880
      Visible         =   0   'False
      Width           =   480
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      AutomaticAllocation=   -1  'True
      SystemType      =   "VGA"
      ProcessingSystem=   1699376
      ProcessingSystemName=   "[Default]"
   End
   Begin MILBUFFERLib.Buffer Buffer1 
      Height          =   480
      Left            =   16320
      TabIndex        =   46
      Top             =   8880
      Visible         =   0   'False
      Width           =   480
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      AutomaticAllocation=   -1  'True
      OwnerSystem     =   "System1"
      SizeX           =   640
      SizeY           =   480
      NumberOfBands   =   3
      AbsoluteValue   =   252
      Saturation      =   252
      ChildRegionEndX =   639
      ChildRegionEndY =   479
      ChildRegionCenterX=   319
      ChildRegionCenterY=   239
      ChildRegionSizeX=   640
      ChildRegionSizeY=   480
      ChildRegionMode =   1
      CanDisplay      =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   16920
      Top             =   8880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Caption         =   "測試"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12360
      TabIndex        =   45
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "分析樣本所在位置"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9960
      TabIndex        =   42
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   3
      Left            =   8880
      TabIndex        =   41
      Top             =   1320
      Width           =   350
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   3
      Left            =   8160
      TabIndex        =   40
      Top             =   2040
      Width           =   350
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   3
      Left            =   7440
      TabIndex        =   39
      Top             =   2760
      Width           =   350
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   4
      Left            =   8880
      TabIndex        =   38
      Top             =   1800
      Width           =   350
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   4
      Left            =   8160
      TabIndex        =   37
      Top             =   2520
      Width           =   350
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   4
      Left            =   7440
      TabIndex        =   36
      Top             =   3240
      Width           =   350
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   4
      Left            =   8040
      TabIndex        =   35
      Top             =   1080
      Width           =   850
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   5
      Left            =   7320
      TabIndex        =   34
      Top             =   1800
      Width           =   850
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   6
      Left            =   6600
      TabIndex        =   33
      Top             =   2400
      Width           =   850
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   2
      Left            =   1440
      TabIndex        =   32
      Top             =   1440
      Width           =   350
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   2
      Left            =   2160
      TabIndex        =   31
      Top             =   2040
      Width           =   350
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   2
      Left            =   2880
      TabIndex        =   30
      Top             =   2640
      Width           =   350
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   1440
      TabIndex        =   29
      Top             =   1920
      Width           =   350
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   2160
      TabIndex        =   28
      Top             =   2520
      Width           =   350
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   2880
      TabIndex        =   27
      Top             =   3120
      Width           =   350
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   4
      Left            =   4440
      TabIndex        =   26
      Top             =   3600
      Width           =   350
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   4
      Left            =   5160
      TabIndex        =   25
      Top             =   3600
      Width           =   350
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   4
      Left            =   5880
      TabIndex        =   24
      Top             =   3600
      Width           =   350
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   6
      Left            =   6360
      TabIndex        =   23
      Top             =   5640
      Width           =   850
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   4
      Left            =   6360
      TabIndex        =   22
      Top             =   4200
      Width           =   850
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   5
      Left            =   6360
      TabIndex        =   21
      Top             =   4920
      Width           =   850
   End
   Begin VB.CommandButton Command5 
      Caption         =   "暫停"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   20
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "開始"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   19
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   17400
      Top             =   8160
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   3
      Left            =   5880
      TabIndex        =   18
      Top             =   3120
      Width           =   350
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   2
      Left            =   5880
      TabIndex        =   17
      Top             =   6720
      Width           =   350
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   5880
      TabIndex        =   16
      Top             =   6240
      Width           =   350
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   3
      Left            =   5160
      TabIndex        =   15
      Top             =   3120
      Width           =   350
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   2
      Left            =   5160
      TabIndex        =   14
      Top             =   6720
      Width           =   350
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   5160
      TabIndex        =   13
      Top             =   6240
      Width           =   350
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   3
      Left            =   3360
      TabIndex        =   12
      Top             =   2400
      Width           =   850
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   2
      Left            =   2520
      TabIndex        =   11
      Top             =   1800
      Width           =   850
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   1800
      TabIndex        =   10
      Top             =   1080
      Width           =   850
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   3
      Left            =   4440
      TabIndex        =   9
      Top             =   3120
      Width           =   350
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   2
      Left            =   4440
      TabIndex        =   8
      Top             =   6720
      Width           =   350
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   4440
      TabIndex        =   7
      Top             =   6240
      Width           =   350
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   2
      Left            =   3360
      TabIndex        =   6
      Top             =   4920
      Width           =   850
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   3360
      TabIndex        =   5
      Top             =   4200
      Width           =   850
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   3
      Left            =   3360
      TabIndex        =   4
      Top             =   5640
      Width           =   850
   End
   Begin VB.CommandButton Command4 
      Caption         =   "關閉檔案"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   3
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "下一步"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   2
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "記錄旋轉指令"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9960
      TabIndex        =   1
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "判斷開始"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9960
      TabIndex        =   0
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   2
      Left            =   240
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   1
      Left            =   240
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H80000000&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   0
      Left            =   240
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   1
      Left            =   2760
      Shape           =   1  '正方形
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   4
      Left            =   2040
      Shape           =   1  '正方形
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   7
      Left            =   1320
      Shape           =   1  '正方形
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   2
      Left            =   2760
      Shape           =   1  '正方形
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   5
      Left            =   2040
      Shape           =   1  '正方形
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   8
      Left            =   1320
      Shape           =   1  '正方形
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   3
      Left            =   2760
      Shape           =   1  '正方形
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   6
      Left            =   2040
      Shape           =   1  '正方形
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   9
      Left            =   1320
      Shape           =   1  '正方形
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   3
      Left            =   5760
      Shape           =   1  '正方形
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   2
      Left            =   5040
      Shape           =   1  '正方形
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   1
      Left            =   4320
      Shape           =   1  '正方形
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   6
      Left            =   5760
      Shape           =   1  '正方形
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   5
      Left            =   5040
      Shape           =   1  '正方形
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   4
      Left            =   4320
      Shape           =   1  '正方形
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   9
      Left            =   5760
      Shape           =   1  '正方形
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   8
      Left            =   5040
      Shape           =   1  '正方形
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   7
      Left            =   4320
      Shape           =   1  '正方形
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   0
      Left            =   240
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   1
      Left            =   4320
      Shape           =   1  '正方形
      Top             =   7320
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   2
      Left            =   5040
      Shape           =   1  '正方形
      Top             =   7320
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   3
      Left            =   5760
      Shape           =   1  '正方形
      Top             =   7320
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   4
      Left            =   4320
      Shape           =   1  '正方形
      Top             =   8040
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   5
      Left            =   5040
      Shape           =   1  '正方形
      Top             =   8040
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   6
      Left            =   5760
      Shape           =   1  '正方形
      Top             =   8040
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   7
      Left            =   4320
      Shape           =   1  '正方形
      Top             =   8760
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   8
      Left            =   5040
      Shape           =   1  '正方形
      Top             =   8760
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   9
      Left            =   5760
      Shape           =   1  '正方形
      Top             =   8760
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   0
      Left            =   240
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   1
      Left            =   4320
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   2
      Left            =   5040
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   3
      Left            =   5760
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   4
      Left            =   4320
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   5
      Left            =   5040
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   6
      Left            =   5760
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   7
      Left            =   4320
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   8
      Left            =   5040
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   9
      Left            =   5760
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   0
      Left            =   240
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   1
      Left            =   7440
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   2
      Left            =   7440
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   3
      Left            =   7440
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   4
      Left            =   8160
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   5
      Left            =   8160
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   6
      Left            =   8160
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   7
      Left            =   8880
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   8
      Left            =   8880
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   9
      Left            =   8880
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   0
      Left            =   240
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H000080FF&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   0
      Left            =   240
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '實心
      Height          =   495
      Index           =   0
      Left            =   240
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim box(1 To 3, 1 To 3, 1 To 3, 1 To 3) As Integer
Dim box2(1 To 3, 1 To 3, 1 To 3, 1 To 3) As Integer
Dim showBox(1 To 3, 1 To 3, 1 To 3, 1 To 3) As Integer
Dim Third0_1(1 To 3, 1 To 3) As Integer
Dim Third0_2(1 To 3, 1 To 3) As Integer
Dim Third1_1(1 To 3, 1 To 3) As Integer
Dim Third1_2(1 To 3, 1 To 3) As Integer
Dim Third1_3(1 To 3, 1 To 3) As Integer
Dim Third2_1(1 To 3, 1 To 3) As Integer
Dim Third2_2(1 To 3, 1 To 3) As Integer
Dim Third2_3(1 To 3, 1 To 3) As Integer
Dim Third2_4(1 To 3, 1 To 3) As Integer
Dim Third2_5(1 To 3, 1 To 3) As Integer
Dim Third2_6(1 To 3, 1 To 3) As Integer
Dim Third2_7(1 To 3, 1 To 3) As Integer


Dim R, L, G, U, E, W As Integer



Dim term As Integer '選擇判斷
Dim showterm As Integer '旋轉指示方塊顏色變換判斷

Private Sub Command1_Click() '讓程式判斷該怎麼轉
Dim i, j, k, n As Integer
Call colorSET

'先歸零
For i = 1 To 3
    For j = 1 To 3
        For k = 1 To 3
            For n = 1 To 3
            
                box2(i, j, k, n) = 0
                box(i, j, k, n) = 0
            
            Next n
        Next k
    Next j
Next i



'開啟文件輸入顏色到box

Call OpenColorfile(2, 3, 0, 4, 1) '前面

Call OpenColorfile(4, 1, 4, 4, 1) '後面

Call OpenColorfile(1, 0, 1, 4, 2) '左面

Call OpenColorfile(3, 4, 3, 4, 2) '右面

Call OpenColorfile(5, 0, 0, 3, 3) '上面

Call OpenColorfile(6, 4, 0, 1, 3) '下面



'-----------------------初始狀況
For k = 1 To 3
    For i = 1 To 3
        For j = 1 To 3
            
                
                Call printcolor(i, j, k)
          
        Next j
    Next i
Next k

Call Openfile(1)
Call showmethod2
'----------------------------------------------開始第一層判斷
    Debug.Print "第一層邊塊測試"
    Debug.Print "判斷121"
    Call firstEdgeBlock(1, 2, 1, E, 0, L)
    
    Debug.Print "判斷211"
    Call firstEdgeBlock(2, 1, 1, 0, U, L)
    
    Debug.Print "判斷231"
    Call firstEdgeBlock(2, 3, 1, 0, G, L)
    
    Debug.Print "判斷321"
    Call firstEdgeBlock(3, 2, 1, R, 0, L)
    
    Debug.Print "第一層角塊測試"
    Debug.Print "判斷111"
    Call firstCornerBlock(1, 1, 1, E, U, L)
    Call text
    Debug.Print "判斷131"
    Call firstCornerBlock(1, 3, 1, E, G, L)
    
    
    'Call text
    Debug.Print "判斷311"
    
    Call firstCornerBlock(3, 1, 1, R, U, L)
    
    Debug.Print "判斷331"
    Call firstCornerBlock(3, 3, 1, R, G, L)
    
'--------------------第一層復原以後的狀況
    
    
    For k = 1 To 3
    For i = 1 To 3
        For j = 1 To 3
            
                
                Call printcolor(i, j, k)
          
        Next j
    Next i
Next k

    


'----------------------------------------第二層判斷
    Debug.Print "第二層測試"
    Debug.Print "判斷112"
    Call secondmethod(1, 1, 2, E, U, 0)
 
    Debug.Print "判斷132"
    Call secondmethod(1, 3, 2, E, G, 0)
    
    Debug.Print "判斷312"
    Call secondmethod(3, 1, 2, R, U, 0)
    
    Debug.Print "判斷332"
    Call secondmethod(3, 3, 2, R, G, 0)
    
'--------------------第二層復原以後的狀況
For k = 1 To 3
    For i = 1 To 3
        For j = 1 To 3
            
                
                Call printcolor(i, j, k)
          
        Next j
    Next i
Next k

'-----------------------------------------第三層判斷
    Debug.Print "第三層測試"
    Call thirdmethod
    
 '--------------------第三層復原以後的狀況
For k = 1 To 3
    For i = 1 To 3
        For j = 1 To 3
            
                
                'Call printcolor(i, j, k)
          
        Next j
    Next i
Next k
   
Call Openfile(3)
'測試
'Call choose1(0, 0, 3) '旋轉的方法:第三層z軸方向旋轉




End Sub
Private Sub printcolor(i, j, k) '查看方塊顏色
Dim a As Integer

For a = 1 To 3
    If box(i, j, k, a) = 0 Then
    Debug.Print i, j, k, a, "0"
    End If
    
    If box(i, j, k, a) = 1 Then
    Debug.Print i, j, k, a, "紅"
    End If
    
    If box(i, j, k, a) = 2 Then
    Debug.Print i, j, k, a, "黑"
    End If
    
    If box(i, j, k, a) = 3 Then
    Debug.Print i, j, k, a, "綠"
    End If
    
    If box(i, j, k, a) = 4 Then
    Debug.Print i, j, k, a, "藍"
    End If
    
    If box(i, j, k, a) = 5 Then
    Debug.Print i, j, k, a, "黃"
    End If
    
    If box(i, j, k, a) = 6 Then
    Debug.Print i, j, k, a, "橘"
    End If
    
Next a
Debug.Print ""




End Sub
Private Sub OpenColorfile(NUMBER As Integer, a As Integer, b As Integer, c As Integer, d As Integer)
    '判斷開啟的文件
    If NUMBER = 1 Then
        Open "blue.txt" For Input As #1
    ElseIf NUMBER = 2 Then
        Open "red.txt" For Input As #1
    ElseIf NUMBER = 3 Then
        Open "green.txt" For Input As #1
    ElseIf NUMBER = 4 Then
        Open "orange.txt" For Input As #1
    ElseIf NUMBER = 5 Then
        Open "yellow.txt" For Input As #1
    ElseIf NUMBER = 6 Then
        Open "black.txt" For Input As #1
    Else
    End If
    
    Call ColorInput(NUMBER, a, b, c, d)
    
    Close #1
    
End Sub
Private Sub ColorInput(NUMBER As Integer, a, b, c, d)
Dim i As Integer
Dim j As Integer
Dim Color As Integer

    If NUMBER = 1 Then '方塊左面
        For i = 1 To 3
            For j = 1 To 3
                Input #1, Color '讀取一條紀錄
                box(a + i, b, c - j, d) = Color
            Next j
        Next i
        
    ElseIf NUMBER = 3 Then '方塊右面
        For i = 1 To 3
            For j = 1 To 3
                Input #1, Color '讀取一條紀錄
                box(a - i, b, c - j, d) = Color
            Next j
        Next i
    
    ElseIf NUMBER = 2 Then '方塊前面
        For i = 1 To 3
            For j = 1 To 3
                Input #1, Color '讀取一條紀錄
                box(a, b + i, c - j, d) = Color
            Next j
        Next i
        
    
      
    ElseIf NUMBER = 4 Then '方塊後面
        For i = 1 To 3
            For j = 1 To 3
                Input #1, Color '讀取一條紀錄
                box(a, b - i, c - j, d) = Color
            Next j
        Next i
        
    ElseIf NUMBER = 5 Then '方塊上面
        For i = 1 To 3
            For j = 1 To 3
                Input #1, Color '讀取一條紀錄
                box(a + j, b + i, c, d) = Color
            Next j
        Next i
       
    ElseIf NUMBER = 6 Then '方塊下面
        For i = 1 To 3
            For j = 1 To 3
                Input #1, Color '讀取一條紀錄
                box(a - j, b + i, c, d) = Color
            Next j
        Next i
        
    Else
    End If
    
End Sub
Private Sub Openfile(X As Integer) '開啟讀寫文件
    If X = 1 Then
        Open "123.txt" For Output As #1
    End If
    
    If X = 2 Then
        Open "123.txt" For Input As #1
    End If
    
    If X = 3 Then
        Close #1
    End If
End Sub
Private Sub colorSET() '設定顏色參數
R = 1 '紅色
L = 2 '黑色
G = 3 '綠色
U = 4 '藍色
W = 5 '黃色
E = 6 '橘色
End Sub
Private Sub choose(fox() As Integer, a, b, c) '建立選擇旋轉的方向和層數的方法
    If b = 0 And c = 0 Then
        Call X(a, fox) 'x軸旋轉
    End If

    If a = 0 And c = 0 Then
        Call Y(b, fox) 'y軸旋轉
    End If
    
    If a = 0 And b = 0 Then
        Call z(c, fox) 'z軸旋轉
    End If
End Sub
Private Sub choose1(fox() As Integer, a, b, c, show) '建立旋轉正方向的選擇
    Dim d As Integer
    d = 1
    
    Call choose(fox, a, b, c)
    If a <> 0 Then
        Debug.Print "X軸", a, "正轉"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    ElseIf b <> 0 Then
        Debug.Print "Y軸", b, "正轉"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    Else
        Debug.Print "Z軸", c, "正轉"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    End If
End Sub
Private Sub choose2(fox() As Integer, a, b, c, show) '建立旋轉反方向的選擇
    Dim d As Integer
    d = 2
    
    Call choose(fox, a, b, c)
    Call choose(fox, a, b, c)
    Call choose(fox, a, b, c)
    
    If a <> 0 Then
        Debug.Print "X軸", a, "反轉"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    ElseIf b <> 0 Then
        Debug.Print "Y軸", b, "反轉"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    Else
        Debug.Print "Z軸", c, "反轉"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    End If
End Sub
Private Sub X(a, fox() As Integer) '沿著x軸旋轉的方法
Dim i As Integer
Dim j As Integer

For i = 1 To 3
    For j = 1 To 3
        Call rotate(i2, j2, i, j)
        box2(a, i2, j2, 1) = fox(a, i, j, 1)
        box2(a, i2, j2, 2) = fox(a, i, j, 3)
        box2(a, i2, j2, 3) = fox(a, i, j, 2)
                
    Next j
Next i

'在扔回box1存取改變
For i = 1 To 3
    For j = 1 To 3
        fox(a, i, j, 1) = box2(a, i, j, 1)
        fox(a, i, j, 2) = box2(a, i, j, 2)
        fox(a, i, j, 3) = box2(a, i, j, 3)
                
    Next j
Next i
End Sub
Private Sub Y(b, fox() As Integer) '沿著y軸旋轉的方法
Dim i As Integer
Dim j As Integer

For i = 1 To 3
    For j = 1 To 3
        Call rotate(i2, j2, i, j)
        box2(j2, b, i2, 1) = fox(j, b, i, 3)
        box2(j2, b, i2, 2) = fox(j, b, i, 2)
        box2(j2, b, i2, 3) = fox(j, b, i, 1)
    Next j
Next i

'在扔回box1存取改變
For i = 1 To 3
    For j = 1 To 3
        fox(j, b, i, 1) = box2(j, b, i, 1)
        fox(j, b, i, 2) = box2(j, b, i, 2)
        fox(j, b, i, 3) = box2(j, b, i, 3)
                
    Next j
Next i
End Sub
Private Sub z(c, fox() As Integer) '沿著z軸旋轉的方法
Dim i As Integer
Dim j As Integer

'先用box2儲存改變
For i = 1 To 3
    For j = 1 To 3
        Call rotate(i2, j2, i, j) '前兩個是後來的座標
                                  '後兩個是原本的座標
        box2(i2, j2, c, 1) = fox(i, j, c, 2)
        box2(i2, j2, c, 2) = fox(i, j, c, 1)
        box2(i2, j2, c, 3) = fox(i, j, c, 3)
                
    Next j
Next i

For i = 1 To 3
    For j = 1 To 3
        fox(i, j, c, 1) = box2(i, j, c, 1)
        fox(i, j, c, 2) = box2(i, j, c, 2)
        fox(i, j, c, 3) = box2(i, j, c, 3)
                
    Next j
Next i
''

'測試
'For i = 1 To 3
'    For j = 1 To 3
'
'        Debug.Print i, j, "3  ", box2(i, j, 3, 1)
'        Debug.Print i, j, "3  ", box2(i, j, 3, 2)
'        Debug.Print i, j, "3  ", box2(i, j, 3, 3)
'
'    Next j
'Next i
End Sub
Private Sub rotate(a2, b2, a1, b1) '建立旋轉座標轉換的公式
For i = 1 To 3
    For j = 1 To 3
        a2 = b1
        b2 = 4 - a1
    Next j
Next i

End Sub

Private Sub firstEdgeBlock(X, Y, z, cx, cy, cz) '偵測第一層邊塊
    term = 0 '條件判斷變數先初始化
Debug.Print "判斷邊塊是否在正確位置"
    Call firstmethod(X, Y, z, cx, cy, cz)
    
    If term <> 1 And term <> 5 Then
Debug.Print "判斷邊塊是否在第二層"
        Call firstmethod3(X, Y, z, 1, 1, 2, cx, cy, cz)
        Call firstmethod3(X, Y, z, 1, 3, 2, cx, cy, cz)
        Call firstmethod3(X, Y, z, 3, 1, 2, cx, cy, cz)
        Call firstmethod3(X, Y, z, 3, 3, 2, cx, cy, cz)
    
    End If
    
    If term <> 1 And term <> 3 Then
Debug.Print "Z軸正確 判斷是否有在第一層"
        Call firstmethod1(X, Y, z, 1, 2, 1, cx, cy, cz)
        Call firstmethod1(X, Y, z, 2, 1, 1, cx, cy, cz)
        Call firstmethod1(X, Y, z, 2, 3, 1, cx, cy, cz)
        Call firstmethod1(X, Y, z, 3, 2, 1, cx, cy, cz)
Debug.Print "Z軸正確 判斷是否有在第三層"
        Call firstmethod1(X, Y, z, 1, 2, 3, cx, cy, cz)
        Call firstmethod1(X, Y, z, 2, 1, 3, cx, cy, cz)
        Call firstmethod1(X, Y, z, 2, 3, 3, cx, cy, cz)
        Call firstmethod1(X, Y, z, 3, 2, 3, cx, cy, cz)
    End If
    
    
    If term <> 1 And term <> 3 Then
Debug.Print "Z軸顛倒 判斷是否有在第一層"
        Call firstmethod2(X, Y, z, 1, 2, 1, cx, cy, cz)
        Call firstmethod2(X, Y, z, 2, 1, 1, cx, cy, cz)
        Call firstmethod2(X, Y, z, 2, 3, 1, cx, cy, cz)
        Call firstmethod2(X, Y, z, 3, 2, 1, cx, cy, cz)
Debug.Print "Z軸顛倒 判斷是否有在第三層"
        Call firstmethod2(X, Y, z, 1, 2, 3, cx, cy, cz)
        Call firstmethod2(X, Y, z, 2, 1, 3, cx, cy, cz)
        Call firstmethod2(X, Y, z, 2, 3, 3, cx, cy, cz)
        Call firstmethod2(X, Y, z, 3, 2, 3, cx, cy, cz)
    End If
  
End Sub

Private Sub firstCornerBlock(X, Y, z, cx, cy, cz) '偵測第一層角塊
    term = 0 '條件初始化
Debug.Print "角塊是否在正確位置"
    Call firstmethod4(X, Y, z, cx, cy, cz)
    
    If term <> 1 And term <> 3 Then
Debug.Print "角塊z軸是否黑色並進行處理"
        Call firstmethod5(X, Y, z, 1, 1, 1, cx, cy, cz)
        Call firstmethod5(X, Y, z, 1, 3, 1, cx, cy, cz)
        Call firstmethod5(X, Y, z, 3, 1, 1, cx, cy, cz)
        Call firstmethod5(X, Y, z, 3, 3, 1, cx, cy, cz)
        
        Call firstmethod5(X, Y, z, 1, 1, 3, cx, cy, cz)
        Call firstmethod5(X, Y, z, 1, 3, 3, cx, cy, cz)
        Call firstmethod5(X, Y, z, 3, 1, 3, cx, cy, cz)
        Call firstmethod5(X, Y, z, 3, 3, 3, cx, cy, cz)
        
    End If
    
    If term <> 1 And term <> 4 Then
Debug.Print "角塊是否z軸不為黑色並進行處理"
        Call firstmethod6(X, Y, z, 1, 1, 1, cx, cy, cz)
        Call firstmethod6(X, Y, z, 1, 3, 1, cx, cy, cz)
        Call firstmethod6(X, Y, z, 3, 1, 1, cx, cy, cz)
        Call firstmethod6(X, Y, z, 3, 3, 1, cx, cy, cz)
        
        Call firstmethod6(X, Y, z, 1, 1, 3, cx, cy, cz)
        Call firstmethod6(X, Y, z, 1, 3, 3, cx, cy, cz)
        Call firstmethod6(X, Y, z, 3, 1, 3, cx, cy, cz)
        Call firstmethod6(X, Y, z, 3, 3, 3, cx, cy, cz)
    End If

End Sub
Private Sub firstmethod(X, Y, z, cx, cy, cz)
    '判斷是否在原位
    Debug.Print "判斷是否在原位"
    
    If box(X, Y, z, 1) = cx And box(X, Y, z, 2) = cy And box(X, Y, z, 3) = cz Then
        term = 1
        Debug.Print "有在原位"
    End If
End Sub
Private Sub firstmethod1(x1, y1, z1, x2, y2, z2, cx, cy, cz)
    '判斷在第一層或第三層
    '判斷條件z軸方向為必為黑色
    If term <> 3 Then
        If box(x1, y1, z1, 1) <> 0 Then
        
            If cx = box(x2, y2, z2, 1) And cz = box(x2, y2, z2, 3) Then
                term = 2
            End If
    
            If cx = box(x2, y2, z2, 2) And cz = box(x2, y2, z2, 3) Then
                term = 2
            End If
        End If
        
        If box(x1, y1, z1, 2) <> 0 Then
        
            If cy = box(x2, y2, z2, 1) And cz = box(x2, y2, z2, 3) Then
                term = 2
            End If
    
            If cy = box(x2, y2, z2, 2) And cz = box(x2, y2, z2, 3) Then
                term = 2
            End If
        End If
    End If
    
    If term = 2 Then
        Call firstmethod1_1(x1, y1, z1, x2, y2, z2)
    End If

End Sub
Private Sub firstmethod1_1(x1, y1, z1, x2, y2, z2)
Debug.Print "進入旋轉處理firstmethod1_1"
        term = 3
        '轉到第三層處理
        If x2 <> 2 And z2 <> 3 Then
            Call secondchoose1(box, x2, 0, 0, 1)
            Call secondchoose1(box, x2, 0, 0, 1)
        End If

        If y2 <> 2 And z2 <> 3 Then
            Call secondchoose1(box, 0, y2, 0, 1)
            Call secondchoose1(box, 0, y2, 0, 1)
        End If
            
        '在第三層旋轉
        Call firstmethod2_2(x1, x2, y1, y2)
            
        '轉回第一層
        If x1 <> 2 Then
            Call secondchoose1(box, x1, 0, 0, 1)
            Call secondchoose1(box, x1, 0, 0, 1)
        End If
        
        If y1 <> 2 Then
            Call secondchoose1(box, 0, y1, 0, 1)
            Call secondchoose1(box, 0, y1, 0, 1)
        End If
End Sub
Private Sub firstmethod2(x1, y1, z1, x2, y2, z2, cx, cy, cz)
    '判斷在第一層或第三層
    '但是z軸顏色顛倒
    If term <> 3 Then
        If box(x1, y1, z1, 1) <> 0 Then
        
            If cx = box(x2, y2, z2, 3) And cz = box(x2, y2, z2, 1) Then
                term = 2
            End If
            
            If cx = box(x2, y2, z2, 3) And cz = box(x2, y2, z2, 2) Then
                term = 2
            End If
        End If
        
        If box(x1, y1, z1, 2) <> 0 Then
        
            If cy = box(x2, y2, z2, 3) And cz = box(x2, y2, z2, 1) Then
                term = 2
            End If
 
            If cy = box(x2, y2, z2, 3) And cz = box(x2, y2, z2, 2) Then
                term = 2
            End If
        End If
        
    End If
    
    If term = 2 Then
        Call firstmethod2_1(x1, y1, z1, x2, y2, z2)
    End If
    
End Sub
Private Sub firstmethod2_1(x1, y1, z1, x2, y2, z2)
Debug.Print "進入旋轉處理firstmethod2_1"
        term = 3
        '轉到第三層處理抓取位置座標
        If x2 <> 2 And z2 <> 3 Then
            Call secondchoose1(box, x2, 0, 0, 1)
            Call secondchoose1(box, x2, 0, 0, 1)
        End If

        If y2 <> 2 And z2 <> 3 Then
            Call secondchoose1(box, 0, y2, 0, 1)
            Call secondchoose1(box, 0, y2, 0, 1)
        End If
            
        '在第三層旋轉
        Call firstmethod2_2(x1, x2, y1, y2)
            
        '轉回第一層正確位置座標
        If x1 <> 2 Then
            
            Call secondchoose2(box, x1, 0, 0, 1)
            Call choose2(box, 0, 0, 2, 1)
            Call secondchoose1(box, x1, 0, 0, 1)
            Call choose1(box, 0, 0, 2, 1)
        End If
        
        If y1 <> 2 Then
            Call secondchoose2(box, 0, y1, 0, 1)
            Call choose2(box, 0, 0, 2, 1)
            Call secondchoose1(box, 0, y1, 0, 1)
            Call choose1(box, 0, 0, 2, 1)
        End If
        
End Sub
Private Sub firstmethod2_2(x1, x2, y1, y2) '副程式----------由firstmethod2_1-呼叫
'目的將在第三層的邊塊轉到第一層的位置上的xy軸   以利於置入第二層中
'1為正確位置 2為抓取位置
    If x1 = x2 And y1 = y2 Then
        
    ElseIf x1 = x2 Or y1 = y2 Then
        Call choose1(box, 0, 0, 3, 1)
        Call choose1(box, 0, 0, 3, 1)
    ElseIf x1 <> x2 And y1 <> y2 Then '
        If x2 = 1 Then
            If y1 < y2 Then
                Call choose2(box, 0, 0, 3, 1)
            Else
                Call choose1(box, 0, 0, 3, 1)
            End If
        End If
            
        If x2 = 3 Then
            If y1 < y2 Then
                Call choose1(box, 0, 0, 3, 1)
            Else
                Call choose2(box, 0, 0, 3, 1)
            End If
        End If
            
        If y2 = 1 Then
            If x1 < x2 Then
                Call choose1(box, 0, 0, 3, 1)
            Else
                Call choose2(box, 0, 0, 3, 1)
            End If
        End If
        
        If y2 = 3 Then
            If x1 < x2 Then
                Call choose2(box, 0, 0, 3, 1)
            Else
                Call choose1(box, 0, 0, 3, 1)
            End If
        End If
    Else
        Debug.Print "firstmethod2_2出錯"
            
        
    End If
End Sub
Private Sub firstmethod3(x1, y1, z1, x2, y2, z2, cx, cy, cz)
    '判斷邊塊在第二層
    '座標1為正確位置 座標2為抓取位置
    If (box(x2, y2, z2, 1) = cx And box(x2, y2, z2, 2) = cz) Or (box(x2, y2, z2, 1) = cz And box(x2, y2, z2, 2) = cx) Then
        term = 2
    End If
    
    If (box(x2, y2, z2, 1) = cy And box(x2, y2, z2, 2) = cz) Or (box(x2, y2, z2, 1) = cz And box(x2, y2, z2, 2) = cy) Then
        term = 2
    End If
        
    If term = 2 Then
        Call firstmethod3_1(x1, y1, z1, x2, y2, z2, cx, cy, cz)
    End If

End Sub
Private Sub firstmethod3_1(x1, y1, z1, x2, y2, z2, cx, cy, cz)
'將在第二層的邊塊轉到第三層
Debug.Print "進入旋轉處理firstmethod3_1"
    term = 5
    If x2 = 1 Then
    
        If y2 = 1 Then
            Call secondchoose2(box, 1, 0, 0, 1)
            Call choose2(box, 0, 0, 3, 1)
            Call secondchoose1(box, 1, 0, 0, 1)
        ElseIf y2 = 3 Then
            Call secondchoose1(box, 1, 0, 0, 1)
            Call choose2(box, 0, 0, 3, 1)
            Call secondchoose2(box, 1, 0, 0, 1)
        Else
Debug.Print "firstmethod3-1 X2=1 Y2錯誤"
        End If
        
    ElseIf x2 = 3 Then
    
        If y2 = 1 Then
            Call secondchoose1(box, 3, 0, 0, 1)
            Call choose2(box, 0, 0, 3, 1)
            Call secondchoose2(box, 3, 0, 0, 1)
        ElseIf y2 = 3 Then
            Call secondchoose2(box, 3, 0, 0, 1)
            Call choose2(box, 0, 0, 3, 1)
            Call secondchoose1(box, 3, 0, 0, 1)
        Else
Debug.Print "firstmethod3-1 X2=1 Y2 錯誤"
        End If
        
    Else
Debug.Print "firstmethod3-1 X2錯誤"
    End If
    
End Sub
Private Sub firstmethod4(x1, y1, z1, cx, cy, cz) '角塊
    '是否在正確的位置上
    If box(x1, y1, z1, 1) = cx And box(x1, y1, z1, 2) = cy And box(x1, y1, z1, 3) = cz Then
        term = 1
    End If
End Sub
Private Sub firstmethod5(x1, y1, z1, x2, y2, z2, cx, cy, cz) '角塊
'判斷條件Z軸為黑色
    If box(x2, y2, z2, 1) = cx And box(x2, y2, z2, 2) = cy And box(x2, y2, z2, 3) = cz Then
        term = 2
    End If
    If box(x2, y2, z2, 1) = cy And box(x2, y2, z2, 2) = cx And box(x2, y2, z2, 3) = cz Then
        term = 2
    End If
    
    If term = 2 Then
        Call firstmethod5_1(x1, y1, z1, x2, y2, z2, cx, cy, cz)
    End If
    
End Sub
Private Sub firstmethod5_1(x1, y1, z1, x2, y2, z2, cx, cy, cz) '角塊
'條件Z軸為黑色
'處理:邊塊z軸不為黑色
Debug.Print "進入旋轉處理firstmethod5_1"
    term = 3
    If z2 = 1 Then
        If x2 = y2 Then
            Call secondchoose1(box, 0, y2, 0, 1)
            Call choose1(box, 0, 0, 3, 1)
            Call secondchoose2(box, 0, y2, 0, 1)
        
        ElseIf x2 <> y2 Then
            Call secondchoose2(box, 0, y2, 0, 1)
            Call choose2(box, 0, 0, 3, 1)
            Call secondchoose1(box, 0, y2, 0, 1)
            
        Else
Debug.Print "firstmethod5_1 z2=1錯誤"

        End If
    End If
    
    If z2 = 3 Then
        If x2 = 1 And y2 = 1 Then
            Call choose1(box, 0, 0, 3, 1)
            Call choose1(box, 0, 0, 3, 1)
            
        ElseIf x2 = 1 And y2 = 3 Then
            Call choose1(box, 0, 0, 3, 1)
        
        ElseIf x2 = 2 And y2 = 1 Then
            Call choose2(box, 0, 0, 3, 1)
        Else
        
        End If
        
        Call secondchoose1(box, 0, 3, 0, 1)
        Call choose2(box, 0, 0, 3, 1)
        Call secondchoose2(box, 0, 3, 0, 1)
            
        Else
Debug.Print "firstmethod5_1 z2=3錯誤"

        
        
    End If
    

End Sub
Private Sub firstmethod6(x1, y1, z1, x2, y2, z2, cx, cy, cz) '角塊
'判斷條件Z軸不為黑色
    If box(x2, y2, z2, 1) = cz And box(x2, y2, z2, 2) = cy And box(x2, y2, z2, 3) = cx Then
        term = 2
    End If
    If box(x2, y2, z2, 1) = cy And box(x2, y2, z2, 2) = cz And box(x2, y2, z2, 3) = cx Then
        term = 2
    End If
    If box(x2, y2, z2, 1) = cx And box(x2, y2, z2, 2) = cz And box(x2, y2, z2, 3) = cy Then
        term = 2
    End If
    If box(x2, y2, z2, 1) = cz And box(x2, y2, z2, 2) = cx And box(x2, y2, z2, 3) = cy Then
        term = 2
    End If
    
    If term = 2 Then
        Call FIRSTMETHOD6_1(x1, y1, z1, x2, y2, z2, cx, cy, cz)
    End If
End Sub
Private Sub FIRSTMETHOD6_1(x1, y1, z1, x2, y2, z2, cx, cy, cz) '角塊
'如果在第一層先轉到第三層
Debug.Print "進入旋轉處理firstmethod6_1"
    term = 4
    If z2 = 1 Then
Debug.Print "如果在第一層先轉到第三層"
        If x2 = y2 Then
            If box(x2, y2, z2, 1) = cz Then
                Call secondchoose2(box, x2, 0, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose1(box, x2, 0, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
            
            
            ElseIf box(x2, y2, z2, 2) = cz Then
                Call secondchoose1(box, 0, y2, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose2(box, 0, y2, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
            End If
            
        Else
            If box(x2, y2, z2, 1) = cz Then
                Call secondchoose1(box, x2, 0, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose2(box, x2, 0, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
            
            
            ElseIf box(x2, y2, z2, 2) = cz Then
                Call secondchoose2(box, 0, y2, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose1(box, 0, y2, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
            End If
        End If
            
    End If
    
'在第三層旋轉到適當位置
Debug.Print "在第三層旋轉到適當位置"
    If x1 <> x2 And y1 <> y2 Then
        Call choose1(box, 0, 0, 3, 1)
        Call choose1(box, 0, 0, 3, 1)
    ElseIf x1 = x2 And y1 = y2 Then
    
    ElseIf x1 <> x2 Or y1 <> y2 Then
        If x1 = x2 And x1 = 1 Then
        
            If y1 < y2 Then
                Call choose2(box, 0, 0, 3, 1)
            Else
                Call choose1(box, 0, 0, 3, 1)
            End If
        ElseIf x1 = x2 And x1 = 3 Then
        
            If y1 < y2 Then
                Call choose1(box, 0, 0, 3, 1)
            Else
                Call choose2(box, 0, 0, 3, 1)
            End If
        ElseIf y1 = y2 And y1 = 1 Then
        
            If x1 < x2 Then
                Call choose1(box, 0, 0, 3, 1)
            Else
                Call choose2(box, 0, 0, 3, 1)
            End If
        ElseIf y1 = y2 And y1 = 3 Then
        
            If x1 < x2 Then
                Call choose2(box, 0, 0, 3, 1)
            Else
                Call choose1(box, 0, 0, 3, 1)
            End If
        End If
    Else
Debug.Print "firstmethod6_1 cx cy錯誤"
    End If
'將邊塊轉到正確的位置
    If x1 = y1 Then
        If box(x1, y1, 3, 2) = cz Then
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, x1, 0, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, x1, 0, 0, 1)
            
            Else
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, 0, y1, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, 0, y1, 0, 1)
            End If
    Else
        
        If box(x1, y1, 3, 2) = cz Then
            Call choose1(box, 0, 0, 3, 1)
            Call secondchoose1(box, x1, 0, 0, 1)
            Call choose2(box, 0, 0, 3, 1)
            Call secondchoose2(box, x1, 0, 0, 1)
            
            Else
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, 0, y1, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, 0, y1, 0, 1)
            End If
    End If
    
            
            
        
End Sub


Private Sub secondmethod(X, Y, z, cx, cy, cz) '副程式----------由主程式呼叫
'第二層單一邊塊位置判斷

    term = 0 '條件判斷變數先初始化
    
    '判斷是否在原位
    Debug.Print "判斷是否在原位"
    
    If box(X, Y, z, 1) = cx And box(X, Y, z, 2) = cy Then
        term = 1
        Debug.Print "有在原位"
    End If
    
    
    '判斷是否在第二層的某個位置
    Debug.Print "判斷是否在第二層"
    
    If term <> 1 Then
        Debug.Print "判斷在第二層的哪裡"
        Call secondmethod1_1(1, 1, 2, cx, cy, cz)
            Debug.Print "透過secondmethod1-1判斷112"
        Call secondmethod1_1(1, 3, 2, cx, cy, cz)
            Debug.Print "透過secondmethod1-1判斷132"
        Call secondmethod1_1(3, 1, 2, cx, cy, cz)
            Debug.Print "透過secondmethod1-1判斷312"
        Call secondmethod1_1(3, 3, 2, cx, cy, cz)
            Debug.Print "透過secondmethod1-1判斷332"
        
    End If
    
    '判斷是否在第三層的某個位置
    Debug.Print "判斷是否在第三層"
    
    If term <> 1 Then
        Debug.Print "判斷在第三層的哪裡"
            Debug.Print "透過secondmethod1-2判斷123"
        Call secondmethod1_2(1, 2, 3, cx, cy, cz)
            Debug.Print "透過secondmethod1-2判斷213"
        Call secondmethod1_2(2, 1, 3, cx, cy, cz)
            Debug.Print "透過secondmethod1-2判斷233"
        Call secondmethod1_2(2, 3, 3, cx, cy, cz)
            Debug.Print "透過secondmethod1-2判斷323"
        Call secondmethod1_2(3, 2, 3, cx, cy, cz)
            
        '置入第二層
            Debug.Print "透過secondmethod1-3將邊塊置入第二層"
        Call secondmethod1_3(X, Y, z)
             
        
    
    End If




End Sub
Private Sub secondmethod1_1(X, Y, z, cx, cy, cz) '副程式---------由secondmethod系列-呼叫
    '判斷是否在第二層的某個位置上
    Debug.Print "副程式判斷邊塊在第二層的哪裡"
    
    Dim termsecond1_1 As Integer
    If (box(X, Y, z, 1) = cx And box(X, Y, z, 2) = cy) Or (box(X, Y, z, 1) = cy And box(X, Y, z, 2) = cx) Then
        term = 2
        termsecond1_1 = 2
    End If
    
    '如果抓到位置再對其進行處理
    '處理方式為將邊塊移至第三層
    If termsecond1_1 = 2 Then
        Debug.Print "第二層移至第三層的旋轉步驟"
        If X = Y Then
            Call secondchoose1(box, 0, Y, 0, 1)
            Call choose2(box, 0, 0, 3, 1)
            Call secondchoose2(box, 0, Y, 0, 1)
            Call choose2(box, 0, 0, 3, 1)
            Call secondchoose2(box, X, 0, 0, 1)
            Call choose1(box, 0, 0, 3, 1)
            Call secondchoose1(box, X, 0, 0, 1)
        
        Else
            Call secondchoose1(box, X, 0, 0, 1)
            Call choose2(box, 0, 0, 3, 1)
            Call secondchoose2(box, X, 0, 0, 1)
            Call choose2(box, 0, 0, 3, 1)
            Call secondchoose2(box, 0, Y, 0, 1)
            Call choose1(box, 0, 0, 3, 1)
            Call secondchoose1(box, 0, Y, 0, 1)
  
        End If
    End If
    
    
End Sub
Private Sub secondmethod1_2(X, Y, z, cx, cy, cz) '副程式----------由secondmethod系列-呼叫
    '判斷是否在第三層的某個位置上
    Debug.Print "副程式判斷邊塊在第三層的哪裡"
    
    Dim termsecond1_2 As Integer
    If (box(X, Y, z, 1) = cx And box(X, Y, z, 3) = cy) Or (box(X, Y, z, 1) = cy And box(X, Y, z, 3) = cx) Then
        termsecond1_2 = 2
    End If
    
    If (box(X, Y, z, 2) = cx And box(X, Y, z, 3) = cy) Or (box(X, Y, z, 2) = cy And box(X, Y, z, 3) = cx) Then
        termsecond1_2 = 2
    End If
    
    
    Debug.Print "判斷變數termsecond1_2", termsecond1_2
    
    '如果抓到位置(在第三層)再對其進行處理
    If termsecond1_2 = 2 Then
        Debug.Print "進入第三層的適當位置旋轉步驟"
        
        '先讓在第三層的邊塊旋轉到第三層的適當位置
        If box(X, Y, z, 3) = box(1, 2, 2, 1) Then
            Call secondmethod1_2_1(X, 1, Y, 2)
            term = 3 '先轉x軸
            
        ElseIf box(X, Y, z, 3) = box(2, 1, 2, 2) Then
            Call secondmethod1_2_1(X, 2, Y, 1)
            term = 4 '先轉y軸
            
        ElseIf box(X, Y, z, 3) = box(2, 3, 2, 2) Then
            Call secondmethod1_2_1(X, 2, Y, 3)
            term = 4 '先轉y軸
            
        ElseIf box(X, Y, z, 3) = box(3, 2, 2, 1) Then
            Call secondmethod1_2_1(X, 3, Y, 2)
            term = 3 '先轉x軸
            
        Else
            Debug.Print "secondmethod1_2出錯"
        End If
       
    End If
End Sub
Private Sub secondmethod1_2_1(x1, x2, y1, y2) '副程式----------由secondmethod1_2-呼叫
'目的將在第三層的邊塊轉到正確的位置上
'以利於置入第二層中
    If x1 = x2 And y1 = y2 Then
        Call choose1(box, 0, 0, 3, 1)
        Call choose1(box, 0, 0, 3, 1)
    ElseIf x1 = x2 Or y1 = y2 Then
        
    ElseIf x1 <> x2 And y1 <> y2 Then
        If x1 = 1 Then
            If y1 < y2 Then
                Call choose2(box, 0, 0, 3, 1)
            Else
                Call choose1(box, 0, 0, 3, 1)
            End If
        End If
        
        If x1 = 3 Then
            If y1 < y2 Then
                Call choose1(box, 0, 0, 3, 1)
            Else
                Call choose2(box, 0, 0, 3, 1)
            End If
        End If
        
        If y1 = 1 Then
            If x1 < x2 Then
                Call choose1(box, 0, 0, 3, 1)
            Else
                Call choose2(box, 0, 0, 3, 1)
            End If
        End If
        
        If y1 = 3 Then
            If x1 < x2 Then
                Call choose2(box, 0, 0, 3, 1)
            Else
                Call choose1(box, 0, 0, 3, 1)
            End If
        End If
        
    Else
        Debug.Print "secondmethod1_2_1出錯"
        
    End If
End Sub
Private Sub secondmethod1_3(X, Y, z)
'再讓邊塊置入第二層
        If X = Y Then
            If term = 3 Then    '先轉X軸
                Call secondchoose2(box, X, 0, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, X, 0, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, 0, Y, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, 0, Y, 0, 1)
            
            End If
            
            If term = 4 Then    '先轉Y軸
                Call secondchoose1(box, 0, Y, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, 0, Y, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, X, 0, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, X, 0, 0, 1)
                
            End If
            
        Else
            If term = 3 Then    '先轉X軸
                Call secondchoose1(box, X, 0, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, X, 0, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, 0, Y, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, 0, Y, 0, 1)
            
            End If
            
            If term = 4 Then    '先轉Y軸
                Call secondchoose2(box, 0, Y, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, 0, Y, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, X, 0, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, X, 0, 0, 1)
                
            End If
        
        End If
End Sub
Private Sub secondchoose(fox() As Integer, a, b, c) '副程式由secondmethod-系列呼叫
'建立側面常規順時針轉動
    If a = 1 Or b = 1 Then
        Call choose(fox, a, b, c)
        Call choose(fox, a, b, c)
        Call choose(fox, a, b, c)
        
    ElseIf a = 3 Or b = 3 Then
        Call choose(fox, a, b, c)
          
    Else
        Debug.Print "secondchoose出現錯誤"
    End If
    
End Sub
Private Sub secondchoose1(fox() As Integer, a, b, c, show) '副程式由secondmethod-系列呼叫
'建立側面常規逆時針轉動
    Dim d As Integer
    d = 3
    Call secondchoose(fox, a, b, c)
    
    If a <> 0 Then
        Debug.Print "X軸", a, "側面順時針轉"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
    End If
    
    If b <> 0 Then
        Debug.Print "Y軸", b, "側面順時針轉"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
    End If
End Sub
Private Sub secondchoose2(fox() As Integer, a, b, c, show) '副程式由secondmethod-系列呼叫
'建立側面常規逆時針轉動
    Dim d As Integer
    d = 4
    Call secondchoose(fox, a, b, c)
    Call secondchoose(fox, a, b, c)
    Call secondchoose(fox, a, b, c)
    
    If a <> 0 Then
        
        Debug.Print "X軸", a, "側面逆時針轉"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
    End If
    
    If b <> 0 Then
        Debug.Print "Y軸", b, "側面逆時針轉"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    End If
End Sub


Private Sub thirdSET() '設定遮罩
    '先初始化
    For i = 1 To 3
        For j = 1 To 3
            Third0_1(i, j) = 0
            Third0_2(i, j) = 0
            Third1_1(i, j) = 0
            Third1_2(i, j) = 0
            Third1_3(i, j) = 0
            Third2_1(i, j) = 0
            Third2_2(i, j) = 0
            Third2_3(i, j) = 0
            Third2_4(i, j) = 0
            Third2_5(i, j) = 0
            Third2_6(i, j) = 0
            Third2_7(i, j) = 0
        Next j
    Next i
    
    '邊塊Z軸復原
        'Third1_1直角
    Third1_1(1, 2) = W
    Third1_1(2, 1) = W
    Third1_1(2, 2) = W
    
        'Third1_2直線
    Third1_2(1, 2) = W
    Third1_2(2, 2) = W
    Third1_2(3, 2) = W
    
        'Third1_3中心點
    Third1_3(2, 2) = W
    
    
    '頂面Z軸復原
        'Third2_1 c1和c2
    Third2_1(1, 1) = W
    Third2_1(1, 2) = W
    Third2_1(2, 1) = W
    Third2_1(2, 2) = W
    Third2_1(2, 3) = W
    Third2_1(3, 2) = W
    
        'Third2_2 c3和c4
    Third2_2(1, 2) = W
    Third2_2(2, 1) = W
    Third2_2(2, 2) = W
    Third2_2(2, 3) = W
    Third2_2(3, 1) = W
    Third2_2(3, 2) = W
    Third2_2(3, 3) = W
    
        'Third2_3 c5
    Third2_3(1, 2) = W
    Third2_3(1, 3) = W
    Third2_3(2, 1) = W
    Third2_3(2, 2) = W
    Third2_3(2, 3) = W
    Third2_3(3, 1) = W
    Third2_3(3, 2) = W
    
        'Third2_4 c6和c7
    Third2_4(1, 2) = W
    Third2_4(2, 1) = W
    Third2_4(2, 2) = W
    Third2_4(2, 3) = W
    Third2_4(3, 2) = W

End Sub
Private Sub thirdchoose1() '將通用模組Third0_1旋轉
Dim i As Integer
Dim j As Integer
    For i = 1 To 3
        For j = 1 To 3
            Call thirdchoose2(i2, j2, i, j)
            Third0_2(i2, j2) = Third0_1(i, j)
            
        Next j
    Next i
    
    For i = 1 To 3
        For j = 1 To 3
            Third0_1(i, j) = Third0_2(i, j)
            
        Next j
    Next i
    
End Sub

Private Sub thirdchoose2(i2, j2, i1, j1) '由Thirdchoose1呼叫
'進行座標轉換程序
    i2 = j1
    j2 = 4 - i1
End Sub
Private Sub thirdCopy(X() As Integer, Y() As Integer)
Dim i As Integer
Dim j As Integer
    For i = 1 To 3
        For j = 1 To 3
            X(i, j) = Y(i, j)
        Next j
    Next i
End Sub

Private Sub thirdmethod()
    Call colorSET
    Call thirdSET
Debug.Print "呼叫thirdmethod1_1"
    Call thirdmethod1_1
Debug.Print "呼叫thirdmethod2_1"
    Call thirdmethod2_1
Debug.Print "第三層角塊復原"
    Call thirdmethod3_1
Debug.Print "第三層邊塊復原"
    Call thirdmethod4_1
    
    
End Sub
Private Sub thirdmethod1_1() '將第三層邊塊z軸恢復
    term = 0
    
    Call thirdmethod1_3(Third2_4)
    
    If term <> 1 Then
        Call thirdmethod1_3(Third1_1)
        If term = 1 Then
            Call thirdmethod1_4
        End If
        
    End If
    
    If term <> 1 Then
        Call thirdmethod1_3(Third1_2)
        If term = 1 Then
            Call thirdmethod1_4
            Call thirdmethod1_4
        End If
    End If
    
    If term <> 1 Then
        Call thirdmethod1_3(Third1_3)
        If term = 1 Then
            Call thirdmethod1_4
            Call choose1(box, 0, 0, 3, 1)
            Call thirdmethod1_4
            Call thirdmethod1_4
        End If
    End If
    
End Sub
Private Sub thirdmethod2_1() '將第三層z軸面復原
    term = 0
    
    If term <> 1 Then
        Call thirdmethod2_3(Third2_2) 'c3 c4
        If term = 1 Then
            If box(1, 1, 3, 1) = W And box(1, 3, 3, 1) = W Then
                Call thirdmethod2_4
            
            ElseIf box(1, 1, 3, 2) = W And box(1, 3, 3, 2) = W Then
                Call choose2(box, 0, 0, 3, 1)
                Call thirdmethod2_4
            End If
            
        End If
    End If
    
    If term <> 1 Then
        Call thirdmethod2_3(Third2_3) 'c5
        If term = 1 Then
            If box(1, 1, 3, 1) = W And box(3, 3, 3, 2) = W Then
                Call thirdmethod2_4
            
            
            ElseIf box(1, 1, 3, 2) = W And box(3, 3, 3, 1) = W Then
                 Call choose1(box, 0, 0, 3, 1)
                 Call choose1(box, 0, 0, 3, 1)
                 Call thirdmethod2_4
            End If
            
        End If
    End If
    
    If term <> 1 Then
    
        Call thirdmethod2_3(Third2_4) 'c6 c7
        If term = 1 Then
            If box(1, 1, 3, 1) = W And box(1, 3, 3, 1) = W And box(3, 1, 3, 1) = W And box(3, 3, 3, 1) = W Then 'c7
                Call choose1(box, 0, 0, 3, 1)
                Call thirdmethod2_4
            
            ElseIf box(1, 1, 3, 2) = W And box(3, 3, 3, 2) = W And box(1, 1, 3, 2) = W And box(3, 3, 3, 2) = W Then 'c7
                
                Call thirdmethod2_4
                
            ElseIf box(1, 1, 3, 2) = W And box(1, 3, 3, 2) = W Then 'c6
                Call thirdmethod2_1
            
            ElseIf box(1, 1, 3, 2) <> W And box(1, 3, 3, 2) = W Then 'c6
                Call choose1(box, 0, 0, 3, 1)
                Call thirdmethod2_4
            
            ElseIf box(1, 1, 3, 2) <> W And box(1, 3, 3, 2) <> W Then 'c6
                Call choose1(box, 0, 0, 3, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call thirdmethod2_4
            
            
            ElseIf box(1, 1, 3, 2) = W And box(1, 3, 3, 2) <> W Then
                Call choose2(box, 0, 0, 3, 1)
                Call thirdmethod2_4
            
            End If

        End If
    End If
    
    term = 0
        Call thirdmethod2_3(Third2_1) 'c1 c2
        If term = 1 Then
            If box(1, 3, 3, 1) = W And box(3, 1, 3, 1) = W Then
                Call thirdmethod2_4
            End If
            
            If box(1, 3, 3, 2) = W And box(3, 1, 3, 2) = W Then
                Call thirdmethod2_5
            End If
                        
        End If
        
   
    
End Sub
Private Sub thirdmethod3_1() '將第三層角塊歸位
    term = 0
    If box(1, 1, 3, 1) = box(1, 3, 3, 1) And box(3, 1, 3, 1) = box(3, 3, 3, 1) Then
        term = 1
    End If
    
    If term <> 1 Then
        If box(1, 1, 3, 1) = box(1, 3, 3, 1) Then '上
            Call choose2(box, 0, 0, 3, 1)
            Call thirdmethod3_2
            
            Debug.Print "角塊1抓到"
        ElseIf box(1, 3, 3, 2) = box(3, 3, 3, 2) Then '右
            Call choose2(box, 0, 0, 3, 1)
            Call choose2(box, 0, 0, 3, 1)
            Call thirdmethod3_2
            
             Debug.Print "角塊2抓到"
        ElseIf box(3, 1, 3, 1) = box(3, 3, 3, 1) Then '下
            Call choose1(box, 0, 0, 3, 1)
            Call thirdmethod3_2
            
             Debug.Print "角塊3抓到"
        ElseIf box(1, 1, 3, 2) = box(3, 1, 3, 2) Then '左
            Call thirdmethod3_2
        Else
            Call thirdmethod3_2
            Call choose2(box, 0, 0, 3, 1)
            Call thirdmethod3_2
        End If
    End If

End Sub
Private Sub thirdmethod4_1() '將第三層的邊塊歸位
    term = 1
    
    Call thirdmethod4_4
    
    If term = 0 Then
        Call thirdmethod4_2
    
        If term = 2 Then
            Call thirdmethod4_2
        End If
    End If
    
    Call thirdmethod4_5
    
       
End Sub
Private Sub thirdmethod1_2(Thirdterm, X() As Integer)
'由THIRDMETHOD1_3呼叫
'與遮罩比較
Dim i As Integer
Dim j As Integer
    For i = 1 To 3
        For j = 1 To 3
            If X(i, j) = W And Thirdterm <> 2 Then
                If box(i, j, 3, 3) = W Then
                    Thirdterm = 1
                Else
                    Thirdterm = 2
                End If
            End If
        Next j
    Next i
End Sub
Private Sub thirdmethod1_3(X() As Integer)
'由THIRDMETHOD1_1呼叫
'呼叫此方法尋找適合的遮罩
Dim i As Integer
Dim i2 As Integer
Dim j As Integer
Dim k As Integer

 
        For i = 0 To 4
Debug.Print "運行第", i, "次"

            Call thirdmethod1_2(Thirdterm, X) '與遮罩比較
            If Thirdterm = 1 Then
Debug.Print "抓取成功"
                     
            End If
            
            '測試
           ' For j = 1 To 3
           '     For k = 1 To 3
           '         Debug.Print j, k, x(j, k)
           '     Next k
           ' Next j
            Debug.Print ""
            
            
            '旋轉
            Call thirdCopy(Third0_1, X)
            Call thirdchoose1
            Call thirdCopy(X, Third0_1)
            
            If Thirdterm = 1 Then term = 1
            If Thirdterm = 1 Then Exit For
            Thirdterm = 0
        Next i
        
        Debug.Print "測試I", i
        While (i <> 0 And i < 4)
            i = i - 1
            Call choose2(box, 0, 0, 3, 1)
        Wend
        
End Sub
Private Sub thirdmethod1_4() '第三層邊塊復原z面公式
    Call secondchoose2(box, 0, 3, 0, 1)
    Call choose2(box, 0, 0, 3, 1)
    Call secondchoose2(box, 3, 0, 0, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
    Call secondchoose1(box, 0, 3, 0, 1)
End Sub
Private Sub thirdmethod2_2(Thirdterm, X() As Integer)
'由THIRDMETHOD1_3呼叫
'與遮罩比較
Dim i As Integer
Dim j As Integer
    For i = 1 To 3
        For j = 1 To 3
            If (X(i, j) = W Or box(i, j, 3, 3) = W) And Thirdterm <> 2 Then
                If X(i, j) = box(i, j, 3, 3) Then
                    Thirdterm = 1
                Else
                    Thirdterm = 2
                End If
                
                
                
            End If
        Next j
    Next i
End Sub
Private Sub thirdmethod2_3(X() As Integer)
'由THIRDMETHOD1_1呼叫
'呼叫此方法尋找適合的遮罩
Dim i As Integer
Dim i2 As Integer
Dim j As Integer
Dim k As Integer

 
        For i = 0 To 4
Debug.Print "運行第", i, "次"

            Call thirdmethod2_2(Thirdterm, X) '與遮罩比較
            If Thirdterm = 1 Then
Debug.Print "抓取成功"
                     
            End If
            
            '測試
           ' For j = 1 To 3
           '     For k = 1 To 3
           '         Debug.Print j, k, x(j, k)
           '     Next k
           ' Next j
            Debug.Print ""
            
            
            '旋轉
            Call thirdCopy(Third0_1, X)
            Call thirdchoose1
            Call thirdCopy(X, Third0_1)
            
            If Thirdterm = 1 Then term = 1
            If Thirdterm = 1 Then Exit For
            
            Thirdterm = 0
        Next i
        
        Debug.Print "測試I", i
        While (i <> 0 And i < 4)
            i = i - 1
            Call choose2(box, 0, 0, 3, 1)
        Wend
        
End Sub
Private Sub thirdmethod2_4() 'c1第三層復原z面公式
    Call secondchoose2(box, 0, 3, 0, 1)
    Call choose2(box, 0, 0, 3, 1)
    Call secondchoose1(box, 0, 3, 0, 1)
    Call choose2(box, 0, 0, 3, 1)
    Call secondchoose2(box, 0, 3, 0, 1)
    Call choose2(box, 0, 0, 3, 1)
    Call choose2(box, 0, 0, 3, 1)
    Call secondchoose1(box, 0, 3, 0, 1)

End Sub
Private Sub thirdmethod2_5() 'c2第三層復原z面公式
    Call secondchoose1(box, 3, 0, 0, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call secondchoose2(box, 3, 0, 0, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call secondchoose2(box, 3, 0, 0, 1)
End Sub
Private Sub thirdmethod3_2() '第三層換角公式
    Call secondchoose1(box, 3, 0, 0, 1)
    Call secondchoose2(box, 0, 3, 0, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
    Call secondchoose2(box, 0, 1, 0, 1)
    Call secondchoose2(box, 0, 1, 0, 1)
    Call secondchoose2(box, 3, 0, 0, 1)
    Call secondchoose1(box, 0, 3, 0, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
    Call secondchoose1(box, 0, 1, 0, 1)
    Call secondchoose1(box, 0, 1, 0, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
    Call secondchoose1(box, 3, 0, 0, 1)

End Sub
Private Sub thirdmethod4_2()
    If box(1, 1, 3, 1) = box(1, 2, 3, 1) And box(1, 3, 3, 1) = box(1, 2, 3, 1) Then '上
            Call thirdmethod4_3
        
        ElseIf box(1, 3, 3, 2) = box(2, 3, 3, 2) And box(3, 3, 3, 2) = box(2, 3, 3, 2) Then '右
            Call choose2(box, 0, 0, 3, 1)
            Call thirdmethod4_3
        
        ElseIf box(3, 1, 3, 1) = box(3, 2, 3, 1) And box(3, 3, 3, 1) = box(3, 2, 3, 1) Then '下
            Call choose1(box, 0, 0, 3, 1)
            Call choose1(box, 0, 0, 3, 1)
            Call thirdmethod4_3
    
        ElseIf box(1, 1, 3, 2) = box(2, 1, 3, 2) And box(3, 1, 3, 2) = box(2, 1, 3, 2) Then '左
            Call choose1(box, 0, 0, 3, 1)
            Call thirdmethod4_3
            
        Else
            term = 2
            Call thirdmethod4_3_1
            
        End If
End Sub
Private Sub thirdmethod4_3()
    If box(2, 1, 3, 2) = box(3, 1, 3, 1) Then '逆時針觸發
        Call thirdmethod4_3_1
    End If
    
    If box(2, 3, 3, 2) = box(3, 1, 3, 1) Then '順時針觸發
        Call thirdmethod4_3_2
    End If
    
End Sub
Private Sub thirdmethod4_3_1() '第三層逆時針三角換邊
    Call secondchoose1(box, 3, 0, 0, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
    Call choose2(box, 0, 0, 3, 1)
    Call choose1(box, 0, 2, 0, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call choose2(box, 0, 2, 0, 1)
    Call choose2(box, 0, 0, 3, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
End Sub
Private Sub thirdmethod4_3_2() '第三層順時針三角換邊
    Call secondchoose1(box, 3, 0, 0, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call choose1(box, 0, 2, 0, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call choose2(box, 0, 2, 0, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
End Sub

Private Sub thirdmethod4_4()
    If box(1, 1, 3, 1) <> box(1, 2, 3, 1) Then '上
        term = 0
    End If
        
    If box(1, 3, 3, 2) <> box(2, 3, 3, 2) Then '右
        term = 0
            
    End If
        
    If box(3, 1, 3, 1) <> box(3, 2, 3, 1) Then '下
        term = 0
            
    End If
        
    If box(1, 1, 3, 2) <> box(2, 1, 3, 2) Then  '左
        term = 0
            
    End If
End Sub
Private Sub thirdmethod4_5()
    While box(1, 2, 3, 1) <> box(1, 2, 2, 1)
        
        Call choose1(box, 0, 0, 3, 1)
    Wend
    
    
      
End Sub
Private Sub Command2_Click()
    
    Timer1.Enabled = True
    
End Sub
Private Sub cmdShow_Click()
'測試用
    
    Call Openfile(2) '開啟檔案
    
    Call showmethod3
    
    
    
End Sub
Private Sub showmethod1()
    Dim a, b, c, d As Integer
    
    Input #1, a, b, c, d '讀取一條紀錄
    Call showmethod4(a, b, c, d)
        
    If d = 1 Then
        Call choose1(showBox, a, b, c, 0)
    End If
    
    If d = 2 Then
        Call choose2(showBox, a, b, c, 0)
    End If
    
    If d = 3 Then
        Call secondchoose1(showBox, a, b, c, 0)
    End If
    
    If d = 4 Then
        Call secondchoose2(showBox, a, b, c, 0)
    End If
    
    Call showmethod3
    
End Sub
Private Sub showmethod2() '紀錄初始狀態
    Dim i, j, k As Integer
    For k = 1 To 3
        For i = 1 To 3
            For j = 1 To 3
                showBox(i, j, k, 1) = box(i, j, k, 1)
                showBox(i, j, k, 2) = box(i, j, k, 2)
                showBox(i, j, k, 3) = box(i, j, k, 3)
            Next j
        Next i
    Next k
End Sub
Private Sub showmethod3() '介面顏色顯示

    Call showmethod3_1
End Sub
Private Sub showmethod3_1()
    
    'X前面
    Call showmethod3_2(showBox(3, 1, 3, 1), Shape1(1))
    Call showmethod3_2(showBox(3, 2, 3, 1), Shape1(2))
    Call showmethod3_2(showBox(3, 3, 3, 1), Shape1(3))
    Call showmethod3_2(showBox(3, 1, 2, 1), Shape1(4))
    Call showmethod3_2(showBox(3, 2, 2, 1), Shape1(5))
    Call showmethod3_2(showBox(3, 3, 2, 1), Shape1(6))
    Call showmethod3_2(showBox(3, 1, 1, 1), Shape1(7))
    Call showmethod3_2(showBox(3, 2, 1, 1), Shape1(8))
    Call showmethod3_2(showBox(3, 3, 1, 1), Shape1(9))
    'Y右面
    Call showmethod3_2(showBox(3, 3, 3, 2), Shape3(1))
    Call showmethod3_2(showBox(2, 3, 3, 2), Shape3(2))
    Call showmethod3_2(showBox(1, 3, 3, 2), Shape3(3))
    Call showmethod3_2(showBox(3, 3, 2, 2), Shape3(4))
    Call showmethod3_2(showBox(2, 3, 2, 2), Shape3(5))
    Call showmethod3_2(showBox(1, 3, 2, 2), Shape3(6))
    Call showmethod3_2(showBox(3, 3, 1, 2), Shape3(7))
    Call showmethod3_2(showBox(2, 3, 1, 2), Shape3(8))
    Call showmethod3_2(showBox(1, 3, 1, 2), Shape3(9))
    'X後面
    Call showmethod3_2(showBox(1, 1, 3, 1), Shape6(1))
    Call showmethod3_2(showBox(1, 2, 3, 1), Shape6(2))
    Call showmethod3_2(showBox(1, 3, 3, 1), Shape6(3))
    Call showmethod3_2(showBox(1, 1, 2, 1), Shape6(4))
    Call showmethod3_2(showBox(1, 2, 2, 1), Shape6(5))
    Call showmethod3_2(showBox(1, 3, 2, 1), Shape6(6))
    Call showmethod3_2(showBox(1, 1, 1, 1), Shape6(7))
    Call showmethod3_2(showBox(1, 2, 1, 1), Shape6(8))
    Call showmethod3_2(showBox(1, 3, 1, 1), Shape6(9))
    'Y左面
    Call showmethod3_2(showBox(3, 1, 3, 2), Shape4(1))
    Call showmethod3_2(showBox(2, 1, 3, 2), Shape4(2))
    Call showmethod3_2(showBox(1, 1, 3, 2), Shape4(3))
    Call showmethod3_2(showBox(3, 1, 2, 2), Shape4(4))
    Call showmethod3_2(showBox(2, 1, 2, 2), Shape4(5))
    Call showmethod3_2(showBox(1, 1, 2, 2), Shape4(6))
    Call showmethod3_2(showBox(3, 1, 1, 2), Shape4(7))
    Call showmethod3_2(showBox(2, 1, 1, 2), Shape4(8))
    Call showmethod3_2(showBox(1, 1, 1, 2), Shape4(9))
    'Z面
    Call showmethod3_2(showBox(1, 1, 3, 3), Shape2(1))
    Call showmethod3_2(showBox(1, 2, 3, 3), Shape2(2))
    Call showmethod3_2(showBox(1, 3, 3, 3), Shape2(3))
    Call showmethod3_2(showBox(2, 1, 3, 3), Shape2(4))
    Call showmethod3_2(showBox(2, 2, 3, 3), Shape2(5))
    Call showmethod3_2(showBox(2, 3, 3, 3), Shape2(6))
    Call showmethod3_2(showBox(3, 1, 3, 3), Shape2(7))
    Call showmethod3_2(showBox(3, 2, 3, 3), Shape2(8))
    Call showmethod3_2(showBox(3, 3, 3, 3), Shape2(9))
    

End Sub
Private Sub showmethod3_2(Color, Colorshape)
    If Color = 1 Then
        Colorshape.FillColor = Shape1(0).FillColor
    End If
    If Color = 2 Then
        Colorshape.FillColor = Shape2(0).FillColor
    End If
    If Color = 3 Then
        Colorshape.FillColor = Shape3(0).FillColor
    End If
    If Color = 4 Then
        Colorshape.FillColor = Shape4(0).FillColor
    End If
    If Color = 5 Then
        Colorshape.FillColor = Shape5(0).FillColor
    End If
    If Color = 6 Then
        Colorshape.FillColor = Shape6(0).FillColor
    End If
    
    
End Sub
Private Sub showmethod4(a, b, c, d) '轉動箭頭指示介面顯示

    If showterm = 0 Then showterm = 1
    Call showmethod4_1
    Call showmethod4_2(a, b, c, d)
    
    If showterm = 1 Then
        showterm = 2
    Else
        showterm = 1
    End If

End Sub
Private Sub showmethod4_1() '清空指示欄位
    Dim i As Integer
    
    For i = 1 To 3
        Text1(i).text = ""
        Text1(i + 3).text = ""
        Text3(i).text = ""
        Text3(i + 3).text = ""
        Call showmethod4_3_1(Text1(i))
        Call showmethod4_3_1(Text1(i + 3))
        Call showmethod4_3_1(Text3(i))
        Call showmethod4_3_1(Text3(i + 3))
    Next i
    
    For i = 1 To 4
        Text2(i).text = ""
        Text4(i).text = ""
        Text5(i).text = ""
        Text6(i).text = ""
        Text7(i).text = ""
        Text8(i).text = ""
        Call showmethod4_3_1(Text2(i))
        Call showmethod4_3_1(Text4(i))
        Call showmethod4_3_1(Text5(i))
        Call showmethod4_3_1(Text6(i))
        Call showmethod4_3_1(Text7(i))
        Call showmethod4_3_1(Text8(i))
        
    Next i
    
End Sub
Private Sub showmethod4_2(a, b, c, d) '選擇顯示正在被旋轉的軸
    If a <> 0 Then
        Call showmethod4_2_1(a, d)
    End If
    
    If b <> 0 Then
        Call showmethod4_2_2(b, d)
    End If
    
    If c <> 0 Then
        Call showmethod4_2_3(c, d)
    End If

    
End Sub
Private Sub showmethod4_2_1(a, d) 'x軸旋轉顯示介面
    If a = 1 Then
        If d = 1 Or d = 4 Then
            Text1(1).text = "--->"
            Text1(4).text = "--->"
        End If
        
        If d = 2 Or d = 3 Then
            Text1(1).text = "<---"
            Text1(4).text = "<---"
        End If
            
        Call showmethod4_3(Text1(1))
        Call showmethod4_3(Text1(4))
    End If
    
    If a = 2 Then
        If d = 1 Then
            Text1(2).text = "--->"
            Text1(5).text = "--->"
        End If
        If d = 2 Then
            Text1(2).text = "<---"
            Text1(5).text = "<---"
        End If
        
        Call showmethod4_3(Text1(2))
        Call showmethod4_3(Text1(5))
    End If
    
    If a = 3 Then
        If d = 1 Or d = 3 Then
            Text1(3).text = "--->"
            Text1(6).text = "--->"
        End If
        
        If d = 2 Or d = 4 Then
            Text1(3).text = "<---"
            Text1(6).text = "<---"
        End If
        
        Call showmethod4_3(Text1(3))
        Call showmethod4_3(Text1(6))
    End If
    
End Sub
Private Sub showmethod4_2_2(b, d) 'y軸旋轉顯示介面
     If b = 1 Then
        If d = 1 Or d = 4 Then
            Text2(1).text = "^"
            Text2(2).text = "|"
            Text2(3).text = "^"
            Text2(4).text = "|"
        End If
        
        If d = 2 Or d = 3 Then
            Text2(1).text = "|"
            Text2(2).text = "v"
            Text2(3).text = "|"
            Text2(4).text = "v"
        End If
        
        Call showmethod4_3(Text2(1))
        Call showmethod4_3(Text2(2))
        Call showmethod4_3(Text2(3))
        Call showmethod4_3(Text2(4))
    End If
    
    If b = 2 Then
        If d = 1 Then
            Text4(1).text = "^"
            Text4(2).text = "|"
            Text4(3).text = "^"
            Text4(4).text = "|"
           
        End If
        
        If d = 2 Then
            Text4(1).text = "|"
            Text4(2).text = "v"
            Text4(3).text = "|"
            Text4(4).text = "v"
        End If
        
        Call showmethod4_3(Text4(1))
        Call showmethod4_3(Text4(2))
        Call showmethod4_3(Text4(3))
        Call showmethod4_3(Text4(4))
    End If
    
    If b = 3 Then
        If d = 1 Or d = 3 Then
            Text5(1).text = "^"
            Text5(2).text = "|"
            Text5(3).text = "^"
            Text5(4).text = "|"
        End If
        
        If d = 2 Or d = 4 Then
            Text5(1).text = "|"
            Text5(2).text = "v"
            Text5(3).text = "|"
            Text5(4).text = "v"
        End If
        
        Call showmethod4_3(Text5(1))
        Call showmethod4_3(Text5(2))
        Call showmethod4_3(Text5(3))
        Call showmethod4_3(Text5(4))
    End If
            
End Sub
Private Sub showmethod4_2_3(c, d) 'z軸旋轉顯示介面
    If c = 1 Then
        If d = 1 Then
            Text3(1).text = "--->"
            Text3(4).text = "--->"
            Text6(1).text = "|"
            Text6(2).text = "^"
            Text6(3).text = "|"
            Text6(4).text = "v"
            
        End If
        
        If d = 2 Then
            Text3(1).text = "<---"
            Text3(4).text = "<---"
            Text6(1).text = "v"
            Text6(2).text = "|"
            Text6(3).text = "^"
            Text6(4).text = "|"
        End If
        
        Call showmethod4_3(Text3(1))
        Call showmethod4_3(Text3(4))
        Call showmethod4_3(Text6(1))
        Call showmethod4_3(Text6(2))
        Call showmethod4_3(Text6(3))
        Call showmethod4_3(Text6(4))
    End If
    
    If c = 2 Then
        If d = 1 Then
            Text3(2).text = "--->"
            Text3(5).text = "--->"
            Text7(1).text = "|"
            Text7(2).text = "^"
            Text7(3).text = "|"
            Text7(4).text = "v"
            
        End If
        
        If d = 2 Then
            Text3(2).text = "<---"
            Text3(5).text = "<---"
            Text7(1).text = "v"
            Text7(2).text = "|"
            Text7(3).text = "^"
            Text7(4).text = "|"
        End If
        
        Call showmethod4_3(Text3(2))
        Call showmethod4_3(Text3(5))
        Call showmethod4_3(Text7(1))
        Call showmethod4_3(Text7(2))
        Call showmethod4_3(Text7(3))
        Call showmethod4_3(Text7(4))
    End If
    
    If c = 3 Then
        If d = 1 Then
            Text3(3).text = "--->"
            Text3(6).text = "--->"
            Text8(1).text = "|"
            Text8(2).text = "^"
            Text8(3).text = "|"
            Text8(4).text = "v"
            
        End If
        
        If d = 2 Then
            Text3(3).text = "<---"
            Text3(6).text = "<---"
            Text8(1).text = "v"
            Text8(2).text = "|"
            Text8(3).text = "^"
            Text8(4).text = "|"
        End If
        
        Call showmethod4_3(Text3(3))
        Call showmethod4_3(Text3(6))
        Call showmethod4_3(Text8(1))
        Call showmethod4_3(Text8(2))
        Call showmethod4_3(Text8(3))
        Call showmethod4_3(Text8(4))
    End If
      
End Sub
Private Sub showmethod4_3(block) '旋轉指示箭頭顏色轉換
    If showterm = 1 Then
        block.BackColor = Shape7(1).FillColor
    End If
    
    If showterm = 2 Then
        block.BackColor = Shape7(2).FillColor
    End If
End Sub
Private Sub showmethod4_3_1(block)
    block.BackColor = Shape7(0).FillColor
End Sub

Private Sub text()
Dim i, j, k As Integer
For k = 1 To 3
    For i = 1 To 3
        For j = 1 To 3
            
                
                Call printcolor(i, j, k)
          
        Next j
    Next i
Next k

   

End Sub

Private Sub Command3_Click()
    Call showmethod1

End Sub

Private Sub Command4_Click()
    Call Openfile(3)
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Command5_Click()
'Call showmethod4_1
    Timer1.Enabled = False
End Sub



Private Sub Command6_Click()
    CommonDialog1.DialogTitle = "load an image"
    CommonDialog1.Filter = "*|*.*|mim|*.mim|tif|*.tif|"
    CommonDialog1.ShowOpen
    Buffer1.Load CommonDialog1.FileName, True
End Sub

Private Sub Command7_Click()
Dim i, j, k, n As Integer
Call colorSET




'先歸零
For i = 1 To 3
    For j = 1 To 3
        For k = 1 To 3
            For n = 1 To 3
            
                box2(i, j, k, n) = 0
                box(i, j, k, n) = 0
            
            Next n
        Next k
    Next j
Next i



'開啟文件輸入顏色到box

Call OpenColorfile(2, 3, 0, 4, 1) '前面

Call OpenColorfile(4, 1, 4, 4, 1) '後面

Call OpenColorfile(1, 0, 1, 4, 2) '左面

Call OpenColorfile(3, 4, 3, 4, 2) '右面

Call OpenColorfile(5, 0, 0, 3, 3) '上面

Call OpenColorfile(6, 4, 0, 1, 3) '下面


'-----------------------初始狀況
For k = 1 To 3
    For i = 1 To 3
        For j = 1 To 3
            
                
                Call printcolor(i, j, k)
          
        Next j
    Next i
Next k
End Sub

Private Sub Timer1_Timer()
    Call showmethod1
End Sub
>>>>>>> 5c6781aae922b293ac3c06c17d26eff6686b8936
