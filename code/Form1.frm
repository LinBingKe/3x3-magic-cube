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
      Name            =   "�s�ө���"
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
   StartUpPosition =   3  '�t�ιw�]��
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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Caption         =   "���R�˥��Ҧb��m"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
      Caption         =   "�Ȱ�"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Caption         =   "�}�l"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
      Caption         =   "�����ɮ�"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Caption         =   "�U�@�B"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Caption         =   "�O��������O"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Caption         =   "�P�_�}�l"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      FillStyle       =   0  '���
      Height          =   495
      Index           =   2
      Left            =   240
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   1
      Left            =   240
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H80000000&
      FillStyle       =   0  '���
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
      FillStyle       =   0  '���
      Height          =   495
      Index           =   1
      Left            =   2760
      Shape           =   1  '�����
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   4
      Left            =   2040
      Shape           =   1  '�����
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   7
      Left            =   1320
      Shape           =   1  '�����
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   2
      Left            =   2760
      Shape           =   1  '�����
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   5
      Left            =   2040
      Shape           =   1  '�����
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   8
      Left            =   1320
      Shape           =   1  '�����
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   3
      Left            =   2760
      Shape           =   1  '�����
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   6
      Left            =   2040
      Shape           =   1  '�����
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   9
      Left            =   1320
      Shape           =   1  '�����
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   3
      Left            =   5760
      Shape           =   1  '�����
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   2
      Left            =   5040
      Shape           =   1  '�����
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   1
      Left            =   4320
      Shape           =   1  '�����
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   6
      Left            =   5760
      Shape           =   1  '�����
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   5
      Left            =   5040
      Shape           =   1  '�����
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   4
      Left            =   4320
      Shape           =   1  '�����
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   9
      Left            =   5760
      Shape           =   1  '�����
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   8
      Left            =   5040
      Shape           =   1  '�����
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   7
      Left            =   4320
      Shape           =   1  '�����
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  '���
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
      FillStyle       =   0  '���
      Height          =   495
      Index           =   1
      Left            =   4320
      Shape           =   1  '�����
      Top             =   7320
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   2
      Left            =   5040
      Shape           =   1  '�����
      Top             =   7320
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   3
      Left            =   5760
      Shape           =   1  '�����
      Top             =   7320
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   4
      Left            =   4320
      Shape           =   1  '�����
      Top             =   8040
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   5
      Left            =   5040
      Shape           =   1  '�����
      Top             =   8040
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   6
      Left            =   5760
      Shape           =   1  '�����
      Top             =   8040
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   7
      Left            =   4320
      Shape           =   1  '�����
      Top             =   8760
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   8
      Left            =   5040
      Shape           =   1  '�����
      Top             =   8760
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   9
      Left            =   5760
      Shape           =   1  '�����
      Top             =   8760
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  '���
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
      FillStyle       =   0  '���
      Height          =   495
      Index           =   1
      Left            =   4320
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   2
      Left            =   5040
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   3
      Left            =   5760
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   4
      Left            =   4320
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   5
      Left            =   5040
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   6
      Left            =   5760
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   7
      Left            =   4320
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   8
      Left            =   5040
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   9
      Left            =   5760
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   0
      Left            =   240
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   1
      Left            =   7440
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   2
      Left            =   7440
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   3
      Left            =   7440
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   4
      Left            =   8160
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   5
      Left            =   8160
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   6
      Left            =   8160
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   7
      Left            =   8880
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   8
      Left            =   8880
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   9
      Left            =   8880
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   0
      Left            =   240
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H000080FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   0
      Left            =   240
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '���
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



Dim term As Integer '��ܧP�_
Dim showterm As Integer '������ܤ���C���ܴ��P�_

Private Sub Command1_Click() '���{���P�_�ӫ����
Dim i, j, k, n As Integer
Call colorSET

'���k�s
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



'�}�Ҥ���J�C���box

Call OpenColorfile(2, 3, 0, 4, 1) '�e��

Call OpenColorfile(4, 1, 4, 4, 1) '�᭱

Call OpenColorfile(1, 0, 1, 4, 2) '����

Call OpenColorfile(3, 4, 3, 4, 2) '�k��

Call OpenColorfile(5, 0, 0, 3, 3) '�W��

Call OpenColorfile(6, 4, 0, 1, 3) '�U��



'-----------------------��l���p
For k = 1 To 3
    For i = 1 To 3
        For j = 1 To 3
            
                
                Call printcolor(i, j, k)
          
        Next j
    Next i
Next k

Call Openfile(1)
Call showmethod2
'----------------------------------------------�}�l�Ĥ@�h�P�_
    Debug.Print "�Ĥ@�h�������"
    Debug.Print "�P�_121"
    Call firstEdgeBlock(1, 2, 1, E, 0, L)
    
    Debug.Print "�P�_211"
    Call firstEdgeBlock(2, 1, 1, 0, U, L)
    
    Debug.Print "�P�_231"
    Call firstEdgeBlock(2, 3, 1, 0, G, L)
    
    Debug.Print "�P�_321"
    Call firstEdgeBlock(3, 2, 1, R, 0, L)
    
    Debug.Print "�Ĥ@�h��������"
    Debug.Print "�P�_111"
    Call firstCornerBlock(1, 1, 1, E, U, L)
    Call text
    Debug.Print "�P�_131"
    Call firstCornerBlock(1, 3, 1, E, G, L)
    
    
    'Call text
    Debug.Print "�P�_311"
    
    Call firstCornerBlock(3, 1, 1, R, U, L)
    
    Debug.Print "�P�_331"
    Call firstCornerBlock(3, 3, 1, R, G, L)
    
'--------------------�Ĥ@�h�_��H�᪺���p
    
    
    For k = 1 To 3
    For i = 1 To 3
        For j = 1 To 3
            
                
                Call printcolor(i, j, k)
          
        Next j
    Next i
Next k

    


'----------------------------------------�ĤG�h�P�_
    Debug.Print "�ĤG�h����"
    Debug.Print "�P�_112"
    Call secondmethod(1, 1, 2, E, U, 0)
 
    Debug.Print "�P�_132"
    Call secondmethod(1, 3, 2, E, G, 0)
    
    Debug.Print "�P�_312"
    Call secondmethod(3, 1, 2, R, U, 0)
    
    Debug.Print "�P�_332"
    Call secondmethod(3, 3, 2, R, G, 0)
    
'--------------------�ĤG�h�_��H�᪺���p
For k = 1 To 3
    For i = 1 To 3
        For j = 1 To 3
            
                
                Call printcolor(i, j, k)
          
        Next j
    Next i
Next k

'-----------------------------------------�ĤT�h�P�_
    Debug.Print "�ĤT�h����"
    Call thirdmethod
    
 '--------------------�ĤT�h�_��H�᪺���p
For k = 1 To 3
    For i = 1 To 3
        For j = 1 To 3
            
                
                'Call printcolor(i, j, k)
          
        Next j
    Next i
Next k
   
Call Openfile(3)
'����
'Call choose1(0, 0, 3) '���઺��k:�ĤT�hz�b��V����




End Sub
Private Sub printcolor(i, j, k) '�d�ݤ���C��
Dim a As Integer

For a = 1 To 3
    If box(i, j, k, a) = 0 Then
    Debug.Print i, j, k, a, "0"
    End If
    
    If box(i, j, k, a) = 1 Then
    Debug.Print i, j, k, a, "��"
    End If
    
    If box(i, j, k, a) = 2 Then
    Debug.Print i, j, k, a, "��"
    End If
    
    If box(i, j, k, a) = 3 Then
    Debug.Print i, j, k, a, "��"
    End If
    
    If box(i, j, k, a) = 4 Then
    Debug.Print i, j, k, a, "��"
    End If
    
    If box(i, j, k, a) = 5 Then
    Debug.Print i, j, k, a, "��"
    End If
    
    If box(i, j, k, a) = 6 Then
    Debug.Print i, j, k, a, "��"
    End If
    
Next a
Debug.Print ""




End Sub
Private Sub OpenColorfile(NUMBER As Integer, a As Integer, b As Integer, c As Integer, d As Integer)
    '�P�_�}�Ҫ����
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

    If NUMBER = 1 Then '�������
        For i = 1 To 3
            For j = 1 To 3
                Input #1, Color 'Ū���@������
                box(a + i, b, c - j, d) = Color
            Next j
        Next i
        
    ElseIf NUMBER = 3 Then '����k��
        For i = 1 To 3
            For j = 1 To 3
                Input #1, Color 'Ū���@������
                box(a - i, b, c - j, d) = Color
            Next j
        Next i
    
    ElseIf NUMBER = 2 Then '����e��
        For i = 1 To 3
            For j = 1 To 3
                Input #1, Color 'Ū���@������
                box(a, b + i, c - j, d) = Color
            Next j
        Next i
        
    
      
    ElseIf NUMBER = 4 Then '����᭱
        For i = 1 To 3
            For j = 1 To 3
                Input #1, Color 'Ū���@������
                box(a, b - i, c - j, d) = Color
            Next j
        Next i
        
    ElseIf NUMBER = 5 Then '����W��
        For i = 1 To 3
            For j = 1 To 3
                Input #1, Color 'Ū���@������
                box(a + j, b + i, c, d) = Color
            Next j
        Next i
       
    ElseIf NUMBER = 6 Then '����U��
        For i = 1 To 3
            For j = 1 To 3
                Input #1, Color 'Ū���@������
                box(a - j, b + i, c, d) = Color
            Next j
        Next i
        
    Else
    End If
    
End Sub
Private Sub Openfile(X As Integer) '�}��Ū�g���
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
Private Sub colorSET() '�]�w�C��Ѽ�
R = 1 '����
L = 2 '�¦�
G = 3 '���
U = 4 '�Ŧ�
W = 5 '����
E = 6 '���
End Sub
Private Sub choose(fox() As Integer, a, b, c) '�إ߿�ܱ��઺��V�M�h�ƪ���k
    If b = 0 And c = 0 Then
        Call X(a, fox) 'x�b����
    End If

    If a = 0 And c = 0 Then
        Call Y(b, fox) 'y�b����
    End If
    
    If a = 0 And b = 0 Then
        Call z(c, fox) 'z�b����
    End If
End Sub
Private Sub choose1(fox() As Integer, a, b, c, show) '�إ߱��ॿ��V�����
    Dim d As Integer
    d = 1
    
    Call choose(fox, a, b, c)
    If a <> 0 Then
        Debug.Print "X�b", a, "����"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    ElseIf b <> 0 Then
        Debug.Print "Y�b", b, "����"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    Else
        Debug.Print "Z�b", c, "����"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    End If
End Sub
Private Sub choose2(fox() As Integer, a, b, c, show) '�إ߱���Ϥ�V�����
    Dim d As Integer
    d = 2
    
    Call choose(fox, a, b, c)
    Call choose(fox, a, b, c)
    Call choose(fox, a, b, c)
    
    If a <> 0 Then
        Debug.Print "X�b", a, "����"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    ElseIf b <> 0 Then
        Debug.Print "Y�b", b, "����"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    Else
        Debug.Print "Z�b", c, "����"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    End If
End Sub
Private Sub X(a, fox() As Integer) '�u��x�b���઺��k
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

'�b���^box1�s������
For i = 1 To 3
    For j = 1 To 3
        fox(a, i, j, 1) = box2(a, i, j, 1)
        fox(a, i, j, 2) = box2(a, i, j, 2)
        fox(a, i, j, 3) = box2(a, i, j, 3)
                
    Next j
Next i
End Sub
Private Sub Y(b, fox() As Integer) '�u��y�b���઺��k
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

'�b���^box1�s������
For i = 1 To 3
    For j = 1 To 3
        fox(j, b, i, 1) = box2(j, b, i, 1)
        fox(j, b, i, 2) = box2(j, b, i, 2)
        fox(j, b, i, 3) = box2(j, b, i, 3)
                
    Next j
Next i
End Sub
Private Sub z(c, fox() As Integer) '�u��z�b���઺��k
Dim i As Integer
Dim j As Integer

'����box2�x�s����
For i = 1 To 3
    For j = 1 To 3
        Call rotate(i2, j2, i, j) '�e��ӬO��Ӫ��y��
                                  '���ӬO�쥻���y��
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

'����
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
Private Sub rotate(a2, b2, a1, b1) '�إ߱���y���ഫ������
For i = 1 To 3
    For j = 1 To 3
        a2 = b1
        b2 = 4 - a1
    Next j
Next i

End Sub

Private Sub firstEdgeBlock(X, Y, z, cx, cy, cz) '�����Ĥ@�h���
    term = 0 '����P�_�ܼƥ���l��
Debug.Print "�P�_����O�_�b���T��m"
    Call firstmethod(X, Y, z, cx, cy, cz)
    
    If term <> 1 And term <> 5 Then
Debug.Print "�P�_����O�_�b�ĤG�h"
        Call firstmethod3(X, Y, z, 1, 1, 2, cx, cy, cz)
        Call firstmethod3(X, Y, z, 1, 3, 2, cx, cy, cz)
        Call firstmethod3(X, Y, z, 3, 1, 2, cx, cy, cz)
        Call firstmethod3(X, Y, z, 3, 3, 2, cx, cy, cz)
    
    End If
    
    If term <> 1 And term <> 3 Then
Debug.Print "Z�b���T �P�_�O�_���b�Ĥ@�h"
        Call firstmethod1(X, Y, z, 1, 2, 1, cx, cy, cz)
        Call firstmethod1(X, Y, z, 2, 1, 1, cx, cy, cz)
        Call firstmethod1(X, Y, z, 2, 3, 1, cx, cy, cz)
        Call firstmethod1(X, Y, z, 3, 2, 1, cx, cy, cz)
Debug.Print "Z�b���T �P�_�O�_���b�ĤT�h"
        Call firstmethod1(X, Y, z, 1, 2, 3, cx, cy, cz)
        Call firstmethod1(X, Y, z, 2, 1, 3, cx, cy, cz)
        Call firstmethod1(X, Y, z, 2, 3, 3, cx, cy, cz)
        Call firstmethod1(X, Y, z, 3, 2, 3, cx, cy, cz)
    End If
    
    
    If term <> 1 And term <> 3 Then
Debug.Print "Z�b�A�� �P�_�O�_���b�Ĥ@�h"
        Call firstmethod2(X, Y, z, 1, 2, 1, cx, cy, cz)
        Call firstmethod2(X, Y, z, 2, 1, 1, cx, cy, cz)
        Call firstmethod2(X, Y, z, 2, 3, 1, cx, cy, cz)
        Call firstmethod2(X, Y, z, 3, 2, 1, cx, cy, cz)
Debug.Print "Z�b�A�� �P�_�O�_���b�ĤT�h"
        Call firstmethod2(X, Y, z, 1, 2, 3, cx, cy, cz)
        Call firstmethod2(X, Y, z, 2, 1, 3, cx, cy, cz)
        Call firstmethod2(X, Y, z, 2, 3, 3, cx, cy, cz)
        Call firstmethod2(X, Y, z, 3, 2, 3, cx, cy, cz)
    End If
  
End Sub

Private Sub firstCornerBlock(X, Y, z, cx, cy, cz) '�����Ĥ@�h����
    term = 0 '�����l��
Debug.Print "�����O�_�b���T��m"
    Call firstmethod4(X, Y, z, cx, cy, cz)
    
    If term <> 1 And term <> 3 Then
Debug.Print "����z�b�O�_�¦�öi��B�z"
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
Debug.Print "�����O�_z�b�����¦�öi��B�z"
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
    '�P�_�O�_�b���
    Debug.Print "�P�_�O�_�b���"
    
    If box(X, Y, z, 1) = cx And box(X, Y, z, 2) = cy And box(X, Y, z, 3) = cz Then
        term = 1
        Debug.Print "���b���"
    End If
End Sub
Private Sub firstmethod1(x1, y1, z1, x2, y2, z2, cx, cy, cz)
    '�P�_�b�Ĥ@�h�βĤT�h
    '�P�_����z�b��V�������¦�
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
Debug.Print "�i�J����B�zfirstmethod1_1"
        term = 3
        '���ĤT�h�B�z
        If x2 <> 2 And z2 <> 3 Then
            Call secondchoose1(box, x2, 0, 0, 1)
            Call secondchoose1(box, x2, 0, 0, 1)
        End If

        If y2 <> 2 And z2 <> 3 Then
            Call secondchoose1(box, 0, y2, 0, 1)
            Call secondchoose1(box, 0, y2, 0, 1)
        End If
            
        '�b�ĤT�h����
        Call firstmethod2_2(x1, x2, y1, y2)
            
        '��^�Ĥ@�h
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
    '�P�_�b�Ĥ@�h�βĤT�h
    '���Oz�b�C���A��
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
Debug.Print "�i�J����B�zfirstmethod2_1"
        term = 3
        '���ĤT�h�B�z�����m�y��
        If x2 <> 2 And z2 <> 3 Then
            Call secondchoose1(box, x2, 0, 0, 1)
            Call secondchoose1(box, x2, 0, 0, 1)
        End If

        If y2 <> 2 And z2 <> 3 Then
            Call secondchoose1(box, 0, y2, 0, 1)
            Call secondchoose1(box, 0, y2, 0, 1)
        End If
            
        '�b�ĤT�h����
        Call firstmethod2_2(x1, x2, y1, y2)
            
        '��^�Ĥ@�h���T��m�y��
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
Private Sub firstmethod2_2(x1, x2, y1, y2) '�Ƶ{��----------��firstmethod2_1-�I�s
'�ت��N�b�ĤT�h��������Ĥ@�h����m�W��xy�b   �H�Q��m�J�ĤG�h��
'1�����T��m 2�������m
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
        Debug.Print "firstmethod2_2�X��"
            
        
    End If
End Sub
Private Sub firstmethod3(x1, y1, z1, x2, y2, z2, cx, cy, cz)
    '�P�_����b�ĤG�h
    '�y��1�����T��m �y��2�������m
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
'�N�b�ĤG�h��������ĤT�h
Debug.Print "�i�J����B�zfirstmethod3_1"
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
Debug.Print "firstmethod3-1 X2=1 Y2���~"
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
Debug.Print "firstmethod3-1 X2=1 Y2 ���~"
        End If
        
    Else
Debug.Print "firstmethod3-1 X2���~"
    End If
    
End Sub
Private Sub firstmethod4(x1, y1, z1, cx, cy, cz) '����
    '�O�_�b���T����m�W
    If box(x1, y1, z1, 1) = cx And box(x1, y1, z1, 2) = cy And box(x1, y1, z1, 3) = cz Then
        term = 1
    End If
End Sub
Private Sub firstmethod5(x1, y1, z1, x2, y2, z2, cx, cy, cz) '����
'�P�_����Z�b���¦�
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
Private Sub firstmethod5_1(x1, y1, z1, x2, y2, z2, cx, cy, cz) '����
'����Z�b���¦�
'�B�z:���z�b�����¦�
Debug.Print "�i�J����B�zfirstmethod5_1"
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
Debug.Print "firstmethod5_1 z2=1���~"

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
Debug.Print "firstmethod5_1 z2=3���~"

        
        
    End If
    

End Sub
Private Sub firstmethod6(x1, y1, z1, x2, y2, z2, cx, cy, cz) '����
'�P�_����Z�b�����¦�
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
Private Sub FIRSTMETHOD6_1(x1, y1, z1, x2, y2, z2, cx, cy, cz) '����
'�p�G�b�Ĥ@�h�����ĤT�h
Debug.Print "�i�J����B�zfirstmethod6_1"
    term = 4
    If z2 = 1 Then
Debug.Print "�p�G�b�Ĥ@�h�����ĤT�h"
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
    
'�b�ĤT�h�����A���m
Debug.Print "�b�ĤT�h�����A���m"
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
Debug.Print "firstmethod6_1 cx cy���~"
    End If
'�N�����쥿�T����m
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


Private Sub secondmethod(X, Y, z, cx, cy, cz) '�Ƶ{��----------�ѥD�{���I�s
'�ĤG�h��@�����m�P�_

    term = 0 '����P�_�ܼƥ���l��
    
    '�P�_�O�_�b���
    Debug.Print "�P�_�O�_�b���"
    
    If box(X, Y, z, 1) = cx And box(X, Y, z, 2) = cy Then
        term = 1
        Debug.Print "���b���"
    End If
    
    
    '�P�_�O�_�b�ĤG�h���Y�Ӧ�m
    Debug.Print "�P�_�O�_�b�ĤG�h"
    
    If term <> 1 Then
        Debug.Print "�P�_�b�ĤG�h������"
        Call secondmethod1_1(1, 1, 2, cx, cy, cz)
            Debug.Print "�z�Lsecondmethod1-1�P�_112"
        Call secondmethod1_1(1, 3, 2, cx, cy, cz)
            Debug.Print "�z�Lsecondmethod1-1�P�_132"
        Call secondmethod1_1(3, 1, 2, cx, cy, cz)
            Debug.Print "�z�Lsecondmethod1-1�P�_312"
        Call secondmethod1_1(3, 3, 2, cx, cy, cz)
            Debug.Print "�z�Lsecondmethod1-1�P�_332"
        
    End If
    
    '�P�_�O�_�b�ĤT�h���Y�Ӧ�m
    Debug.Print "�P�_�O�_�b�ĤT�h"
    
    If term <> 1 Then
        Debug.Print "�P�_�b�ĤT�h������"
            Debug.Print "�z�Lsecondmethod1-2�P�_123"
        Call secondmethod1_2(1, 2, 3, cx, cy, cz)
            Debug.Print "�z�Lsecondmethod1-2�P�_213"
        Call secondmethod1_2(2, 1, 3, cx, cy, cz)
            Debug.Print "�z�Lsecondmethod1-2�P�_233"
        Call secondmethod1_2(2, 3, 3, cx, cy, cz)
            Debug.Print "�z�Lsecondmethod1-2�P�_323"
        Call secondmethod1_2(3, 2, 3, cx, cy, cz)
            
        '�m�J�ĤG�h
            Debug.Print "�z�Lsecondmethod1-3�N����m�J�ĤG�h"
        Call secondmethod1_3(X, Y, z)
             
        
    
    End If




End Sub
Private Sub secondmethod1_1(X, Y, z, cx, cy, cz) '�Ƶ{��---------��secondmethod�t�C-�I�s
    '�P�_�O�_�b�ĤG�h���Y�Ӧ�m�W
    Debug.Print "�Ƶ{���P�_����b�ĤG�h������"
    
    Dim termsecond1_1 As Integer
    If (box(X, Y, z, 1) = cx And box(X, Y, z, 2) = cy) Or (box(X, Y, z, 1) = cy And box(X, Y, z, 2) = cx) Then
        term = 2
        termsecond1_1 = 2
    End If
    
    '�p�G����m�A���i��B�z
    '�B�z�覡���N������ܲĤT�h
    If termsecond1_1 = 2 Then
        Debug.Print "�ĤG�h���ܲĤT�h������B�J"
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
Private Sub secondmethod1_2(X, Y, z, cx, cy, cz) '�Ƶ{��----------��secondmethod�t�C-�I�s
    '�P�_�O�_�b�ĤT�h���Y�Ӧ�m�W
    Debug.Print "�Ƶ{���P�_����b�ĤT�h������"
    
    Dim termsecond1_2 As Integer
    If (box(X, Y, z, 1) = cx And box(X, Y, z, 3) = cy) Or (box(X, Y, z, 1) = cy And box(X, Y, z, 3) = cx) Then
        termsecond1_2 = 2
    End If
    
    If (box(X, Y, z, 2) = cx And box(X, Y, z, 3) = cy) Or (box(X, Y, z, 2) = cy And box(X, Y, z, 3) = cx) Then
        termsecond1_2 = 2
    End If
    
    
    Debug.Print "�P�_�ܼ�termsecond1_2", termsecond1_2
    
    '�p�G����m(�b�ĤT�h)�A���i��B�z
    If termsecond1_2 = 2 Then
        Debug.Print "�i�J�ĤT�h���A���m����B�J"
        
        '�����b�ĤT�h����������ĤT�h���A���m
        If box(X, Y, z, 3) = box(1, 2, 2, 1) Then
            Call secondmethod1_2_1(X, 1, Y, 2)
            term = 3 '����x�b
            
        ElseIf box(X, Y, z, 3) = box(2, 1, 2, 2) Then
            Call secondmethod1_2_1(X, 2, Y, 1)
            term = 4 '����y�b
            
        ElseIf box(X, Y, z, 3) = box(2, 3, 2, 2) Then
            Call secondmethod1_2_1(X, 2, Y, 3)
            term = 4 '����y�b
            
        ElseIf box(X, Y, z, 3) = box(3, 2, 2, 1) Then
            Call secondmethod1_2_1(X, 3, Y, 2)
            term = 3 '����x�b
            
        Else
            Debug.Print "secondmethod1_2�X��"
        End If
       
    End If
End Sub
Private Sub secondmethod1_2_1(x1, x2, y1, y2) '�Ƶ{��----------��secondmethod1_2-�I�s
'�ت��N�b�ĤT�h�������쥿�T����m�W
'�H�Q��m�J�ĤG�h��
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
        Debug.Print "secondmethod1_2_1�X��"
        
    End If
End Sub
Private Sub secondmethod1_3(X, Y, z)
'�A������m�J�ĤG�h
        If X = Y Then
            If term = 3 Then    '����X�b
                Call secondchoose2(box, X, 0, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, X, 0, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, 0, Y, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, 0, Y, 0, 1)
            
            End If
            
            If term = 4 Then    '����Y�b
                Call secondchoose1(box, 0, Y, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, 0, Y, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, X, 0, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, X, 0, 0, 1)
                
            End If
            
        Else
            If term = 3 Then    '����X�b
                Call secondchoose1(box, X, 0, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, X, 0, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, 0, Y, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, 0, Y, 0, 1)
            
            End If
            
            If term = 4 Then    '����Y�b
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
Private Sub secondchoose(fox() As Integer, a, b, c) '�Ƶ{����secondmethod-�t�C�I�s
'�إ߰����`�W���ɰw���
    If a = 1 Or b = 1 Then
        Call choose(fox, a, b, c)
        Call choose(fox, a, b, c)
        Call choose(fox, a, b, c)
        
    ElseIf a = 3 Or b = 3 Then
        Call choose(fox, a, b, c)
          
    Else
        Debug.Print "secondchoose�X�{���~"
    End If
    
End Sub
Private Sub secondchoose1(fox() As Integer, a, b, c, show) '�Ƶ{����secondmethod-�t�C�I�s
'�إ߰����`�W�f�ɰw���
    Dim d As Integer
    d = 3
    Call secondchoose(fox, a, b, c)
    
    If a <> 0 Then
        Debug.Print "X�b", a, "�������ɰw��"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
    End If
    
    If b <> 0 Then
        Debug.Print "Y�b", b, "�������ɰw��"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
    End If
End Sub
Private Sub secondchoose2(fox() As Integer, a, b, c, show) '�Ƶ{����secondmethod-�t�C�I�s
'�إ߰����`�W�f�ɰw���
    Dim d As Integer
    d = 4
    Call secondchoose(fox, a, b, c)
    Call secondchoose(fox, a, b, c)
    Call secondchoose(fox, a, b, c)
    
    If a <> 0 Then
        
        Debug.Print "X�b", a, "�����f�ɰw��"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
    End If
    
    If b <> 0 Then
        Debug.Print "Y�b", b, "�����f�ɰw��"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    End If
End Sub


Private Sub thirdSET() '�]�w�B�n
    '����l��
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
    
    '���Z�b�_��
        'Third1_1����
    Third1_1(1, 2) = W
    Third1_1(2, 1) = W
    Third1_1(2, 2) = W
    
        'Third1_2���u
    Third1_2(1, 2) = W
    Third1_2(2, 2) = W
    Third1_2(3, 2) = W
    
        'Third1_3�����I
    Third1_3(2, 2) = W
    
    
    '����Z�b�_��
        'Third2_1 c1�Mc2
    Third2_1(1, 1) = W
    Third2_1(1, 2) = W
    Third2_1(2, 1) = W
    Third2_1(2, 2) = W
    Third2_1(2, 3) = W
    Third2_1(3, 2) = W
    
        'Third2_2 c3�Mc4
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
    
        'Third2_4 c6�Mc7
    Third2_4(1, 2) = W
    Third2_4(2, 1) = W
    Third2_4(2, 2) = W
    Third2_4(2, 3) = W
    Third2_4(3, 2) = W

End Sub
Private Sub thirdchoose1() '�N�q�μҲ�Third0_1����
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

Private Sub thirdchoose2(i2, j2, i1, j1) '��Thirdchoose1�I�s
'�i��y���ഫ�{��
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
Debug.Print "�I�sthirdmethod1_1"
    Call thirdmethod1_1
Debug.Print "�I�sthirdmethod2_1"
    Call thirdmethod2_1
Debug.Print "�ĤT�h�����_��"
    Call thirdmethod3_1
Debug.Print "�ĤT�h����_��"
    Call thirdmethod4_1
    
    
End Sub
Private Sub thirdmethod1_1() '�N�ĤT�h���z�b��_
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
Private Sub thirdmethod2_1() '�N�ĤT�hz�b���_��
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
Private Sub thirdmethod3_1() '�N�ĤT�h�����k��
    term = 0
    If box(1, 1, 3, 1) = box(1, 3, 3, 1) And box(3, 1, 3, 1) = box(3, 3, 3, 1) Then
        term = 1
    End If
    
    If term <> 1 Then
        If box(1, 1, 3, 1) = box(1, 3, 3, 1) Then '�W
            Call choose2(box, 0, 0, 3, 1)
            Call thirdmethod3_2
            
            Debug.Print "����1���"
        ElseIf box(1, 3, 3, 2) = box(3, 3, 3, 2) Then '�k
            Call choose2(box, 0, 0, 3, 1)
            Call choose2(box, 0, 0, 3, 1)
            Call thirdmethod3_2
            
             Debug.Print "����2���"
        ElseIf box(3, 1, 3, 1) = box(3, 3, 3, 1) Then '�U
            Call choose1(box, 0, 0, 3, 1)
            Call thirdmethod3_2
            
             Debug.Print "����3���"
        ElseIf box(1, 1, 3, 2) = box(3, 1, 3, 2) Then '��
            Call thirdmethod3_2
        Else
            Call thirdmethod3_2
            Call choose2(box, 0, 0, 3, 1)
            Call thirdmethod3_2
        End If
    End If

End Sub
Private Sub thirdmethod4_1() '�N�ĤT�h������k��
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
'��THIRDMETHOD1_3�I�s
'�P�B�n���
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
'��THIRDMETHOD1_1�I�s
'�I�s����k�M��A�X���B�n
Dim i As Integer
Dim i2 As Integer
Dim j As Integer
Dim k As Integer

 
        For i = 0 To 4
Debug.Print "�B���", i, "��"

            Call thirdmethod1_2(Thirdterm, X) '�P�B�n���
            If Thirdterm = 1 Then
Debug.Print "������\"
                     
            End If
            
            '����
           ' For j = 1 To 3
           '     For k = 1 To 3
           '         Debug.Print j, k, x(j, k)
           '     Next k
           ' Next j
            Debug.Print ""
            
            
            '����
            Call thirdCopy(Third0_1, X)
            Call thirdchoose1
            Call thirdCopy(X, Third0_1)
            
            If Thirdterm = 1 Then term = 1
            If Thirdterm = 1 Then Exit For
            Thirdterm = 0
        Next i
        
        Debug.Print "����I", i
        While (i <> 0 And i < 4)
            i = i - 1
            Call choose2(box, 0, 0, 3, 1)
        Wend
        
End Sub
Private Sub thirdmethod1_4() '�ĤT�h����_��z������
    Call secondchoose2(box, 0, 3, 0, 1)
    Call choose2(box, 0, 0, 3, 1)
    Call secondchoose2(box, 3, 0, 0, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
    Call secondchoose1(box, 0, 3, 0, 1)
End Sub
Private Sub thirdmethod2_2(Thirdterm, X() As Integer)
'��THIRDMETHOD1_3�I�s
'�P�B�n���
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
'��THIRDMETHOD1_1�I�s
'�I�s����k�M��A�X���B�n
Dim i As Integer
Dim i2 As Integer
Dim j As Integer
Dim k As Integer

 
        For i = 0 To 4
Debug.Print "�B���", i, "��"

            Call thirdmethod2_2(Thirdterm, X) '�P�B�n���
            If Thirdterm = 1 Then
Debug.Print "������\"
                     
            End If
            
            '����
           ' For j = 1 To 3
           '     For k = 1 To 3
           '         Debug.Print j, k, x(j, k)
           '     Next k
           ' Next j
            Debug.Print ""
            
            
            '����
            Call thirdCopy(Third0_1, X)
            Call thirdchoose1
            Call thirdCopy(X, Third0_1)
            
            If Thirdterm = 1 Then term = 1
            If Thirdterm = 1 Then Exit For
            
            Thirdterm = 0
        Next i
        
        Debug.Print "����I", i
        While (i <> 0 And i < 4)
            i = i - 1
            Call choose2(box, 0, 0, 3, 1)
        Wend
        
End Sub
Private Sub thirdmethod2_4() 'c1�ĤT�h�_��z������
    Call secondchoose2(box, 0, 3, 0, 1)
    Call choose2(box, 0, 0, 3, 1)
    Call secondchoose1(box, 0, 3, 0, 1)
    Call choose2(box, 0, 0, 3, 1)
    Call secondchoose2(box, 0, 3, 0, 1)
    Call choose2(box, 0, 0, 3, 1)
    Call choose2(box, 0, 0, 3, 1)
    Call secondchoose1(box, 0, 3, 0, 1)

End Sub
Private Sub thirdmethod2_5() 'c2�ĤT�h�_��z������
    Call secondchoose1(box, 3, 0, 0, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call secondchoose2(box, 3, 0, 0, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call secondchoose2(box, 3, 0, 0, 1)
End Sub
Private Sub thirdmethod3_2() '�ĤT�h��������
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
    If box(1, 1, 3, 1) = box(1, 2, 3, 1) And box(1, 3, 3, 1) = box(1, 2, 3, 1) Then '�W
            Call thirdmethod4_3
        
        ElseIf box(1, 3, 3, 2) = box(2, 3, 3, 2) And box(3, 3, 3, 2) = box(2, 3, 3, 2) Then '�k
            Call choose2(box, 0, 0, 3, 1)
            Call thirdmethod4_3
        
        ElseIf box(3, 1, 3, 1) = box(3, 2, 3, 1) And box(3, 3, 3, 1) = box(3, 2, 3, 1) Then '�U
            Call choose1(box, 0, 0, 3, 1)
            Call choose1(box, 0, 0, 3, 1)
            Call thirdmethod4_3
    
        ElseIf box(1, 1, 3, 2) = box(2, 1, 3, 2) And box(3, 1, 3, 2) = box(2, 1, 3, 2) Then '��
            Call choose1(box, 0, 0, 3, 1)
            Call thirdmethod4_3
            
        Else
            term = 2
            Call thirdmethod4_3_1
            
        End If
End Sub
Private Sub thirdmethod4_3()
    If box(2, 1, 3, 2) = box(3, 1, 3, 1) Then '�f�ɰwĲ�o
        Call thirdmethod4_3_1
    End If
    
    If box(2, 3, 3, 2) = box(3, 1, 3, 1) Then '���ɰwĲ�o
        Call thirdmethod4_3_2
    End If
    
End Sub
Private Sub thirdmethod4_3_1() '�ĤT�h�f�ɰw�T������
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
Private Sub thirdmethod4_3_2() '�ĤT�h���ɰw�T������
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
    If box(1, 1, 3, 1) <> box(1, 2, 3, 1) Then '�W
        term = 0
    End If
        
    If box(1, 3, 3, 2) <> box(2, 3, 3, 2) Then '�k
        term = 0
            
    End If
        
    If box(3, 1, 3, 1) <> box(3, 2, 3, 1) Then '�U
        term = 0
            
    End If
        
    If box(1, 1, 3, 2) <> box(2, 1, 3, 2) Then  '��
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
'���ե�
    
    Call Openfile(2) '�}���ɮ�
    
    Call showmethod3
    
    
    
End Sub
Private Sub showmethod1()
    Dim a, b, c, d As Integer
    
    Input #1, a, b, c, d 'Ū���@������
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
Private Sub showmethod2() '������l���A
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
Private Sub showmethod3() '�����C�����

    Call showmethod3_1
End Sub
Private Sub showmethod3_1()
    
    'X�e��
    Call showmethod3_2(showBox(3, 1, 3, 1), Shape1(1))
    Call showmethod3_2(showBox(3, 2, 3, 1), Shape1(2))
    Call showmethod3_2(showBox(3, 3, 3, 1), Shape1(3))
    Call showmethod3_2(showBox(3, 1, 2, 1), Shape1(4))
    Call showmethod3_2(showBox(3, 2, 2, 1), Shape1(5))
    Call showmethod3_2(showBox(3, 3, 2, 1), Shape1(6))
    Call showmethod3_2(showBox(3, 1, 1, 1), Shape1(7))
    Call showmethod3_2(showBox(3, 2, 1, 1), Shape1(8))
    Call showmethod3_2(showBox(3, 3, 1, 1), Shape1(9))
    'Y�k��
    Call showmethod3_2(showBox(3, 3, 3, 2), Shape3(1))
    Call showmethod3_2(showBox(2, 3, 3, 2), Shape3(2))
    Call showmethod3_2(showBox(1, 3, 3, 2), Shape3(3))
    Call showmethod3_2(showBox(3, 3, 2, 2), Shape3(4))
    Call showmethod3_2(showBox(2, 3, 2, 2), Shape3(5))
    Call showmethod3_2(showBox(1, 3, 2, 2), Shape3(6))
    Call showmethod3_2(showBox(3, 3, 1, 2), Shape3(7))
    Call showmethod3_2(showBox(2, 3, 1, 2), Shape3(8))
    Call showmethod3_2(showBox(1, 3, 1, 2), Shape3(9))
    'X�᭱
    Call showmethod3_2(showBox(1, 1, 3, 1), Shape6(1))
    Call showmethod3_2(showBox(1, 2, 3, 1), Shape6(2))
    Call showmethod3_2(showBox(1, 3, 3, 1), Shape6(3))
    Call showmethod3_2(showBox(1, 1, 2, 1), Shape6(4))
    Call showmethod3_2(showBox(1, 2, 2, 1), Shape6(5))
    Call showmethod3_2(showBox(1, 3, 2, 1), Shape6(6))
    Call showmethod3_2(showBox(1, 1, 1, 1), Shape6(7))
    Call showmethod3_2(showBox(1, 2, 1, 1), Shape6(8))
    Call showmethod3_2(showBox(1, 3, 1, 1), Shape6(9))
    'Y����
    Call showmethod3_2(showBox(3, 1, 3, 2), Shape4(1))
    Call showmethod3_2(showBox(2, 1, 3, 2), Shape4(2))
    Call showmethod3_2(showBox(1, 1, 3, 2), Shape4(3))
    Call showmethod3_2(showBox(3, 1, 2, 2), Shape4(4))
    Call showmethod3_2(showBox(2, 1, 2, 2), Shape4(5))
    Call showmethod3_2(showBox(1, 1, 2, 2), Shape4(6))
    Call showmethod3_2(showBox(3, 1, 1, 2), Shape4(7))
    Call showmethod3_2(showBox(2, 1, 1, 2), Shape4(8))
    Call showmethod3_2(showBox(1, 1, 1, 2), Shape4(9))
    'Z��
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
Private Sub showmethod4(a, b, c, d) '��ʽb�Y���ܤ������

    If showterm = 0 Then showterm = 1
    Call showmethod4_1
    Call showmethod4_2(a, b, c, d)
    
    If showterm = 1 Then
        showterm = 2
    Else
        showterm = 1
    End If

End Sub
Private Sub showmethod4_1() '�M�ū������
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
Private Sub showmethod4_2(a, b, c, d) '�����ܥ��b�Q���઺�b
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
Private Sub showmethod4_2_1(a, d) 'x�b������ܤ���
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
Private Sub showmethod4_2_2(b, d) 'y�b������ܤ���
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
Private Sub showmethod4_2_3(c, d) 'z�b������ܤ���
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
Private Sub showmethod4_3(block) '������ܽb�Y�C���ഫ
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




'���k�s
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



'�}�Ҥ���J�C���box

Call OpenColorfile(2, 3, 0, 4, 1) '�e��

Call OpenColorfile(4, 1, 4, 4, 1) '�᭱

Call OpenColorfile(1, 0, 1, 4, 2) '����

Call OpenColorfile(3, 4, 3, 4, 2) '�k��

Call OpenColorfile(5, 0, 0, 3, 3) '�W��

Call OpenColorfile(6, 4, 0, 1, 3) '�U��


'-----------------------��l���p
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
      Name            =   "�s�ө���"
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
   StartUpPosition =   3  '�t�ιw�]��
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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Caption         =   "���R�˥��Ҧb��m"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
      Caption         =   "�Ȱ�"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Caption         =   "�}�l"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
      Caption         =   "�����ɮ�"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Caption         =   "�U�@�B"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Caption         =   "�O��������O"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Caption         =   "�P�_�}�l"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      FillStyle       =   0  '���
      Height          =   495
      Index           =   2
      Left            =   240
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   1
      Left            =   240
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H80000000&
      FillStyle       =   0  '���
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
      FillStyle       =   0  '���
      Height          =   495
      Index           =   1
      Left            =   2760
      Shape           =   1  '�����
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   4
      Left            =   2040
      Shape           =   1  '�����
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   7
      Left            =   1320
      Shape           =   1  '�����
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   2
      Left            =   2760
      Shape           =   1  '�����
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   5
      Left            =   2040
      Shape           =   1  '�����
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   8
      Left            =   1320
      Shape           =   1  '�����
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   3
      Left            =   2760
      Shape           =   1  '�����
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   6
      Left            =   2040
      Shape           =   1  '�����
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   9
      Left            =   1320
      Shape           =   1  '�����
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   3
      Left            =   5760
      Shape           =   1  '�����
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   2
      Left            =   5040
      Shape           =   1  '�����
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   1
      Left            =   4320
      Shape           =   1  '�����
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   6
      Left            =   5760
      Shape           =   1  '�����
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   5
      Left            =   5040
      Shape           =   1  '�����
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   4
      Left            =   4320
      Shape           =   1  '�����
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   9
      Left            =   5760
      Shape           =   1  '�����
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   8
      Left            =   5040
      Shape           =   1  '�����
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   7
      Left            =   4320
      Shape           =   1  '�����
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  '���
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
      FillStyle       =   0  '���
      Height          =   495
      Index           =   1
      Left            =   4320
      Shape           =   1  '�����
      Top             =   7320
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   2
      Left            =   5040
      Shape           =   1  '�����
      Top             =   7320
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   3
      Left            =   5760
      Shape           =   1  '�����
      Top             =   7320
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   4
      Left            =   4320
      Shape           =   1  '�����
      Top             =   8040
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   5
      Left            =   5040
      Shape           =   1  '�����
      Top             =   8040
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   6
      Left            =   5760
      Shape           =   1  '�����
      Top             =   8040
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   7
      Left            =   4320
      Shape           =   1  '�����
      Top             =   8760
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   8
      Left            =   5040
      Shape           =   1  '�����
      Top             =   8760
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   9
      Left            =   5760
      Shape           =   1  '�����
      Top             =   8760
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  '���
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
      FillStyle       =   0  '���
      Height          =   495
      Index           =   1
      Left            =   4320
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   2
      Left            =   5040
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   3
      Left            =   5760
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   4
      Left            =   4320
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   5
      Left            =   5040
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   6
      Left            =   5760
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   7
      Left            =   4320
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   8
      Left            =   5040
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   9
      Left            =   5760
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   0
      Left            =   240
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   1
      Left            =   7440
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   2
      Left            =   7440
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   3
      Left            =   7440
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   4
      Left            =   8160
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   5
      Left            =   8160
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   6
      Left            =   8160
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   7
      Left            =   8880
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   8
      Left            =   8880
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   9
      Left            =   8880
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   0
      Left            =   240
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H000080FF&
      FillStyle       =   0  '���
      Height          =   495
      Index           =   0
      Left            =   240
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '���
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



Dim term As Integer '��ܧP�_
Dim showterm As Integer '������ܤ���C���ܴ��P�_

Private Sub Command1_Click() '���{���P�_�ӫ����
Dim i, j, k, n As Integer
Call colorSET

'���k�s
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



'�}�Ҥ���J�C���box

Call OpenColorfile(2, 3, 0, 4, 1) '�e��

Call OpenColorfile(4, 1, 4, 4, 1) '�᭱

Call OpenColorfile(1, 0, 1, 4, 2) '����

Call OpenColorfile(3, 4, 3, 4, 2) '�k��

Call OpenColorfile(5, 0, 0, 3, 3) '�W��

Call OpenColorfile(6, 4, 0, 1, 3) '�U��



'-----------------------��l���p
For k = 1 To 3
    For i = 1 To 3
        For j = 1 To 3
            
                
                Call printcolor(i, j, k)
          
        Next j
    Next i
Next k

Call Openfile(1)
Call showmethod2
'----------------------------------------------�}�l�Ĥ@�h�P�_
    Debug.Print "�Ĥ@�h�������"
    Debug.Print "�P�_121"
    Call firstEdgeBlock(1, 2, 1, E, 0, L)
    
    Debug.Print "�P�_211"
    Call firstEdgeBlock(2, 1, 1, 0, U, L)
    
    Debug.Print "�P�_231"
    Call firstEdgeBlock(2, 3, 1, 0, G, L)
    
    Debug.Print "�P�_321"
    Call firstEdgeBlock(3, 2, 1, R, 0, L)
    
    Debug.Print "�Ĥ@�h��������"
    Debug.Print "�P�_111"
    Call firstCornerBlock(1, 1, 1, E, U, L)
    Call text
    Debug.Print "�P�_131"
    Call firstCornerBlock(1, 3, 1, E, G, L)
    
    
    'Call text
    Debug.Print "�P�_311"
    
    Call firstCornerBlock(3, 1, 1, R, U, L)
    
    Debug.Print "�P�_331"
    Call firstCornerBlock(3, 3, 1, R, G, L)
    
'--------------------�Ĥ@�h�_��H�᪺���p
    
    
    For k = 1 To 3
    For i = 1 To 3
        For j = 1 To 3
            
                
                Call printcolor(i, j, k)
          
        Next j
    Next i
Next k

    


'----------------------------------------�ĤG�h�P�_
    Debug.Print "�ĤG�h����"
    Debug.Print "�P�_112"
    Call secondmethod(1, 1, 2, E, U, 0)
 
    Debug.Print "�P�_132"
    Call secondmethod(1, 3, 2, E, G, 0)
    
    Debug.Print "�P�_312"
    Call secondmethod(3, 1, 2, R, U, 0)
    
    Debug.Print "�P�_332"
    Call secondmethod(3, 3, 2, R, G, 0)
    
'--------------------�ĤG�h�_��H�᪺���p
For k = 1 To 3
    For i = 1 To 3
        For j = 1 To 3
            
                
                Call printcolor(i, j, k)
          
        Next j
    Next i
Next k

'-----------------------------------------�ĤT�h�P�_
    Debug.Print "�ĤT�h����"
    Call thirdmethod
    
 '--------------------�ĤT�h�_��H�᪺���p
For k = 1 To 3
    For i = 1 To 3
        For j = 1 To 3
            
                
                'Call printcolor(i, j, k)
          
        Next j
    Next i
Next k
   
Call Openfile(3)
'����
'Call choose1(0, 0, 3) '���઺��k:�ĤT�hz�b��V����




End Sub
Private Sub printcolor(i, j, k) '�d�ݤ���C��
Dim a As Integer

For a = 1 To 3
    If box(i, j, k, a) = 0 Then
    Debug.Print i, j, k, a, "0"
    End If
    
    If box(i, j, k, a) = 1 Then
    Debug.Print i, j, k, a, "��"
    End If
    
    If box(i, j, k, a) = 2 Then
    Debug.Print i, j, k, a, "��"
    End If
    
    If box(i, j, k, a) = 3 Then
    Debug.Print i, j, k, a, "��"
    End If
    
    If box(i, j, k, a) = 4 Then
    Debug.Print i, j, k, a, "��"
    End If
    
    If box(i, j, k, a) = 5 Then
    Debug.Print i, j, k, a, "��"
    End If
    
    If box(i, j, k, a) = 6 Then
    Debug.Print i, j, k, a, "��"
    End If
    
Next a
Debug.Print ""




End Sub
Private Sub OpenColorfile(NUMBER As Integer, a As Integer, b As Integer, c As Integer, d As Integer)
    '�P�_�}�Ҫ����
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

    If NUMBER = 1 Then '�������
        For i = 1 To 3
            For j = 1 To 3
                Input #1, Color 'Ū���@������
                box(a + i, b, c - j, d) = Color
            Next j
        Next i
        
    ElseIf NUMBER = 3 Then '����k��
        For i = 1 To 3
            For j = 1 To 3
                Input #1, Color 'Ū���@������
                box(a - i, b, c - j, d) = Color
            Next j
        Next i
    
    ElseIf NUMBER = 2 Then '����e��
        For i = 1 To 3
            For j = 1 To 3
                Input #1, Color 'Ū���@������
                box(a, b + i, c - j, d) = Color
            Next j
        Next i
        
    
      
    ElseIf NUMBER = 4 Then '����᭱
        For i = 1 To 3
            For j = 1 To 3
                Input #1, Color 'Ū���@������
                box(a, b - i, c - j, d) = Color
            Next j
        Next i
        
    ElseIf NUMBER = 5 Then '����W��
        For i = 1 To 3
            For j = 1 To 3
                Input #1, Color 'Ū���@������
                box(a + j, b + i, c, d) = Color
            Next j
        Next i
       
    ElseIf NUMBER = 6 Then '����U��
        For i = 1 To 3
            For j = 1 To 3
                Input #1, Color 'Ū���@������
                box(a - j, b + i, c, d) = Color
            Next j
        Next i
        
    Else
    End If
    
End Sub
Private Sub Openfile(X As Integer) '�}��Ū�g���
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
Private Sub colorSET() '�]�w�C��Ѽ�
R = 1 '����
L = 2 '�¦�
G = 3 '���
U = 4 '�Ŧ�
W = 5 '����
E = 6 '���
End Sub
Private Sub choose(fox() As Integer, a, b, c) '�إ߿�ܱ��઺��V�M�h�ƪ���k
    If b = 0 And c = 0 Then
        Call X(a, fox) 'x�b����
    End If

    If a = 0 And c = 0 Then
        Call Y(b, fox) 'y�b����
    End If
    
    If a = 0 And b = 0 Then
        Call z(c, fox) 'z�b����
    End If
End Sub
Private Sub choose1(fox() As Integer, a, b, c, show) '�إ߱��ॿ��V�����
    Dim d As Integer
    d = 1
    
    Call choose(fox, a, b, c)
    If a <> 0 Then
        Debug.Print "X�b", a, "����"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    ElseIf b <> 0 Then
        Debug.Print "Y�b", b, "����"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    Else
        Debug.Print "Z�b", c, "����"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    End If
End Sub
Private Sub choose2(fox() As Integer, a, b, c, show) '�إ߱���Ϥ�V�����
    Dim d As Integer
    d = 2
    
    Call choose(fox, a, b, c)
    Call choose(fox, a, b, c)
    Call choose(fox, a, b, c)
    
    If a <> 0 Then
        Debug.Print "X�b", a, "����"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    ElseIf b <> 0 Then
        Debug.Print "Y�b", b, "����"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    Else
        Debug.Print "Z�b", c, "����"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    End If
End Sub
Private Sub X(a, fox() As Integer) '�u��x�b���઺��k
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

'�b���^box1�s������
For i = 1 To 3
    For j = 1 To 3
        fox(a, i, j, 1) = box2(a, i, j, 1)
        fox(a, i, j, 2) = box2(a, i, j, 2)
        fox(a, i, j, 3) = box2(a, i, j, 3)
                
    Next j
Next i
End Sub
Private Sub Y(b, fox() As Integer) '�u��y�b���઺��k
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

'�b���^box1�s������
For i = 1 To 3
    For j = 1 To 3
        fox(j, b, i, 1) = box2(j, b, i, 1)
        fox(j, b, i, 2) = box2(j, b, i, 2)
        fox(j, b, i, 3) = box2(j, b, i, 3)
                
    Next j
Next i
End Sub
Private Sub z(c, fox() As Integer) '�u��z�b���઺��k
Dim i As Integer
Dim j As Integer

'����box2�x�s����
For i = 1 To 3
    For j = 1 To 3
        Call rotate(i2, j2, i, j) '�e��ӬO��Ӫ��y��
                                  '���ӬO�쥻���y��
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

'����
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
Private Sub rotate(a2, b2, a1, b1) '�إ߱���y���ഫ������
For i = 1 To 3
    For j = 1 To 3
        a2 = b1
        b2 = 4 - a1
    Next j
Next i

End Sub

Private Sub firstEdgeBlock(X, Y, z, cx, cy, cz) '�����Ĥ@�h���
    term = 0 '����P�_�ܼƥ���l��
Debug.Print "�P�_����O�_�b���T��m"
    Call firstmethod(X, Y, z, cx, cy, cz)
    
    If term <> 1 And term <> 5 Then
Debug.Print "�P�_����O�_�b�ĤG�h"
        Call firstmethod3(X, Y, z, 1, 1, 2, cx, cy, cz)
        Call firstmethod3(X, Y, z, 1, 3, 2, cx, cy, cz)
        Call firstmethod3(X, Y, z, 3, 1, 2, cx, cy, cz)
        Call firstmethod3(X, Y, z, 3, 3, 2, cx, cy, cz)
    
    End If
    
    If term <> 1 And term <> 3 Then
Debug.Print "Z�b���T �P�_�O�_���b�Ĥ@�h"
        Call firstmethod1(X, Y, z, 1, 2, 1, cx, cy, cz)
        Call firstmethod1(X, Y, z, 2, 1, 1, cx, cy, cz)
        Call firstmethod1(X, Y, z, 2, 3, 1, cx, cy, cz)
        Call firstmethod1(X, Y, z, 3, 2, 1, cx, cy, cz)
Debug.Print "Z�b���T �P�_�O�_���b�ĤT�h"
        Call firstmethod1(X, Y, z, 1, 2, 3, cx, cy, cz)
        Call firstmethod1(X, Y, z, 2, 1, 3, cx, cy, cz)
        Call firstmethod1(X, Y, z, 2, 3, 3, cx, cy, cz)
        Call firstmethod1(X, Y, z, 3, 2, 3, cx, cy, cz)
    End If
    
    
    If term <> 1 And term <> 3 Then
Debug.Print "Z�b�A�� �P�_�O�_���b�Ĥ@�h"
        Call firstmethod2(X, Y, z, 1, 2, 1, cx, cy, cz)
        Call firstmethod2(X, Y, z, 2, 1, 1, cx, cy, cz)
        Call firstmethod2(X, Y, z, 2, 3, 1, cx, cy, cz)
        Call firstmethod2(X, Y, z, 3, 2, 1, cx, cy, cz)
Debug.Print "Z�b�A�� �P�_�O�_���b�ĤT�h"
        Call firstmethod2(X, Y, z, 1, 2, 3, cx, cy, cz)
        Call firstmethod2(X, Y, z, 2, 1, 3, cx, cy, cz)
        Call firstmethod2(X, Y, z, 2, 3, 3, cx, cy, cz)
        Call firstmethod2(X, Y, z, 3, 2, 3, cx, cy, cz)
    End If
  
End Sub

Private Sub firstCornerBlock(X, Y, z, cx, cy, cz) '�����Ĥ@�h����
    term = 0 '�����l��
Debug.Print "�����O�_�b���T��m"
    Call firstmethod4(X, Y, z, cx, cy, cz)
    
    If term <> 1 And term <> 3 Then
Debug.Print "����z�b�O�_�¦�öi��B�z"
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
Debug.Print "�����O�_z�b�����¦�öi��B�z"
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
    '�P�_�O�_�b���
    Debug.Print "�P�_�O�_�b���"
    
    If box(X, Y, z, 1) = cx And box(X, Y, z, 2) = cy And box(X, Y, z, 3) = cz Then
        term = 1
        Debug.Print "���b���"
    End If
End Sub
Private Sub firstmethod1(x1, y1, z1, x2, y2, z2, cx, cy, cz)
    '�P�_�b�Ĥ@�h�βĤT�h
    '�P�_����z�b��V�������¦�
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
Debug.Print "�i�J����B�zfirstmethod1_1"
        term = 3
        '���ĤT�h�B�z
        If x2 <> 2 And z2 <> 3 Then
            Call secondchoose1(box, x2, 0, 0, 1)
            Call secondchoose1(box, x2, 0, 0, 1)
        End If

        If y2 <> 2 And z2 <> 3 Then
            Call secondchoose1(box, 0, y2, 0, 1)
            Call secondchoose1(box, 0, y2, 0, 1)
        End If
            
        '�b�ĤT�h����
        Call firstmethod2_2(x1, x2, y1, y2)
            
        '��^�Ĥ@�h
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
    '�P�_�b�Ĥ@�h�βĤT�h
    '���Oz�b�C���A��
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
Debug.Print "�i�J����B�zfirstmethod2_1"
        term = 3
        '���ĤT�h�B�z�����m�y��
        If x2 <> 2 And z2 <> 3 Then
            Call secondchoose1(box, x2, 0, 0, 1)
            Call secondchoose1(box, x2, 0, 0, 1)
        End If

        If y2 <> 2 And z2 <> 3 Then
            Call secondchoose1(box, 0, y2, 0, 1)
            Call secondchoose1(box, 0, y2, 0, 1)
        End If
            
        '�b�ĤT�h����
        Call firstmethod2_2(x1, x2, y1, y2)
            
        '��^�Ĥ@�h���T��m�y��
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
Private Sub firstmethod2_2(x1, x2, y1, y2) '�Ƶ{��----------��firstmethod2_1-�I�s
'�ت��N�b�ĤT�h��������Ĥ@�h����m�W��xy�b   �H�Q��m�J�ĤG�h��
'1�����T��m 2�������m
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
        Debug.Print "firstmethod2_2�X��"
            
        
    End If
End Sub
Private Sub firstmethod3(x1, y1, z1, x2, y2, z2, cx, cy, cz)
    '�P�_����b�ĤG�h
    '�y��1�����T��m �y��2�������m
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
'�N�b�ĤG�h��������ĤT�h
Debug.Print "�i�J����B�zfirstmethod3_1"
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
Debug.Print "firstmethod3-1 X2=1 Y2���~"
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
Debug.Print "firstmethod3-1 X2=1 Y2 ���~"
        End If
        
    Else
Debug.Print "firstmethod3-1 X2���~"
    End If
    
End Sub
Private Sub firstmethod4(x1, y1, z1, cx, cy, cz) '����
    '�O�_�b���T����m�W
    If box(x1, y1, z1, 1) = cx And box(x1, y1, z1, 2) = cy And box(x1, y1, z1, 3) = cz Then
        term = 1
    End If
End Sub
Private Sub firstmethod5(x1, y1, z1, x2, y2, z2, cx, cy, cz) '����
'�P�_����Z�b���¦�
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
Private Sub firstmethod5_1(x1, y1, z1, x2, y2, z2, cx, cy, cz) '����
'����Z�b���¦�
'�B�z:���z�b�����¦�
Debug.Print "�i�J����B�zfirstmethod5_1"
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
Debug.Print "firstmethod5_1 z2=1���~"

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
Debug.Print "firstmethod5_1 z2=3���~"

        
        
    End If
    

End Sub
Private Sub firstmethod6(x1, y1, z1, x2, y2, z2, cx, cy, cz) '����
'�P�_����Z�b�����¦�
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
Private Sub FIRSTMETHOD6_1(x1, y1, z1, x2, y2, z2, cx, cy, cz) '����
'�p�G�b�Ĥ@�h�����ĤT�h
Debug.Print "�i�J����B�zfirstmethod6_1"
    term = 4
    If z2 = 1 Then
Debug.Print "�p�G�b�Ĥ@�h�����ĤT�h"
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
    
'�b�ĤT�h�����A���m
Debug.Print "�b�ĤT�h�����A���m"
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
Debug.Print "firstmethod6_1 cx cy���~"
    End If
'�N�����쥿�T����m
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


Private Sub secondmethod(X, Y, z, cx, cy, cz) '�Ƶ{��----------�ѥD�{���I�s
'�ĤG�h��@�����m�P�_

    term = 0 '����P�_�ܼƥ���l��
    
    '�P�_�O�_�b���
    Debug.Print "�P�_�O�_�b���"
    
    If box(X, Y, z, 1) = cx And box(X, Y, z, 2) = cy Then
        term = 1
        Debug.Print "���b���"
    End If
    
    
    '�P�_�O�_�b�ĤG�h���Y�Ӧ�m
    Debug.Print "�P�_�O�_�b�ĤG�h"
    
    If term <> 1 Then
        Debug.Print "�P�_�b�ĤG�h������"
        Call secondmethod1_1(1, 1, 2, cx, cy, cz)
            Debug.Print "�z�Lsecondmethod1-1�P�_112"
        Call secondmethod1_1(1, 3, 2, cx, cy, cz)
            Debug.Print "�z�Lsecondmethod1-1�P�_132"
        Call secondmethod1_1(3, 1, 2, cx, cy, cz)
            Debug.Print "�z�Lsecondmethod1-1�P�_312"
        Call secondmethod1_1(3, 3, 2, cx, cy, cz)
            Debug.Print "�z�Lsecondmethod1-1�P�_332"
        
    End If
    
    '�P�_�O�_�b�ĤT�h���Y�Ӧ�m
    Debug.Print "�P�_�O�_�b�ĤT�h"
    
    If term <> 1 Then
        Debug.Print "�P�_�b�ĤT�h������"
            Debug.Print "�z�Lsecondmethod1-2�P�_123"
        Call secondmethod1_2(1, 2, 3, cx, cy, cz)
            Debug.Print "�z�Lsecondmethod1-2�P�_213"
        Call secondmethod1_2(2, 1, 3, cx, cy, cz)
            Debug.Print "�z�Lsecondmethod1-2�P�_233"
        Call secondmethod1_2(2, 3, 3, cx, cy, cz)
            Debug.Print "�z�Lsecondmethod1-2�P�_323"
        Call secondmethod1_2(3, 2, 3, cx, cy, cz)
            
        '�m�J�ĤG�h
            Debug.Print "�z�Lsecondmethod1-3�N����m�J�ĤG�h"
        Call secondmethod1_3(X, Y, z)
             
        
    
    End If




End Sub
Private Sub secondmethod1_1(X, Y, z, cx, cy, cz) '�Ƶ{��---------��secondmethod�t�C-�I�s
    '�P�_�O�_�b�ĤG�h���Y�Ӧ�m�W
    Debug.Print "�Ƶ{���P�_����b�ĤG�h������"
    
    Dim termsecond1_1 As Integer
    If (box(X, Y, z, 1) = cx And box(X, Y, z, 2) = cy) Or (box(X, Y, z, 1) = cy And box(X, Y, z, 2) = cx) Then
        term = 2
        termsecond1_1 = 2
    End If
    
    '�p�G����m�A���i��B�z
    '�B�z�覡���N������ܲĤT�h
    If termsecond1_1 = 2 Then
        Debug.Print "�ĤG�h���ܲĤT�h������B�J"
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
Private Sub secondmethod1_2(X, Y, z, cx, cy, cz) '�Ƶ{��----------��secondmethod�t�C-�I�s
    '�P�_�O�_�b�ĤT�h���Y�Ӧ�m�W
    Debug.Print "�Ƶ{���P�_����b�ĤT�h������"
    
    Dim termsecond1_2 As Integer
    If (box(X, Y, z, 1) = cx And box(X, Y, z, 3) = cy) Or (box(X, Y, z, 1) = cy And box(X, Y, z, 3) = cx) Then
        termsecond1_2 = 2
    End If
    
    If (box(X, Y, z, 2) = cx And box(X, Y, z, 3) = cy) Or (box(X, Y, z, 2) = cy And box(X, Y, z, 3) = cx) Then
        termsecond1_2 = 2
    End If
    
    
    Debug.Print "�P�_�ܼ�termsecond1_2", termsecond1_2
    
    '�p�G����m(�b�ĤT�h)�A���i��B�z
    If termsecond1_2 = 2 Then
        Debug.Print "�i�J�ĤT�h���A���m����B�J"
        
        '�����b�ĤT�h����������ĤT�h���A���m
        If box(X, Y, z, 3) = box(1, 2, 2, 1) Then
            Call secondmethod1_2_1(X, 1, Y, 2)
            term = 3 '����x�b
            
        ElseIf box(X, Y, z, 3) = box(2, 1, 2, 2) Then
            Call secondmethod1_2_1(X, 2, Y, 1)
            term = 4 '����y�b
            
        ElseIf box(X, Y, z, 3) = box(2, 3, 2, 2) Then
            Call secondmethod1_2_1(X, 2, Y, 3)
            term = 4 '����y�b
            
        ElseIf box(X, Y, z, 3) = box(3, 2, 2, 1) Then
            Call secondmethod1_2_1(X, 3, Y, 2)
            term = 3 '����x�b
            
        Else
            Debug.Print "secondmethod1_2�X��"
        End If
       
    End If
End Sub
Private Sub secondmethod1_2_1(x1, x2, y1, y2) '�Ƶ{��----------��secondmethod1_2-�I�s
'�ت��N�b�ĤT�h�������쥿�T����m�W
'�H�Q��m�J�ĤG�h��
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
        Debug.Print "secondmethod1_2_1�X��"
        
    End If
End Sub
Private Sub secondmethod1_3(X, Y, z)
'�A������m�J�ĤG�h
        If X = Y Then
            If term = 3 Then    '����X�b
                Call secondchoose2(box, X, 0, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, X, 0, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, 0, Y, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, 0, Y, 0, 1)
            
            End If
            
            If term = 4 Then    '����Y�b
                Call secondchoose1(box, 0, Y, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, 0, Y, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, X, 0, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, X, 0, 0, 1)
                
            End If
            
        Else
            If term = 3 Then    '����X�b
                Call secondchoose1(box, X, 0, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, X, 0, 0, 1)
                Call choose2(box, 0, 0, 3, 1)
                Call secondchoose2(box, 0, Y, 0, 1)
                Call choose1(box, 0, 0, 3, 1)
                Call secondchoose1(box, 0, Y, 0, 1)
            
            End If
            
            If term = 4 Then    '����Y�b
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
Private Sub secondchoose(fox() As Integer, a, b, c) '�Ƶ{����secondmethod-�t�C�I�s
'�إ߰����`�W���ɰw���
    If a = 1 Or b = 1 Then
        Call choose(fox, a, b, c)
        Call choose(fox, a, b, c)
        Call choose(fox, a, b, c)
        
    ElseIf a = 3 Or b = 3 Then
        Call choose(fox, a, b, c)
          
    Else
        Debug.Print "secondchoose�X�{���~"
    End If
    
End Sub
Private Sub secondchoose1(fox() As Integer, a, b, c, show) '�Ƶ{����secondmethod-�t�C�I�s
'�إ߰����`�W�f�ɰw���
    Dim d As Integer
    d = 3
    Call secondchoose(fox, a, b, c)
    
    If a <> 0 Then
        Debug.Print "X�b", a, "�������ɰw��"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
    End If
    
    If b <> 0 Then
        Debug.Print "Y�b", b, "�������ɰw��"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
    End If
End Sub
Private Sub secondchoose2(fox() As Integer, a, b, c, show) '�Ƶ{����secondmethod-�t�C�I�s
'�إ߰����`�W�f�ɰw���
    Dim d As Integer
    d = 4
    Call secondchoose(fox, a, b, c)
    Call secondchoose(fox, a, b, c)
    Call secondchoose(fox, a, b, c)
    
    If a <> 0 Then
        
        Debug.Print "X�b", a, "�����f�ɰw��"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
    End If
    
    If b <> 0 Then
        Debug.Print "Y�b", b, "�����f�ɰw��"
        If show = 1 Then
            Write #1, a, b, c, d
        End If
        
    End If
End Sub


Private Sub thirdSET() '�]�w�B�n
    '����l��
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
    
    '���Z�b�_��
        'Third1_1����
    Third1_1(1, 2) = W
    Third1_1(2, 1) = W
    Third1_1(2, 2) = W
    
        'Third1_2���u
    Third1_2(1, 2) = W
    Third1_2(2, 2) = W
    Third1_2(3, 2) = W
    
        'Third1_3�����I
    Third1_3(2, 2) = W
    
    
    '����Z�b�_��
        'Third2_1 c1�Mc2
    Third2_1(1, 1) = W
    Third2_1(1, 2) = W
    Third2_1(2, 1) = W
    Third2_1(2, 2) = W
    Third2_1(2, 3) = W
    Third2_1(3, 2) = W
    
        'Third2_2 c3�Mc4
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
    
        'Third2_4 c6�Mc7
    Third2_4(1, 2) = W
    Third2_4(2, 1) = W
    Third2_4(2, 2) = W
    Third2_4(2, 3) = W
    Third2_4(3, 2) = W

End Sub
Private Sub thirdchoose1() '�N�q�μҲ�Third0_1����
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

Private Sub thirdchoose2(i2, j2, i1, j1) '��Thirdchoose1�I�s
'�i��y���ഫ�{��
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
Debug.Print "�I�sthirdmethod1_1"
    Call thirdmethod1_1
Debug.Print "�I�sthirdmethod2_1"
    Call thirdmethod2_1
Debug.Print "�ĤT�h�����_��"
    Call thirdmethod3_1
Debug.Print "�ĤT�h����_��"
    Call thirdmethod4_1
    
    
End Sub
Private Sub thirdmethod1_1() '�N�ĤT�h���z�b��_
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
Private Sub thirdmethod2_1() '�N�ĤT�hz�b���_��
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
Private Sub thirdmethod3_1() '�N�ĤT�h�����k��
    term = 0
    If box(1, 1, 3, 1) = box(1, 3, 3, 1) And box(3, 1, 3, 1) = box(3, 3, 3, 1) Then
        term = 1
    End If
    
    If term <> 1 Then
        If box(1, 1, 3, 1) = box(1, 3, 3, 1) Then '�W
            Call choose2(box, 0, 0, 3, 1)
            Call thirdmethod3_2
            
            Debug.Print "����1���"
        ElseIf box(1, 3, 3, 2) = box(3, 3, 3, 2) Then '�k
            Call choose2(box, 0, 0, 3, 1)
            Call choose2(box, 0, 0, 3, 1)
            Call thirdmethod3_2
            
             Debug.Print "����2���"
        ElseIf box(3, 1, 3, 1) = box(3, 3, 3, 1) Then '�U
            Call choose1(box, 0, 0, 3, 1)
            Call thirdmethod3_2
            
             Debug.Print "����3���"
        ElseIf box(1, 1, 3, 2) = box(3, 1, 3, 2) Then '��
            Call thirdmethod3_2
        Else
            Call thirdmethod3_2
            Call choose2(box, 0, 0, 3, 1)
            Call thirdmethod3_2
        End If
    End If

End Sub
Private Sub thirdmethod4_1() '�N�ĤT�h������k��
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
'��THIRDMETHOD1_3�I�s
'�P�B�n���
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
'��THIRDMETHOD1_1�I�s
'�I�s����k�M��A�X���B�n
Dim i As Integer
Dim i2 As Integer
Dim j As Integer
Dim k As Integer

 
        For i = 0 To 4
Debug.Print "�B���", i, "��"

            Call thirdmethod1_2(Thirdterm, X) '�P�B�n���
            If Thirdterm = 1 Then
Debug.Print "������\"
                     
            End If
            
            '����
           ' For j = 1 To 3
           '     For k = 1 To 3
           '         Debug.Print j, k, x(j, k)
           '     Next k
           ' Next j
            Debug.Print ""
            
            
            '����
            Call thirdCopy(Third0_1, X)
            Call thirdchoose1
            Call thirdCopy(X, Third0_1)
            
            If Thirdterm = 1 Then term = 1
            If Thirdterm = 1 Then Exit For
            Thirdterm = 0
        Next i
        
        Debug.Print "����I", i
        While (i <> 0 And i < 4)
            i = i - 1
            Call choose2(box, 0, 0, 3, 1)
        Wend
        
End Sub
Private Sub thirdmethod1_4() '�ĤT�h����_��z������
    Call secondchoose2(box, 0, 3, 0, 1)
    Call choose2(box, 0, 0, 3, 1)
    Call secondchoose2(box, 3, 0, 0, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
    Call secondchoose1(box, 0, 3, 0, 1)
End Sub
Private Sub thirdmethod2_2(Thirdterm, X() As Integer)
'��THIRDMETHOD1_3�I�s
'�P�B�n���
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
'��THIRDMETHOD1_1�I�s
'�I�s����k�M��A�X���B�n
Dim i As Integer
Dim i2 As Integer
Dim j As Integer
Dim k As Integer

 
        For i = 0 To 4
Debug.Print "�B���", i, "��"

            Call thirdmethod2_2(Thirdterm, X) '�P�B�n���
            If Thirdterm = 1 Then
Debug.Print "������\"
                     
            End If
            
            '����
           ' For j = 1 To 3
           '     For k = 1 To 3
           '         Debug.Print j, k, x(j, k)
           '     Next k
           ' Next j
            Debug.Print ""
            
            
            '����
            Call thirdCopy(Third0_1, X)
            Call thirdchoose1
            Call thirdCopy(X, Third0_1)
            
            If Thirdterm = 1 Then term = 1
            If Thirdterm = 1 Then Exit For
            
            Thirdterm = 0
        Next i
        
        Debug.Print "����I", i
        While (i <> 0 And i < 4)
            i = i - 1
            Call choose2(box, 0, 0, 3, 1)
        Wend
        
End Sub
Private Sub thirdmethod2_4() 'c1�ĤT�h�_��z������
    Call secondchoose2(box, 0, 3, 0, 1)
    Call choose2(box, 0, 0, 3, 1)
    Call secondchoose1(box, 0, 3, 0, 1)
    Call choose2(box, 0, 0, 3, 1)
    Call secondchoose2(box, 0, 3, 0, 1)
    Call choose2(box, 0, 0, 3, 1)
    Call choose2(box, 0, 0, 3, 1)
    Call secondchoose1(box, 0, 3, 0, 1)

End Sub
Private Sub thirdmethod2_5() 'c2�ĤT�h�_��z������
    Call secondchoose1(box, 3, 0, 0, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call secondchoose2(box, 3, 0, 0, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call secondchoose1(box, 3, 0, 0, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call choose1(box, 0, 0, 3, 1)
    Call secondchoose2(box, 3, 0, 0, 1)
End Sub
Private Sub thirdmethod3_2() '�ĤT�h��������
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
    If box(1, 1, 3, 1) = box(1, 2, 3, 1) And box(1, 3, 3, 1) = box(1, 2, 3, 1) Then '�W
            Call thirdmethod4_3
        
        ElseIf box(1, 3, 3, 2) = box(2, 3, 3, 2) And box(3, 3, 3, 2) = box(2, 3, 3, 2) Then '�k
            Call choose2(box, 0, 0, 3, 1)
            Call thirdmethod4_3
        
        ElseIf box(3, 1, 3, 1) = box(3, 2, 3, 1) And box(3, 3, 3, 1) = box(3, 2, 3, 1) Then '�U
            Call choose1(box, 0, 0, 3, 1)
            Call choose1(box, 0, 0, 3, 1)
            Call thirdmethod4_3
    
        ElseIf box(1, 1, 3, 2) = box(2, 1, 3, 2) And box(3, 1, 3, 2) = box(2, 1, 3, 2) Then '��
            Call choose1(box, 0, 0, 3, 1)
            Call thirdmethod4_3
            
        Else
            term = 2
            Call thirdmethod4_3_1
            
        End If
End Sub
Private Sub thirdmethod4_3()
    If box(2, 1, 3, 2) = box(3, 1, 3, 1) Then '�f�ɰwĲ�o
        Call thirdmethod4_3_1
    End If
    
    If box(2, 3, 3, 2) = box(3, 1, 3, 1) Then '���ɰwĲ�o
        Call thirdmethod4_3_2
    End If
    
End Sub
Private Sub thirdmethod4_3_1() '�ĤT�h�f�ɰw�T������
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
Private Sub thirdmethod4_3_2() '�ĤT�h���ɰw�T������
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
    If box(1, 1, 3, 1) <> box(1, 2, 3, 1) Then '�W
        term = 0
    End If
        
    If box(1, 3, 3, 2) <> box(2, 3, 3, 2) Then '�k
        term = 0
            
    End If
        
    If box(3, 1, 3, 1) <> box(3, 2, 3, 1) Then '�U
        term = 0
            
    End If
        
    If box(1, 1, 3, 2) <> box(2, 1, 3, 2) Then  '��
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
'���ե�
    
    Call Openfile(2) '�}���ɮ�
    
    Call showmethod3
    
    
    
End Sub
Private Sub showmethod1()
    Dim a, b, c, d As Integer
    
    Input #1, a, b, c, d 'Ū���@������
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
Private Sub showmethod2() '������l���A
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
Private Sub showmethod3() '�����C�����

    Call showmethod3_1
End Sub
Private Sub showmethod3_1()
    
    'X�e��
    Call showmethod3_2(showBox(3, 1, 3, 1), Shape1(1))
    Call showmethod3_2(showBox(3, 2, 3, 1), Shape1(2))
    Call showmethod3_2(showBox(3, 3, 3, 1), Shape1(3))
    Call showmethod3_2(showBox(3, 1, 2, 1), Shape1(4))
    Call showmethod3_2(showBox(3, 2, 2, 1), Shape1(5))
    Call showmethod3_2(showBox(3, 3, 2, 1), Shape1(6))
    Call showmethod3_2(showBox(3, 1, 1, 1), Shape1(7))
    Call showmethod3_2(showBox(3, 2, 1, 1), Shape1(8))
    Call showmethod3_2(showBox(3, 3, 1, 1), Shape1(9))
    'Y�k��
    Call showmethod3_2(showBox(3, 3, 3, 2), Shape3(1))
    Call showmethod3_2(showBox(2, 3, 3, 2), Shape3(2))
    Call showmethod3_2(showBox(1, 3, 3, 2), Shape3(3))
    Call showmethod3_2(showBox(3, 3, 2, 2), Shape3(4))
    Call showmethod3_2(showBox(2, 3, 2, 2), Shape3(5))
    Call showmethod3_2(showBox(1, 3, 2, 2), Shape3(6))
    Call showmethod3_2(showBox(3, 3, 1, 2), Shape3(7))
    Call showmethod3_2(showBox(2, 3, 1, 2), Shape3(8))
    Call showmethod3_2(showBox(1, 3, 1, 2), Shape3(9))
    'X�᭱
    Call showmethod3_2(showBox(1, 1, 3, 1), Shape6(1))
    Call showmethod3_2(showBox(1, 2, 3, 1), Shape6(2))
    Call showmethod3_2(showBox(1, 3, 3, 1), Shape6(3))
    Call showmethod3_2(showBox(1, 1, 2, 1), Shape6(4))
    Call showmethod3_2(showBox(1, 2, 2, 1), Shape6(5))
    Call showmethod3_2(showBox(1, 3, 2, 1), Shape6(6))
    Call showmethod3_2(showBox(1, 1, 1, 1), Shape6(7))
    Call showmethod3_2(showBox(1, 2, 1, 1), Shape6(8))
    Call showmethod3_2(showBox(1, 3, 1, 1), Shape6(9))
    'Y����
    Call showmethod3_2(showBox(3, 1, 3, 2), Shape4(1))
    Call showmethod3_2(showBox(2, 1, 3, 2), Shape4(2))
    Call showmethod3_2(showBox(1, 1, 3, 2), Shape4(3))
    Call showmethod3_2(showBox(3, 1, 2, 2), Shape4(4))
    Call showmethod3_2(showBox(2, 1, 2, 2), Shape4(5))
    Call showmethod3_2(showBox(1, 1, 2, 2), Shape4(6))
    Call showmethod3_2(showBox(3, 1, 1, 2), Shape4(7))
    Call showmethod3_2(showBox(2, 1, 1, 2), Shape4(8))
    Call showmethod3_2(showBox(1, 1, 1, 2), Shape4(9))
    'Z��
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
Private Sub showmethod4(a, b, c, d) '��ʽb�Y���ܤ������

    If showterm = 0 Then showterm = 1
    Call showmethod4_1
    Call showmethod4_2(a, b, c, d)
    
    If showterm = 1 Then
        showterm = 2
    Else
        showterm = 1
    End If

End Sub
Private Sub showmethod4_1() '�M�ū������
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
Private Sub showmethod4_2(a, b, c, d) '�����ܥ��b�Q���઺�b
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
Private Sub showmethod4_2_1(a, d) 'x�b������ܤ���
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
Private Sub showmethod4_2_2(b, d) 'y�b������ܤ���
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
Private Sub showmethod4_2_3(c, d) 'z�b������ܤ���
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
Private Sub showmethod4_3(block) '������ܽb�Y�C���ഫ
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




'���k�s
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



'�}�Ҥ���J�C���box

Call OpenColorfile(2, 3, 0, 4, 1) '�e��

Call OpenColorfile(4, 1, 4, 4, 1) '�᭱

Call OpenColorfile(1, 0, 1, 4, 2) '����

Call OpenColorfile(3, 4, 3, 4, 2) '�k��

Call OpenColorfile(5, 0, 0, 3, 3) '�W��

Call OpenColorfile(6, 4, 0, 1, 3) '�U��


'-----------------------��l���p
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
