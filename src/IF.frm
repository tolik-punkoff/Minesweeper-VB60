VERSION 5.00
Begin VB.Form InvisItems 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Smile 
      Height          =   495
      Left            =   2040
      Picture         =   "IF.frx":0000
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   16
      Top             =   2640
      Width           =   615
   End
   Begin VB.PictureBox None 
      Height          =   495
      Left            =   2400
      ScaleHeight     =   435
      ScaleWidth      =   195
      TabIndex        =   15
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Mine 
      Height          =   255
      Left            =   960
      Picture         =   "IF.frx":0282
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   14
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox flag 
      Height          =   255
      Left            =   720
      Picture         =   "IF.frx":0596
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   13
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox errors 
      Height          =   255
      Left            =   360
      Picture         =   "IF.frx":08AA
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   12
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox mineboom 
      Height          =   375
      Left            =   0
      Picture         =   "IF.frx":0C1E
      ScaleHeight     =   315
      ScaleWidth      =   195
      TabIndex        =   11
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox Numers 
      Height          =   375
      Index           =   8
      Left            =   1920
      Picture         =   "IF.frx":0F92
      ScaleHeight     =   315
      ScaleWidth      =   195
      TabIndex        =   10
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Numers 
      Height          =   375
      Index           =   7
      Left            =   1680
      Picture         =   "IF.frx":126A
      ScaleHeight     =   315
      ScaleWidth      =   195
      TabIndex        =   9
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Numers 
      Height          =   375
      Index           =   6
      Left            =   1440
      Picture         =   "IF.frx":1506
      ScaleHeight     =   315
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Numers 
      Height          =   375
      Index           =   5
      Left            =   1200
      Picture         =   "IF.frx":181A
      ScaleHeight     =   315
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Numers 
      Height          =   375
      Index           =   4
      Left            =   960
      Picture         =   "IF.frx":1B5E
      ScaleHeight     =   315
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Numers 
      Height          =   375
      Index           =   3
      Left            =   720
      Picture         =   "IF.frx":1E72
      ScaleHeight     =   315
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Numers 
      Height          =   375
      Index           =   2
      Left            =   480
      Picture         =   "IF.frx":21E6
      ScaleHeight     =   315
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Numers 
      Height          =   375
      Index           =   1
      Left            =   0
      Picture         =   "IF.frx":252A
      ScaleHeight     =   315
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox O 
      Height          =   495
      Left            =   1320
      Picture         =   "IF.frx":282E
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   2640
      Width           =   495
   End
   Begin VB.PictureBox Porajen 
      Height          =   615
      Left            =   600
      Picture         =   "IF.frx":2AB0
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   2520
      Width           =   615
   End
   Begin VB.PictureBox Pobeda 
      Height          =   495
      Left            =   0
      Picture         =   "IF.frx":2D32
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   2640
      Width           =   495
   End
End
Attribute VB_Name = "InvisItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

