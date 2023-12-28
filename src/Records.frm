VERSION 5.00
Begin VB.Form Records 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Рекорды"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Очистить"
      Height          =   375
      Left            =   3960
      TabIndex        =   14
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Список рекордсменов"
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.Label Label6 
         Caption         =   "999"
         Height          =   255
         Index           =   5
         Left            =   3840
         TabIndex        =   12
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "999"
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   11
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "999"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   10
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Время"
         Height          =   255
         Left            =   3840
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Неизвестный"
         Height          =   255
         Index           =   4
         Left            =   1680
         TabIndex        =   8
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Неизвестный"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   7
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Неизвестный"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   6
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Имя"
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Крутой"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Бывалый"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Чайник"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Уровень"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Records"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
DeleteRecord
Form_Load
End Sub

Private Sub Form_Load()
GetRecords
Label6(0) = RecArr(0).NameR
Label6(1) = Val(RecArr(0).Time)
Label6(2) = RecArr(1).NameR
Label6(3) = Val(RecArr(1).Time)
Label6(4) = RecArr(2).NameR
Label6(5) = Val(RecArr(2).Time)
End Sub
