VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "О программе"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   8280
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   3840
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   120
      Picture         =   "frmAbout.frx":0442
      ScaleHeight     =   3435
      ScaleWidth      =   3915
      TabIndex        =   7
      Top             =   120
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      Caption         =   " О программе "
      Height          =   2175
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.Label Label7 
         Caption         =   "Разработано на оборудовании ООО ""Превед"""
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   3615
      End
      Begin VB.Label Label6 
         Caption         =   "Оперативная информация"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label Label5 
         Caption         =   "Internet: http://cityk.onego.ru"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Специально для центризбкрома РФ"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Разработчик: Kuzia _DSL"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Версия: 1.666"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Саперег"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me 'выгрузка объекта формы
End Sub
