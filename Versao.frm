VERSION 5.00
Begin VB.Form formVersao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sobre o Módulo Acessos"
   ClientHeight    =   3330
   ClientLeft      =   3630
   ClientTop       =   1755
   ClientWidth     =   6930
   ControlBox      =   0   'False
   Icon            =   "Versao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "Versao.frx":000C
   ScaleHeight     =   3330
   ScaleWidth      =   6930
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   510
      TabIndex        =   4
      Top             =   2760
      Width           =   5955
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   510
      TabIndex        =   3
      Top             =   450
      Width           =   5955
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   570
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "Versao.frx":0CD6
      Top             =   1440
      Width           =   5835
   End
   Begin VB.CommandButton fcmbEscape 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   5280
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2910
      Width           =   1120
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   570
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Versao.frx":0DAF
      Top             =   600
      Width           =   5835
   End
End
Attribute VB_Name = "formVersao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Private Sub fcmbEscape_Click()
        Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        gintTempoP = 0
        rgsTratarFuncoes KeyCode, Me
End Sub
Private Sub Form_Load()
        gintTempoP = 0
        rgsCentralizarForm Me
        rgsPosicionarAjuda Me, gintForAtu, gbooForLog
        formMDIAce.fsbrModAce.Panels(4).Picture = LoadPicture()
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        gintTempoP = 0
End Sub
