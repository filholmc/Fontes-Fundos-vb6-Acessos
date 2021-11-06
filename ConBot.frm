VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formConBot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relação de Botões de um Form"
   ClientHeight    =   5190
   ClientLeft      =   2760
   ClientTop       =   1755
   ClientWidth     =   8730
   Icon            =   "ConBot.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   8730
   Begin VB.CommandButton fcmbF08Man 
      Caption         =   "Manipular (F8)"
      Height          =   315
      Left            =   150
      TabIndex        =   10
      Top             =   4740
      Visible         =   0   'False
      Width           =   1120
   End
   Begin VB.Frame frmeFormes 
      Caption         =   "  Parâmetros: "
      Height          =   1125
      Left            =   120
      TabIndex        =   4
      Top             =   30
      Width           =   8520
      Begin VB.CommandButton fcmbF09LCA 
         Caption         =   "F9"
         Height          =   255
         Left            =   4920
         TabIndex        =   8
         Top             =   180
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton fcmbF11Tab 
         Caption         =   "F11"
         Height          =   255
         Left            =   5370
         TabIndex        =   7
         Top             =   180
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton fcmbF12Hom 
         Caption         =   "F12"
         Height          =   255
         Left            =   5820
         TabIndex        =   6
         Top             =   180
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton fcmbF05Con 
         Caption         =   "Consultar (F5)"
         Default         =   -1  'True
         Height          =   315
         Left            =   3990
         TabIndex        =   2
         Top             =   450
         Width           =   1120
      End
      Begin VB.ComboBox fcboFormes 
         Height          =   315
         ItemData        =   "ConBot.frx":08CA
         Left            =   2100
         List            =   "ConBot.frx":08CC
         TabIndex        =   1
         Top             =   450
         Width           =   1800
      End
      Begin VB.CommandButton fcmbEscape 
         Caption         =   "Fechar (Esc)"
         Height          =   315
         Left            =   5130
         TabIndex        =   3
         Top             =   450
         Width           =   1120
      End
      Begin VB.ComboBox fcboModulo 
         Height          =   315
         Left            =   220
         TabIndex        =   0
         Top             =   450
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Form"
         Height          =   195
         Left            =   2100
         TabIndex        =   9
         Top             =   240
         Width           =   345
      End
      Begin VB.Label flblModulo 
         AutoSize        =   -1  'True
         Caption         =   "Módulo"
         Height          =   195
         Left            =   220
         TabIndex        =   5
         Top             =   240
         Width           =   525
      End
   End
   Begin MSComctlLib.ListView flsvBotoes 
      Height          =   3825
      Left            =   120
      TabIndex        =   11
      Top             =   1260
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   6747
      View            =   3
      Arrange         =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "formConBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private fbytColInd As Byte

Private fbytColOrd As Byte

Private fbytNumMod As Byte

Private fbytOrdChv As Byte

Private fclsBotoes As clssBotoes

Private fintNumBot As Integer

Private fintNumFor As Integer

Private fstrSmbOrd As String
Private Sub fcboFormes_Click()
        If (fcboFormes.ListIndex = -1) Then Exit Sub

        fintNumFor = CInt(Mid(fcboFormes, 14, Len(fcboFormes) - 13))
End Sub
Private Sub fcboModulo_Click()
        If (fcboModulo.ListIndex = -1) Then Exit Sub

        fbytNumMod = CByte(Mid(fcboModulo, 1, InStr(fcboModulo, " ") - 1))

        rgsCarregarFormesDeUmModuloComBotoes fbytNumMod, fcboFormes
End Sub
Private Sub fcmbEscape_Click()
        Unload Me
End Sub
Private Sub fcmbF05Con_Click()
        If (fcboModulo.ListIndex = -1) Then
            rgfMsgBox "Escolha uma opção do campo 'Módulo'", MsgErr
            fcboModulo.SetFocus
            Exit Sub
        End If

        If (fcboFormes.ListIndex = -1) Then
            rgfMsgBox "Escolha uma opção do campo 'Form'", MsgErr
            fcboFormes.SetFocus
            Exit Sub
        End If

        rlsConsultar
End Sub
Private Sub fcmbF08Man_Click()
        If (formMDIAce.menuBotoes.Enabled And fintNumBot > 0) Then
            gbooConBot = True
            gintNumBot = fintNumBot
            formBotoes.SetFocus
        End If
End Sub
Private Sub fcmbF09LCA_Click()
        If (TypeOf ActiveControl Is ComboBox) Then ActiveControl.Text = ""
End Sub
Private Sub fcmbF11Tab_Click()
        SendKeys "+{TAB}"
End Sub
Private Sub fcmbF12Hom_Click()
        fcboModulo.SetFocus
End Sub
Private Sub flsvBotoes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        gintTempoP = 0
        fbytColInd = ColumnHeader.Index
        fstrSmbOrd = IIf(fbytOrdChv = 1, "/\", "\/")
        fbytOrdChv = IIf(fbytOrdChv = 1, 0, 1)
        rlsMontarColunasNomes
End Sub
Private Sub flsvBotoes_DblClick()
        fcmbF08Man_Click
End Sub
Private Sub flsvBotoes_ItemClick(ByVal Item As MSComctlLib.ListItem)
        gintTempoP = 0
        fintNumBot = Item.SubItems(2)
End Sub
Private Sub flsvBotoes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        gintTempoP = 0
End Sub
Private Sub Form_Activate()
        gintTempoP = 0
        formMDIAce.fsbrModAce.Panels(4).Picture = LoadPicture()
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        gintTempoP = 0
        rgsTratarFuncoes KeyCode, Me
End Sub
Private Sub Form_Load()
        fbytColInd = 1
        fbytOrdChv = 0
        fstrSmbOrd = "/\"

        Me.Width = (Screen.Width) - 120
        Me.Height = Screen.Height - 360 - 1715

        With frmeFormes
            .Top = 90
            .Left = 90
            .Width = Me.Width - 270
            .Height = 1065
         End With

        With flsvBotoes
            .Top = 1245
            .Left = 90
            .Width = (Me.Width) - 270
            .Height = Me.Height - 540 - 1200
            .SortKey = 0
            .SortOrder = fbytOrdChv
         End With

        rgsCentralizarForm Me
        rgsPosicionarAjuda Me, gintForAtu, gbooForLog

        Set fclsBotoes = New clssBotoes

        rgsCarregarModulos fcboModulo
End Sub
Private Sub Form_Unload(Cancel As Integer)
        Set fclsBotoes = Nothing
End Sub
Private Sub rlsConsultar()
        Dim lLsILinhas As MSComctlLib.ListItem

        Dim lRStBotoes As Recordset

        Set lRStBotoes = fclsBotoes.ConsultarBotoesDeUmForm(fbytNumMod, fintNumFor)

            flsvBotoes.ListItems.Clear

        If (lRStBotoes.EOF) Then
            rgfMsgBox "Não há Botões cadastrados para este Form", MsgInf
            flsvBotoes.ToolTipText = ""
            lRStBotoes.Close
            fintNumBot = 0
            Exit Sub
        End If

        If (formMDIAce.menuBotoes.Enabled) Then
            flsvBotoes.ToolTipText = "F8 ou Duplo Clique para Manipular o Registro Selecionado"
        End If

            rlsMontarColunas
            fintNumBot = lRStBotoes!Numero
        Do _
            While (Not ((lRStBotoes.EOF)))
        Set lLsILinhas = flsvBotoes.ListItems.Add(, , lRStBotoes!NomBot)
       With lLsILinhas
           .SubItems(1) = lRStBotoes!Descri
           .SubItems(2) = lRStBotoes!Numero
           .SubItems(3) = Format(lRStBotoes!Numero, "0000")
        End With
            lRStBotoes.MoveNext
        Loop
        flsvBotoes.SetFocus
        lRStBotoes.Close
End Sub
Private Sub rlsMontarColunas()
       With flsvBotoes.ColumnHeaders
           .Clear
           .Add 1, , "", 1300, 0
           .Add 2, , "", IIf(gintTotRes < 1792, 9240, 12600), 0
           .Add 3, , "", 1000, 1
           .Add 4, , "", 0
        End With
        rlsMontarColunasNomes
End Sub
Private Sub rlsMontarColunasNomes()
            fbytColOrd = fbytColInd - 1

        If (fbytColOrd = 2) Then
            fbytColOrd = 3
        End If

       With flsvBotoes
           .SortKey = ((fbytColOrd))
           .SortOrder = fbytOrdChv

            With .ColumnHeaders
                 .Item(1).Text = "Nome"
                 .Item(2).Text = "Descrição"
                 .Item(3).Text = "Número"

      Select Case fbytColInd
             Case 1
                 .Item(1).Text = "Nome  " & fstrSmbOrd
             Case 2
                 .Item(2).Text = "Descrição  " & fstrSmbOrd
             Case 3
                 .Item(3).Text = fstrSmbOrd & "  Número"
        End Select
        End With
        End With
End Sub
