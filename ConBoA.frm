VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formConBoA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuários que acessam um Botão"
   ClientHeight    =   5190
   ClientLeft      =   2760
   ClientTop       =   1755
   ClientWidth     =   8730
   Icon            =   "ConBoA.frx":0000
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
      TabIndex        =   12
      Top             =   4740
      Visible         =   0   'False
      Width           =   1120
   End
   Begin VB.Frame frmeFormes 
      Caption         =   "  Parâmetros: "
      Height          =   1125
      Left            =   120
      TabIndex        =   5
      Top             =   30
      Width           =   8520
      Begin VB.CommandButton fcmbF09LCA 
         Caption         =   "F9"
         Height          =   255
         Left            =   6810
         TabIndex        =   8
         Top             =   180
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton fcmbF11Tab 
         Caption         =   "F11"
         Height          =   255
         Left            =   7260
         TabIndex        =   7
         Top             =   180
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton fcmbF12Hom 
         Caption         =   "F12"
         Height          =   255
         Left            =   7710
         TabIndex        =   6
         Top             =   180
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton fcmbF05Con 
         Caption         =   "Consultar (F5)"
         Default         =   -1  'True
         Height          =   315
         Left            =   5880
         TabIndex        =   3
         Top             =   450
         Width           =   1120
      End
      Begin VB.ComboBox fcboBotoes 
         Height          =   315
         ItemData        =   "ConBoA.frx":08CA
         Left            =   3990
         List            =   "ConBoA.frx":08CC
         TabIndex        =   2
         Top             =   450
         Width           =   1800
      End
      Begin VB.ComboBox fcboFormes 
         Height          =   315
         ItemData        =   "ConBoA.frx":08CE
         Left            =   2100
         List            =   "ConBoA.frx":08D0
         TabIndex        =   1
         Top             =   450
         Width           =   1800
      End
      Begin VB.ComboBox fcboModulo 
         Height          =   315
         ItemData        =   "ConBoA.frx":08D2
         Left            =   220
         List            =   "ConBoA.frx":08D4
         TabIndex        =   0
         Top             =   450
         Width           =   1800
      End
      Begin VB.CommandButton fcmbEscape 
         Caption         =   "Fechar (Esc)"
         Height          =   315
         Left            =   7020
         TabIndex        =   4
         Top             =   450
         Width           =   1120
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Botão"
         Height          =   195
         Left            =   3990
         TabIndex        =   11
         Top             =   240
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Form"
         Height          =   195
         Left            =   2100
         TabIndex        =   10
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Módulo"
         Height          =   195
         Left            =   220
         TabIndex        =   9
         Top             =   240
         Width           =   525
      End
   End
   Begin MSComctlLib.ListView flsvUsuBot 
      Height          =   3795
      Left            =   120
      TabIndex        =   13
      Top             =   1290
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   6694
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
Attribute VB_Name = "formConBoA"
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

Private fintNumUsu As Integer

Private fstrSmbOrd As String
Private Sub fcboBotoes_Click()
        If (fcboBotoes.ListIndex = -1) Then Exit Sub

        fintNumBot = CInt(Mid(fcboBotoes, 14, Len(fcboBotoes) - 13))
End Sub
Private Sub fcboFormes_Click()
        If (fcboFormes.ListIndex = -1) Then Exit Sub

        fintNumFor = CInt(Mid(fcboFormes, 14, Len(fcboFormes) - 13))

        rlsCarregarBotoes
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

        If (fcboBotoes.ListIndex = -1) Then
            rgfMsgBox "Escolha uma opção do campo 'Botão'", MsgErr
            fcboBotoes.SetFocus
            Exit Sub
        End If

        rlsConsultar
End Sub
Private Sub fcmbF08Man_Click()
        If (formMDIAce.menuAceBot.Enabled And fintNumUsu > 0) Then
            gbytConBot = 2
            gbooConBot = True
            gintNumUsu = fintNumUsu
            gbytNumMod = fbytNumMod
            gintNumFor = fintNumFor
            gintNumBot = fintNumBot
            formAceBot.SetFocus
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
Private Sub flsvUsuBot_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        gintTempoP = 0
        fbytColInd = ColumnHeader.Index
        fstrSmbOrd = IIf(fbytOrdChv = 1, "/\", "\/")
        fbytOrdChv = IIf(fbytOrdChv = 1, 0, 1)
        rlsMontarColunasNomes
End Sub
Private Sub flsvUsuBot_DblClick()
        fcmbF08Man_Click
End Sub
Private Sub flsvUsuBot_ItemClick(ByVal Item As MSComctlLib.ListItem)
        gintTempoP = 0
        fintNumUsu = Item.SubItems(1)
End Sub
Private Sub flsvUsuBot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

        With flsvUsuBot
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
Private Sub rlsCarregarBotoes()
        Dim lRStBotoes As Recordset

        Set lRStBotoes = fclsBotoes.ConsultarBotoesDeUmForm(fbytNumMod, fintNumFor)

            fcboBotoes.Clear
        Do _
            While (Not lRStBotoes.EOF)
            fcboBotoes.AddItem lRStBotoes!NomBot & " - " & lRStBotoes!Numero
            lRStBotoes.MoveNext
        Loop
        lRStBotoes.Close
End Sub
Private Sub rlsConsultar()
        Dim lLsILinhas As MSComctlLib.ListItem

        Dim lRStAceBot As Recordset

        Set lRStAceBot = gclsAceBot.ConsultarUsuariosDeUmBotao(fintNumBot)

            flsvUsuBot.ListItems.Clear

        If (lRStAceBot.EOF) Then
            rgfMsgBox "Não há Usuários com acesso a este Botão", MsgInf
            flsvUsuBot.ToolTipText = ""
            lRStAceBot.Close
            fintNumUsu = 0
            Exit Sub
        End If

        If (formMDIAce.menuAceBot.Enabled) Then
            flsvUsuBot.ToolTipText = "F8 ou Duplo Clique para Manipular o Registro Selecionado"
        End If

            rlsMontarColunas
            fintNumUsu = lRStAceBot!Numero
        Do _
            While (Not ((lRStAceBot.EOF)))
        Set lLsILinhas = flsvUsuBot.ListItems.Add(, , lRStAceBot!NomUsu)
       With lLsILinhas
           .SubItems(1) = lRStAceBot!Numero
           .SubItems(2) = Format(lRStAceBot!Numero, "0000")
           .SubItems(3) = lRStAceBot!CodAge
           .SubItems(4) = Format(lRStAceBot!CodAge, "0000")
        End With
            lRStAceBot.MoveNext
        Loop
        flsvUsuBot.SetFocus
        lRStAceBot.Close
End Sub
Private Sub rlsMontarColunas()
       With flsvUsuBot.ColumnHeaders
           .Clear
           .Add 1, , "", IIf(gintTotRes < 1792, 9150, 12510), 0
           .Add 2, , "", 1200, 1
           .Add 3, , "", 0
           .Add 4, , "", 1200, 1
           .Add 5, , "", 0
        End With
        rlsMontarColunasNomes
End Sub
Private Sub rlsMontarColunasNomes()
            fbytColOrd = fbytColInd - 1

        If (fbytColOrd > 0) Then
            fbytColOrd = fbytColOrd + 1
        End If

       With flsvUsuBot
           .SortKey = ((fbytColOrd))
           .SortOrder = fbytOrdChv

            With .ColumnHeaders
                 .Item(1).Text = "Nome"
                 .Item(2).Text = "Número"
                 .Item(4).Text = "Agência"

      Select Case fbytColInd
             Case 1
                 .Item(1).Text = "Nome  " & fstrSmbOrd
             Case 2
                 .Item(2).Text = fstrSmbOrd & "  Número"
             Case 4
                 .Item(4).Text = fstrSmbOrd & "  Agência"
        End Select
        End With
        End With
End Sub
