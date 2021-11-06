VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formConFor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relação de Forms de um Módulo"
   ClientHeight    =   5190
   ClientLeft      =   2790
   ClientTop       =   1755
   ClientWidth     =   8715
   Icon            =   "ConFor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   8715
   Begin VB.CommandButton fcmbF08Man 
      Caption         =   "Manipular (F8)"
      Height          =   315
      Left            =   150
      TabIndex        =   8
      Top             =   4740
      Visible         =   0   'False
      Width           =   1120
   End
   Begin VB.Frame frmeFormes 
      Caption         =   "  Parâmetros: "
      Height          =   1125
      Left            =   120
      TabIndex        =   3
      Top             =   30
      Width           =   8490
      Begin VB.CommandButton fcmbF09LCA 
         Caption         =   "F9"
         Height          =   255
         Left            =   3030
         TabIndex        =   7
         Top             =   180
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton fcmbF11Tab 
         Caption         =   "F11"
         Height          =   255
         Left            =   3480
         TabIndex        =   6
         Top             =   180
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton fcmbF12Hom 
         Caption         =   "F12"
         Height          =   255
         Left            =   3930
         TabIndex        =   5
         Top             =   180
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton fcmbF05Con 
         Caption         =   "Consultar (F5)"
         Default         =   -1  'True
         Height          =   315
         Left            =   2100
         TabIndex        =   1
         Top             =   450
         Width           =   1120
      End
      Begin VB.CommandButton fcmbEscape 
         Caption         =   "Fechar (Esc)"
         Height          =   315
         Left            =   3240
         TabIndex        =   2
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
      Begin VB.Label flblModulo 
         AutoSize        =   -1  'True
         Caption         =   "Módulo"
         Height          =   195
         Left            =   220
         TabIndex        =   4
         Top             =   240
         Width           =   525
      End
   End
   Begin MSComctlLib.ListView flsvFormes 
      Height          =   3825
      Left            =   120
      TabIndex        =   9
      Top             =   1260
      Width           =   8475
      _ExtentX        =   14949
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
Attribute VB_Name = "formConFor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private fbytColInd As Byte

Private fbytColOrd As Byte

Private fbytNumMod As Byte

Private fbytOrdChv As Byte

Private fintNumFor As Integer

Private fstrSmbOrd As String
Private Sub fcboModulo_Click()
        If (fcboModulo.ListIndex = -1) Then Exit Sub

        fbytNumMod = CByte(Mid(fcboModulo, 1, InStr(fcboModulo, " ") - 1))
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

        rlsConsultar
End Sub
Private Sub fcmbF08Man_Click()
        If (formMDIAce.menuFormes.Enabled And fintNumFor > 0) Then
            gbooConFor = True
            gintNumFor = fintNumFor
            formFormes.SetFocus
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
Private Sub flsvFormes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        gintTempoP = 0
        fbytColInd = ColumnHeader.Index
        fstrSmbOrd = IIf(fbytOrdChv = 1, "/\", "\/")
        fbytOrdChv = IIf(fbytOrdChv = 1, 0, 1)
        rlsMontarColunasNomes
End Sub
Private Sub flsvFormes_DblClick()
        fcmbF08Man_Click
End Sub
Private Sub flsvFormes_ItemClick(ByVal Item As MSComctlLib.ListItem)
        gintTempoP = 0
        fintNumFor = Item.SubItems(4)
End Sub
Private Sub flsvFormes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

        With flsvFormes
            .Top = 1245
            .Left = 90
            .Width = (Me.Width) - 270
            .Height = Me.Height - 540 - 1200
            .SortKey = 0
            .SortOrder = fbytOrdChv
         End With

        rgsCentralizarForm Me
        rgsPosicionarAjuda Me, gintForAtu, gbooForLog

        rgsCarregarModulos fcboModulo
End Sub
Private Sub rlsConsultar()
        Dim lLsILinhas As MSComctlLib.ListItem

        Dim lRStFormes As Recordset

        Set lRStFormes = gclsFormes.ConsultarFormsDeUmModulo(fbytNumMod)

            flsvFormes.ListItems.Clear

        If (lRStFormes.EOF) Then
            rgfMsgBox "Não há Forms cadastrados para este Módulo", MsgInf
            flsvFormes.ToolTipText = ""
            lRStFormes.Close
            fintNumFor = 0
            Exit Sub
        End If

        If (formMDIAce.menuFormes.Enabled) Then
            flsvFormes.ToolTipText = "F8 ou Duplo Clique para Manipular o Registro Selecionado"
        End If

        If (gintTotRes < 1792) Then
            gbytScrBar = IIf(lRStFormes.RecordCount < 24, 0, 240)
        Else
            gbytScrBar = IIf(lRStFormes.RecordCount < 36, 0, 240)
        End If

            rlsMontarColunas
            fintNumFor = lRStFormes!Numero
        Do _
            While (Not ((lRStFormes.EOF)))
        Set lLsILinhas = flsvFormes.ListItems.Add(, , lRStFormes!NomFor)
       With lLsILinhas
           .SubItems(1) = lRStFormes!Descri
           .SubItems(2) = IIf(lRStFormes!SemAce, "Livre", "")
           .SubItems(3) = IIf(lRStFormes!TemLog, "Sim", "")
           .SubItems(4) = lRStFormes!Numero
           .SubItems(5) = Format(lRStFormes!Numero, "0000")
           .SubItems(6) = lRStFormes!NumAju
           .SubItems(7) = Format(lRStFormes!NumAju, "0000")
        End With
            lRStFormes.MoveNext
        Loop
        flsvFormes.SetFocus
        lRStFormes.Close
End Sub
Private Sub rlsMontarColunas()
       With flsvFormes.ColumnHeaders
           .Clear
           .Add 1, , "", 1500, 0
           .Add 2, , "", IIf(gintTotRes < 1792, 7140, 9700) - gbytScrBar, 0
           .Add 3, , "", 1000, 0
           .Add 4, , "", 800, 0
           .Add 5, , "", 1000, 1
           .Add 6, , "", 0
           .Add 7, , "", 900, 1
           .Add 8, , "", 0
        End With
        rlsMontarColunasNomes
End Sub
Private Sub rlsMontarColunasNomes()
            fbytColOrd = fbytColInd - 1

        If (fbytColOrd > 3) Then
            fbytColOrd = fbytColOrd + 1
        End If

       With flsvFormes
           .SortKey = ((fbytColOrd))
           .SortOrder = fbytOrdChv

            With .ColumnHeaders
                 .Item(1).Text = "Nome"
                 .Item(2).Text = "Descrição"
                 .Item(3).Text = "Acesso"
                 .Item(4).Text = "Log"
                 .Item(5).Text = "Número"
                 .Item(7).Text = "Ajuda"

      Select Case fbytColInd
             Case 1
                 .Item(1).Text = "Nome  " & fstrSmbOrd
             Case 2
                 .Item(2).Text = "Descrição  " & fstrSmbOrd
             Case 3
                 .Item(3).Text = "Acesso  " & fstrSmbOrd
             Case 4
                 .Item(4).Text = "Log  " & fstrSmbOrd
             Case 5
                 .Item(5).Text = fstrSmbOrd & "  Número"
             Case 7
                 .Item(7).Text = fstrSmbOrd & "  Ajuda"
        End Select
        End With
        End With
End Sub
