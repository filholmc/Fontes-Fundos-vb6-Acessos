VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formConAFu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fundos acessados por um Usuário"
   ClientHeight    =   5190
   ClientLeft      =   2760
   ClientTop       =   1755
   ClientWidth     =   8745
   Icon            =   "ConAFu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   8745
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
      Width           =   8520
      Begin VB.CommandButton fcmbF09LCA 
         Caption         =   "F9"
         Height          =   255
         Left            =   4530
         TabIndex        =   6
         Top             =   180
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton fcmbF11Tab 
         Caption         =   "F11"
         Height          =   255
         Left            =   4980
         TabIndex        =   5
         Top             =   180
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton fcmbF12Hom 
         Caption         =   "F12"
         Height          =   255
         Left            =   5430
         TabIndex        =   4
         Top             =   180
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton fcmbF05Con 
         Caption         =   "Consultar (F5)"
         Default         =   -1  'True
         Height          =   315
         Left            =   3600
         TabIndex        =   1
         Top             =   450
         Width           =   1120
      End
      Begin VB.ComboBox fcboUsuari 
         Height          =   315
         ItemData        =   "ConAFu.frx":08CA
         Left            =   220
         List            =   "ConAFu.frx":08CC
         TabIndex        =   0
         Top             =   450
         Width           =   3300
      End
      Begin VB.CommandButton fcmbEscape 
         Caption         =   "Fechar (Esc)"
         Height          =   315
         Left            =   4740
         TabIndex        =   2
         Top             =   450
         Width           =   1120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuário"
         Height          =   195
         Left            =   220
         TabIndex        =   7
         Top             =   240
         Width           =   540
      End
   End
   Begin MSComctlLib.ListView flsvFunUsu 
      Height          =   3825
      Left            =   120
      TabIndex        =   9
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
Attribute VB_Name = "formConAFu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private fbytColInd As Byte

Private fbytColOrd As Byte

Private fbytNumFun As Byte

Private fbytOrdChv As Byte

Private fclsAceFun As clssAceFun

Private fintNumUsu As Integer

Private fstrSmbOrd As String
Private Sub fcboUsuari_Click()
        If (fcboUsuari.ListIndex = -1) Then Exit Sub

        fintNumUsu = CInt(Mid(fcboUsuari, 1, InStr(fcboUsuari, " ") - 1))
End Sub
Private Sub fcmbEscape_Click()
        Unload Me
End Sub
Private Sub fcmbF05Con_Click()
        If (fcboUsuari.ListIndex = -1) Then
            rgfMsgBox "Escolha uma opção do campo 'Usuário'", MsgErr
            fcboUsuari.SetFocus
            Exit Sub
        End If

        rlsConsultar
End Sub
Private Sub fcmbF08Man_Click()
        If (formMDIAce.menuAceFun.Enabled And fbytNumFun > 0) Then
            gbytConFun = 1
            gbooConFun = True
            gintNumUsu = fintNumUsu
            gbytNumFun = fbytNumFun
            formAceFun.SetFocus
        End If
End Sub
Private Sub fcmbF09LCA_Click()
        If (TypeOf ActiveControl Is ComboBox) Then ActiveControl.Text = ""
End Sub
Private Sub fcmbF11Tab_Click()
        SendKeys "+{TAB}"
End Sub
Private Sub fcmbF12Hom_Click()
        fcboUsuari.SetFocus
End Sub
Private Sub flsvFunUsu_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        gintTempoP = 0
        fbytColInd = ColumnHeader.Index
        fstrSmbOrd = IIf(fbytOrdChv = 1, "/\", "\/")
        fbytOrdChv = IIf(fbytOrdChv = 1, 0, 1)
        rlsMontarColunasNomes
End Sub
Private Sub flsvFunUsu_DblClick()
        fcmbF08Man_Click
End Sub
Private Sub flsvFunUsu_ItemClick(ByVal Item As MSComctlLib.ListItem)
        gintTempoP = 0
        fbytNumFun = Item.SubItems(1)
End Sub
Private Sub flsvFunUsu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

        With flsvFunUsu
            .Top = 1245
            .Left = 90
            .Width = (Me.Width) - 270
            .Height = Me.Height - 540 - 1200
            .SortKey = 0
            .SortOrder = fbytOrdChv
         End With

        rgsCentralizarForm Me
        rgsPosicionarAjuda Me, gintForAtu, gbooForLog

        Set fclsAceFun = New clssAceFun

        rgsCarregarUsuarios fcboUsuari, False
End Sub
Private Sub Form_Unload(Cancel As Integer)
        Set fclsAceFun = Nothing
End Sub
Private Sub rlsConsultar()
        Dim lLsILinhas As MSComctlLib.ListItem

        Dim lRStAceFun As Recordset

        Set lRStAceFun = fclsAceFun.ConsultarFundosDeUmUsuario(fintNumUsu)

            flsvFunUsu.ListItems.Clear

        If (lRStAceFun.EOF) Then
            rgfMsgBox "Não há Fundos acessados por este Usuário", MsgInf
            flsvFunUsu.ToolTipText = ""
            lRStAceFun.Close
            fbytNumFun = 0
            Exit Sub
        End If

        If (formMDIAce.menuAceFun.Enabled) Then
            flsvFunUsu.ToolTipText = "F8 ou Duplo Clique para Manipular o Registro Selecionado"
        End If

            rlsMontarColunas
            fbytNumFun = lRStAceFun!Numero
        Do _
            While (Not ((lRStAceFun.EOF)))
        Set lLsILinhas = flsvFunUsu.ListItems.Add(, , lRStAceFun!Codigo)
       With lLsILinhas
           .SubItems(1) = lRStAceFun!Numero
           .SubItems(2) = Format(lRStAceFun!Numero, "000")
        End With
            lRStAceFun.MoveNext
        Loop
        flsvFunUsu.SetFocus
        lRStAceFun.Close
End Sub
Private Sub rlsMontarColunas()
       With flsvFunUsu.ColumnHeaders
           .Clear
           .Add 1, , "", IIf(gintTotRes < 1792, 10350, 13710), 0
           .Add 2, , "", 1200, 1
           .Add 3, , "", 0
        End With
        rlsMontarColunasNomes
End Sub
Private Sub rlsMontarColunasNomes()
            fbytColOrd = fbytColInd - 1

        If (fbytColOrd = 1) Then
            fbytColOrd = 2
        End If

       With flsvFunUsu
           .SortKey = ((fbytColOrd))
           .SortOrder = fbytOrdChv

            Select Case fbytColInd
                   Case 1
                       .ColumnHeaders.Item(1).Text = "Código  " & fstrSmbOrd
                       .ColumnHeaders.Item(2).Text = "Número"
                   Case 2
                       .ColumnHeaders.Item(1).Text = "Código"
                       .ColumnHeaders.Item(2).Text = fstrSmbOrd & "  Número"
            End Select
        End With
End Sub
