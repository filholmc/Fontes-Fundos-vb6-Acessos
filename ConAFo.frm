VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formConAFo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Forms de um M�dulo acessados por um Usu�rio"
   ClientHeight    =   5190
   ClientLeft      =   2790
   ClientTop       =   1755
   ClientWidth     =   8700
   Icon            =   "ConAFo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   8700
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
      Caption         =   "  Par�metros: "
      Height          =   1125
      Left            =   120
      TabIndex        =   4
      Top             =   30
      Width           =   8490
      Begin VB.CommandButton fcmbF09LCA 
         Caption         =   "F9"
         Height          =   255
         Left            =   6420
         TabIndex        =   7
         Top             =   210
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton fcmbF11Tab 
         Caption         =   "F11"
         Height          =   255
         Left            =   6870
         TabIndex        =   6
         Top             =   210
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton fcmbF12Hom 
         Caption         =   "F12"
         Height          =   255
         Left            =   7320
         TabIndex        =   5
         Top             =   210
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton fcmbF05Con 
         Caption         =   "Consultar (F5)"
         Default         =   -1  'True
         Height          =   315
         Left            =   5490
         TabIndex        =   2
         Top             =   480
         Width           =   1120
      End
      Begin VB.ComboBox fcboModulo 
         Height          =   315
         Left            =   3600
         TabIndex        =   1
         Top             =   480
         Width           =   1800
      End
      Begin VB.ComboBox fcboUsuari 
         Height          =   315
         ItemData        =   "ConAFo.frx":08CA
         Left            =   220
         List            =   "ConAFo.frx":08CC
         TabIndex        =   0
         Top             =   480
         Width           =   3300
      End
      Begin VB.CommandButton fcmbEscape 
         Caption         =   "Fechar (Esc)"
         Height          =   315
         Left            =   6630
         TabIndex        =   3
         Top             =   480
         Width           =   1120
      End
      Begin VB.Label flblModulo 
         AutoSize        =   -1  'True
         Caption         =   "M�dulo"
         Height          =   195
         Left            =   3600
         TabIndex        =   9
         Top             =   270
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usu�rio"
         Height          =   195
         Left            =   220
         TabIndex        =   8
         Top             =   270
         Width           =   540
      End
   End
   Begin MSComctlLib.ListView flsvForUsu 
      Height          =   3825
      Left            =   120
      TabIndex        =   11
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
Attribute VB_Name = "formConAFo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private fbytColInd As Byte

Private fbytColOrd As Byte

Private fbytNumMod As Byte

Private fbytOrdChv As Byte

Private fclsAceFor As clssAceFor

Private fintNumFor As Integer

Private fintNumUsu As Integer

Private fstrSmbOrd As String
Private Sub fcboModulo_Click()
        If (fcboModulo.ListIndex = -1) Then Exit Sub

        fbytNumMod = CByte(Mid(fcboModulo, 1, InStr(fcboModulo, " ") - 1))
End Sub
Private Sub fcboUsuari_Click()
        If (fcboUsuari.ListIndex = -1) Then Exit Sub

        fintNumUsu = CInt(Mid(fcboUsuari, 1, InStr(fcboUsuari, " ") - 1))

        rgsCarregarModulosDeUmUsuario fintNumUsu, fcboModulo
End Sub
Private Sub fcmbEscape_Click()
        Unload Me
End Sub
Private Sub fcmbF05Con_Click()
        If (fcboUsuari.ListIndex = -1) Then
            rgfMsgBox "Escolha uma op��o do campo 'Usu�rio'", MsgErr
            fcboUsuari.SetFocus
            Exit Sub
        End If

        If (fcboModulo.ListIndex = -1) Then
            rgfMsgBox "Escolha uma op��o do campo 'M�dulo'", MsgErr
            fcboModulo.SetFocus
            Exit Sub
        End If

        rlsConsultar
End Sub
Private Sub fcmbF08Man_Click()
        If (formMDIAce.menuAceFor.Enabled And fintNumFor > 0) Then
            gbytConFor = 1
            gbooConFor = True
            gintNumUsu = fintNumUsu
            gbytNumMod = fbytNumMod
            gintNumFor = fintNumFor
            formAceFor.SetFocus
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
Private Sub flsvForUsu_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        gintTempoP = 0
        fbytColInd = ColumnHeader.Index
        fstrSmbOrd = IIf(fbytOrdChv = 1, "/\", "\/")
        fbytOrdChv = IIf(fbytOrdChv = 1, 0, 1)
        rlsMontarColunasNomes
End Sub
Private Sub flsvForUsu_DblClick()
        fcmbF08Man_Click
End Sub
Private Sub flsvForUsu_ItemClick(ByVal Item As MSComctlLib.ListItem)
        gintTempoP = 0
        fintNumFor = Item.SubItems(2)
End Sub
Private Sub flsvForUsu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

        With flsvForUsu
            .Top = 1245
            .Left = 90
            .Width = (Me.Width) - 270
            .Height = Me.Height - 540 - 1200
            .SortKey = 0
            .SortOrder = fbytOrdChv
         End With

        rgsCentralizarForm Me
        rgsPosicionarAjuda Me, gintForAtu, gbooForLog

        Set fclsAceFor = New clssAceFor

        rgsCarregarUsuarios fcboUsuari, False
End Sub
Private Sub Form_Unload(Cancel As Integer)
        Set fclsAceFor = Nothing
End Sub
Private Sub rlsConsultar()
        Dim lLsILinhas As MSComctlLib.ListItem

        Dim lRStAceFor As Recordset

        Set lRStAceFor = fclsAceFor.ConsultarFormsDeUmUsuarioPorModulo(fintNumUsu, fbytNumMod)

            flsvForUsu.ListItems.Clear

        If (lRStAceFor.EOF) Then
            rgfMsgBox "N�o h� Forms acessados por este Usu�rio neste M�dulo", MsgInf
            flsvForUsu.ToolTipText = ""
            lRStAceFor.Close
            fintNumFor = 0
            Exit Sub
        End If

        If (formMDIAce.menuAceFor.Enabled) Then
            flsvForUsu.ToolTipText = "F8 ou Duplo Clique para Manipular o Registro Selecionado"
        End If

        If (gintTotRes < 1792) Then
            gbytScrBar = IIf(lRStAceFor.RecordCount < 24, 0, 240)
        Else
            gbytScrBar = IIf(lRStAceFor.RecordCount < 36, 0, 240)
        End If

            rlsMontarColunas
            fintNumFor = lRStAceFor!Numero
        Do _
            While (Not ((lRStAceFor.EOF)))
        Set lLsILinhas = flsvForUsu.ListItems.Add(, , lRStAceFor!NomFor)
       With lLsILinhas
           .SubItems(1) = lRStAceFor!Descri
           .SubItems(2) = lRStAceFor!Numero
           .SubItems(3) = Format(lRStAceFor!Numero, "0000")
        End With
            lRStAceFor.MoveNext
        Loop
        flsvForUsu.SetFocus
        lRStAceFor.Close
End Sub
Private Sub rlsMontarColunas()
       With flsvForUsu.ColumnHeaders
           .Clear
           .Add 1, , "", 1500, 0
           .Add 2, , "", IIf(gintTotRes < 1792, 9050, 12400) - gbytScrBar, 0
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

       With flsvForUsu
           .SortKey = ((fbytColOrd))
           .SortOrder = fbytOrdChv

            With .ColumnHeaders
                 .Item(1).Text = "Nome"
                 .Item(2).Text = "Descri��o"
                 .Item(3).Text = "N�mero"

      Select Case fbytColInd
             Case 1
                 .Item(1).Text = "Nome  " & fstrSmbOrd
             Case 2
                 .Item(2).Text = "Descri��o  " & fstrSmbOrd
             Case 3
                 .Item(3).Text = fstrSmbOrd & "  N�mero"
        End Select
        End With
        End With
End Sub
