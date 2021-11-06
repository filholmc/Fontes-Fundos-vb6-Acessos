VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formConFun 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relação de Fundos"
   ClientHeight    =   5190
   ClientLeft      =   2775
   ClientTop       =   1755
   ClientWidth     =   8700
   Icon            =   "ConFun.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   8700
   Begin VB.CommandButton fcmbF05Con 
      Caption         =   "Consultar (F5)"
      Height          =   315
      Left            =   150
      TabIndex        =   2
      Top             =   4740
      Visible         =   0   'False
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF08Man 
      Caption         =   "Manipular (F8)"
      Height          =   315
      Left            =   1290
      TabIndex        =   1
      Top             =   4740
      Visible         =   0   'False
      Width           =   1120
   End
   Begin VB.CommandButton fcmbEscape 
      Caption         =   "Fechar (Esc)"
      Height          =   315
      Left            =   2430
      TabIndex        =   0
      Top             =   4740
      Visible         =   0   'False
      Width           =   1120
   End
   Begin MSComctlLib.ListView flsvFundos 
      Height          =   4965
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   8758
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
Attribute VB_Name = "formConFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private fbytColInd As Byte

Private fbytColOrd As Byte

Private fbytNumFun As Byte

Private fbytOrdChv As Byte

Private fstrSmbOrd As String
Private Sub fcmbEscape_Click()
        Unload Me
End Sub
Private Sub fcmbF05Con_Click()
        rlsConsultar
End Sub
Private Sub fcmbF08Man_Click()
        If (formMDIAce.menuFundos.Enabled And fbytNumFun > 0) Then
            gbooConFun = True
            gbytNumFun = fbytNumFun
            formFundos.SetFocus
        End If
End Sub
Private Sub flsvFundos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        gintTempoP = 0
        fbytColInd = ColumnHeader.Index
        fstrSmbOrd = IIf(fbytOrdChv = 1, "/\", "\/")
        fbytOrdChv = IIf(fbytOrdChv = 1, 0, 1)
        rlsMontarColunasNomes
End Sub
Private Sub flsvFundos_DblClick()
        fcmbF08Man_Click
End Sub
Private Sub flsvFundos_ItemClick(ByVal Item As MSComctlLib.ListItem)
        gintTempoP = 0
        fbytNumFun = Item.SubItems(1)
End Sub
Private Sub flsvFundos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

        With flsvFundos
            .Top = 90
            .Left = 90
            .Width = (Me.Width) - 270
            .Height = Me.Height - 540
            .SortKey = 0
            .SortOrder = fbytOrdChv
         End With

        rgsCentralizarForm Me
        rgsPosicionarAjuda Me, gintForAtu, gbooForLog

        rlsConsultar
End Sub
Private Sub rlsConsultar()
        Dim lLsILinhas As MSComctlLib.ListItem

        Dim lRStFundos As Recordset

        Set lRStFundos = gclsFundos.ConsultarTodos

            flsvFundos.ListItems.Clear

        If (lRStFundos.EOF) Then
            rgfMsgBox "Não há Fundos cadastrados", MsgInf
            flsvFundos.ToolTipText = ""
            lRStFundos.Close
            fbytNumFun = 0
            Exit Sub
        End If

        If (formMDIAce.menuFundos.Enabled) Then
            flsvFundos.ToolTipText = "F8 ou Duplo Clique para Manipular o Registro Selecionado"
        End If

            rlsMontarColunas
            fbytNumFun = lRStFundos!Numero
        Do _
            While (Not ((lRStFundos.EOF)))
        Set lLsILinhas = flsvFundos.ListItems.Add(, , lRStFundos!Codigo)
       With lLsILinhas
           .SubItems(1) = lRStFundos!Numero
           .SubItems(2) = Format(lRStFundos!Numero, "000")
        End With
            lRStFundos.MoveNext
        Loop
        lRStFundos.Close
End Sub
Private Sub rlsMontarColunas()
       With flsvFundos.ColumnHeaders
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

       With flsvFundos
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
