VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formConUsu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relação de Usuários"
   ClientHeight    =   5190
   ClientLeft      =   2790
   ClientTop       =   1755
   ClientWidth     =   8700
   Icon            =   "ConUsu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   8700
   Begin VB.CommandButton fcmbEscape 
      Caption         =   "Fechar (Esc)"
      Height          =   315
      Left            =   2430
      TabIndex        =   2
      Top             =   4740
      Visible         =   0   'False
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF05Con 
      Caption         =   "Consultar (F5)"
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Top             =   4740
      Visible         =   0   'False
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF08Man 
      Caption         =   "Manipular (F8)"
      Height          =   315
      Left            =   1290
      TabIndex        =   0
      Top             =   4740
      Visible         =   0   'False
      Width           =   1120
   End
   Begin MSComctlLib.ListView flsvUsuari 
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
Attribute VB_Name = "formConUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private fbytColInd As Byte

Private fbytColOrd As Byte

Private fbytOrdChv As Byte

Private fintNumUsu As Integer

Private fstrSmbOrd As String
Private Sub fcmbEscape_Click()
        Unload Me
End Sub
Private Sub fcmbF05Con_Click()
        rlsConsultar
End Sub
Private Sub fcmbF08Man_Click()
        If (formMDIAce.menuUsuari.Enabled And fintNumUsu > 0) Then
            gbooConUsu = True
            gintNumUsu = fintNumUsu
            formUsuari.SetFocus
        End If
End Sub
Private Sub flsvUsuari_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        gintTempoP = 0
        fbytColInd = ColumnHeader.Index
        fstrSmbOrd = IIf(fbytOrdChv = 1, "/\", "\/")
        fbytOrdChv = IIf(fbytOrdChv = 1, 0, 1)
        rlsMontarColunasNomes
End Sub
Private Sub flsvUsuari_DblClick()
        fcmbF08Man_Click
End Sub
Private Sub flsvUsuari_ItemClick(ByVal Item As MSComctlLib.ListItem)
        gintTempoP = 0
        fintNumUsu = Item.SubItems(1)
End Sub
Private Sub flsvUsuari_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
        fbytColInd = 2
        fbytOrdChv = 0
        fstrSmbOrd = "/\"

        formMDIAce.ftbrModAce.Buttons("RelUsu").Value = tbrPressed

        Me.Width = (Screen.Width) - 120
        Me.Height = Screen.Height - 360 - 1715

        With flsvUsuari
            .Top = 90
            .Left = 90
            .Width = Me.Width - 270
            .Height = Me.Height - 540
            .SortKey = 0
            .SortOrder = fbytOrdChv
         End With

        rgsCentralizarForm Me
        rgsPosicionarAjuda Me, gintForAtu, gbooForLog

        rlsConsultar
End Sub
Private Sub Form_Unload(Cancel As Integer)
        formMDIAce.ftbrModAce.Buttons("RelUsu").Value = tbrUnpressed
End Sub
Private Sub rlsConsultar()
        Dim lLsILinhas As MSComctlLib.ListItem

        Dim lRStUsuari As Recordset

        Set lRStUsuari = gclsUsuari.ConsultarTodosPorNumero

            flsvUsuari.ListItems.Clear

        If (lRStUsuari.EOF) Then
            rgfMsgBox "Não há Usuários cadastrados", MsgInf
            flsvUsuari.ToolTipText = ""
            lRStUsuari.Close
            fintNumUsu = 0
            Exit Sub
        End If

        If (formMDIAce.menuUsuari.Enabled) Then
            flsvUsuari.ToolTipText = "F8 ou Duplo Clique para Manipular o Registro Selecionado"
        End If

            rlsMontarColunas
            fintNumUsu = lRStUsuari!Numero
        Do _
            While (Not ((lRStUsuari.EOF)))
        Set lLsILinhas = flsvUsuari.ListItems.Add(, , lRStUsuari!NomUsu)
       With lLsILinhas
           .SubItems(1) = lRStUsuari!Numero
           .SubItems(2) = Format(lRStUsuari!Numero, "0000")
           .SubItems(3) = IIf(lRStUsuari!TemLog, "Sim", "")
           .SubItems(4) = IIf(lRStUsuari!Status, "Bloqueado", "Ativo")
           .SubItems(5) = lRStUsuari!CodAge
           .SubItems(6) = Format(lRStUsuari!CodAge, "0000")
           .SubItems(7) = lRStUsuari!Funcao
           .SubItems(8) = Format(lRStUsuari!DatVal, "dd/mm/yyyy")
           .SubItems(9) = Format(lRStUsuari!DatVal, "yyyy/mm/dd")
           .SubItems(10) = lRStUsuari!E_Mail
        End With
            lRStUsuari.MoveNext
        Loop
        lRStUsuari.Close
End Sub
Private Sub rlsMontarColunas()
       With flsvUsuari.ColumnHeaders
           .Clear
           .Add 1, , "", IIf(gintTotRes < 1792, 3220, 4500), 0
           .Add 2, , "", 1000, 1
           .Add 3, , "", 0
           .Add 4, , "", 760, 0
           .Add 5, , "", 1200, 0
           .Add 6, , "", 1000, 1
           .Add 7, , "", 0
           .Add 8, , "", 1300, 0
           .Add 9, , "", 1300, 1
           .Add 10, , "", 0
           .Add 11, , "", 3820, 0
        End With
        rlsMontarColunasNomes
End Sub
Private Sub rlsMontarColunasNomes()
        fbytColOrd = fbytColInd - 1

         Select Case fbytColOrd
                Case 1, 5, 8
                     fbytColOrd = fbytColOrd + 1
         End Select

        With flsvUsuari
            .SortKey = ((fbytColOrd))
            .SortOrder = fbytOrdChv

             With .ColumnHeaders
                  .Item(1).Text = "Nome"
                  .Item(2).Text = "Número"
                  .Item(4).Text = "Log"
                  .Item(5).Text = "Status"
                  .Item(6).Text = "Agência"
                  .Item(8).Text = "Função"
                  .Item(9).Text = "Acesso até"
                  .Item(11).Text = "e-mail"

       Select Case fbytColInd
              Case 1
                  .Item(1).Text = "Nome  " & fstrSmbOrd
              Case 2
                  .Item(2).Text = fstrSmbOrd & "  Número"
              Case 4
                  .Item(4).Text = "Log  " & fstrSmbOrd
              Case 5
                  .Item(5).Text = "Status  " & fstrSmbOrd
              Case 6
                  .Item(6).Text = fstrSmbOrd & "  Agência"
              Case 8
                  .Item(8).Text = "Função  " & fstrSmbOrd
              Case 9
                  .Item(9).Text = fstrSmbOrd & "  Acesso Até"
              Case 11
                  .Item(11).Text = "e-mail  " & fstrSmbOrd
        End Select
        End With
        End With
End Sub
