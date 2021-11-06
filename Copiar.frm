VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form formCopiar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cópia de Acessos"
   ClientHeight    =   4410
   ClientLeft      =   4320
   ClientTop       =   1755
   ClientWidth     =   5640
   Icon            =   "Copiar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "Copiar.frx":030A
   ScaleHeight     =   4410
   ScaleWidth      =   5640
   Begin VB.Frame Frame3 
      Caption         =   "Copiar:"
      Height          =   615
      Left            =   570
      TabIndex        =   2
      Top             =   1860
      Width           =   4515
      Begin VB.OptionButton fopcCopItm 
         Caption         =   "por Item"
         Height          =   225
         Left            =   3480
         TabIndex        =   4
         Top             =   240
         Width           =   885
      End
      Begin VB.OptionButton fopcCopAll 
         Caption         =   "Tudo"
         Height          =   225
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.ComboBox fcboModulo 
      Height          =   315
      ItemData        =   "Copiar.frx":0BD4
      Left            =   570
      List            =   "Copiar.frx":0BD6
      TabIndex        =   6
      Top             =   3420
      Visible         =   0   'False
      Width           =   4515
   End
   Begin VB.ComboBox fcboItmAce 
      Height          =   315
      ItemData        =   "Copiar.frx":0BD8
      Left            =   570
      List            =   "Copiar.frx":0BE8
      TabIndex        =   5
      Top             =   2790
      Visible         =   0   'False
      Width           =   4515
   End
   Begin VB.ComboBox fcboUsuDes 
      Height          =   315
      ItemData        =   "Copiar.frx":0C0C
      Left            =   570
      List            =   "Copiar.frx":0C0E
      TabIndex        =   1
      ToolTipText     =   "Usuário para quem os Acessos serão copiados"
      Top             =   1440
      Width           =   4515
   End
   Begin VB.ComboBox fcboUsuOri 
      Height          =   315
      ItemData        =   "Copiar.frx":0C10
      Left            =   570
      List            =   "Copiar.frx":0C12
      TabIndex        =   0
      ToolTipText     =   "Usuário de quem os Acessos serão copiados"
      Top             =   810
      Width           =   4515
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   0
      TabIndex        =   14
      Top             =   3840
      Width           =   5625
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   510
      TabIndex        =   13
      Top             =   450
      Width           =   5115
   End
   Begin VB.CommandButton fcmbF09LCA 
      Caption         =   "F9"
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF11Tab 
      Caption         =   "F11"
      Height          =   255
      Left            =   4650
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF12Hom 
      Caption         =   "F12"
      Height          =   255
      Left            =   5100
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbEscape 
      Caption         =   "Fechar (Esc)"
      Height          =   315
      Left            =   4410
      TabIndex        =   8
      Top             =   3990
      Width           =   1125
   End
   Begin VB.CommandButton fcmbF02Cop 
      Caption         =   "Copiar (F2)"
      Default         =   -1  'True
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   3990
      Width           =   1120
   End
   Begin ComctlLib.ProgressBar fprbCopiar 
      Height          =   180
      Left            =   570
      TabIndex        =   12
      Top             =   150
      Visible         =   0   'False
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   318
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label flblModulo 
      AutoSize        =   -1  'True
      Caption         =   "Módulo ao qual pertencem os"
      Height          =   195
      Left            =   570
      LinkTimeout     =   0
      TabIndex        =   18
      Top             =   3210
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label flblItmAce 
      AutoSize        =   -1  'True
      Caption         =   "Itens que serão copiados"
      Height          =   195
      Left            =   570
      LinkTimeout     =   0
      TabIndex        =   17
      Top             =   2580
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Label flblUsuDes 
      AutoSize        =   -1  'True
      Caption         =   "Usuário de Destino"
      Height          =   195
      Left            =   570
      LinkTimeout     =   0
      TabIndex        =   16
      Top             =   1230
      Width           =   1350
   End
   Begin VB.Label flblUsuOri 
      AutoSize        =   -1  'True
      Caption         =   "Usuário de Origem"
      Height          =   195
      Left            =   570
      LinkTimeout     =   0
      TabIndex        =   15
      Top             =   600
      Width           =   1305
   End
End
Attribute VB_Name = "formCopiar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private fbooForLog As Boolean

Private fbytNumMod As Byte

Private fbytQtdInc As Byte

Private fclsAceFun As clssAceFun

Private fclsAceFor As clssAceFor

Private fintForAtu As Integer

Private fintUsuDes As Integer, fintUsuOri As Integer
Private Sub fcboItmAce_Click()
        flblModulo.Caption = "Módulo ao qual petencem os " & fcboItmAce
        flblModulo.Visible = IIf(fcboItmAce.ListIndex > 1, True, False)
        fcboModulo.Visible = IIf(fcboItmAce.ListIndex > 1, True, False)
End Sub
Private Sub fcboModulo_Click()
        If (fcboModulo.ListIndex = -1) Then Exit Sub

        fbytNumMod = CByte(Mid(fcboModulo, 1, InStr(fcboModulo, " ") - 1))
End Sub
Private Sub fcboUsuDes_Click()
        If (fcboUsuDes.ListIndex = -1) Then Exit Sub

        fintUsuDes = CInt(Mid(fcboUsuDes, 1, InStr(fcboUsuDes, " ") - 1))
End Sub
Private Sub fcboUsuOri_Click()
        If (fcboUsuOri.ListIndex = -1) Then Exit Sub

        fintUsuOri = CInt(Mid(fcboUsuOri, 1, InStr(fcboUsuOri, " ") - 1))

        rgsCarregarModulosDeUmUsuario fintUsuOri, fcboModulo

        rlsCarregarUsuariosDeDestino
End Sub
Private Sub fcmbEscape_Click()
        Unload Me
End Sub
Private Sub fcmbF02Cop_Click()
        If (fcboUsuOri.ListIndex = -1) Then
            rgfMsgBox "Escolha um Usuário de Origem", MsgErr
            fcboUsuOri.SetFocus
            Exit Sub
        End If

        If (fcboUsuDes.ListIndex = -1) Then
            rgfMsgBox "Escolha um Usuário de Destino", MsgErr
            fcboUsuDes.SetFocus
            Exit Sub
        End If

        If (fopcCopItm) Then
        If (fcboItmAce.ListIndex = -1) Then
            rgfMsgBox "Escolha os Itens que serão copiados", MsgErr
            fcboItmAce.SetFocus
            Exit Sub
        End If

        If (fcboItmAce.ListIndex > 1) Then
        If (fcboModulo.ListIndex = -1) Then
            rgfMsgBox "Escolha o " & Mid(flblModulo, 1, Len(flblModulo) - 1), MsgErr
            fcboModulo.SetFocus
            Exit Sub
        Else
        If (gclsAceMod.Ausente(fintUsuDes, fbytNumMod)) Then
            rgfMsgBox "Usuário de Destino não possui Acesso ao Módulo especificado", MsgErr
            fcboModulo.SetFocus
            Exit Sub
        End If
        End If
        End If
        End If

        If (fopcCopAll) Then
            rlsCopiarTudo
        Else
            rlsCopiarPorItem
        End If
            gclsDiario.Incluir fbooForLog, gintUsuLog, gstrNomCmp, 1, fintForAtu, _
                                                         "Copiou", 0, Format(fintUsuOri, "0") & "; " & _
                                                                      Format(fintUsuDes, "0"), gstrCteudo
End Sub
Private Sub fcmbF09LCA_Click()
        If (Not TypeOf ActiveControl Is CommandButton) Then ActiveControl.Text = ""
End Sub
Private Sub fcmbF11Tab_Click()
        SendKeys "+{TAB}"
End Sub
Private Sub fcmbF12Hom_Click()
        fcboUsuOri.SetFocus
End Sub
Private Sub fopcCopAll_Click()
        flblItmAce.Visible = False
        fcboItmAce.Visible = False
        flblModulo.Visible = False
        fcboModulo.Visible = False
End Sub
Private Sub fopcCopItm_Click()
        flblItmAce.Visible = True
        fcboItmAce.Visible = True
End Sub
Private Sub Form_Activate()
        gintTempoP = 0
        formMDIAce.fsbrModAce.Panels(4).Picture = IIf(gbooUsuLog Or fbooForLog, _
                                                      formMDIAce.fimlStaBar.ListImages(1).Picture, LoadPicture())
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        gintTempoP = 0
        rgsTratarFuncoes KeyCode, Me
End Sub
Private Sub Form_Load()
        fopcCopAll = True
        formMDIAce.ftbrModAce.Buttons("Copiar").Value = tbrPressed

        rgsCentralizarForm Me
        rgsPosicionarAjuda Me, fintForAtu, fbooForLog

        Set fclsAceFun = New clssAceFun
        Set fclsAceFor = New clssAceFor

        rgsCarregarUsuarios fcboUsuOri, False
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        gintTempoP = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
        Set fclsAceFun = Nothing
        Set fclsAceFor = Nothing
        formMDIAce.fsbrModAce.Panels(4).Picture = LoadPicture()
        formMDIAce.ftbrModAce.Buttons("Copiar").Value = tbrUnpressed
End Sub
Private Sub rlsCarregarUsuariosDeDestino()
        Dim lRStUsuari As Recordset

        Set lRStUsuari = gclsUsuari.ConsultarTodosPorNumero

            fcboUsuDes.Clear
        Do _
            While (Not (((lRStUsuari.EOF))))
        If (fintUsuOri <> lRStUsuari!Numero) Then
            fcboUsuDes.AddItem lRStUsuari!Numero & " - " & lRStUsuari!NomUsu
        End If
            lRStUsuari.MoveNext
        Loop
        lRStUsuari.Close
End Sub
Private Sub rlsCopiarAcessosAosBotoes()
        Dim lRStAceBot As Recordset

        Set lRStAceBot = gclsAceBot.ConsultarBotoesDeUmUsuarioPorModulo(fintUsuOri, fbytNumMod)

        If (lRStAceBot.EOF) Then
            rgfMsgBox "Não há Botões acessados pelo Usuário de Origem", MsgInf
            lRStAceBot.Close
            Exit Sub
        End If

        Do _
            While (Not lRStAceBot.EOF)
        If (gclsAceBot.Ausente(fintUsuDes, lRStAceBot!NumBot)) Then
            gclsAceBot.Incluir fintUsuDes, lRStAceBot!NumBot

        If (gDBCFundos.Errors.Count > 0) Then Exit Sub

            fbytQtdInc = fbytQtdInc + 1
        End If
            gbytQtdAce = gbytQtdAce + 1
            fprbCopiar = gbytQtdAce / lRStAceBot.RecordCount * 100

            lRStAceBot.MoveNext
        Loop
        lRStAceBot.Close
End Sub
Private Sub rlsCopiarAcessosAosForms()
        Dim lRStAceFor As Recordset

        Set lRStAceFor = fclsAceFor.ConsultarFormsDeUmUsuarioPorModulo(fintUsuOri, fbytNumMod)

        If (lRStAceFor.EOF) Then
            rgfMsgBox "Não há Forms acessados pelo Usuário de Origem", MsgInf
            lRStAceFor.Close
            Exit Sub
        End If

        Do _
            While (Not lRStAceFor.EOF)
        If (fclsAceFor.Ausente(fintUsuDes, lRStAceFor!Numero)) Then
            fclsAceFor.Incluir fintUsuDes, lRStAceFor!Numero

        If (gDBCFundos.Errors.Count > 0) Then Exit Sub

            fbytQtdInc = fbytQtdInc + 1
        End If
            gbytQtdAce = gbytQtdAce + 1
            fprbCopiar = gbytQtdAce / lRStAceFor.RecordCount * 100

            lRStAceFor.MoveNext
        Loop
        lRStAceFor.Close
End Sub
Private Sub rlsCopiarAcessosAosFundos()
        Dim lRStAceFun As Recordset

        Set lRStAceFun = fclsAceFun.ConsultarFundosDeUmUsuario(fintUsuOri)

        If (lRStAceFun.EOF) Then
            rgfMsgBox "Não há Fundos acessados pelo Usuário de Origem", MsgInf
            lRStAceFun.Close
            Exit Sub
        End If

        Do _
            While (Not lRStAceFun.EOF)
        If (fclsAceFun.Ausente(fintUsuDes, lRStAceFun!Numero)) Then
            fclsAceFun.Incluir fintUsuDes, lRStAceFun!Numero

        If (gDBCFundos.Errors.Count > 0) Then Exit Sub

            fbytQtdInc = fbytQtdInc + 1
        End If
            gbytQtdAce = gbytQtdAce + 1
            fprbCopiar = gbytQtdAce / lRStAceFun.RecordCount * 100

            lRStAceFun.MoveNext
        Loop
        lRStAceFun.Close
End Sub
Private Sub rlsCopiarAcessosAosModulos()
        Dim lRStAceMod As Recordset

        Set lRStAceMod = gclsAceMod.ConsultarModulosDeUmUsuario(fintUsuOri)

        If (lRStAceMod.EOF) Then
            rgfMsgBox "Não há Módulos acessados pelo Usuário de Origem", MsgInf
            lRStAceMod.Close
            Exit Sub
        End If

        Do _
            While (Not lRStAceMod.EOF)
        If (gclsAceMod.Ausente(fintUsuDes, lRStAceMod!Numero)) Then
            gclsAceMod.Incluir fintUsuDes, lRStAceMod!Numero

        If (gDBCFundos.Errors.Count > 0) Then Exit Sub

            fbytQtdInc = fbytQtdInc + 1
        End If
            gbytQtdAce = gbytQtdAce + 1
            fprbCopiar = gbytQtdAce / lRStAceMod.RecordCount * 100

            lRStAceMod.MoveNext
        Loop
        lRStAceMod.Close
End Sub
Private Sub rlsCopiarPorItem()
        fbytQtdInc = 0
        gbytQtdAce = 0
        fprbCopiar.Visible = True

        Select Case fcboItmAce.ListIndex
               Case 0
                    gstrCteudo = "Módulos"
                    rlsCopiarAcessosAosModulos
               Case 1
                    gstrCteudo = "Fundos"
                    rlsCopiarAcessosAosFundos
               Case 2
                    gstrCteudo = "Forms do Módulo " & Mid(fcboModulo, 5, Len(fcboModulo) - 4)
                    rlsCopiarAcessosAosForms
               Case 3
                    gstrCteudo = "Botões do Módulo " & Mid(fcboModulo, 5, Len(fcboModulo) - 4)
                    rlsCopiarAcessosAosBotoes
        End Select

        If (gDBCFundos.Errors.Count > 0) Then
            rgsTratarErro Err, gDBCFundos.Errors, Me
        Else
            rgfMsgBox fbytQtdInc & " Acesso(s) copiado(s)", MsgInf

            fprbCopiar.Visible = False
            fcboUsuOri.SetFocus
        End If
End Sub
Private Sub rlsCopiarTudo()
            fbytQtdInc = 0
            gbytQtdAce = 0
            gstrCteudo = "Tudo"
            fprbCopiar.Visible = True

            rlsCopiarAcessosAosModulos

        If (gDBCFundos.Errors.Count > 0) Then GoTo Erro_DB

            rgfMsgBox fbytQtdInc & " Acesso(s) a Módulos copiado(s)", MsgInf

            fbytQtdInc = 0
            gbytQtdAce = 0

            rlsCopiarAcessosAosFundos

        If (gDBCFundos.Errors.Count > 0) Then GoTo Erro_DB

            rgfMsgBox fbytQtdInc & " Acesso(s) a Fundos copiado(s)", MsgInf

        For gintNumItm = 0 To fcboModulo.ListCount - 1
            fbytQtdInc = 0
            gbytQtdAce = 0
            fcboModulo.ListIndex = gintNumItm
            fbytNumMod = CByte(Mid(fcboModulo, 1, InStr(fcboModulo, " ") - 1))

            rlsCopiarAcessosAosForms

        If (gDBCFundos.Errors.Count > 0) Then GoTo Erro_DB

            rgfMsgBox fbytQtdInc & " Acesso(s) a Forms do Módulo " & Mid(fcboModulo, 5, Len(fcboModulo)) & " copiado(s)", MsgInf
        Next

        For gintNumItm = 0 To fcboModulo.ListCount - 1
            fbytQtdInc = 0
            gbytQtdAce = 0
            fcboModulo.ListIndex = gintNumItm
            fbytNumMod = CByte(Mid(fcboModulo, 1, InStr(fcboModulo, " ") - 1))

            rlsCopiarAcessosAosBotoes

        If (gDBCFundos.Errors.Count > 0) Then GoTo Erro_DB

            rgfMsgBox fbytQtdInc & " Acesso(s) a Botões do Módulo " & Mid(fcboModulo, 5, Len(fcboModulo)) & " copiado(s)", MsgInf
        Next
        Exit Sub
Erro_DB:
        rgsTratarErro Err, gDBCFundos.Errors, Me
End Sub
