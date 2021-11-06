Attribute VB_Name = "moduDclara"
Option Explicit

'      Administradora
Public gintCodAdm As Integer

Public gstrNomApl As String

'      Ambiente
Public gbooCancel As Boolean

Public gbooProAce As Boolean

Public gintTempoE As Integer

Public gintTempoP As Integer

Public gintTotRes As Integer

'      Banco de Dados
Public gdatServBD As Date

Public gstrServBD As String, gstrNomeBD As String

Public gstrUsuaBD As String, gstrSenhBD As String

'      Chaves do  Form  de  Usuários
Public gbooConUsu As Boolean

Public gintNumUsu As Integer

'      Chaves do  Form  de  Fundos
Public gbooConFun As Boolean

Public gbytNumFun As Byte

'      Chaves do  Form  de  Módulos
Public gbooConMod As Boolean

Public gbytNumMod As Byte

'      Chaves do  Form  de  Formes
Public gbooConFor As Boolean

Public gintNumFor As Integer

'      Chaves do  Form  de  Botões
Public gbooConBot As Boolean

Public gintNumBot As Integer

'      Chaves do  Form  de  Acesso a Fundos
Public gbytConFun As Byte

'      Chaves do  Form  de  Acesso a Módulos
Public gbytConMod As Byte

'      Chaves do  Form  de  Acesso a Forms
Public gbytConFor As Byte

'      Chaves do  Form  de  Acesso a Botões
Public gbytConBot As Byte

'      Classes
Public gclsModulo As clssModulo

Public gclsFundos As clssFundos

Public gclsUsuari As clssUsuari

Public gclsFormes As clssFormes

Public gclsAceMod As clssAceMod

Public gclsAceBot As clssAceBot

Public gclsLogado As clssLogado

Public gclsDiario As clssDiario

Public gclsInsAdm As clssInsAdm

'      Conexão
Public gDBCFundos As Connection

'      Diário
Public gbooForLog As Boolean

Public gbooUsuLog As Boolean

Public gintForAtu As Integer

Public gstrCteudo As String

'      Mensagens
Public gbooAjuHab As Boolean

Public Enum NumMsg
       MsgErr = 1
       MsgInf = 2
       MsgNao = 3
End Enum

'      Miscelânea
Public gbytQtdAce As Byte

Public gbytScrBar As Byte

Public gdatDatFim As Date

Public gdatDatIni As Date

Public gintNumItm As Integer

'      Rotas
Public gstrPthAAj As String

Public gstrPthAju As String

Public gstrPthAtu As String

Public gstrPthExe As String

'      Sessão
Public gbytModAtv As Byte

Public gintAgeLog As Integer

Public gintUsuLog As Integer

Public gstrNomCmp As String

Public gstrNomUsu As String, gstrSenhas As String
