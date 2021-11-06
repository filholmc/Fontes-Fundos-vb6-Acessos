VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clssAgeCad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Function ConsultarTodas() As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarTodas = New Recordset
              ConsultarTodas.Open _
 _
             "SELECT Codigo, NomAge FROM TBAgeCad ORDER BY Codigo", _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
