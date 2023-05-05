VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObjData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@
'@ Autor: Juan Agustín Oliva
'@ UserForos/Discord: agush/Thorkes
'@ juancho_isap14@hotmail.com
'@ desc: nuevo sistema de areas para Argentum Online
'@

Option Explicit

Private objIndex    As Integer
Private amount      As Integer
Private pos_x       As Byte
Private pos_y       As Byte

Public Property Get p_objIndex() As Integer
    p_objIndex = objIndex
End Property

Public Property Get p_amount() As Integer
    p_amount = amount
End Property

Public Property Get p_posX() As Byte
    p_posX = pos_x
End Property

Public Property Get p_posY() As Byte
    p_posY = pos_y
End Property

Public Property Let p_objIndex(ByVal obj As Integer)
    objIndex = obj
End Property

Public Property Let p_amount(ByVal am As Integer)
    amount = am
End Property

Public Property Let p_posX(ByVal p As Byte)
    pos_x = p
End Property

Public Property Let p_posY(ByVal p As Byte)
    pos_y = p
End Property
