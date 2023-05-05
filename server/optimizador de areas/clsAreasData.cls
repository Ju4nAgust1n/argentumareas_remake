VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAreasData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public npcArea      As Dictionary
Public userArea     As Dictionary
Public objArea      As Dictionary

'***************************************************
'Autor: Juan Agustín Oliva - UserForos: Agushh/Thorkes
'***************************************************
Private Sub Class_Initialize()

    Set npcArea = New Dictionary
    Set userArea = New Dictionary
    Set objArea = New Dictionary
    
End Sub

'***************************************************
'Autor: Juan Agustín Oliva - UserForos: Agushh/Thorkes
'***************************************************
Private Sub Class_Terminate()
    
    Set npcArea = Nothing
    Set userArea = Nothing
    Set objArea = Nothing
    
End Sub

