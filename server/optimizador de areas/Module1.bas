Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long


'cantidad máxima de objetos permitidos en el piso
Private Const MAX_OBJ_DROP  As Integer = 10000

'visión para una resolución de 800x600
Private Const USER_VISION_X As Integer = 20
Private Const USER_VISION_Y As Integer = 20

'dimensiones del mapa
Private Const MAP_X_SIZE    As Integer = 100
Private Const MAP_Y_SIZE    As Integer = 100

'variables de áreas
Public areasData()          As New clsAreasData
Public areasSize            As Long
Public areasAmount          As Long

'enumerados de array de areas
Private Enum eAreasAdyacent
    NORTH = 1
    NORTHWEST = 2
    NORTHEAST = 3
    SOUTH = 4
    SOUTHWEST = 5
    SOUTHEAST = 6
    EAST = 7
    WEST = 8
    ACTUAL_AREA = 9
End Enum

' @
' @ desc; inicializamos las variables vinculaas a las areas
' @
Public Sub areasInitialize()

    'calculos iniciales
    areasSize = USER_VISION_X * USER_VISION_Y
    areasAmount = (MAP_X_SIZE * MAP_Y_SIZE) / areasSize
    
    Debug.Print "initAreas> areasSize: " & areasSize
    
    Debug.Print "initAreas> areasAmount: " & areasAmount
    
    'redimensionamos e inicializamos las areas
    ReDim areasData(1 To 2, 0 To areasAmount) As New clsAreasData
    
    Debug.Print "initAreas> areasMap: " & UBound(areasData)
    
End Sub

' @
' @ desc; obtenemos el areaID a partir de una posición
' @
Public Function posToAreaID(ByVal map As Integer, ByVal x As Byte, ByVal y As Byte) As Long

    posToAreaID = (x / USER_VISION_X) * (y / USER_VISION_Y)

End Function

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Escribe VAR en un archivo
'*****************************************************************

writeprivateprofilestring Main, Var, value, file
    
End Sub
