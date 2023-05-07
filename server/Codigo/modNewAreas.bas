Attribute VB_Name = "modNewAreas"
'@
'@ Autor: Juan Agustín Oliva
'@ UserForos/Discord: agush/Thorkes
'@ juancho_isap14@hotmail.com
'@ desc: nuevo sistema de areas para Argentum Online
'@

Option Explicit

'Tamaño del mapa
Public Const XMaxMapSize As Integer = 100
Public Const XMinMapSize As Integer = 1
Public Const YMaxMapSize As Integer = 100
Public Const YMinMapSize As Integer = 1

'cantidad máxima de objetos permitidos en el piso
Private Const MAX_OBJ_DROP                                                  As Integer = 10000

'visión para una resolución de 800x600
Private Const USER_VISION_X                                                 As Integer = 20
Private Const USER_VISION_Y                                                 As Integer = 20

'variables de áreas
Public areasData()                                                          As New clsAreasData
Public adyacentArea()                                                       As tAdyacents
Public areasSize                                                            As Long
Public areasAmount                                                          As Long
Public posToAreaID(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize)  As Long

Public Const MAX_ADY                                                        As Byte = 9 'máximo de nueve porque incluye el propio area principal

Public Const AREAS_AMOUNT                                                   As Byte = 25 ' 10000 (tiles cuadrados del mapa) /400 (tiles cuadrados del área)

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

Type tAdyacents
    ady(1 To MAX_ADY)                                                       As Long
End Type

Public Type tNewAreas
    areaID                                                                  As Long
End Type

' @
' @ Autor: Juan Agustín Oliva
' @ desc; inicializamos el sistema de áreas y sus variables
' @
Public Sub areasInitialize()

    Dim getter          As New clsIniReader
    Dim fileAmountAreas As Long
    Dim i               As Long
    Dim X               As Long
    Dim Y               As Long
    Dim z               As Long

    'calculos iniciales
    areasSize = USER_VISION_X * USER_VISION_Y
    areasAmount = (XMaxMapSize * YMaxMapSize) / areasSize
    
    'redimensionamos e inicializamos las areas
    ReDim areasData(1 To NumMaps, areasAmount) As New clsAreasData
    
    Call getter.Initialize(DatPath & "areas.dat")
    
    fileAmountAreas = CLng(getter.GetValue("INIT", "NumAreas"))
    
    ReDim adyacentArea(1 To fileAmountAreas) As tAdyacents
    
    For i = 1 To fileAmountAreas
    
        For X = 1 To MAX_ADY
            adyacentArea(i).ady(X) = getter.GetValue("AREA" & i, "Ad" & X)
        Next X
        
    Next i
    
    Set getter = Nothing
    
    Call getter.Initialize(DatPath & "areaspos.ini")
    
    Dim m_area(1 To AREAS_AMOUNT)       As Long
    Dim minX(1 To AREAS_AMOUNT)         As Long
    Dim minY(1 To AREAS_AMOUNT)         As Long
    Dim maxX(1 To AREAS_AMOUNT)         As Long
    Dim maxY(1 To AREAS_AMOUNT)         As Long
    
    For X = 1 To fileAmountAreas
    
        m_area(X) = X
    
        minX(X) = CLng(getter.GetValue("AREA" & X, "minX"))
        maxX(X) = CLng(getter.GetValue("AREA" & X, "maxX"))
        
        minY(X) = CLng(getter.GetValue("AREA" & X, "minY"))
        maxY(X) = CLng(getter.GetValue("AREA" & X, "maxY"))

    Next X
    
    For X = XMinMapSize To XMaxMapSize
    
        For Y = YMinMapSize To YMaxMapSize
        
            For z = 1 To AREAS_AMOUNT
            
                If X >= minX(z) And X <= maxX(z) And Y >= minY(z) And Y <= maxY(z) Then
                    posToAreaID(X, Y) = m_area(z)
                End If
            
            Next z
        
        Next Y
        
    Next X
    
    Set getter = Nothing
    
End Sub

' @
' @ Autor: Juan Agustín Oliva
' @ desc; agregamos un NPC al area
' @
Public Sub addNpcArea(ByVal NpcIndex As Integer, ByVal map As Integer, ByVal X As Byte, ByVal Y As Byte)

With areasData(map, posToAreaID(X, Y))

    Npclist(NpcIndex).newAreas.areaID = -1
    
    Call refreshNPCArea(NpcIndex)

End With

End Sub

' @
' @ Autor: Juan Agustín Oliva
' @ desc; agregamos un USER al area
' @
Public Sub addUserArea(ByVal UserIndex As Integer, ByVal map As Integer, ByVal X As Byte, ByVal Y As Byte)

With areasData(map, posToAreaID(X, Y))

    UserList(UserIndex).newAreas.areaID = -1
    
    Call refreshUserArea(UserIndex)

End With

End Sub

' @
' @ Autor: Juan Agustín Oliva
' @ desc; agregamos un OBJETO al area
' @
Public Sub addObjArea(ByVal objIndex As Integer, ByVal amount As Integer, _
                      ByVal map As Integer, ByVal X As Byte, ByVal Y As Byte)

If objIndex > 0 And amount > 0 And amount <= MAX_OBJ_DROP Then

    Dim tempObjData As New clsObjData
    
    tempObjData.p_objIndex = objIndex
    tempObjData.p_amount = amount
    tempObjData.p_posX = X
    tempObjData.p_posY = Y
    
    With areasData(map, posToAreaID(X, Y))
    
        Call .objArea.Add(tempObjData, tempObjData)
    
    End With
    
    Set tempObjData = Nothing

End If

End Sub

' @
' @ Autor: Juan Agustín Oliva
' @ desc; quitamos un NPC del area
' @
Public Sub removeNpcArea(ByVal NpcIndex As Integer, ByVal map As Integer, ByVal X As Byte, ByVal Y As Byte)

With areasData(map, posToAreaID(X, Y))

    Call .npcArea.Remove(NpcIndex)

End With

End Sub

' @
' @ Autor: Juan Agustín Oliva
' @ desc; quitamos un USER del area
' @
Public Sub removeUserArea(ByVal UserIndex As Integer, ByVal map As Integer, ByVal X As Byte, ByVal Y As Byte)

With areasData(map, posToAreaID(X, Y))

    Call .userArea.Remove(UserIndex)

End With

End Sub

' @
' @ Autor: Juan Agustín Oliva
' @ desc; quitamos un OBJETO del area
' @
Public Sub removeObjArea(ByVal objIndex As Integer, ByVal map As Integer, ByVal X As Byte, ByVal Y As Byte)

With areasData(map, posToAreaID(X, Y))

    Dim i
    
    For Each i In .objArea.Items
    
        If i.p_objIndex = objIndex Then
            Call .objArea.Remove(i)
            Exit For
        End If
    
    Next

End With

End Sub

'@
'@ Autor: Juan Agustín Oliva
'@ enviamos un paquete al area pero no a gms
'@
Public Sub new_SendToUserAreaButGMByPos(ByVal data As String, _
                                        ByVal map As Integer, ByVal X As Byte, ByVal Y As Byte)

Dim areaID              As Long
Dim i                   As Long
Dim UserIndex           As Integer
Dim T

areaID = posToAreaID(X, Y)

For i = 1 To MAX_ADY
    
    For Each T In areasData(map, adyacentArea(areaID).ady(i)).userArea.Items
                
            UserIndex = T
                
            If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
                
                Call EnviarDatosASlot(UserIndex, data)
                    
            End If
            
    Next

Next i

End Sub

'@
'@ Autor: Juan Agustín Oliva
'@ enviamos un paquete al area solo a gms
'@
Public Sub new_SendToGMsAreaByPos(ByVal data As String, _
                                   ByVal map As Integer, ByVal X As Byte, ByVal Y As Byte)

Dim areaID              As Long
Dim areasAdyacent()     As Long
Dim i                   As Long
Dim UserIndex           As Integer
Dim T

areaID = posToAreaID(X, Y)
 
For i = 1 To MAX_ADY
    
        For Each T In areasData(map, adyacentArea(areaID).ady(i)).userArea.Items
                
                UserIndex = T
                
                If UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            
                    Call EnviarDatosASlot(UserIndex, data)
                    
                End If
            
        Next

Next i

End Sub

'@
'@ Autor: Juan Agustín Oliva
'@ enviamos un paquete al area solo a gente del party
'@
Public Sub new_SendToPartyAreaByPos(ByVal partyIndex As Integer, ByVal data As String, _
                                   ByVal map As Integer, ByVal X As Byte, ByVal Y As Byte)

Dim areaID              As Long
Dim areasAdyacent()     As Long
Dim i                   As Long
Dim UserIndex           As Integer
Dim T

areaID = posToAreaID(X, Y)
 
For i = 1 To MAX_ADY
    
        For Each T In areasData(map, adyacentArea(areaID).ady(i)).userArea.Items
        
            If partyIndex <> 0 Then
                
                UserIndex = T
                
                If partyIndex = UserList(UserIndex).partyIndex Then
            
                    Call EnviarDatosASlot(UserIndex, data)
                    
                End If
                
            End If
            
        Next
        
Next i

End Sub

'@
'@ Autor: Juan Agustín Oliva
'@ enviamos un paquete al area del npc en funcion de su posición
'@
Public Sub new_SendToNPCAreaByPos(ByVal data As String, _
                                   ByVal map As Integer, ByVal X As Byte, ByVal Y As Byte, _
                                   Optional ByVal walk As Boolean = False)
                                   
Dim areaID              As Long
Dim i                   As Long
Dim UserIndex           As Integer
Dim T

areaID = posToAreaID(X, Y)

For i = 1 To MAX_ADY
        
    For Each T In areasData(map, adyacentArea(areaID).ady(i)).userArea.Items
                   
        UserIndex = T
                
        Call EnviarDatosASlot(UserIndex, data)
                
    Next

Next i
                                   
End Sub

'@
'@ Autor: Juan Agustín Oliva
'@ enviamos un paquete al area de gms pero no a consejeros
'@
Public Sub new_SendToGMAreaButCouncilByPos(ByVal data As String, _
                                   ByVal map As Integer, ByVal X As Byte, ByVal Y As Byte)
                                   
Dim areaID              As Long
Dim i                   As Long
Dim UserIndex           As Integer
Dim T

areaID = posToAreaID(X, Y)
 
For i = 1 To MAX_ADY

    For Each T In areasData(map, adyacentArea(areaID).ady(i)).userArea.Items
                
        UserIndex = T
                
        If UserList(UserIndex).flags.Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin) Then
            
            Call EnviarDatosASlot(UserIndex, data)
                    
        End If
            
    Next

Next i
                                   
End Sub

'@
'@ Autor: Juan Agustín Oliva
'@ enviamos un paquete al area pero solo al clan
'@
Public Sub new_SendToGuildAreaByPos(ByVal GuildIndex As Integer, ByVal data As String, _
                                   ByVal map As Integer, ByVal X As Byte, ByVal Y As Byte)

Dim areaID              As Long
Dim areasAdyacent()     As Long
Dim i                   As Long
Dim UserIndex           As Integer
Dim T

areaID = posToAreaID(X, Y)
 
For i = 1 To MAX_ADY
    
    For Each T In areasData(map, adyacentArea(areaID).ady(i)).userArea.Items
        
            If GuildIndex <> 0 Then
                
                UserIndex = T
                
                If GuildIndex = UserList(UserIndex).GuildIndex Then
            
                    Call EnviarDatosASlot(UserIndex, data)
                    
                End If
                
            End If
            
        Next

Next i

End Sub

'@
'@ Autor: Juan Agustín Oliva
'@ enviamos un paquete al area pero solo a caspers
'@
Public Sub new_SendToDeadAreaByPos(ByVal data As String, _
                                   ByVal map As Integer, ByVal X As Byte, ByVal Y As Byte)

Dim areaID              As Long
Dim areasAdyacent()     As Long
Dim i                   As Long
Dim UserIndex           As Integer
Dim T

areaID = posToAreaID(X, Y)
 
For i = 1 To MAX_ADY
    
        For Each T In areasData(map, adyacentArea(areaID).ady(i)).userArea.Items
                
                UserIndex = T
                
                If UserList(UserIndex).flags.Muerto Then
            
                    Call EnviarDatosASlot(UserIndex, data)
                    
                End If
            
        Next

Next i

End Sub

'@
'@ Autor: Juan Agustín Oliva
'@ enviamos un paquete al area en función de las coordenadas
'@
Public Sub new_SendToAreaByPos(ByVal data As String, _
                                ByVal map As Integer, ByVal X As Byte, ByVal Y As Byte)

Dim areaID              As Long
Dim tIndex              As Integer
Dim T
Dim i                   As Long

areaID = posToAreaID(X, Y)
    
For i = 1 To MAX_ADY
    
    For Each T In areasData(map, adyacentArea(areaID).ady(i)).userArea.Items
            
        tIndex = T
                
        Call EnviarDatosASlot(tIndex, data)
                
    Next
    
Next i

End Sub

'@
'@ Autor: Juan Agustín Oliva
'@ enviamos un paquete al area en función de las coordenadas, menos a userIndex
'@
Public Sub new_SendToAreaButIndexByPos(ByVal UserIndex As Integer, ByVal data As String, _
                                ByVal map As Integer, ByVal X As Byte, ByVal Y As Byte)

Dim areaID              As Long
Dim tIndex              As Integer
Dim T
Dim i                   As Long

areaID = posToAreaID(X, Y)

For i = 1 To MAX_ADY
    
    For Each T In areasData(map, adyacentArea(areaID).ady(i)).userArea.Items
            
        tIndex = T
        
        If tIndex <> UserIndex Then
            Call EnviarDatosASlot(tIndex, data)
        End If
                
    Next
    
Next i

End Sub

'@
'@ Autor: Juan Agustín Oliva
'@ este método se llama cuando un usuario cambia de área. Refresca info y envía al cliente
'@ en el viejo sistema de áreas, este método sería el 'CheckUpdateNeededUser'
'@
Public Sub refreshUserArea(ByVal UserIndex As Integer)

With UserList(UserIndex)
    
    Dim areaID  As Long
    areaID = posToAreaID(.Pos.X, .Pos.Y)

    If .newAreas.areaID <> areaID Then
    
        Dim i   As Long
    
        'borramos el user del viejo area
        If .newAreas.areaID <> -1 Then Call areasData(.Pos.map, .newAreas.areaID).userArea.Remove(UserIndex)
        
        'agregamos el user al nuevo area
        Call areasData(.Pos.map, areaID).userArea.Add(UserIndex, UserIndex)
        
        Dim T
        Dim tIndex  As Integer
        
        For i = 1 To MAX_ADY
        
            For Each T In areasData(.Pos.map, adyacentArea(areaID).ady(i)).userArea.Items
            
                tIndex = T 'lo manejamos dentro de un integer para no convertirlo cada vez que se use
                
                If tIndex <> UserIndex Or .newAreas.areaID = -1 Then
            
                    Call MakeUserChar(False, UserIndex, tIndex, UserList(tIndex).Pos.map, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y)
                    Call MakeUserChar(False, tIndex, UserIndex, .Pos.map, .Pos.X, .Pos.Y)
                    
                End If
            
            Next
            
            For Each T In areasData(.Pos.map, adyacentArea(areaID).ady(i)).npcArea.Items
            
                tIndex = T
            
                Call MakeNPCChar(False, UserIndex, tIndex, Npclist(tIndex).Pos.map, Npclist(tIndex).Pos.X, Npclist(tIndex).Pos.Y)
            
            Next
            
            For Each T In areasData(.Pos.map, adyacentArea(areaID).ady(i)).objArea.Items
            
                Call WriteObjectCreate(UserIndex, ObjData(T.p_objIndex).GrhIndex, T.p_posX, T.p_posY)
                            
                If ObjData(T.p_objIndex).OBJType = eOBJType.otPuertas Then
                   Call Bloquear(False, UserIndex, T.p_posX, T.p_posY, MapData(.Pos.map, T.p_posX, T.p_posY).Blocked)
                   Call Bloquear(False, UserIndex, T.p_posX - 1, T.p_posY, MapData(.Pos.map, T.p_posX - 1, T.p_posY).Blocked)
                End If
            
            Next
        
        Next i
        
        .newAreas.areaID = areaID
    
    End If

End With

End Sub

'@
'@ Autor: Juan Agustín Oliva
'@ este método se llama cuando un npc cambia de área. Refresca info y envía a los usuarios de corresponder
'@
Public Sub refreshNPCArea(ByVal NpcIndex As Integer)

With Npclist(NpcIndex)

    Dim areaID  As Long
    areaID = posToAreaID(.Pos.X, .Pos.Y)

    If .newAreas.areaID <> areaID Then
    
        Dim i   As Long
    
        'borramos el npc del viejo area
        If .newAreas.areaID <> -1 Then Call areasData(.Pos.map, .newAreas.areaID).npcArea.Remove(NpcIndex)
        
        'agregamos el user al nuevo area
        Call areasData(.Pos.map, areaID).npcArea.Add(NpcIndex, NpcIndex)
        
        Dim T
        Dim tIndex  As Integer
        
        For i = 1 To MAX_ADY
        
            For Each T In areasData(.Pos.map, adyacentArea(areaID).ady(i)).userArea.Items
            
                tIndex = T 'lo manejamos dentro de un integer por las dudas
            
                Call MakeNPCChar(False, tIndex, NpcIndex, .Pos.map, .Pos.X, .Pos.Y)
            
            Next
            
        Next i
        
        .newAreas.areaID = areaID
        
    End If

End With

End Sub


