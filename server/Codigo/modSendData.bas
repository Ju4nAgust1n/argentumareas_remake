Attribute VB_Name = "modSendData"
'**************************************************************
' SendData.bas - Has all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' Implemented by Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

''
' Contains all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' @author Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20070107

Option Explicit

Public Enum SendTarget
    ToAll = 1
    toMap
    ToPCArea
    ToAllButIndex
    ToMapButIndex
    ToGM
    ToNPCArea
    ToGuildMembers
    ToAdmins
    ToPCAreaButIndex
    ToAdminsAreaButConsejeros
    ToDiosesYclan
    ToConsejo
    ToClanArea
    ToConsejoCaos
    ToRolesMasters
    ToDeadArea
    ToCiudadanos
    ToCriminales
    ToPartyArea
    ToReal
    ToCaos
    ToCiudadanosYRMs
    ToCriminalesYRMs
    ToRealYRMs
    ToCaosYRMs
    ToHigherAdmins
    ToGMsArea
    ToUsersAreaButGMs
End Enum

Public Sub SendData(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndData As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus) - Rewrite of original
'Last Modify Date: 01/08/2007
'Last modified by: (liquid)
'**************************************************************
On Error Resume Next
    Dim LoopC As Long
    Dim map As Integer
    
    Select Case sndRoute
        Case SendTarget.ToPCArea
            Call SendToUserArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToAdmins
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnID <> -1 Then
                    If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
                        Call EnviarDatosASlot(LoopC, sndData)
                   End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToAll
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnID <> -1 Then
                    If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToAllButIndex
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) And (LoopC <> sndIndex) Then
                    If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.toMap
            Call SendToMap(sndIndex, sndData)
            Exit Sub
          
        Case SendTarget.ToMapButIndex
            Call SendToMapButIndex(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToGuildMembers
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            While LoopC > 0
                If (UserList(LoopC).ConnID <> -1) Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
                LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            Wend
            Exit Sub
        
        Case SendTarget.ToDeadArea
            Call SendToDeadUserArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToPCAreaButIndex
            Call SendToUserAreaButindex(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToClanArea
            Call SendToUserGuildArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToPartyArea
            Call SendToUserPartyArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToAdminsAreaButConsejeros
            Call SendToAdminsButConsejerosArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToNPCArea
            Call SendToNpcArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToDiosesYclan
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            While LoopC > 0
                If (UserList(LoopC).ConnID <> -1) Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
                LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            Wend
            
            LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
            While LoopC > 0
                If (UserList(LoopC).ConnID <> -1) Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
                LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
            Wend
            
            Exit Sub
        
        Case SendTarget.ToConsejo
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).flags.Privilegios And PlayerType.RoyalCouncil Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToConsejoCaos
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).flags.Privilegios And PlayerType.ChaosCouncil Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToRolesMasters
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToCiudadanos
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If Not criminal(LoopC) Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToCriminales
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If criminal(LoopC) Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToReal
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).Faccion.ArmadaReal = 1 Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToCaos
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).Faccion.FuerzasCaos = 1 Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToCiudadanosYRMs
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If Not criminal(LoopC) Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToCriminalesYRMs
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If criminal(LoopC) Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToRealYRMs
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).Faccion.ArmadaReal = 1 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToCaosYRMs
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).Faccion.FuerzasCaos = 1 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToHigherAdmins
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnID <> -1 Then
                    If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
                        Call EnviarDatosASlot(LoopC, sndData)
                   End If
                End If
            Next LoopC
            Exit Sub
            
        Case SendTarget.ToGMsArea
            Call SendToGMsArea(sndIndex, sndData)
            Exit Sub
            
        Case SendTarget.ToUsersAreaButGMs
            Call SendToUsersAreaButGMs(sndIndex, sndData)
            Exit Sub
            
    End Select
End Sub

'**************************************************************
'Author: Juan Agustín Oliva (Agushh/Thorkes)
'**************************************************************
Private Sub SendToUserArea(ByVal userIndex As Integer, ByVal sdData As String)

With UserList(userIndex)
    Call modNewAreas.new_SendToAreaByPos(sdData, .Pos.map, .Pos.X, .Pos.Y)
End With

End Sub

'**************************************************************
'Author: Juan Agustín Oliva (Agushh/Thorkes)
'**************************************************************
Private Sub SendToUserAreaButindex(ByVal userIndex As Integer, ByVal sdData As String)

With UserList(userIndex)
    Call modNewAreas.new_SendToAreaButIndexByPos(userIndex, sdData, .Pos.map, .Pos.X, .Pos.Y)
End With

End Sub

'**************************************************************
'Author: Juan Agustín Oliva (Agushh/Thorkes)
'**************************************************************
Private Sub SendToDeadUserArea(ByVal userIndex As Integer, ByVal sdData As String)

With UserList(userIndex)

    Call new_SendToDeadAreaByPos(sdData, .Pos.map, .Pos.X, .Pos.Y)

End With

End Sub

'**************************************************************
'Author: Juan Agustín Oliva (Agushh/Thorkes)
'**************************************************************
Private Sub SendToUserGuildArea(ByVal userIndex As Integer, ByVal sdData As String)

With UserList(userIndex)

    Call new_SendToGuildAreaByPos(.GuildIndex, sdData, .Pos.map, .Pos.X, .Pos.Y)

End With

End Sub

'**************************************************************
'Author: Juan Agustín Oliva (Agushh/Thorkes)
'**************************************************************
Private Sub SendToUserPartyArea(ByVal userIndex As Integer, ByVal sdData As String)

With UserList(userIndex)

    Call new_SendToPartyAreaByPos(.partyIndex, sdData, .Pos.map, .Pos.X, .Pos.Y)

End With

End Sub

'**************************************************************
'Author: Juan Agustín Oliva (Agushh/Thorkes)
'**************************************************************
Private Sub SendToAdminsButConsejerosArea(ByVal userIndex As Integer, ByVal sdData As String)

With UserList(userIndex)

    Call new_SendToGMAreaButCouncilByPos(sdData, .Pos.map, .Pos.X, .Pos.Y)

End With

End Sub

'**************************************************************
'Author: Juan Agustín Oliva (Agushh/Thorkes)
'**************************************************************
Private Sub SendToNpcArea(ByVal NpcIndex As Long, ByVal sdData As String)

With Npclist(NpcIndex)

    Call new_SendToNPCAreaByPos(sdData, .Pos.map, .Pos.X, .Pos.Y)

End With

End Sub

'**************************************************************
'Author: Juan Agustín Oliva (Agushh/Thorkes)
'**************************************************************
Public Sub SendToAreaByPos(ByVal sdData As String, ByVal map As Integer, _
                            ByVal X As Byte, ByVal Y As Byte)

Call new_SendToAreaByPos(sdData, map, X, Y)

End Sub

'**************************************************************
'Author: Juan Agustín Oliva (Agushh/Thorkes)
'Enviamos un paquete a todos los usuarios de las areas del mapa
'**************************************************************
Public Sub SendToMap(ByVal map As Integer, ByVal sdData As String)

    Dim i           As Long
    Dim UI
    Dim userIndex   As Integer
    
    For i = 1 To AREAS_AMOUNT
    
        For Each UI In areasData(map, i).userArea.Items
            
            userIndex = UI
            
            If UserList(userIndex).ConnIDValida Then
                Call EnviarDatosASlot(userIndex, sdData)
            End If
            
        Next
        
    Next i
    
End Sub

'**************************************************************
'Author: Juan Agustín Oliva (Agushh/Thorkes)
'Enviamos un paquete a todos los usuarios de las areas del mapa menos a uno (userIndex)
'**************************************************************
Public Sub SendToMapButIndex(ByVal userIndex As Integer, ByVal sdData As String)

    Dim i           As Long
    Dim UI
    Dim tIndex      As Integer
    
    For i = 1 To AREAS_AMOUNT
    
        For Each UI In areasData(UserList(userIndex).Pos.map, i).userArea.Items
            
            tIndex = UI
            
            If UserList(userIndex).ConnIDValida Then
            
                If tIndex <> userIndex Then
                    Call EnviarDatosASlot(tIndex, sdData)
                End If
                
            End If
            
        Next
        
    Next i
    
End Sub

'**************************************************************
'Author: Juan Agustín Oliva (Agushh/Thorkes)
'**************************************************************
Private Sub SendToGMsArea(ByVal userIndex As Integer, ByVal sdData As String)

With UserList(userIndex)

    Call new_SendToGMsAreaByPos(sdData, .Pos.map, .Pos.X, .Pos.Y)

End With

End Sub

'**************************************************************
'Author: Juan Agustín Oliva (Agushh/Thorkes)
'**************************************************************
Private Sub SendToUsersAreaButGMs(ByVal userIndex As Integer, ByVal sdData As String)

With UserList(userIndex)

    Call new_SendToUserAreaButGMByPos(sdData, .Pos.map, .Pos.X, .Pos.Y)

End With

End Sub
