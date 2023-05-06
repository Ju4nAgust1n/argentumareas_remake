Attribute VB_Name = "InvUsuario"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
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
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean

'17/09/02
'Agregue que la función se asegure que el objeto no es un barco

On Error Resume Next

Dim i As Integer
Dim objIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS
    objIndex = UserList(UserIndex).Invent.Object(i).objIndex
    If objIndex > 0 Then
            If (ObjData(objIndex).OBJType <> eOBJType.otLlaves And _
                ObjData(objIndex).OBJType <> eOBJType.otBarcos) Then
                  TieneObjetosRobables = True
                  Exit Function
            End If
    
    End If
Next i


End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, ByVal objIndex As Integer) As Boolean
On Error GoTo manejador

'Call LogTarea("ClasePuedeUsarItem")

Dim flag As Boolean

'Admins can use ANYTHING!
If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
    If ObjData(objIndex).ClaseProhibida(1) <> 0 Then
        Dim i As Integer
        For i = 1 To NUMCLASES
            If ObjData(objIndex).ClaseProhibida(i) = UserList(UserIndex).clase Then
                ClasePuedeUsarItem = False
                Exit Function
            End If
        Next i
    End If
End If

ClasePuedeUsarItem = True

Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")
End Function

Sub QuitarNewbieObj(ByVal UserIndex As Integer)
Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(j).objIndex > 0 Then
             
             If ObjData(UserList(UserIndex).Invent.Object(j).objIndex).Newbie = 1 Then _
                    Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
                    Call UpdateUserInv(False, UserIndex, j)
        
        End If
Next j

'[Barrin 17-12-03] Si el usuario dejó de ser Newbie, y estaba en el Newbie Dungeon
'es transportado a su hogar de origen ;)
If UCase$(MapInfo(UserList(UserIndex).Pos.map).Restringir) = "NEWBIE" Then
    
    Dim DeDonde As WorldPos
    
    Select Case UserList(UserIndex).Hogar
        Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
            DeDonde = Lindos
        Case eCiudad.cUllathorpe
            DeDonde = Ullathorpe
        Case eCiudad.cBanderbill
            DeDonde = Banderbill
        Case Else
            DeDonde = Nix
    End Select
    
    Call WarpUserChar(UserIndex, DeDonde.map, DeDonde.X, DeDonde.Y, True)

End If
'[/Barrin]

End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)


Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        UserList(UserIndex).Invent.Object(j).objIndex = 0
        UserList(UserIndex).Invent.Object(j).amount = 0
        UserList(UserIndex).Invent.Object(j).Equipped = 0
        
Next

UserList(UserIndex).Invent.NroItems = 0

UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
UserList(UserIndex).Invent.ArmourEqpSlot = 0

UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
UserList(UserIndex).Invent.WeaponEqpSlot = 0

UserList(UserIndex).Invent.CascoEqpObjIndex = 0
UserList(UserIndex).Invent.CascoEqpSlot = 0

UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
UserList(UserIndex).Invent.EscudoEqpSlot = 0

UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
UserList(UserIndex).Invent.AnilloEqpSlot = 0

UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
UserList(UserIndex).Invent.MunicionEqpSlot = 0

UserList(UserIndex).Invent.BarcoObjIndex = 0
UserList(UserIndex).Invent.BarcoSlot = 0

End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal UserIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 23/01/2007
'23/01/2007 -> Pablo (ToxicWaste): Billetera invertida y explotar oro en el agua.
'***************************************************
On Error GoTo Errhandler

'If Cantidad > 100000 Then Exit Sub

'SI EL Pjta TIENE ORO LO TIRAMOS
If (Cantidad > 0) And (Cantidad <= UserList(UserIndex).Stats.GLD) Then
        Dim i As Byte
        Dim MiObj As obj
        'info debug
        Dim loops As Integer
        
        'Seguridad Alkon (guardo el oro tirado si supera los 50k)
        If Cantidad > 50000 Then
            Dim j As Integer
            Dim k As Integer
            Dim M As Integer
            Dim Cercanos As String
            M = UserList(UserIndex).Pos.map
            For j = UserList(UserIndex).Pos.X - 10 To UserList(UserIndex).Pos.X + 10
                For k = UserList(UserIndex).Pos.Y - 10 To UserList(UserIndex).Pos.Y + 10
                    If InMapBounds(M, j, k) Then
                        If MapData(M, j, k).UserIndex > 0 Then
                            Cercanos = Cercanos & UserList(MapData(M, j, k).UserIndex).name & ","
                        End If
                    End If
                Next k
            Next j
            Call LogDesarrollo(UserList(UserIndex).name & " tira oro. Cercanos: " & Cercanos)
        End If
        '/Seguridad
        Dim Extra As Long
        Dim TeniaOro As Long
        TeniaOro = UserList(UserIndex).Stats.GLD
        If Cantidad > 500000 Then 'Para evitar explotar demasiado
            Extra = Cantidad - 500000
            Cantidad = 500000
        End If
        
        Do While (Cantidad > 0)
            
            If Cantidad > MAX_INVENTORY_OBJS And UserList(UserIndex).Stats.GLD > MAX_INVENTORY_OBJS Then
                MiObj.amount = MAX_INVENTORY_OBJS
                Cantidad = Cantidad - MiObj.amount
            Else
                MiObj.amount = Cantidad
                Cantidad = Cantidad - MiObj.amount
            End If

            MiObj.objIndex = iORO
            
            If EsGM(UserIndex) Then Call LogGM(UserList(UserIndex).name, "Tiro cantidad:" & MiObj.amount & " Objeto:" & ObjData(MiObj.objIndex).name)
            Dim AuxPos As WorldPos
            
            If UserList(UserIndex).clase = eClass.Pirat And UserList(UserIndex).Invent.BarcoObjIndex = 476 Then
                AuxPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj, False)
                If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MiObj.amount
                End If
            Else
                AuxPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj, True)
                If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MiObj.amount
                End If
            End If
            
            'info debug
            loops = loops + 1
            If loops > 100 Then
                LogError ("Error en tiraroro")
                Exit Sub
            End If
            
        Loop
        If TeniaOro = UserList(UserIndex).Stats.GLD Then Extra = 0
        If Extra > 0 Then
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Extra
        End If
    
End If

Exit Sub

Errhandler:

End Sub

Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)

On Error GoTo Errhandler

    If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub
    
    With UserList(UserIndex).Invent.Object(Slot)
        If .amount <= Cantidad And .Equipped = 1 Then
            Call Desequipar(UserIndex, Slot)
        End If
        
        'Quita un objeto
        .amount = .amount - Cantidad
        '¿Quedan mas?
        If .amount <= 0 Then
            UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
            .objIndex = 0
            .amount = 0
        End If
    End With

Exit Sub

Errhandler:
    Call LogError("Error en QuitarUserInvItem. Error " & err.Number & " : " & err.description)
    
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)

On Error GoTo Errhandler

Dim NullObj As UserOBJ
Dim LoopC As Long

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(UserIndex).Invent.Object(Slot).objIndex > 0 Then
        Call ChangeUserInv(UserIndex, Slot, UserList(UserIndex).Invent.Object(Slot))
    Else
        Call ChangeUserInv(UserIndex, Slot, NullObj)
    End If

Else

'Actualiza todos los slots
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        'Actualiza el inventario
        If UserList(UserIndex).Invent.Object(LoopC).objIndex > 0 Then
            Call ChangeUserInv(UserIndex, LoopC, UserList(UserIndex).Invent.Object(LoopC))
        Else
            Call ChangeUserInv(UserIndex, LoopC, NullObj)
        End If
    Next LoopC
End If

Exit Sub

Errhandler:
    Call LogError("Error en UpdateUserInv. Error " & err.Number & " : " & err.description)

End Sub

Sub DropObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal num As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Error GoTo err

10 Dim obj As obj

20 If num > 0 Then
  
30  If num > UserList(UserIndex).Invent.Object(Slot).amount Then num = UserList(UserIndex).Invent.Object(Slot).amount
  
  'Check objeto en el suelo
40  If MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.objIndex = 0 Or MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.objIndex = UserList(UserIndex).Invent.Object(Slot).objIndex Then
50        obj.objIndex = UserList(UserIndex).Invent.Object(Slot).objIndex
        
60        If num + MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.amount > MAX_INVENTORY_OBJS Then
70            num = MAX_INVENTORY_OBJS - MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.amount
80        End If
        
90        obj.amount = num
        
100        Call MakeObj(obj, map, X, Y)
110        Call QuitarUserInvItem(UserIndex, Slot, num)
120        Call UpdateUserInv(False, UserIndex, Slot)
        
130        If ObjData(obj.objIndex).OBJType = eOBJType.otBarcos Then
140            Call WriteConsoleMsg(UserIndex, "¡¡ATENCION!! ¡ACABAS DE TIRAR TU BARCA!", FontTypeNames.FONTTYPE_TALK)
150        End If
        
160        If Not UserList(UserIndex).flags.Privilegios And PlayerType.User Then Call LogGM(UserList(UserIndex).name, "Tiro cantidad:" & num & " Objeto:" & ObjData(obj.objIndex).name)
        
        'Log de Objetos que se tiran al piso. Pablo (ToxicWaste) 07/09/07
        'Es un Objeto que tenemos que loguear?
170        If ObjData(obj.objIndex).Log = 1 Then
180            Call LogDesarrollo(UserList(UserIndex).name & " tiró al piso " & obj.amount & " " & ObjData(obj.objIndex).name & " Mapa: " & map & " X: " & X & " Y: " & Y)
190        ElseIf obj.amount > 5000 Then 'Es mucha cantidad? > Subí a 5000 el minimo porque si no se llenaba el log de cosas al pedo. (NicoNZ)
        'Si no es de los prohibidos de loguear, lo logueamos.
200            If ObjData(obj.objIndex).NoLog <> 1 Then
210                Call LogDesarrollo(UserList(UserIndex).name & " tiró al piso " & obj.amount & " " & ObjData(obj.objIndex).name & " Mapa: " & map & " X: " & X & " Y: " & Y)
220            End If
230        End If
240  Else
250    Call WriteConsoleMsg(UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
260  End If
    
270 End If

Exit Sub

err:

Debug.Print "Error: " & Erl

End Sub

Sub EraseObj(ByVal num As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)

MapData(map, X, Y).ObjInfo.amount = MapData(map, X, Y).ObjInfo.amount - num

If MapData(map, X, Y).ObjInfo.amount <= 0 Then

    Call removeObjArea(MapData(map, X, Y).ObjInfo.objIndex, map, X, Y)

    MapData(map, X, Y).ObjInfo.objIndex = 0
    MapData(map, X, Y).ObjInfo.amount = 0
    
    Call modSendData.SendToAreaByPos(PrepareMessageObjectDelete(X, Y), map, X, Y)
End If

End Sub

Sub MakeObj(ByRef obj As obj, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Error GoTo err

10 If obj.objIndex > 0 And obj.objIndex <= UBound(ObjData) Then

20    If MapData(map, X, Y).ObjInfo.objIndex = obj.objIndex Then
30        MapData(map, X, Y).ObjInfo.amount = MapData(map, X, Y).ObjInfo.amount + obj.amount
40    Else
    
50        MapData(map, X, Y).ObjInfo = obj
        
60        Call addObjArea(obj.objIndex, obj.amount, map, X, Y)
        
70        Call modSendData.SendToAreaByPos(PrepareMessageObjectCreate(ObjData(obj.objIndex).GrhIndex, X, Y), map, X, Y)
80    End If
90 End If

Exit Sub

err:

Debug.Print "Error makeObj: " & Erl

End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As obj) As Boolean
On Error GoTo Errhandler

'Call LogTarea("MeterItemEnInventario")
 
Dim X As Integer
Dim Y As Integer
Dim Slot As Byte

'¿el user ya tiene un objeto del mismo tipo?
Slot = 1
Do Until UserList(UserIndex).Invent.Object(Slot).objIndex = MiObj.objIndex And _
         UserList(UserIndex).Invent.Object(Slot).amount + MiObj.amount <= MAX_INVENTORY_OBJS
   Slot = Slot + 1
   If Slot > MAX_INVENTORY_SLOTS Then
         Exit Do
   End If
Loop
    
'Sino busca un slot vacio
If Slot > MAX_INVENTORY_SLOTS Then
   Slot = 1
   Do Until UserList(UserIndex).Invent.Object(Slot).objIndex = 0
       Slot = Slot + 1
       If Slot > MAX_INVENTORY_SLOTS Then
           Call WriteConsoleMsg(UserIndex, "No podes cargar mas objetos.", FontTypeNames.FONTTYPE_FIGHT)
           MeterItemEnInventario = False
           Exit Function
       End If
   Loop
   UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
End If
    
'Mete el objeto
If UserList(UserIndex).Invent.Object(Slot).amount + MiObj.amount <= MAX_INVENTORY_OBJS Then
   'Menor que MAX_INV_OBJS
   UserList(UserIndex).Invent.Object(Slot).objIndex = MiObj.objIndex
   UserList(UserIndex).Invent.Object(Slot).amount = UserList(UserIndex).Invent.Object(Slot).amount + MiObj.amount
Else
   UserList(UserIndex).Invent.Object(Slot).amount = MAX_INVENTORY_OBJS
End If
    
MeterItemEnInventario = True
       
Call UpdateUserInv(False, UserIndex, Slot)


Exit Function
Errhandler:

End Function


Sub GetObj(ByVal UserIndex As Integer)

Dim obj As ObjData
Dim MiObj As obj
Dim ObjPos As String

'¿Hay algun obj?
If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.objIndex > 0 Then
    '¿Esta permitido agarrar este obj?
    If ObjData(MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.objIndex).Agarrable <> 1 Then
        Dim X As Integer
        Dim Y As Integer
        Dim Slot As Byte
        
        X = UserList(UserIndex).Pos.X
        Y = UserList(UserIndex).Pos.Y
        obj = ObjData(MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.objIndex)
        MiObj.amount = MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.amount
        MiObj.objIndex = MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.objIndex
        
        If MeterItemEnInventario(UserIndex, MiObj) Then
            'Quitamos el objeto
            Call EraseObj(MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.amount, UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
            If Not UserList(UserIndex).flags.Privilegios And PlayerType.User Then Call LogGM(UserList(UserIndex).name, "Agarro:" & MiObj.amount & " Objeto:" & ObjData(MiObj.objIndex).name)
            
            'Log de Objetos que se agarran del piso. Pablo (ToxicWaste) 07/09/07
            'Es un Objeto que tenemos que loguear?
            If ObjData(MiObj.objIndex).Log = 1 Then
                ObjPos = " Mapa: " & UserList(UserIndex).Pos.map & " X: " & UserList(UserIndex).Pos.X & " Y: " & UserList(UserIndex).Pos.Y
                Call LogDesarrollo(UserList(UserIndex).name & " juntó del piso " & MiObj.amount & " " & ObjData(MiObj.objIndex).name & ObjPos)
            ElseIf MiObj.amount > MAX_INVENTORY_OBJS - 1000 Then 'Es mucha cantidad?
                'Si no es de los prohibidos de loguear, lo logueamos.
                If ObjData(MiObj.objIndex).NoLog <> 1 Then
                    ObjPos = " Mapa: " & UserList(UserIndex).Pos.map & " X: " & UserList(UserIndex).Pos.X & " Y: " & UserList(UserIndex).Pos.Y
                    Call LogDesarrollo(UserList(UserIndex).name & " juntó del piso " & MiObj.amount & " " & ObjData(MiObj.objIndex).name & ObjPos)
                End If
            End If
        End If
        
    End If
Else
    Call WriteConsoleMsg(UserIndex, "No hay nada aquí.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Sub Desequipar(ByVal UserIndex As Integer, ByVal Slot As Byte)

On Error GoTo Errhandler

'Desequipa el item slot del inventario
Dim obj As ObjData


If (Slot < LBound(UserList(UserIndex).Invent.Object)) Or (Slot > UBound(UserList(UserIndex).Invent.Object)) Then
    Exit Sub
ElseIf UserList(UserIndex).Invent.Object(Slot).objIndex = 0 Then
    Exit Sub
End If

obj = ObjData(UserList(UserIndex).Invent.Object(Slot).objIndex)

Select Case obj.OBJType
    Case eOBJType.otWeapon
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
        UserList(UserIndex).Invent.WeaponEqpSlot = 0
        If Not UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).Char.WeaponAnim = NingunArma
            Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        End If
    
    Case eOBJType.otFlechas
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
        UserList(UserIndex).Invent.MunicionEqpSlot = 0
    
    Case eOBJType.otAnillo
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
        UserList(UserIndex).Invent.AnilloEqpSlot = 0
    
    Case eOBJType.otArmadura
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
        UserList(UserIndex).Invent.ArmourEqpSlot = 0
        Call DarCuerpoDesnudo(UserIndex, UserList(UserIndex).flags.Mimetizado = 1)
        Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            
    Case eOBJType.otCASCO
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.CascoEqpObjIndex = 0
        UserList(UserIndex).Invent.CascoEqpSlot = 0
        If Not UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).Char.CascoAnim = NingunCasco
            Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        End If
    
    Case eOBJType.otESCUDO
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
        UserList(UserIndex).Invent.EscudoEqpSlot = 0
        If Not UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).Char.ShieldAnim = NingunEscudo
            Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        End If
End Select

Call WriteUpdateUserStats(UserIndex)
Call UpdateUserInv(False, UserIndex, Slot)

Exit Sub

Errhandler:
    Call LogError("Error en Desquipar. Error " & err.Number & " : " & err.description)

End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, ByVal objIndex As Integer) As Boolean
On Error GoTo Errhandler

If ObjData(objIndex).Mujer = 1 Then
    SexoPuedeUsarItem = UserList(UserIndex).genero <> eGenero.Hombre
ElseIf ObjData(objIndex).Hombre = 1 Then
    SexoPuedeUsarItem = UserList(UserIndex).genero <> eGenero.Mujer
Else
    SexoPuedeUsarItem = True
End If

Exit Function
Errhandler:
    Call LogError("SexoPuedeUsarItem")
End Function


Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, ByVal objIndex As Integer) As Boolean

If ObjData(objIndex).Real = 1 Then
    If Not criminal(UserIndex) Then
        FaccionPuedeUsarItem = esArmada(UserIndex)
    Else
        FaccionPuedeUsarItem = False
    End If
ElseIf ObjData(objIndex).Caos = 1 Then
    If criminal(UserIndex) Then
        FaccionPuedeUsarItem = esCaos(UserIndex)
    Else
        FaccionPuedeUsarItem = False
    End If
Else
    FaccionPuedeUsarItem = True
End If

End Function

Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
'*************************************************
'Author: Unknown
'Last modified: 01/08/2009
'01/08/2009: ZaMa - Now it's not sent any sound made by an invisible admin
'*************************************************

On Error GoTo Errhandler

'Equipa un item del inventario
Dim obj As ObjData
Dim objIndex As Integer

objIndex = UserList(UserIndex).Invent.Object(Slot).objIndex
obj = ObjData(objIndex)

If obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
     Call WriteConsoleMsg(UserIndex, "Solo los newbies pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)
     Exit Sub
End If
        
Select Case obj.OBJType
    Case eOBJType.otWeapon
       If ClasePuedeUsarItem(UserIndex, objIndex) And _
          FaccionPuedeUsarItem(UserIndex, objIndex) Then
            'Si esta equipado lo quita
            If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                'Quitamos del inv el item
                Call Desequipar(UserIndex, Slot)
                'Animacion por defecto
                If UserList(UserIndex).flags.Mimetizado = 1 Then
                    UserList(UserIndex).CharMimetizado.WeaponAnim = NingunArma
                Else
                    UserList(UserIndex).Char.WeaponAnim = NingunArma
                    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                End If
                Exit Sub
            End If
            
            'Quitamos el elemento anterior
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
            End If
            
            UserList(UserIndex).Invent.Object(Slot).Equipped = 1
            UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).objIndex
            UserList(UserIndex).Invent.WeaponEqpSlot = Slot
            
            'El sonido solo se envia si no lo produce un admin invisible
            If Not (UserList(UserIndex).flags.AdminInvisible = 1) Then _
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SACARARMA, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
            
            If UserList(UserIndex).flags.Mimetizado = 1 Then
                UserList(UserIndex).CharMimetizado.WeaponAnim = obj.WeaponAnim
            Else
                UserList(UserIndex).Char.WeaponAnim = obj.WeaponAnim
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            End If
       Else
            Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
    
    Case eOBJType.otAnillo
       If ClasePuedeUsarItem(UserIndex, objIndex) And _
          FaccionPuedeUsarItem(UserIndex, objIndex) Then
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.AnilloEqpSlot)
                End If
        
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.AnilloEqpObjIndex = objIndex
                UserList(UserIndex).Invent.AnilloEqpSlot = Slot
                
       Else
            Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
    
    Case eOBJType.otFlechas
       If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).objIndex) And _
          FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).objIndex) Then
                
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
                End If
        
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).objIndex
                UserList(UserIndex).Invent.MunicionEqpSlot = Slot
                
       Else
            Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
    
    Case eOBJType.otArmadura
        If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
        'Nos aseguramos que puede usarla
        If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).objIndex) And _
           SexoPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).objIndex) And _
           CheckRazaUsaRopa(UserIndex, UserList(UserIndex).Invent.Object(Slot).objIndex) And _
           FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).objIndex) Then
           
           'Si esta equipado lo quita
            If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                Call Desequipar(UserIndex, Slot)
                Call DarCuerpoDesnudo(UserIndex, UserList(UserIndex).flags.Mimetizado = 1)
                If Not UserList(UserIndex).flags.Mimetizado = 1 Then
                    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                End If
                Exit Sub
            End If
    
            'Quita el anterior
            If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
            End If
    
            'Lo equipa
            UserList(UserIndex).Invent.Object(Slot).Equipped = 1
            UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).objIndex
            UserList(UserIndex).Invent.ArmourEqpSlot = Slot
                
            If UserList(UserIndex).flags.Mimetizado = 1 Then
                UserList(UserIndex).CharMimetizado.body = obj.Ropaje
            Else
                UserList(UserIndex).Char.body = obj.Ropaje
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            End If
            UserList(UserIndex).flags.Desnudo = 0
            

        Else
            Call WriteConsoleMsg(UserIndex, "Tu clase,genero o raza no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
        End If
    
    Case eOBJType.otCASCO
        If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
        If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).objIndex) Then
            'Si esta equipado lo quita
            If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                Call Desequipar(UserIndex, Slot)
                If UserList(UserIndex).flags.Mimetizado = 1 Then
                    UserList(UserIndex).CharMimetizado.CascoAnim = NingunCasco
                Else
                    UserList(UserIndex).Char.CascoAnim = NingunCasco
                    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                End If
                Exit Sub
            End If
    
            'Quita el anterior
            If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)
            End If
    
            'Lo equipa
            
            UserList(UserIndex).Invent.Object(Slot).Equipped = 1
            UserList(UserIndex).Invent.CascoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).objIndex
            UserList(UserIndex).Invent.CascoEqpSlot = Slot
            If UserList(UserIndex).flags.Mimetizado = 1 Then
                UserList(UserIndex).CharMimetizado.CascoAnim = obj.CascoAnim
            Else
                UserList(UserIndex).Char.CascoAnim = obj.CascoAnim
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
        End If
    
    Case eOBJType.otESCUDO
        If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
         If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).objIndex) And _
             FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).objIndex) Then

             'Si esta equipado lo quita
             If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                 Call Desequipar(UserIndex, Slot)
                 If UserList(UserIndex).flags.Mimetizado = 1 Then
                     UserList(UserIndex).CharMimetizado.ShieldAnim = NingunEscudo
                 Else
                     UserList(UserIndex).Char.ShieldAnim = NingunEscudo
                     Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                 End If
                 Exit Sub
             End If
     
             'Quita el anterior
             If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
                 Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
             End If
     
             'Lo equipa
             
             UserList(UserIndex).Invent.Object(Slot).Equipped = 1
             UserList(UserIndex).Invent.EscudoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).objIndex
             UserList(UserIndex).Invent.EscudoEqpSlot = Slot
             
             If UserList(UserIndex).flags.Mimetizado = 1 Then
                 UserList(UserIndex).CharMimetizado.ShieldAnim = obj.ShieldAnim
             Else
                 UserList(UserIndex).Char.ShieldAnim = obj.ShieldAnim
                 
                 Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
             End If
         Else
             Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
         End If
End Select

'Actualiza
Call UpdateUserInv(False, UserIndex, Slot)

Exit Sub
Errhandler:
Call LogError("EquiparInvItem Slot:" & Slot & " - Error: " & err.Number & " - Error Description : " & err.description)
End Sub

Private Function CheckRazaUsaRopa(ByVal UserIndex As Integer, ItemIndex As Integer) As Boolean
On Error GoTo Errhandler

'Verifica si la raza puede usar la ropa
If UserList(UserIndex).raza = eRaza.Humano Or _
   UserList(UserIndex).raza = eRaza.Elfo Or _
   UserList(UserIndex).raza = eRaza.Drow Then
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
Else
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
End If

'Solo se habilita la ropa exclusiva para Drows por ahora. Pablo (ToxicWaste)
If (UserList(UserIndex).raza <> eRaza.Drow) And ObjData(ItemIndex).RazaDrow Then
    CheckRazaUsaRopa = False
End If

Exit Function
Errhandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
'*************************************************
'Author: Unknown
'Last modified: 01/08/2009
'Handels the usage of items from inventory box.
'24/01/2007 Pablo (ToxicWaste) - Agrego el Cuerno de la Armada y la Legión.
'24/01/2007 Pablo (ToxicWaste) - Utilización nueva de Barco en lvl 20 por clase Pirata y Pescador.
'01/08/2009: ZaMa - Now it's not sent any sound made by an invisible admin, except to its own client
'*************************************************

Dim obj As ObjData
Dim objIndex As Integer
Dim TargObj As ObjData
Dim MiObj As obj

With UserList(UserIndex)

If .Invent.Object(Slot).amount = 0 Then Exit Sub

obj = ObjData(.Invent.Object(Slot).objIndex)

If obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
    Call WriteConsoleMsg(UserIndex, "Solo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If obj.OBJType = eOBJType.otWeapon Then
    If obj.proyectil = 1 Then
        If Not .flags.ModoCombate Then
            Call WriteConsoleMsg(UserIndex, "No estás en modo de combate, presiona la tecla ""C"" para pasar al modo combate.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
        If Not IntervaloPermiteUsar(UserIndex, False) Then Exit Sub
    Else
        'dagas
        If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
    End If
Else
    If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
End If

objIndex = .Invent.Object(Slot).objIndex
.flags.TargetObjInvIndex = objIndex
.flags.TargetObjInvSlot = Slot

Select Case obj.OBJType
    Case eOBJType.otUseOnce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Usa el item
        .Stats.MinHam = .Stats.MinHam + obj.MinHam
        If .Stats.MinHam > .Stats.MaxHam Then _
            .Stats.MinHam = .Stats.MaxHam
        .flags.Hambre = 0
        Call WriteUpdateHungerAndThirst(UserIndex)
        'Sonido
        
        If objIndex = e_ObjetosCriticos.Manzana Or objIndex = e_ObjetosCriticos.Manzana2 Or objIndex = e_ObjetosCriticos.ManzanaNewbie Then
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MORFAR_MANZANA)
        Else
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.SOUND_COMIDA)
        End If
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        
        Call UpdateUserInv(False, UserIndex, Slot)

    Case eOBJType.otGuita
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        .Stats.GLD = .Stats.GLD + .Invent.Object(Slot).amount
        .Invent.Object(Slot).amount = 0
        .Invent.Object(Slot).objIndex = 0
        .Invent.NroItems = .Invent.NroItems - 1
        
        Call UpdateUserInv(False, UserIndex, Slot)
        Call WriteUpdateGold(UserIndex)
        
    Case eOBJType.otWeapon
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Not .Stats.MinSta > 0 Then
            If .genero = eGenero.Hombre Then
                Call WriteConsoleMsg(UserIndex, "Estas muy cansado", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Estas muy cansada", FontTypeNames.FONTTYPE_INFO)
            End If
            Exit Sub
        End If
        
        
        If ObjData(objIndex).proyectil = 1 Then
            If .Invent.Object(Slot).Equipped = 0 Then
                Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            'liquid: muevo esto aca adentro, para que solo pida modo combate si estamos por usar el arco
            If Not .flags.ModoCombate Then
                Call WriteConsoleMsg(UserIndex, "No estás en modo de combate, presiona la tecla ""C"" para pasar al modo combate.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            Call WriteWorkRequestTarget(UserIndex, Proyectiles)
        Else
            If .flags.TargetObj = Leña Then
                If .Invent.Object(Slot).objIndex = DAGA Then
                    If .Invent.Object(Slot).Equipped = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    Call TratarDeHacerFogata(.flags.TargetObjMap, _
                         .flags.TargetObjX, .flags.TargetObjY, UserIndex)
                End If
            End If
        End If

        
        Select Case objIndex
            Case CAÑA_PESCA, RED_PESCA
                Call WriteWorkRequestTarget(UserIndex, eSkill.Pesca)
            Case HACHA_LEÑADOR
                Call WriteWorkRequestTarget(UserIndex, eSkill.Talar)
            Case PIQUETE_MINERO
                Call WriteWorkRequestTarget(UserIndex, eSkill.Mineria)
            Case MARTILLO_HERRERO
                Call WriteWorkRequestTarget(UserIndex, eSkill.Herreria)
            Case SERRUCHO_CARPINTERO
                Call EnivarObjConstruibles(UserIndex)
                Call WriteShowCarpenterForm(UserIndex)
        End Select
        
    
    Case eOBJType.otPociones
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Not IntervaloPermiteGolpeUsar(UserIndex, False) Then
            Call WriteConsoleMsg(UserIndex, "¡¡Debes esperar unos momentos para tomar otra poción!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        .flags.TomoPocion = True
        .flags.TipoPocion = obj.TipoPocion
                
        Select Case .flags.TipoPocion
        
            Case 1 'Modif la agilidad
                .flags.DuracionEfecto = obj.DuracionEfecto
        
                'Usa el item
                .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(obj.MinModificador, obj.MaxModificador)
                If .Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then _
                    .Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
                If .Stats.UserAtributos(eAtributos.Agilidad) > 2 * .Stats.UserAtributosBackUP(Agilidad) Then .Stats.UserAtributos(eAtributos.Agilidad) = 2 * .Stats.UserAtributosBackUP(Agilidad)
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                End If
        
            Case 2 'Modif la fuerza
                .flags.DuracionEfecto = obj.DuracionEfecto
        
                'Usa el item
                .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(obj.MinModificador, obj.MaxModificador)
                If .Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then _
                    .Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
                If .Stats.UserAtributos(eAtributos.Fuerza) > 2 * .Stats.UserAtributosBackUP(Fuerza) Then .Stats.UserAtributos(eAtributos.Fuerza) = 2 * .Stats.UserAtributosBackUP(Fuerza)
                
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                End If
                
            Case 3 'Pocion roja, restaura HP
                'Usa el item
                .Stats.MinHP = .Stats.MinHP + RandomNumber(obj.MinModificador, obj.MaxModificador)
                If .Stats.MinHP > .Stats.MaxHP Then _
                    .Stats.MinHP = .Stats.MaxHP
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                End If
            
            Case 4 'Pocion azul, restaura MANA
                'Usa el item
                'nuevo calculo para recargar mana
                .Stats.MinMAN = .Stats.MinMAN + Porcentaje(.Stats.MaxMAN, 4) + .Stats.ELV \ 2 + 40 / .Stats.ELV
                If .Stats.MinMAN > .Stats.MaxMAN Then _
                    .Stats.MinMAN = .Stats.MaxMAN
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                End If
                
            Case 5 ' Pocion violeta
                If .flags.Envenenado = 1 Then
                    .flags.Envenenado = 0
                    Call WriteConsoleMsg(UserIndex, "Te has curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)
                End If
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                End If
                
            Case 6  ' Pocion Negra
                If .flags.Privilegios And PlayerType.User Then
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call UserDie(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Sientes un gran mareo y pierdes el conocimiento.", FontTypeNames.FONTTYPE_FIGHT)
                End If
       End Select
       Call WriteUpdateUserStats(UserIndex)
       Call UpdateUserInv(False, UserIndex, Slot)

     Case eOBJType.otBebidas
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        .Stats.MinAGU = .Stats.MinAGU + obj.MinSed
        If .Stats.MinAGU > .Stats.MaxAGU Then _
            .Stats.MinAGU = .Stats.MaxAGU
        .flags.Sed = 0
        Call WriteUpdateHungerAndThirst(UserIndex)
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        
        ' Los admin invisibles solo producen sonidos a si mismos
        If .flags.AdminInvisible = 1 Then
            Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
        End If
        
        Call UpdateUserInv(False, UserIndex, Slot)
    
    Case eOBJType.otLlaves
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.TargetObj = 0 Then Exit Sub
        TargObj = ObjData(.flags.TargetObj)
        '¿El objeto clickeado es una puerta?
        If TargObj.OBJType = eOBJType.otPuertas Then
            '¿Esta cerrada?
            If TargObj.Cerrada = 1 Then
                  '¿Cerrada con llave?
                  If TargObj.Llave > 0 Then
                     If TargObj.clave = obj.clave Then
         
                        MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.objIndex _
                        = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.objIndex).IndexCerrada
                        .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.objIndex
                        Call WriteConsoleMsg(UserIndex, "Has abierto la puerta.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     Else
                        Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     End If
                  Else
                     If TargObj.clave = obj.clave Then
                        MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.objIndex _
                        = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.objIndex).IndexCerradaLlave
                        Call WriteConsoleMsg(UserIndex, "Has cerrado con llave la puerta.", FontTypeNames.FONTTYPE_INFO)
                        .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.objIndex
                        Exit Sub
                     Else
                        Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     End If
                  End If
            Else
                  Call WriteConsoleMsg(UserIndex, "No esta cerrada.", FontTypeNames.FONTTYPE_INFO)
                  Exit Sub
            End If
        End If
    
    Case eOBJType.otBotellaVacia
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Not HayAgua(.Pos.map, .flags.TargetX, .flags.TargetY) Then
            Call WriteConsoleMsg(UserIndex, "No hay agua allí.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        MiObj.amount = 1
        MiObj.objIndex = ObjData(.Invent.Object(Slot).objIndex).IndexAbierta
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)
        End If
        
        Call UpdateUserInv(False, UserIndex, Slot)
    
    Case eOBJType.otBotellaLlena
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        .Stats.MinAGU = .Stats.MinAGU + obj.MinSed
        If .Stats.MinAGU > .Stats.MaxAGU Then _
            .Stats.MinAGU = .Stats.MaxAGU
        .flags.Sed = 0
        Call WriteUpdateHungerAndThirst(UserIndex)
        MiObj.amount = 1
        MiObj.objIndex = ObjData(.Invent.Object(Slot).objIndex).IndexCerrada
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)
        End If
        
        Call UpdateUserInv(False, UserIndex, Slot)
    
    Case eOBJType.otPergaminos
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .Stats.MaxMAN > 0 Then
            If .flags.Hambre = 0 And _
                .flags.Sed = 0 Then
                Call AgregarHechizo(UserIndex, Slot)
                Call UpdateUserInv(False, UserIndex, Slot)
            Else
                Call WriteConsoleMsg(UserIndex, "Estas demasiado hambriento y sediento.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de las Artes Arcanas.", FontTypeNames.FONTTYPE_INFO)
        End If
    Case eOBJType.otMinerales
        If .flags.Muerto = 1 Then
             Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        Call WriteWorkRequestTarget(UserIndex, FundirMetal)
       
    Case eOBJType.otInstrumentos
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If obj.Real Then '¿Es el Cuerno Real?
            If FaccionPuedeUsarItem(UserIndex, objIndex) Then
                If MapInfo(.Pos.map).Pk = False Then
                    Call WriteConsoleMsg(UserIndex, "No hay Peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.toMap, .Pos.map, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                End If
                
                Exit Sub
            Else
                Call WriteConsoleMsg(UserIndex, "Solo Miembros de la Armada Real pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        ElseIf obj.Caos Then '¿Es el Cuerno Legión?
            If FaccionPuedeUsarItem(UserIndex, objIndex) Then
                If MapInfo(.Pos.map).Pk = False Then
                    Call WriteConsoleMsg(UserIndex, "No hay Peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.toMap, .Pos.map, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                End If
                
                Exit Sub
            Else
                Call WriteConsoleMsg(UserIndex, "Solo Miembros de la Legión Oscura pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        'Si llega aca es porque es o Laud o Tambor o Flauta
        ' Los admin invisibles solo producen sonidos a si mismos
        If .flags.AdminInvisible = 1 Then
            Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
        End If
       
    Case eOBJType.otBarcos
        'Verifica si esta aproximado al agua antes de permitirle navegar
        If .Stats.ELV < 25 Then
            If .clase <> eClass.Fisher And .clase <> eClass.Pirat Then
                Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 25 o superior.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            Else
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 20 o superior.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
        End If
        
        If ((LegalPos(.Pos.map, .Pos.X - 1, .Pos.Y, True, False) _
                Or LegalPos(.Pos.map, .Pos.X, .Pos.Y - 1, True, False) _
                Or LegalPos(.Pos.map, .Pos.X + 1, .Pos.Y, True, False) _
                Or LegalPos(.Pos.map, .Pos.X, .Pos.Y + 1, True, False)) _
                And .flags.Navegando = 0) _
                Or .flags.Navegando = 1 Then
            Call DoNavega(UserIndex, obj, Slot)
        Else
            Call WriteConsoleMsg(UserIndex, "¡Debes aproximarte al agua para usar el barco!", FontTypeNames.FONTTYPE_INFO)
        End If
End Select

End With

End Sub

Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)

Call WriteBlacksmithWeapons(UserIndex)

End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)

Call WriteCarpenterObjects(UserIndex)

End Sub

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)

Call WriteBlacksmithArmors(UserIndex)

End Sub

Sub TirarTodo(ByVal UserIndex As Integer)
On Error Resume Next

If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub

Call TirarTodosLosItems(UserIndex)

Dim Cantidad As Long
Cantidad = UserList(UserIndex).Stats.GLD - CLng(UserList(UserIndex).Stats.ELV) * 10000

If Cantidad > 0 Then _
    Call TirarOro(Cantidad, UserIndex)

End Sub

Public Function ItemSeCae(ByVal index As Integer) As Boolean

ItemSeCae = (ObjData(index).Real <> 1 Or ObjData(index).NoSeCae = 0) And _
            (ObjData(index).Caos <> 1 Or ObjData(index).NoSeCae = 0) And _
            ObjData(index).OBJType <> eOBJType.otLlaves And _
            ObjData(index).OBJType <> eOBJType.otBarcos And _
            ObjData(index).NoSeCae = 0


End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)
    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As obj
    Dim ItemIndex As Integer
    
    For i = 1 To MAX_INVENTORY_SLOTS
        ItemIndex = UserList(UserIndex).Invent.Object(i).objIndex
        If ItemIndex > 0 Then
             If ItemSeCae(ItemIndex) Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
                
                'Creo el Obj
                MiObj.amount = UserList(UserIndex).Invent.Object(i).amount
                MiObj.objIndex = ItemIndex
                'Pablo (ToxicWaste) 24/01/2007
                'Si es pirata y usa un Galeón entonces no explota los items. (en el agua)
                If UserList(UserIndex).clase = eClass.Pirat And UserList(UserIndex).Invent.BarcoObjIndex = 476 Then
                    Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj, False, True
                Else
                    Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj, True, True
                End If
                
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.X, NuevaPos.Y)
                End If
             End If
        End If
    Next i
End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean

ItemNewbie = ObjData(ItemIndex).Newbie = 1

End Function

Sub TirarTodosLosItemsNoNewbies(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 07/11/09
'07/11/09: Pato - Fix bug #2819911
'***************************************************
Dim i As Byte
Dim NuevaPos As WorldPos
Dim MiObj As obj
Dim ItemIndex As Integer

If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub

For i = 1 To MAX_INVENTORY_SLOTS
    ItemIndex = UserList(UserIndex).Invent.Object(i).objIndex
    If ItemIndex > 0 Then
        If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
            NuevaPos.X = 0
            NuevaPos.Y = 0
            
            'Creo MiObj
            MiObj.amount = UserList(UserIndex).Invent.Object(i).amount
            MiObj.objIndex = ItemIndex
            'Pablo (ToxicWaste) 24/01/2007
            'Tira los Items no newbies en todos lados.
            Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj, True, True
            If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.X, NuevaPos.Y)
            End If
        End If
    End If
Next i

End Sub
