VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Optimizador de areas"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "y"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "x"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inserte las coordenadas y luego presione 'aceptar'"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const X_LIMITER = 20
Private Const Y_LIMITER = 20


Private Sub Command1_Click()

Dim x       As Long
Dim y       As Long
Dim minX    As Long
Dim minY    As Long
Dim countY  As Long
Dim countX  As Long
Dim writeX  As Boolean
Dim writeY  As Boolean
Dim lastX   As Long
Dim lastY   As Long

Dim dm  As Single
Dim ent()   As String

For x = 1 To 100

    For y = 1 To 100
    
        countY = countY + 1
        
        If countY >= Y_LIMITER Then
            countY = 0
            writeY = True
        End If
    
        If writeY Then
        
            If lastY >= 100 Then lastY = 0
        
            minY = lastY + 1
            lastY = y
            
            dm = y / Y_LIMITER
            ent = Split(CStr(dm), ".")
            
            WriteVar App.Path & "\areaspos.ini", "POSTOAREA", "y-" & minY & "-" & lastY, CStr(ent(0))
        
            writeY = False
        End If
    
    Next y
    
        countX = countX + 1
        
        If countX >= X_LIMITER Then
            countX = 0
            writeX = True
        End If
    
        If writeX Then
        
            If lastX >= 100 Then lastX = 0
        
            minX = lastX + 1
            lastX = x
            
            dm = y / Y_LIMITER
            ent = Split(CStr(dm), ".")
            
            WriteVar App.Path & "\areaspos.ini", "POSTOAREA", "x-" & minY & "-" & lastY, CStr(ent(0))
        
            writeX = False
        End If

Next x


End Sub

Private Sub Form_Load()
    areasInitialize
End Sub
