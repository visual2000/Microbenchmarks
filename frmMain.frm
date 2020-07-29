VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLog 
      Height          =   5415
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const iterations = 10000000
Dim i As Long

Private Sub Form_Load()
    log ("Starting up. # of Iterations = " & CStr(iterations))
    Call DoBenchmarks
End Sub

Private Sub DoBenchmarks()
    Call StructVsPrimitive
    Call LocalVsFunction
    Call testGoTo
End Sub

Private Sub LocalVsFunction()
' We'll initialise a struct in a local loop, as well as with our TVec3_init() function
    Dim ticks As Long, totalTime As Long
    
    log ("Local...")
    ticks = GetTickCount
    For i = 0 To iterations
        Dim t As TVec3
        t.x = 2#
        t.y = 3#
        t.z = 4#
    Next i
    totalTime = GetTickCount - ticks
    log ("Local loop took " & CStr(totalTime) & " ms.")
    
    
    log ("Init function...")
    ticks = GetTickCount
    For i = 0 To iterations
        Dim t2 As TVec3
        t2 = TVec3_init(2#, 3#, 4#)
    Next i
    totalTime = GetTickCount - ticks
    log ("Init function took " & CStr(totalTime) & " ms.")
End Sub


Private Sub StructVsPrimitive()
    Dim ticks As Long, totalTime As Long
    
    log ("Structs...")
    ticks = GetTickCount
    For i = 0 To iterations
        Dim t As TVec3
        Dim x As Double, y As Double, z As Double
        t.x = 2
        t.y = 3
        t.z = 4
        x = t.x
        y = t.y
        z = t.z
    Next i
    totalTime = GetTickCount - ticks
    log ("Setting structs took " & CStr(totalTime) & " ms.")
    
    log ("Primitives...")
    ticks = GetTickCount
    For i = 0 To iterations
        Dim xx As Double, yy As Double, zz As Double
        xx = 2
        yy = 3
        zz = 4
        xx = yy
        yy = zz
        zz = xx
    Next i
    totalTime = GetTickCount - ticks
    log ("Setting primitives took " & CStr(totalTime) & " ms.")

End Sub

Private Sub testGoTo()
    Dim something_ret As Double
    Dim something_x As Double, something_y As Double

    Dim value As Double
    
    For i = 1 To 5
        something_x = i * 1#
        GoTo SomethingBody
CallSiteSomething1:
        value = something_ret
        log ("The return value for loop " & CStr(i) & " is " & CStr(value))
    Next i
    GoTo TheEnd
    
SomethingBody:
    something_ret = something_x
    GoTo CallSiteSomething1
    
TheEnd:
End Sub

