Attribute VB_Name = "types"
Option Explicit

Public Type TVec3
    x As Double
    y As Double
    z As Double
End Type
    
Public Function TVec3_init(x As Double, y As Double, z As Double) As TVec3
    Dim t As TVec3
    t.x = x
    t.y = y
    t.z = z
    TVec3_init = t
End Function

    
