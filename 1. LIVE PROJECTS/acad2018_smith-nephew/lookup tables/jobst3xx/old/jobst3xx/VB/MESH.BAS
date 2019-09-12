Attribute VB_Name = "Module1"
   'XY data type to represent points
    Type XY
        x As Double
        y As Double
    End Type
    
    Public Type BiArc
        xyStart     As XY
        xyTangent   As XY
        xyEnd       As XY
        xyR1        As XY
        xyR2        As XY
        nR1         As Double
        nR2         As Double
    End Type

