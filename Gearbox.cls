VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Gearbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public blueShaft As Double
Public blueGear As Double
Public redShaft As Double
Public redGear As Double

Public bShaftDim As Dimension
Public bGearDim As Dimension
Public rShaftDim As Dimension
Public rGearDim As Dimension

Public Sub ChangeDimensions()
    Call SetDimVal(bShaftDim, blueShaft)
    Call SetDimVal(bGearDim, blueGear)
    Call SetDimVal(rShaftDim, redShaft)
    Call SetDimVal(rGearDim, redGear)
End Sub

Private Sub SetDimVal(myDimension As Object, newValue As Double)
    myDimension.SystemValue = newValue * 0.0254 'value in meters
End Sub
