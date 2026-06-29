Attribute VB_Name = "changeDims1"
' ******************************************************************************
' C:\Users\IQ1\AppData\Local\Temp\swx30192\Macro1.swb - macro recorded on 06/28/26 by IQ1
' ******************************************************************************
Dim swApp As Object

Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
Const m2i As Double = 0.0254
Dim myGearbox As New Gearbox

Sub main()
    
    Set swApp = Application.SldWorks
    
    Set Part = swApp.ActiveDoc
    
    Dim value As String
    Dim numVal As Double
    
    With myGearbox
        Set .bShaftDim = Part.Parameter("ShaftDia1@Sketch1@bevelgear.Part")
        myGearbox.blueShaft = .bShaftDim.value
        
        Set .bGearDim = Part.Parameter("BevDia1@Sketch1@bevelgear.Part")
        myGearbox.blueGear = .bGearDim.value
        
        Set .rShaftDim = Part.Parameter("ShaftDia2@Sketch1@bevelgear2.Part")
        myGearbox.redShaft = .rShaftDim.value
        
        Set .rGearDim = Part.Parameter("BevDia2@Sketch1@bevelgear2.Part")
        myGearbox.redGear = .rGearDim.value
    End With
    
    'Stop
    
    Dim gf As New GearboxForm
    Set gf.myGearbox = myGearbox
    gf.Show
    'when they click OK, it continues here
    If gf.myGearbox Is Nothing Then Exit Sub 'the user might have closed the dialog instead of OK

    'process the changes
    myGearbox.ChangeDimensions
    
    boolstatus = Part.EditRebuild3()
    Part.ClearSelection2 True

End Sub

Private Sub SetDimVal(myDimension As Object, message As String)
    value = InputBox(message)
    numVal = CDbl(value)
    myDimension.SystemValue = numVal * m2i 'value in meters
End Sub
