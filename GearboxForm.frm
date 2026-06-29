VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GearboxForm 
   Caption         =   "Gearbox Builder"
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5040
   OleObjectBlob   =   "GearboxForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GearboxForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myGearbox As Gearbox

Private Sub OKButton_Click()
    Me.Hide
End Sub

Private Sub txtBlueGearOD_Change()
    myGearbox.blueGear = CDbl(txtBlueGearOD)
End Sub

Private Sub txtBlueGearShaftDia_Change()
    myGearbox.blueShaft = CDbl(txtBlueGearShaftDia)
End Sub

Private Sub txtRedGearOD_Change()
On Error GoTo CatchBlock
    myGearbox.redGear = CDbl(txtRedGearOD)
CatchBlock:
End Sub


Private Sub txtRedGearShaftDia_Change()
On Error GoTo CatchBlock
    myGearbox.redShaft = CDbl(txtRedGearShaftDia)
CatchBlock:
    'ignore the problem for now
End Sub

Private Sub UserForm_Activate()
    'load in the myGearbox values
    With myGearbox
       txtBlueGearOD = .blueGear
       txtBlueGearShaftDia = .blueShaft
       txtRedGearOD = .redGear
       txtRedGearShaftDia = .redShaft
    End With
End Sub





