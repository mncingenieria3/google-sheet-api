VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton3_Click()
    CallAPI ("REQUISICION_OS")
    ExportOC ("ORDEN DE SERVICIO")
End Sub

Private Sub CommandButton4_Click()
    CheckAndWakeupAPI ("REQUISICION_OS")
End Sub
