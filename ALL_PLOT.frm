VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ALL_PLOT 
   Caption         =   "ALLE LAYOUTS PLOTTEN"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2850
   OleObjectBlob   =   "ALL_PLOT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ALL_PLOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub LAUNCH_FORM_ALLPLOT()


'Dim DATE0 As String

'Call ThisDrawing.CHECK_DATE(DATE0)
'If DATE0 = 0 Then
'ALL_PLOT.hide
'Else
'End

SHOW_UF_VAR = True


ALL_PLOT.show
'End If

End Sub
Private Sub CommandButton1_Click()
Dim PLOTTER As String
Dim size As String
PLOTTER = "ZWART_PLOT"
Call ThisDrawing.PlotLayouts(PLOTTER, size)

End Sub

Private Sub CommandButton2_Click()
Dim PLOTTER As String
Dim size As String

PLOTTER = "KLEUR_PLOT"
Call ThisDrawing.PlotLayouts(PLOTTER, size)

End Sub

Private Sub CommandButton3_Click()
Dim PLOTTER As String
Dim size As String
PLOTTER = "PDF"
Call ThisDrawing.PlotLayouts(PLOTTER, size)

End Sub

Private Sub CommandButton4_Click()
Dim PLOTTER As String
Dim size As String
PLOTTER = "DWF"
Call ThisDrawing.PlotLayouts(PLOTTER, size)
End Sub

Private Sub CommandButton5_Click()
Dim PLOTTER As String
Dim size As String
PLOTTER = "JPEG"
Call ThisDrawing.PlotLayouts(PLOTTER, size)
End Sub

Private Sub CommandButton6_Click()
Dim PLOTTER As String
Dim size As String
PLOTTER = "TIFF"
Call ThisDrawing.PlotLayouts(PLOTTER, size)
End Sub

Private Sub PDF_Click()

End Sub

Private Sub UserForm_Click()

End Sub
