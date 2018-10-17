VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CLEAN_UP_XREF 
   Caption         =   "CLEAN_UP_XREF_00"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10125
   OleObjectBlob   =   "CLEAN_UP_XREF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CLEAN_UP_XREF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************
'--- SET_UP_DRAWING ---
'BUTTON FUNCTION
'1  UCS_WORLD
'1  PLANVIEW_UCSWORLD
'2  SET_UP_VARIABLES
'2  SET_UP_UNITS
'3  SCALE_DRAWING
'4  LTSCALE_ALL_EXCLUDE_BLOCKS (30) 'WAARDE VAN LTSCALE MEEGEVEN
'5  LTSCALE_ALL_BLOCKS (15) 'WAARDE VAN LTSCALE MEEGEVEN
'6  FLATTEN_DRAWING

'--- SET_UP_LAYERS ---
'7  STORE_LAYERSTATE
'8  RESTORE_LAYERSTATE
'9  SET_LAYER_XREF_TEMP_CURRENT
'10 ALL_LAYERS_ON
'11 UNLOCK_ALL_LAYERS
'12 UNFREEZE_ALL_LAYERS
'13 LAYERS_LINEWEIGHT_DEFAULT
'14 REMOVE_LAYER_FILTERS

'--- CLEANING_STUFF ---
'15 PURGE_DRAWING
'16 AUDIT_MODELSPACE
'17 DELETE_LAYOUTS
'18 DELETE_VIEWPORTS
'19 DELETE_VIEWS
'20 DELETE_DIMENSIONS
'21 CLEAN_SCALELIST

'--- LAYER_STUFF ---
'22 CLEAR_LAYER0
'   DELETE_LAYER ("E-61-ELEKTRISCHE BORDEN") 'HIER DIEN JE NAAM VAN LAYER IN TE VULLEN
'23 DELETE_FROZEN_LAYERS
'24 ALL_BY_LAYER
'25 <<NIET GEDEFINIEERD>>


'--- BLOCKS_STUFF ---
'26 ALL_BLOCKS_ENTITIES_ON_LAYER0
'27 BURST_EXPLODE_BLOCKS_BY_SELECTION
'28 BURST_EXPLODE_DRAWING_LEVEL_FAST
'29 BURST_EXPLODE_DRAWING_LEVEL_SLOW


'--- XREF_STUFF ---
'30 RELOAD_XREFS
'31 BIND_AND_EXPLODE_XREFS_LOADED
'32 DETACH_XREFS_UNLOADED
'33 DETACH_XREFS_NOT_FOUND
'34 DETACH_XREFS_ALL

'--- EXPORT_STUFF ---
'35 CREATE_BACKUP_FILE
'36 CREATE_WBLOCK
'37 COPY_BLOCK_INTO_TEMPLATE
'38 SELECT_NULPUNT_ADD_HYPERLINK
'
'39 GO BUTTON
'************************************************************************
'-- VOORGEDEFINIEERDE WAARDEN ---

Dim SHOW_UF_VAR As Boolean

Sub LAUNCH_FORM()

'Dim DATE0 As String
'DATE0 = 1
'Call ThisDrawing.CHECK_DATE(DATE0)
'If DATE0 = 0 Then
'CLEAN_UP_XREF.hide
'Else

'SHOW_UF_VAR = True
Frame1.Visible = False
CLEAN_UP_XREF.show
'End If

End Sub


Private Sub CommandButton1_Click()
''SHOW_UF_VAR = False
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.UCS_WORLD
Call ThisDrawing.PLANVIEW_UCSWORLD
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton2_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.SET_UP_VARIABLES
Call ThisDrawing.SET_UP_UNITS
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton3_Click()
Dim Factor As Double
Factor = CLEAN_UP_XREF.TextBox6.Text
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.UNLOCK_ALL_LAYERS
Call ThisDrawing.SCALE_DRAWING(Factor)
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton4_Click()
Dim LTscale As Double
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
LTscale = CLEAN_UP_XREF.TextBox1.Text
Call ThisDrawing.LTSCALE_ALL_EXCLUDE_BLOCKS(LTscale)
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton41_Click()

'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.ALL_Mtext_BYLAYER
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton42_Click()

'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.DELETE_DIMENSIONS_IN_BLOCK
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show

End Sub

Private Sub CommandButton43_Click()

'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
LTscale = CLEAN_UP_XREF.TextBox2.Text

Call ThisDrawing.SEND_EXPORT

Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show

End Sub

Private Sub CommandButton44_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False

Call ThisDrawing.A2T

Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton45_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False

Call ThisDrawing.db2sb

Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.show
End Sub

Private Sub CommandButton46_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False


Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton49_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.X_Mtext
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show

End Sub

Private Sub CommandButton5_Click()
Dim LTscale As Double
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
LTscale = CLEAN_UP_XREF.TextBox2.Text
Call ThisDrawing.LTSCALE_ALL_BLOCKS(LTscale)
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton6_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.FLATTEN_DRAWING
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton7_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.UNLOCK_ALL_LAYERS
Call ThisDrawing.STORE_LAYERSTATE
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show

End Sub

Private Sub CommandButton8_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.UNLOCK_ALL_LAYERS
Call ThisDrawing.RESTORE_LAYERSTATE_ALL
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton9_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.SET_LAYER_XREF_TEMP_CURRENT
'If SHOW_UF_VAR Then CLEAN_UP_XREF.Show
Frame1.Visible = False
MultiPage1.Visible = True
End Sub
Private Sub CommandButton10_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.ALL_LAYERS_ON
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton11_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.UNLOCK_ALL_LAYERS
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton12_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.UNFREEZE_ALL_LAYERS
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton13_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.LAYERS_LINEWEIGHT_DEFAULT
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton14_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.REMOVE_LAYER_FILTERS
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton15_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.PURGE_DRAWING
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton16_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.AUDIT_MODELSPACE
Frame1.Visible = False
MultiPage1.Visible = True
'If SHOW_UF_VAR Then CLEAN_UP_XREF.Show

End Sub

Private Sub CommandButton17_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.DELETE_LAYOUTS
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton18_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.DELETE_VIEWPORTS
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton19_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.DELETE_VIEWS
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton20_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.DELETE_DIMENSIONS
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton21_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.CLEAN_SCALELIST
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton22_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.CLEAR_LAYER0
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton23_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
'MultiPage1.Visible = False
Call ThisDrawing.DELETE_FROZEN_LAYERS

Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.HIDE
End Sub

Private Sub CommandButton24_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
'Call ThisDrawing.STORE_LAYERSTATE
Call ThisDrawing.ALL_LAYERS_ON
Call ThisDrawing.UNLOCK_ALL_LAYERS
Call ThisDrawing.UNFREEZE_ALL_LAYERS
Call ThisDrawing.ALL_BY_LAYER 'MAIN FUNCTIE
'Call ThisDrawing.RESTORE_LAYERSTATE_ALL
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton25_Click()
'<<NIET GEDEFINIEERD>>
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.MTEXT_IN_BLOCK
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton26_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.ALL_BLOCKS_ENTITIES_ON_LAYER0
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton27_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.BURST_EXPLODE_BLOCKS_BY_SELECTION
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
CLEAN_UP_XREF.show
End Sub

Private Sub CommandButton28_Click()
Dim level As Integer
level = CLEAN_UP_XREF.TextBox3.Text
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.BURST_EXPLODE_DRAWING_LEVEL_FAST(level)
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton29_Click()
Dim level As Integer
level = CLEAN_UP_XREF.TextBox3.Text
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.BURST_EXPLODE_DRAWING_LEVEL_SLOW(level)
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton30_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.RELOAD_XREFS
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton31_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.BIND_AND_EXPLODE_XREFS_LOADED
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton32_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.DETACH_XREFS_UNLOADED
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton33_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.DETACH_XREFS_NOT_FOUND
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton34_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.DETACH_XREFS_ALL
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton35_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.CREATE_BACKUP_FILE
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton36_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.CREATE_WBLOCK
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton37_Click()
Dim templatefullpath As String
Dim xrefnaam As String
templatefullpath = CLEAN_UP_XREF.TextBox5.Text
xrefnaam = CLEAN_UP_XREF.TextBox7.Text & "-" & CLEAN_UP_XREF.TextBox11.Text & "-SBT-" & CLEAN_UP_XREF.TextBox12.Text & "-P-" & CLEAN_UP_XREF.TextBox8.Text & "-XREF.dwg"
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.COPY_WBLOCK_INTO_TEMPLATE(templatefullpath, xrefnaam)
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.Show
End Sub

Private Sub CommandButton38_Click()
'CLEAN_UP_XREF.HIDE
Frame1.Visible = True
MultiPage1.Visible = False
Call ThisDrawing.SELECT_NULPUNT_ADD_HYPERLINK
'If SHOW_UF_VAR Then
Frame1.Visible = False
MultiPage1.Visible = True
'CLEAN_UP_XREF.show

End Sub
Private Sub CommandButton39_Click()

Dim strInput As String
Dim Splitsymbol As String
Dim outarray() As Variant
Dim arrsize As Integer

strInput = CLEAN_UP_XREF.TextBox4.Text
Splitsymbol = "-"

'--START--
'This Function divides a string and returns all substrings that are hold together by the SplitSymbol
'The default Split Symbol is "-" but this can be overruled by every symbol you like
While InStr(strInput, Splitsymbol) > 0
    ReDim Preserve outarray(0 To arrsize) As Variant
    outarray(arrsize) = Left(strInput, InStr(strInput, Splitsymbol) - 1)
    strInput = Mid(strInput, InStr(strInput, Splitsymbol) + Len(Splitsymbol))
    arrsize = arrsize + 1
Wend

ReDim Preserve outarray(0 To arrsize) As Variant
outarray(arrsize) = strInput
'--STOP--

'SHOW_UF_VAR = False
'CLEAN_UP_XREF.HIDE
For i = 0 To arrsize

ThisDrawing.Utility.Prompt vbCrLf & "*" & outarray(i) & "*" & vbCrLf
   
Select Case outarray(i)
    
    Case "1"
    CommandButton1_Click
    
    Case "2"
    CommandButton2_Click
    
    Case "3"
    CommandButton3_Click
    
    Case "4"
    CommandButton4_Click
    
    Case "5"
    CommandButton5_Click
    
    Case "6"
    CommandButton6_Click
    
    Case "7"
    CommandButton7_Click
    
    Case "8"
    CommandButton8_Click
    
    Case "9"
    CommandButton9_Click
    
    Case "10"
    CommandButton10_Click
    
    Case "11"
    CommandButton11_Click
    
    Case "12"
    CommandButton12_Click
    
    Case "13"
    CommandButton13_Click
    
    Case "14"
    CommandButton14_Click
    
    Case "15"
    CommandButton15_Click
    
    Case "16"
    CommandButton16_Click
    
    Case "17"
    CommandButton17_Click
    
    Case "18"
    CommandButton18_Click
    
    Case "19"
    CommandButton19_Click
    
    Case "20"
    CommandButton20_Click
    
    Case "20b"
    CommandButton42_Click
    
    Case "21"
    CommandButton21_Click
    
    Case "22"
    CommandButton22_Click
    
    Case "23"
    CommandButton23_Click
    
    Case "24"
    CommandButton24_Click
    
    Case "25"
    CommandButton25_Click
    
    Case "25b"
    CommandButton41_Click
    
    Case "26"
    CommandButton26_Click
    
    Case "27"
    CommandButton27_Click
    
    Case "28"
    CommandButton28_Click
    
    Case "29"
    CommandButton29_Click
    
    Case "30"
    CommandButton30_Click
    
    Case "31"
    CommandButton31_Click
    
    Case "32"
    CommandButton32_Click
    
    Case "33"
    CommandButton33_Click
    
    Case "34"
    CommandButton34_Click
    
    Case "35"
    CommandButton35_Click
    
    Case "36"
    CommandButton36_Click
    
    Case "37"
    CommandButton37_Click
    
    Case "38"
    CommandButton38_Click
    
    Case "39"
    CommandButton43_Click
    
    Case "40"
    
    
    Case "41"
    CommandButton45_Click
    
    Case "42"
    
    Case "43"
    
    Case "44"
    
    
    
    Case "END"
          
        CLEAN_UP_XREF.MultiPage1.SetFocus
    CLEAN_UP_XREF.hide
        MsgBox ("OPERATIE IS VOLTOOID")
    
    End Select

Next i

'SHOW_UF_VAR = True
'If SHOW_UF_VAR Then CLEAN_UP_XREF.Show
'CLEAN_UP_XREF.show
End Sub

Private Sub CommandButton40_Click()
CLEAN_UP_XREF.hide
End Sub

Private Sub UserForm_Click()

End Sub
