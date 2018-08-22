VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text extents"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' // https://cairographics.org/samples/
' // Adapted by The trick

Option Explicit

Private Sub Form_Load()
    Dim psfcForm    As Long
    Dim pCr         As Long
    Dim tExtents    As cairo_text_extents_t
    Dim dX          As Double
    Dim dY          As Double
    Dim sUtf8       As String
    
    sUtf8 = "cairo"
    
    psfcForm = cairo_win32_surface_create(Me.hDC)
    pCr = cairo_create(psfcForm)

    cairo_select_font_face pCr, "Sans", _
        CAIRO_FONT_SLANT_NORMAL, _
        CAIRO_FONT_WEIGHT_NORMAL
    
    cairo_set_font_size pCr, 100#
    cairo_text_extents pCr, sUtf8, tExtents
    
    dX = 25#
    dY = 150#
    
    cairo_move_to pCr, dX, dY
    cairo_show_text pCr, sUtf8
    
    ' /* draw helping lines */
    cairo_set_source_rgba pCr, 1, 0.2, 0.2, 0.6
    cairo_set_line_width pCr, 6#
    cairo_arc pCr, dX, dY, 10#, 0, 2 * M_PI
    cairo_fill pCr
    cairo_move_to pCr, dX, dY
    cairo_rel_line_to pCr, 0, -tExtents.Height
    cairo_rel_line_to pCr, tExtents.Width, 0
    cairo_rel_line_to pCr, tExtents.x_bearing, -tExtents.y_bearing
    cairo_stroke pCr

    cairo_destroy pCr
    cairo_surface_destroy psfcForm

End Sub

