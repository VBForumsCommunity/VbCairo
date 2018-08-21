VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gradient"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   3675
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
    Dim pPat        As Long
    
    psfcForm = cairo_win32_surface_create(Me.hDC)
    pCr = cairo_create(psfcForm)
    
    pPat = cairo_pattern_create_linear(0#, 0#, 0#, 256#)
    cairo_pattern_add_color_stop_rgba pPat, 1, 0, 0, 0, 1
    cairo_pattern_add_color_stop_rgba pPat, 0, 1, 1, 1, 1
    cairo_rectangle pCr, 0, 0, 256, 256
    cairo_set_source pCr, pPat
    cairo_fill pCr
    cairo_pattern_destroy pPat
    
    pPat = cairo_pattern_create_radial(115.2, 102.4, 25.6, _
                                       102.4, 102.4, 128#)
    cairo_pattern_add_color_stop_rgba pPat, 0, 1, 1, 1, 1
    cairo_pattern_add_color_stop_rgba pPat, 1, 0, 0, 0, 1
    cairo_set_source pCr, pPat
    cairo_arc pCr, 128#, 128#, 76.8, 0, 2 * M_PI
    cairo_fill pCr
    cairo_pattern_destroy pPat

    cairo_destroy pCr
    cairo_surface_destroy psfcForm

End Sub



