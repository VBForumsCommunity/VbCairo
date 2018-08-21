VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fill Style"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   6435
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
    
    psfcForm = cairo_win32_surface_create(Me.hDC)
    pCr = cairo_create(psfcForm)
    
    cairo_set_line_width pCr, 6
    
    cairo_rectangle pCr, 12, 12, 232, 70
    cairo_new_sub_path pCr: cairo_arc pCr, 64, 64, 40, 0, 2 * M_PI
    cairo_new_sub_path pCr: cairo_arc_negative pCr, 192, 64, 40, 0, -2 * M_PI
    
    cairo_set_fill_rule pCr, CAIRO_FILL_RULE_EVEN_ODD
    cairo_set_source_rgb pCr, 0, 0.7, 0: cairo_fill_preserve pCr
    cairo_set_source_rgb pCr, 0, 0, 0:   cairo_stroke pCr
    
    cairo_translate pCr, 0, 128
    cairo_rectangle pCr, 12, 12, 232, 70
    cairo_new_sub_path pCr: cairo_arc pCr, 64, 64, 40, 0, 2 * M_PI
    cairo_new_sub_path pCr: cairo_arc_negative pCr, 192, 64, 40, 0, -2 * M_PI
    
    cairo_set_fill_rule pCr, CAIRO_FILL_RULE_WINDING
    cairo_set_source_rgb pCr, 0, 0, 0.9:    cairo_fill_preserve pCr
    cairo_set_source_rgb pCr, 0, 0, 0:      cairo_stroke pCr

    cairo_destroy pCr
    cairo_surface_destroy psfcForm

End Sub




