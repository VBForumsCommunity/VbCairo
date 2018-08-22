VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   3930
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
        
    cairo_select_font_face pCr, "Sans", CAIRO_FONT_SLANT_NORMAL, _
                                   CAIRO_FONT_WEIGHT_BOLD
    cairo_set_font_size pCr, 90#
    
    cairo_move_to pCr, 10#, 135#
    cairo_show_text pCr, "Hello"
    
    cairo_move_to pCr, 70#, 165#
    cairo_text_path pCr, "void"
    cairo_set_source_rgb pCr, 0.5, 0.5, 1
    cairo_fill_preserve pCr
    cairo_set_source_rgb pCr, 0, 0, 0
    cairo_set_line_width pCr, 2.56
    cairo_stroke pCr
    
    ' /* draw helping lines */
    cairo_set_source_rgba pCr, 1, 0.2, 0.2, 0.6
    cairo_arc pCr, 10#, 135#, 5.12, 0, 2 * M_PI
    cairo_close_path pCr
    cairo_arc pCr, 70#, 165#, 5.12, 0, 2 * M_PI
    cairo_fill pCr

    cairo_destroy pCr
    cairo_surface_destroy psfcForm

End Sub

