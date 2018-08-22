VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set line cap"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   3780
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
    
    cairo_set_line_width pCr, 30#
    cairo_set_line_cap pCr, CAIRO_LINE_CAP_BUTT  ' /* default */
    cairo_move_to pCr, 64#, 50#:    cairo_line_to pCr, 64#, 200#
    cairo_stroke pCr
    cairo_set_line_cap pCr, CAIRO_LINE_CAP_ROUND
    cairo_move_to pCr, 128#, 50#:   cairo_line_to pCr, 128#, 200#
    cairo_stroke pCr
    cairo_set_line_cap pCr, CAIRO_LINE_CAP_SQUARE
    cairo_move_to pCr, 192#, 50#:   cairo_line_to pCr, 192#, 200#
    cairo_stroke pCr
    
    ' /* draw helping lines */
    cairo_set_source_rgb pCr, 1, 0.2, 0.2
    cairo_set_line_width pCr, 2.56
    cairo_move_to pCr, 64#, 50#:    cairo_line_to pCr, 64#, 200#
    cairo_move_to pCr, 128#, 50#:   cairo_line_to pCr, 128#, 200#
    cairo_move_to pCr, 192#, 50#:   cairo_line_to pCr, 192#, 200#
    cairo_stroke pCr

    cairo_destroy pCr
    cairo_surface_destroy psfcForm

End Sub

