VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set line join"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
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
    
    cairo_set_line_width pCr, 40.96
    cairo_move_to pCr, 76.8, 84.48
    cairo_rel_line_to pCr, 51.2, -51.2
    cairo_rel_line_to pCr, 51.2, 51.2
    cairo_set_line_join pCr, CAIRO_LINE_JOIN_MITER ' /* default */
    cairo_stroke pCr
    
    cairo_move_to pCr, 76.8, 161.28
    cairo_rel_line_to pCr, 51.2, -51.2
    cairo_rel_line_to pCr, 51.2, 51.2
    cairo_set_line_join pCr, CAIRO_LINE_JOIN_BEVEL
    cairo_stroke pCr
    
    cairo_move_to pCr, 76.8, 238.08
    cairo_rel_line_to pCr, 51.2, -51.2
    cairo_rel_line_to pCr, 51.2, 51.2
    cairo_set_line_join pCr, CAIRO_LINE_JOIN_ROUND
    cairo_stroke pCr

    cairo_destroy pCr
    cairo_surface_destroy psfcForm

End Sub

