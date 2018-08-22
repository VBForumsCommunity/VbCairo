VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fill nad stroke 2"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7005
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
    
    cairo_move_to pCr, 128#, 25.6
    cairo_line_to pCr, 230.4, 230.4
    cairo_rel_line_to pCr, -102.4, 0#
    cairo_curve_to pCr, 51.2, 230.4, 51.2, 128#, 128#, 128#
    cairo_close_path pCr
    
    cairo_move_to pCr, 64#, 25.6
    cairo_rel_line_to pCr, 51.2, 51.2
    cairo_rel_line_to pCr, -51.2, 51.2
    cairo_rel_line_to pCr, -51.2, -51.2
    cairo_close_path pCr
    
    cairo_set_line_width pCr, 10#
    cairo_set_source_rgb pCr, 0, 0, 1
    cairo_fill_preserve pCr
    cairo_set_source_rgb pCr, 0, 0, 0
    cairo_stroke pCr

    cairo_destroy pCr
    cairo_surface_destroy psfcForm

End Sub



