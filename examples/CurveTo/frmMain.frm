VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Curve to"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   5700
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
    Dim dx(3)       As Double
    Dim dy(3)       As Double
    
    psfcForm = cairo_win32_surface_create(Me.hDC)
    pCr = cairo_create(psfcForm)
    
    dx(0) = 25.6:   dy(0) = 128#
    dx(1) = 102.4:  dy(1) = 230.4
    dx(2) = 153.6:  dy(2) = 25.6
    dx(3) = 230.4:  dy(3) = 128#

    
    cairo_move_to pCr, dx(0), dy(0)
    cairo_curve_to pCr, dx(1), dy(1), dx(2), dy(2), dx(3), dy(3)
    
    cairo_set_line_width pCr, 10#
    cairo_stroke pCr
    
    cairo_set_source_rgba pCr, 1, 0.2, 0.2, 0.6
    cairo_set_line_width pCr, 6#
    cairo_move_to pCr, dx(0), dy(0): cairo_line_to pCr, dx(1), dy(1)
    cairo_move_to pCr, dx(2), dy(2): cairo_line_to pCr, dx(3), dy(3)
    cairo_stroke pCr

    cairo_destroy pCr
    cairo_surface_destroy psfcForm

End Sub

