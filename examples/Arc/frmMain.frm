VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Arc"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   5730
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
    Dim dXc         As Double
    Dim dYc         As Double
    Dim dRadius     As Double
    Dim dAngle1     As Double
    Dim dAngle2     As Double
    
    dXc = 128#: dYc = 128#
    dRadius = 100#
    dAngle1 = 45# * (M_PI / 180#)
    dAngle2 = 180# * (M_PI / 180#)
    
    psfcForm = cairo_win32_surface_create(Me.hDC)
    pCr = cairo_create(psfcForm)
    
    cairo_set_line_width pCr, 10#
    cairo_arc pCr, dXc, dYc, dRadius, dAngle1, dAngle2
    
    cairo_stroke pCr
    
    '/* draw helping lines */
    cairo_set_source_rgba pCr, 1, 0.2, 0.2, 0.6
    cairo_set_line_width pCr, 6#

    cairo_arc pCr, dXc, dYc, 10#, 0, 2 * M_PI
    cairo_fill pCr

    cairo_arc pCr, dXc, dYc, dRadius, dAngle1, dAngle1
    cairo_line_to pCr, dXc, dYc
    cairo_arc pCr, dXc, dYc, dRadius, dAngle2, dAngle2
    cairo_line_to pCr, dXc, dYc
    cairo_stroke pCr

    cairo_destroy pCr
    cairo_surface_destroy psfcForm

End Sub
