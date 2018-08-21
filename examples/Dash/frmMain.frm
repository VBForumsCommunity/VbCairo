VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dash"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6090
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
    Dim dDashes(3)  As Double
    Dim dOffset     As Double
    
    psfcForm = cairo_win32_surface_create(Me.hDC)
    pCr = cairo_create(psfcForm)
    
    dDashes(0) = 50 ' /* ink */
    dDashes(1) = 10 ' /* skip */
    dDashes(2) = 10 ' /* ink */
    dDashes(3) = 10 ' /* skip*/
    
    dOffset = -50#

    cairo_set_dash pCr, dDashes(0), UBound(dDashes) + 1, dOffset
    cairo_set_line_width pCr, 10#
    
    cairo_move_to pCr, 128#, 25.6
    cairo_line_to pCr, 230.4, 230.4
    cairo_rel_line_to pCr, -102.4, 0#
    cairo_curve_to pCr, 51.2, 230.4, 51.2, 128#, 128#, 128#
    
    cairo_stroke pCr

    cairo_destroy pCr
    cairo_surface_destroy psfcForm

End Sub


