VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi segment caps"
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
    
    cairo_move_to pCr, 50#, 75#
    cairo_line_to pCr, 200#, 75#
    
    cairo_move_to pCr, 50#, 125#
    cairo_line_to pCr, 200#, 125#
    
    cairo_move_to pCr, 50#, 175#
    cairo_line_to pCr, 200#, 175#
    
    cairo_set_line_width pCr, 30#
    cairo_set_line_cap pCr, CAIRO_LINE_CAP_ROUND
    cairo_stroke pCr

    cairo_destroy pCr
    cairo_surface_destroy psfcForm

End Sub

