VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clip"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4815
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
    
    cairo_arc pCr, 128#, 128#, 76.8, 0, 2 * M_PI

    cairo_clip pCr

    cairo_new_path pCr  ' /* current path is not
                        '  consumed by cairo_clip() */
                        
    cairo_rectangle pCr, 0, 0, 256, 256
    cairo_fill pCr
    cairo_set_source_rgb pCr, 0, 1, 0
    cairo_move_to pCr, 0, 0
    cairo_line_to pCr, 256, 256
    cairo_move_to pCr, 256, 0
    cairo_line_to pCr, 0, 256
    cairo_set_line_width pCr, 10#
    cairo_stroke pCr

    cairo_destroy pCr
    cairo_surface_destroy psfcForm

End Sub

