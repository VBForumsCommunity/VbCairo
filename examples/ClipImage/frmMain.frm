VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clip image"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5325
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
    Dim lW          As Long
    Dim lH          As Long
    Dim psfcPng     As Long
    
    psfcForm = cairo_win32_surface_create(Me.hDC)
    pCr = cairo_create(psfcForm)

    cairo_arc pCr, 128#, 128#, 76.8, 0, 2 * M_PI
    cairo_clip pCr
    cairo_new_path pCr ' /* path not consumed by clip()*/

    psfcPng = cairo_image_surface_create_from_png(App.Path & "\icon.png")
    lW = cairo_image_surface_get_width(psfcPng)
    lH = cairo_image_surface_get_height(psfcPng)

    cairo_scale pCr, 256# / lW, 256# / lH

    cairo_set_source_surface pCr, psfcPng, 0, 0
    cairo_paint pCr

    cairo_surface_destroy psfcPng

    cairo_destroy pCr
    cairo_surface_destroy psfcForm

End Sub


