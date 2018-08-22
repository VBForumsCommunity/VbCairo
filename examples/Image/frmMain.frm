VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
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
    Dim lW          As Long
    Dim lH          As Long
    Dim psfcImg     As Long
    
    psfcForm = cairo_win32_surface_create(Me.hDC)
    pCr = cairo_create(psfcForm)

    psfcImg = cairo_image_surface_create_from_png(App.Path & "\icon.png")
    lW = cairo_image_surface_get_width(psfcImg)
    lH = cairo_image_surface_get_height(psfcImg)
    
    cairo_translate pCr, 128#, 128#
    cairo_rotate pCr, 45 * M_PI / 180
    cairo_scale pCr, 256# / lW, 256# / lH
    cairo_translate pCr, -0.5 * lW, -0.5 * lH
    
    cairo_set_source_surface pCr, psfcImg, 0, 0
    cairo_paint pCr
    cairo_surface_destroy psfcImg

    cairo_destroy pCr
    cairo_surface_destroy psfcForm

End Sub





