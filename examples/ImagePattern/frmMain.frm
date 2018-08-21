VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image pattern"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   3840
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
    Dim pPattern    As Long
    Dim tMatrix     As cairo_matrix_t
    
    psfcForm = cairo_win32_surface_create(Me.hDC)
    pCr = cairo_create(psfcForm)

    psfcImg = cairo_image_surface_create_from_png(App.Path & "\icon.png")
    lW = cairo_image_surface_get_width(psfcImg)
    lH = cairo_image_surface_get_height(psfcImg)

    pPattern = cairo_pattern_create_for_surface(psfcImg)
    cairo_pattern_set_extend pPattern, CAIRO_EXTEND_REPEAT
    
    cairo_translate pCr, 128#, 128#
    cairo_rotate pCr, M_PI / 4
    cairo_scale pCr, 1 / Sqr(2), 1 / Sqr(2)
    cairo_translate pCr, -128#, -128#
    
    cairo_matrix_init_scale tMatrix, lW / 256# * 5#, lH / 256# * 5#
    cairo_pattern_set_matrix pPattern, tMatrix
    
    cairo_set_source pCr, pPattern
    
    cairo_rectangle pCr, 0, 0, 256#, 256#
    cairo_fill pCr
    
    cairo_pattern_destroy pPattern
    cairo_surface_destroy psfcImg

    cairo_destroy pCr
    cairo_surface_destroy psfcForm

End Sub






