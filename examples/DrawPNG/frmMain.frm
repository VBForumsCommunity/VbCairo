VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Draw PNG"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim psfcFrm As cairo_surface_t
    Dim psfcImg As cairo_surface_t
    Dim pCr     As cairo_t

    psfcFrm = cairo_win32_surface_create(Me.hDC)
    pCr = cairo_create(psfcFrm)

    psfcImg = cairo_image_surface_create_from_png(App.Path & "\icon.png")
    
    cairo_set_source_surface pCr, psfcImg, 0, 0
    cairo_paint pCr
    
    cairo_surface_destroy psfcImg
    cairo_destroy pCr
    cairo_surface_destroy psfcFrm

End Sub




