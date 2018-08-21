VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Curve rectangle"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   6315
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
    
    DrawRoundRect pCr, 30, 30, 300, 200, 102
    
    cairo_destroy pCr
    cairo_surface_destroy psfcForm

End Sub

Private Function DrawRoundRect( _
                 ByVal pCr As Long, _
                 ByVal dX0 As Double, _
                 ByVal dY0 As Double, _
                 ByVal dRect_Width As Double, _
                 ByVal dRect_Height As Double, _
                 ByVal dRadius As Double) As Boolean
    Dim dX1 As Double, dY1  As Double
    
    dX1 = dX0 + dRect_Width
    dY1 = dY0 + dRect_Height
           
    If dRect_Height = 0 Or dRect_Width = 9 Then Exit Function
    
    If dRect_Width / 2 < dRadius Then
        If dRect_Height / 2 < dRadius Then
            cairo_move_to pCr, dX0, (dY0 + dY1) / 2
            cairo_curve_to pCr, dX0, dY0, dX0, dY0, (dX0 + dX1) / 2, dY0
            cairo_curve_to pCr, dX1, dY0, dX1, dY0, dX1, (dY0 + dY1) / 2
            cairo_curve_to pCr, dX1, dY1, dX1, dY1, (dX1 + dX0) / 2, dY1
            cairo_curve_to pCr, dX0, dY1, dX0, dY1, dX0, (dY0 + dY1) / 2
        Else
            cairo_move_to pCr, dX0, dY0 + dRadius
            cairo_curve_to pCr, dX0, dY0, dX0, dY0, (dX0 + dX1) / 2, dY0
            cairo_curve_to pCr, dX1, dY0, dX1, dY0, dX1, dY0 + dRadius
            cairo_line_to pCr, dX1, dY1 - dRadius
            cairo_curve_to pCr, dX1, dY1, dX1, dY1, (dX1 + dX0) / 2, dY1
            cairo_curve_to pCr, dX0, dY1, dX0, dY1, dX0, dY1 - dRadius
        End If
    Else
        If dRect_Height / 2 < dRadius Then
            cairo_move_to pCr, dX0, (dY0 + dY1) / 2
            cairo_curve_to pCr, dX0, dY0, dX0, dY0, dX0 + dRadius, dY0
            cairo_line_to pCr, dX1 - dRadius, dY0
            cairo_curve_to pCr, dX1, dY0, dX1, dY0, dX1, (dY0 + dY1) / 2
            cairo_curve_to pCr, dX1, dY1, dX1, dY1, dX1 - dRadius, dY1
            cairo_line_to pCr, dX0 + dRadius, dY1
            cairo_curve_to pCr, dX0, dY1, dX0, dY1, dX0, (dY0 + dY1) / 2
        Else
            cairo_move_to pCr, dX0, dY0 + dRadius
            cairo_curve_to pCr, dX0, dY0, dX0, dY0, dX0 + dRadius, dY0
            cairo_line_to pCr, dX1 - dRadius, dY0
            cairo_curve_to pCr, dX1, dY0, dX1, dY0, dX1, dY0 + dRadius
            cairo_line_to pCr, dX1, dY1 - dRadius
            cairo_curve_to pCr, dX1, dY1, dX1, dY1, dX1 - dRadius, dY1
            cairo_line_to pCr, dX0 + dRadius, dY1
            cairo_curve_to pCr, dX0, dY1, dX0, dY1, dX0, dY1 - dRadius
        End If
    End If
    
    cairo_close_path pCr
    
    cairo_set_source_rgb pCr, 0.5, 0.5, 1
    cairo_fill_preserve pCr
    cairo_set_source_rgba pCr, 0.5, 0, 0, 0.5
    cairo_set_line_width pCr, 10#
    cairo_stroke pCr
       
End Function


