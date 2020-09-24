VERSION 5.00
Begin VB.UserControl FlatProgress 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   16
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Shape shpProgress 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00F99C62&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Left            =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "FlatProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'FlatProgress.ctl
'=================
'Version:       1.0.0b
'Author:        Daniel Elkins (DigiRev)
'Copyright:     (c) 2008 Daniel Elkins
'Website:       http://www.DigiRev.org
'E-mail:        Daniel@DigiRev.org
'Created:       March 19th, 2008
'Last-updated:  March 19th, 2008

'Description: The most simple, lightweight progress bar I could come up with that
'             can support Currency min/max/value.

'License:     Do whatever you want with it...
Option Explicit

Private p_Min As Currency
Private p_Value As Currency
Private p_Max As Currency

Private curPercentValue As Currency
Private curPercentPixels As Currency

Private Sub Draw()
    With shpBorder
        .Width = UserControl.ScaleWidth
        .Height = UserControl.ScaleHeight
    End With
    
    If p_Value > p_Min Then
        'Get p_Value% of p_Max
        curPercentValue = (p_Value / p_Max) * 100
        'Get curPercentValue% of ScaleWidth.
        curPercentPixels = UserControl.ScaleWidth * (curPercentValue / 100)
        With shpProgress
            If .Visible = False Then .Visible = True
            .Height = UserControl.ScaleHeight
            .Width = curPercentPixels
        End With
    Else
        shpProgress.Visible = False
    End If
End Sub

Public Property Get Value() As Currency
    Value = p_Value
End Property

Public Property Let Value(ByVal NewValue As Currency)
    p_Value = NewValue
    PropertyChanged "Value"
    Draw
End Property

Public Property Get Max() As Currency
    Max = p_Max
End Property

Public Property Let Max(ByVal NewValue As Currency)
    p_Max = NewValue
    PropertyChanged "Max"
    Draw
End Property

Public Property Get Min() As Currency
    Min = p_Min
End Property

Public Property Let Min(ByVal NewValue As Currency)
    p_Min = NewValue
    PropertyChanged "Min"
    Draw
End Property

Private Sub UserControl_InitProperties()
    p_Min = 0
    p_Max = 100
    p_Value = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        p_Min = .ReadProperty("Min", 0)
        p_Max = .ReadProperty("Max", 100)
        p_Value = .ReadProperty("Value", 0)
    End With
End Sub

Private Sub UserControl_Resize()
    Draw
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Min", p_Min, 0
        .WriteProperty "Max", p_Max, 100
        .WriteProperty "Value", p_Value, 0
    End With
End Sub
