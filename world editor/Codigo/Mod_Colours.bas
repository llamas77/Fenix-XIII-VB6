Attribute VB_Name = "Mod_Colours"
'Some awesome colours would be placed here in the future
Option Explicit

Public White(3) As Long
Public Red(3) As Long
Public Cyan(3) As Long
Public Black(3) As Long
Public Yellow(3) As Long
Public Gray(3) As Long
Public Transparent(3) As Long
Public Green(3) As Long

Public Sub InitColours()
        
    White(0) = D3DColorXRGB(255, 255, 255)
    White(1) = White(0)
    White(2) = White(0)
    White(3) = White(0)
    
    Red(0) = D3DColorXRGB(255, 0, 0)
    Red(1) = Red(0)
    Red(2) = Red(0)
    Red(3) = Red(0)
    
    Cyan(0) = D3DColorXRGB(0, 255, 255)
    Cyan(1) = Cyan(0)
    Cyan(2) = Cyan(0)
    Cyan(3) = Cyan(0)
    
    Black(0) = D3DColorARGB(255, 0, 0, 0)
    Black(1) = Black(0)
    Black(2) = Black(0)
    Black(3) = Black(0)
    
    Yellow(0) = D3DColorXRGB(255, 255, 0)
    Yellow(1) = Yellow(0)
    Yellow(2) = Yellow(0)
    Yellow(3) = Yellow(0)
    
    Gray(0) = D3DColorXRGB(150, 150, 150)
    Gray(1) = Gray(0)
    Gray(2) = Gray(0)
    Gray(3) = Gray(0)
    
    Transparent(0) = D3DColorARGB(0, 0, 0, 0)
    Transparent(1) = Transparent(0)
    Transparent(2) = Transparent(0)
    Transparent(3) = Transparent(0)
    
    Green(0) = D3DColorXRGB(0, 255, 0)
    Green(1) = Green(0)
    Green(2) = Green(0)
    Green(3) = Green(0)
    
End Sub
