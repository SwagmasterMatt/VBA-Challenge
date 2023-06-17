Attribute VB_Name = "GlobalsColor"
Option Explicit

' Yellow Shades
Public LightYellow As Long
Public Yellow As Long
Public DarkYellow As Long

' Orange Shades
Public LightOrange As Long
Public Orange As Long
Public DarkOrange As Long

' Green Shades
Public LightGreen As Long
Public Green As Long
Public DarkGreen As Long

' Cyan Shades
Public LightCyan As Long
Public Cyan As Long
Public DarkCyan As Long

' Blue Shades
Public LightBlue As Long
Public Blue As Long
Public DarkBlue As Long

' Purple Shades
Public LightPurple As Long
Public Purple As Long
Public DarkPurple As Long

' Magenta Shades
Public LightMagenta As Long
Public Magenta As Long
Public DarkMagenta As Long

' Red Shades
Public LightRed As Long
Public Red As Long
Public DarkRed As Long

' Brown Shades
Public LightBrown As Long
Public Brown As Long
Public DarkBrown As Long

' Gray Shades
Public LightGray As Long
Public Gray As Long
Public DarkGray As Long

' Additional Colors
Public White As Long
Public Black As Long
Public NoColor As Long

Public Sub CallColors()

    ' Yellow Shades
    LightYellow = RGB(255, 255, 224)
    Yellow = RGB(255, 255, 0)
    DarkYellow = RGB(128, 128, 0)

    ' Orange Shades
    LightOrange = RGB(255, 165, 0)
    Orange = RGB(255, 140, 0)
    DarkOrange = RGB(255, 69, 0)

    ' Green Shades
    LightGreen = RGB(144, 238, 144)
    Green = RGB(0, 128, 0)
    DarkGreen = RGB(0, 100, 0)

    ' Cyan Shades
    LightCyan = RGB(224, 255, 255)
    Cyan = RGB(0, 255, 255)
    DarkCyan = RGB(0, 139, 139)

    ' Blue Shades
    LightBlue = RGB(173, 216, 230)
    Blue = RGB(0, 0, 255)
    DarkBlue = RGB(0, 0, 139)

    ' Purple Shades
    LightPurple = RGB(147, 112, 219)
    Purple = RGB(128, 0, 128)
    DarkPurple = RGB(85, 26, 139)

    ' Magenta Shades
    LightMagenta = RGB(255, 105, 180)
    Magenta = RGB(255, 0, 255)
    DarkMagenta = RGB(139, 0, 139)

    ' Red Shades
    LightRed = RGB(255, 182, 193)
    Red = RGB(255, 0, 0)
    DarkRed = RGB(139, 0, 0)

    ' Brown Shades
    LightBrown = RGB(181, 101, 29)
    Brown = RGB(165, 42, 42)
    DarkBrown = RGB(101, 33, 18)

    ' Gray Shades
    LightGray = RGB(211, 211, 211)
    Gray = RGB(128, 128, 128)
    DarkGray = RGB(169, 169, 169)

    ' Additional Colors
    White = RGB(255, 255, 255)
    Black = RGB(0, 0, 0)
    NoColor = RGB(255, 255, 255)
    
End Sub

