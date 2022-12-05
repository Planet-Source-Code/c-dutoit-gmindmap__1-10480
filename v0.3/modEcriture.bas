Attribute VB_Name = "modEcriture"
'modEcriture : Tout ce qui a trait aux écritures exotiques : penchées, ...
'Par C.Dutoit, 5 Août 2000 (dutoitc@hotmail.com)
'http://www.home.ch/~spaw4758
'merci à KPD pour leur précieuse aide !
    'KPD-Team 1998
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net

Option Explicit


'In general section
Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long



Const LF_FACESIZE = 32
Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type



'Imprimer le texte <text> à la position (x,y) avec un angle <angle>
'Size : taille en points
Sub PrintRotfrmMap(X As Long, Y As Long, Angle As Long, text As String, size As Byte)
    Dim RotateMe As LOGFONT
    Dim rfont, cfont
   
    'Initialisations
    RotateMe.lfEscapement = Angle * 10  'Set the rotation degree
    RotateMe.lfHeight = (size * -20) / Screen.TwipsPerPixelY 'Set the height of the font
    
    'Create the font
    rfont = CreateFontIndirect(RotateMe)
        
    'Select the font within the Form's device context
    cfont = SelectObject(frmMap.hdc, rfont)
        
        
    'Calculer la pos du début du texte
    Dim NewX, NewY, AngleRad, DistX, DistY
    AngleRad = Angle * 3.1415926535 / 180
  
    'Print the text
    frmMap.CurrentX = X - LongueurTexteRot(text, size) * Cos(AngleRad) / 2
    frmMap.CurrentY = Y + LongueurTexteRot(text, size) * Sin(AngleRad) / 2
    frmMap.Print text
    
    'DeleteObject (rfont)
    ' restoring everything to normal
    Dim dum
    dum = SelectObject(frmMap.hdc, cfont)
    dum = DeleteObject(rfont)
End Sub


'Imprimer le texte <text> à la position (x,y) avec un angle <angle>
'Size : taille en points
Sub PrintRotprinter(X As Long, Y As Long, Angle As Long, text As String, size As Byte)
    Dim RotateMe As LOGFONT
    Dim rfont, cfont
   
    'Initialisations
    RotateMe.lfEscapement = Angle * 10  'Set the rotation degree
    RotateMe.lfHeight = (size * -20) / Printer.TwipsPerPixelY 'Set the height of the font
    
    'Create the font
    rfont = CreateFontIndirect(RotateMe)
        
    'Select the font within the Form's device context
    cfont = SelectObject(Printer.hdc, rfont)
        
        
    'Calculer la pos du début du texte
    Dim NewX, NewY, AngleRad, DistX, DistY
    AngleRad = Angle * 3.1415926535 / 180
  
    'Print the text
    Printer.CurrentX = X - PrinterLongueurTexteRot(text, size) * Cos(AngleRad) / 2
    Printer.CurrentY = Y + PrinterLongueurTexteRot(text, size) * Sin(AngleRad) / 2
    Printer.FillStyle = vbTransparent
    Printer.Print text
    
    'DeleteObject (rfont)
    ' restoring everything to normal
    Dim dum
    dum = SelectObject(Printer.hdc, cfont)
    dum = DeleteObject(rfont)
End Sub



'Retourner la largeur d'un texte, une fois la rotation effectuée
Function oldLargeurTexteRot(texte As String, Angle As Long, size As Byte)
    Dim RotateMe As LOGFONT
    Dim rfont, cfont
   
    'Initialisations
    RotateMe.lfEscapement = Angle * 10  'Set the rotation degree
    RotateMe.lfHeight = (size * -20) / Screen.TwipsPerPixelY 'Set the height of the font
    
    'Create the font
    rfont = CreateFontIndirect(RotateMe)
        
    'Select the font within the Form's device context
    cfont = SelectObject(frmMap.hdc, rfont)
        
        
    'Calculer la pos du début du texte
    Dim NewX, NewY, AngleRad, DistX, DistY
    AngleRad = Angle * 3.1415926535 / 180
    DistX = frmMap.TextWidth(texte)
  
    'DeleteObject (rfont)
    ' restoring everything to normal
    Dim dum
    dum = SelectObject(frmMap.hdc, cfont)
    dum = DeleteObject(rfont)
    
    'LargeurTexteRot = DistX
End Function 'LargeurTexteRot



'Retourner la hauteur d'un texte, une fois la rotation effectuée
Function oldHauteurTexteRot(texte As String, Angle As Long, size As Byte)
    Dim RotateMe As LOGFONT
    Dim rfont, cfont
    
     If Angle Mod 90 < 5 Or Angle Mod 90 > 85 Then
        'HauteurTexteRot = frmMap.TextHeight(texte)
        Exit Function
    End If
   
    'Initialisations
    RotateMe.lfEscapement = Angle * 10  'Set the rotation degree
    RotateMe.lfHeight = (size * -20) / Screen.TwipsPerPixelY 'Set the height of the font
    
    'Create the font
    rfont = CreateFontIndirect(RotateMe)
        
    'Select the font within the Form's device context
    cfont = SelectObject(frmMap.hdc, rfont)
        
        
    'Calculer la pos du début du texte
    Dim NewX, NewY, AngleRad, DistX, DistY
    AngleRad = Angle * 3.1415926535 / 180
    DistX = frmMap.TextWidth(texte)
    DistY = DistX * Tan(AngleRad)
  
    'DeleteObject (rfont)
    ' restoring everything to normal
    Dim dum
    dum = SelectObject(frmMap.hdc, cfont)
    dum = DeleteObject(rfont)
    
    'HauteurTexteRot = DistY
End Function 'HauteurTexteRot



'Retourner la hauteur d'un texte, une fois la rotation effectuée
Function LongueurTexteRot(texte As String, size As Byte)
    Dim RotateMe As LOGFONT
    Dim rfont, cfont
       
    'Initialisations
    RotateMe.lfEscapement = 0  'Set the rotation degree
    RotateMe.lfHeight = (size * -20) / Screen.TwipsPerPixelY 'Set the height of the font
    
    'Create the font
    rfont = CreateFontIndirect(RotateMe)
        
    'Select the font within the Form's device context
    cfont = SelectObject(frmMap.hdc, rfont)
        
        
    'Calculer la pos du début du texte
    Dim dist
    dist = frmMap.TextWidth(texte)
  
    'DeleteObject (rfont)
    ' restoring everything to normal
    Dim dum
    dum = SelectObject(frmMap.hdc, cfont)
    dum = DeleteObject(rfont)
    
    LongueurTexteRot = dist
End Function 'LongueurTexteRot


'Retourner la hauteur d'un texte, une fois la rotation effectuée
Function PrinterLongueurTexteRot(texte As String, size As Byte)
    Dim RotateMe As LOGFONT
    Dim rfont, cfont
       
    'Initialisations
    RotateMe.lfEscapement = 0  'Set the rotation degree
    RotateMe.lfHeight = (size * -20) / Printer.TwipsPerPixelY 'Set the height of the font
    
    'Create the font
    rfont = CreateFontIndirect(RotateMe)
        
    'Select the font within the Form's device context
    cfont = SelectObject(Printer.hdc, rfont)
        
        
    'Calculer la pos du début du texte
    Dim dist
    dist = Printer.TextWidth(texte)
  
    'DeleteObject (rfont)
    ' restoring everything to normal
    Dim dum
    dum = SelectObject(Printer.hdc, cfont)
    dum = DeleteObject(rfont)
    
    PrinterLongueurTexteRot = dist
End Function 'PrinterLongueurTexteRot


