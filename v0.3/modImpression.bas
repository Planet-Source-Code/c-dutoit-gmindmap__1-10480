Attribute VB_Name = "modImpression"
'modImpression : gestion de l'impression
'Par C.Dutoit, 3 Août 2000 (dutoitc@hotmail.com)
'http://www.home.ch/~spaw4758
Option Explicit


'Préparer la feuille pour l'impression
Private Sub DessinerCartouche()
    Dim BordGauche, BordHaut
    Dim HauteurLigne 'Hauteur d'une ligne
    Dim Intervale    'Hauteur entre 2 lignes
    
    BordGauche = Printer.ScaleWidth - Printer.TextWidth("OOOOOOOOOOOOOOOOOOOOOOOOOOOOOO") '(30 caractères de large)
    HauteurLigne = Printer.TextHeight("O")
    Intervale = HauteurLigne / 2
    BordHaut = Printer.ScaleHeight - (HauteurLigne + Intervale) * 3 'place pour 3 lignes
    
    'Cartouche
    Printer.Line (BordGauche, BordHaut)- _
                 (Printer.ScaleWidth, Printer.ScaleHeight), , B
                
    'Trait horizontal entre le titre et "G-Mindmap..."
    Printer.Line (BordGauche, BordHaut + HauteurLigne + Intervale)- _
                 (Printer.ScaleWidth, BordHaut + HauteurLigne + Intervale)
                
    'Trait horizontal entre "G-Mindmap..." et l'auteur+date
    Printer.Line (BordGauche, BordHaut + (HauteurLigne + Intervale) * 2)- _
                 (Printer.ScaleWidth, BordHaut + (HauteurLigne + Intervale) * 2)
     
                
    'Trait vertical entre la version et (date-auteur)
    Printer.Line ((Printer.ScaleWidth + BordGauche) / 2, BordHaut + (HauteurLigne + Intervale) * 3)- _
                 ((Printer.ScaleWidth + BordGauche) / 2, BordHaut + (HauteurLigne + Intervale) * 3)
                                
    'Afficher le titre
    Printer.CurrentX = BordGauche + Intervale
    Printer.CurrentY = BordHaut + Intervale
    If Len(Arbre(0).Legende) > 20 Then '20 premiers car. uniquement
        Printer.Print Left$(Arbre(0).Legende, 20)
    Else
        Printer.Print Arbre(0).Legende
    End If
                             
    'Afficher la version
    Printer.CurrentX = BordGauche + Intervale
    Printer.CurrentY = BordHaut + (HauteurLigne + Intervale)
    Printer.Print "G-Mindmap v" & App.Major & "." & App.Minor & "." & App.Revision
                     
    'Afficher l'auteur
    Printer.CurrentX = BordGauche + Intervale
    Printer.CurrentY = BordHaut + (HauteurLigne + Intervale) * 2
    Printer.Print InputBox("Entrez le nom de l'auteur", "Impression", "Anonyme")
        
    'Afficher la date
    Printer.CurrentX = (BordGauche + Printer.ScaleWidth) / 2 + Intervale
    Printer.CurrentY = BordHaut + (HauteurLigne + Intervale) * 2
    Printer.Print Date
End Sub 'DessinerCartouche


'Imprimer le mindmap
Sub ImprimerMindmap()
    'DessinerCartouche
    
    'frmImpression.Picture = frmMap.Picture
    'frmImpression.Show
    
    
    'test
    'frmMDI.cmDlgImprimer.Flags = cdlPDPrintSetup
    
    Dim NbreCopies As Integer
    Dim i As Integer
    frmMDI.cmDlgImprimer.ShowPrinter
    'NbreCopies = frmMDI.cmDlgImprimer.Copies
    'For i = 1 To NbreCopies
        ImprimerUnMindMap
    'Next i
End Sub 'ImprimerMindmap



'Impression d'un mindmap, avec les options de la boîte de dialogue d'impression
Private Sub ImprimerUnMindMap()
    'frmMap.PrintForm
    PrinterDessinerAllMindMap
End Sub 'ImprimerUnMindMap






'Dessiner un noeud
Private Sub PrinterDessinerNoeud(X, Y, Index As Long)
    Dim txtW As Long
    Dim txtH As Long
    Dim w As Long           'Largeur
    Dim h As Long           'Hauteur
    
    'Calculer la hauteur et la largeur
    txtW = Printer.TextWidth(Arbre(Index).Legende)
    txtH = Printer.TextHeight(Arbre(Index).Legende)
    w = txtW * 0.5 + Printer.TextWidth("OO")
    h = txtH * 0.5 + Printer.TextHeight("O") / 2
    
    'Dessiner le centre
    Printer.FillColor = RGB(255, 255, 255)
    Printer.FillStyle = vbSolid
    Printer.DrawWidth = 2
    Printer.Circle (X, Y), w, , , , h / w
    Printer.DrawWidth = 1
    Printer.FillStyle = vbTransparent
    
    
    'Afficher le label
    Printer.CurrentX = X - txtW / 2
    Printer.CurrentY = Y - txtH / 2
    Printer.ForeColor = 0 'Couleur du cadre
    Printer.Print Arbre(Index).Legende & vbCrLf & Arbre(Index).URL
    
    'Enregistrer la position
    Arbre(Index).X = X
    Arbre(Index).Y = Y
End Sub 'PrinterDessinerNoeud


Private Sub PrinterDessinerNoeudEtFils(NoeudDepart As Long, AngleDeb, AngleFin, X, Y, Etape)
    Dim Etalon1 As Long
    Etalon1 = Printer.ScaleWidth / 20

    'Dessiner les suivants
    If Arbre(NoeudDepart).NbSuivants > 0 Then
        'Normaliser les angles
        Dim IncAngle
        If AngleDeb < 0 Then AngleDeb = AngleDeb + 360
        If AngleFin < AngleDeb Then AngleFin = AngleFin + 360
    
        'Calculer l'incrément
        If Arbre(NoeudDepart).NbSuivants = 1 Then
            IncAngle = 0
            AngleDeb = (AngleDeb + AngleFin) / 2
        Else
            If AngleDeb Mod 360 = AngleFin Mod 360 Then
                IncAngle = (AngleFin - AngleDeb) / (Arbre(NoeudDepart).NbSuivants)
            Else
                IncAngle = (AngleFin - AngleDeb) / (Arbre(NoeudDepart).NbSuivants - 1)
            End If
        End If
    
        Dim i
        Dim NewAngleDeb
        Dim NewAngleFin
        Dim Delta
        Dim NewX, NewY
        Dim dist, Angle As Single '***modifié
        Dim Xp, Yp

    
        'Afficher chaque suivant
        For i = 0 To Arbre(NoeudDepart).NbSuivants - 1
            'Calculer les angles limites
            Delta = (90 - Etape * 9)
            NewAngleDeb = IncAngle * i + AngleDeb - Delta / 2
            NewAngleFin = IncAngle * i + AngleDeb + Delta / 2
        
            'Calculer l'angle (en radian)
            Angle = (IncAngle * i + AngleDeb) / 180 * 3.1415926535
            
            'Calculer la pos. finale
            Dim texte As String
            Dim AngleTexte As Long
            Dim HCar As Byte
            AngleTexte = Angle * 180 / 3.1415926535 '-Atn((NewY - Y) / (NewX - X)) * 180 / 3.1415926535
            If AngleTexte Mod 360 > 90 And AngleTexte Mod 360 < 270 Then AngleTexte = AngleTexte Mod 360 - 180
            texte = Arbre(Arbre(NoeudDepart).Suivants(i)).Legende
            HCar = ((HauteurArbre(0) - Etape) * 3 / HauteurArbre(0)) ^ 2 + 8
        
            NewX = X + PrinterLongueurTexteRot(texte & "OO", HCar) * Cos(Angle)
            NewY = Y - PrinterLongueurTexteRot(texte & "OO", HCar) * Sin(Angle)

            If NoeudDepart = 0 Then 'fils de racine ? => agrandir
                NewX = NewX + Cos(Angle) * Etalon1
                NewY = NewY - Sin(Angle) * Etalon1
            End If
            
            Printer.ForeColor = RGB(Etape * 64 Mod 256, Etape * 128 Mod 256, Etape * 32 Mod 256)
            PrintRotprinter (X + NewX) / 2, (Y + NewY) / 2, AngleTexte, texte, HCar
            
            'Tracer une ligne
            Printer.DrawWidth = (HauteurArbre(0) - Etape) + 1
            Printer.Line (X, Y)-(NewX, NewY)
            Printer.DrawWidth = 1
                   
            PrinterDessinerNoeudEtFils Arbre(NoeudDepart).Suivants(i), _
               NewAngleDeb, NewAngleFin, _
                NewX, NewY, Etape + 1
        Next i
    End If
    
    'Dessiner la racine
    If Etape = 1 Then PrinterDessinerNoeud X, Y, NoeudDepart
End Sub 'PrinterDessinerNoeudEtFils


Private Sub OldPrinterDessinerNoeudEtFils(NoeudDepart As Long, AngleDeb, AngleFin, X, Y, Etape)
    Dim Etalon1 As Long
    Etalon1 = Printer.ScaleWidth / 7

    'Dessiner les suivants
    If Arbre(NoeudDepart).NbSuivants > 0 Then
    Dim IncAngle
    If AngleDeb < 0 Then AngleDeb = AngleDeb + 360
    If AngleFin < AngleDeb Then AngleFin = AngleFin + 360
    
    'Calculer l'incrément
    If Arbre(NoeudDepart).NbSuivants = 1 Then
        IncAngle = 0
        AngleDeb = (AngleDeb + AngleFin) / 2
    Else
        If AngleDeb Mod 360 = AngleFin Mod 360 Then
            IncAngle = (AngleFin - AngleDeb) / (Arbre(NoeudDepart).NbSuivants)
        Else
            IncAngle = (AngleFin - AngleDeb) / (Arbre(NoeudDepart).NbSuivants - 1)
        End If
    End If
    
    Dim i
    Dim NewAngleDeb
    Dim NewAngleFin
    Dim Delta
    Dim NewX, NewY
    
    For i = 0 To Arbre(NoeudDepart).NbSuivants - 1
        Delta = (90 - Etape * 9)
        NewAngleDeb = IncAngle * i + AngleDeb - Delta / 2
        NewAngleFin = IncAngle * i + AngleDeb + Delta / 2
        
        NewX = X + Cos((IncAngle * i + AngleDeb) / 180 * 3.1415926535) * ((HauteurArbre(0) - Etape + 1) / HauteurArbre(0) * Etalon1 + 10)
        NewY = Y + Sin((IncAngle * i + AngleDeb) / 180 * 3.1415926535) * ((HauteurArbre(0) - Etape + 1) / HauteurArbre(0) * Etalon1 + 10)
           
        'Tracer une ligne
        Printer.ForeColor = RGB(Etape * 64 Mod 256, Etape * 128 Mod 256, Etape * 32 Mod 256)
        Printer.DrawWidth = (HauteurArbre(0) - Etape) ^ 2
        Printer.Line (X, Y)-(NewX, NewY)
        Printer.DrawWidth = 1
           
        '***
        'PrintRot frmMap.hdc, (x + NewX) / 2, (y + NewY) / 2, Atn((NewY - y) / (NewX - x)), Arbre(Arbre(NoeudDepart).Suivants(i)).Legende
           
        PrinterDessinerNoeudEtFils Arbre(NoeudDepart).Suivants(i), _
           NewAngleDeb, NewAngleFin, _
           NewX, NewY, Etape + 1
    Next i
    End If
    
    'Dessiner la racine
    PrinterDessinerNoeud X, Y, NoeudDepart
End Sub 'PrinterDessinerNoeudEtFils



'Dessiner tous le mindmap
Private Sub PrinterDessinerAllMindMap()
    'Printer.NewPage '*** nécessaire ?
    Printer.ScaleMode = vbPixels
    'Printer.Orientation = vbPRORLandscape
    Printer.PrintQuality = vbPRPQHigh
    
    'Dessiner un cadre
    Printer.FillStyle = vbTransparent
    Printer.FillColor = RGB(255, 255, 255)
    Printer.Line (0, 0)-(Printer.ScaleWidth - 1, Printer.ScaleHeight - 1), , B
    DessinerCartouche
    PrinterDessinerNoeudEtFils 0, 0, 360, Printer.ScaleWidth / 2, Printer.ScaleHeight / 2, 1
    Printer.EndDoc
End Sub 'PrinterDessinerAllMindMap

