Attribute VB_Name = "modMap"
'modMap : Gestion de l'affichage du mindmap + structure de donnée
'Par C.Dutoit, 1er Août 2000 (dutoitc@hotmail.com)
'http://www.home.ch/~spaw4758
Option Explicit

'max 10 fils !


'Un Noeud
Type TNoeud
    Legende As String       'Légende du noeud
    URL As String           'URL
    X As Long
    Y As Long               'Position centrale
    NbSuivants As Byte      'Nombre de fils
    Suivants() As Long      'Liste des fils
End Type 'TNoeud


Global Arbre() As TNoeud         'L'arbre du mindmap
Global NoeudSelectionne As Long  'Noeud sélectionné



'Dessiner un noeud
Private Sub DessinerNoeud(X, Y, Index As Long)
    Dim txtW As Long
    Dim txtH As Long
    Dim w As Long           'Largeur
    Dim h As Long           'Hauteur
    
    'Calculer la hauteur et la largeur
    txtW = frmMap.TextWidth(Arbre(Index).Legende)
    txtH = frmMap.TextHeight(Arbre(Index).Legende)
    w = txtW * 0.5 + frmMap.TextWidth("OO")
    h = txtH * 0.5 + frmMap.TextHeight("O") / 2
    
    'Dessiner le centre
    frmMap.FillColor = RGB(255, 255, 255)
    frmMap.FillStyle = 0 'solide
    frmMap.DrawWidth = 2
    frmMap.Circle (X, Y), w, , , , h / w
    frmMap.DrawWidth = 1
    
    'Sélectionné ? => tracer un cadre traitillé autour de l'ellipse
    If Index = NoeudSelectionne Then
        frmMap.ForeColor = 0
        frmMap.DrawStyle = 2
        frmMap.FillStyle = 1 'transparent
        frmMap.Line (X - txtW / 2 - 2, Y - txtH / 2 - 2)-(X + txtW / 2 + 2, Y + txtH / 2 + 2), , B
        frmMap.DrawStyle = 0
    End If
    
    'Afficher le label
    frmMap.CurrentX = X - txtW / 2
    frmMap.CurrentY = Y - txtH / 2
    frmMap.ForeColor = 0 'Couleur du cadre
    'frmMap.BackColor = RGB(255, 255, 200)
    'frmMap.FillColor = RGB(0, 255, 0)
    frmMap.Print Arbre(Index).Legende & vbCrLf & Arbre(Index).URL
    
    'Enregistrer la position
    Arbre(Index).X = X
    Arbre(Index).Y = Y
End Sub 'DessinerNoeud



Private Sub DessinerNoeudEtFils(NoeudDepart As Long, AngleDeb, AngleFin, X, Y, Etape)
    Dim Etalon1 As Long
    Etalon1 = frmMap.ScaleWidth / 20

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
            'Dist = frmMap.TextWidth(Arbre(Arbre(NoeudDepart).Suivants(i)).Legende) * 2
            'If NoeudDepart = 0 Then Dist = Dist + frmMap.TextWidth(Arbre(0).Legende)
            'Dist = Dist * 1.1
            
            'Calculer la pos. finale
            Dim texte As String
            Dim AngleTexte As Long
            Dim HCar As Byte
            AngleTexte = Angle * 180 / 3.1415926535 '-Atn((NewY - Y) / (NewX - X)) * 180 / 3.1415926535
            If AngleTexte Mod 360 > 90 And AngleTexte Mod 360 < 270 Then AngleTexte = AngleTexte Mod 360 - 180
            texte = Arbre(Arbre(NoeudDepart).Suivants(i)).Legende
            HCar = ((HauteurArbre(0) - Etape) * 3 / HauteurArbre(0)) ^ 2 + 8
        
            NewX = X + LongueurTexteRot(texte & "OO", HCar) * Cos(Angle)  ' * Dist '((HauteurArbre(0) - Etape + 1) / HauteurArbre(0) * Dist + 10)
            NewY = Y - LongueurTexteRot(texte & "OO", HCar) * Sin(Angle)  '* Dist '((HauteurArbre(0) - Etape + 1) / HauteurArbre(0) * Dist + 10)

            If NoeudDepart = 0 Then 'fils de racine ? => agrandir
                NewX = NewX + Cos(Angle) * Etalon1
                NewY = NewY - Sin(Angle) * Etalon1
            End If
            
           
            'Tracer une ligne
            frmMap.ForeColor = RGB(Etape * 64 Mod 256, Etape * 128 Mod 256, Etape * 32 Mod 256)
            frmMap.DrawWidth = (HauteurArbre(0) - Etape) + 1
            frmMap.Line (X, Y)-(NewX, NewY)
            frmMap.DrawWidth = 1
           
            '***
            PrintRotfrmMap (X + NewX) / 2, (Y + NewY) / 2, AngleTexte, texte, HCar
                   
            DessinerNoeudEtFils Arbre(NoeudDepart).Suivants(i), _
               NewAngleDeb, NewAngleFin, _
                NewX, NewY, Etape + 1
        Next i
    End If
    
    'Dessiner la racine
    If Etape = 1 Then DessinerNoeud X, Y, NoeudDepart
    
    
    'Enregistrer la position
    Arbre(NoeudDepart).X = X
    Arbre(NoeudDepart).Y = Y
    If NoeudSelectionne = NoeudDepart And NoeudSelectionne <> 0 Then frmMap.Circle (X, Y), 5, RGB(255, 0, 0)
End Sub 'DessinerNoeudEtFils



'Dessiner tous le mindmap
Sub DessinerAllMindMap()
    frmMap.Cls
    DessinerNoeudEtFils 0, 0, 360, frmMap.ScaleWidth / 2, frmMap.ScaleHeight / 2, 1
End Sub 'DessinerAllMindMap



Function HauteurArbre(Racine) As Long
    Dim h As Long       'Hauteur de l'arbre
    h = 0               'Hauteur à 0
    
    'Hauteur des fils
    Dim i, HTemp
    For i = 0 To Arbre(Racine).NbSuivants - 1
        HTemp = HauteurArbre(Arbre(Racine).Suivants(i))
        If HTemp > h Then h = HTemp
    Next i
    
    'Retourner la hauteur + 1 pour cet étage
    HauteurArbre = h + 1
End Function 'HauteurArbre



'Retourner le N° du noeud le plus proche
Function NoeudLePlusProche(X As Long, Y As Long) As Long
    Dim i As Long      'Variable de boucle
    Dim dist As Long, DistTemp As Long 'Distance au point
    Dim Noeud As Long  'Noeud le plus proche
    
    'Initialisation
    dist = -1
    
    'Chercher le point le plus proche
    For i = 0 To UBound(Arbre)
        'Calculer la distance au point
        DistTemp = Sqr((Arbre(i).X - X) ^ 2 + (Arbre(i).Y - Y) ^ 2)
        
        'Distance plus petite ? => on enregistre le point et la distance
        If dist = -1 Or DistTemp < dist Then
            dist = DistTemp
            Noeud = i
        End If
    Next i
    
    'Retourner le noeud le plus proche
    NoeudLePlusProche = Noeud
End Function 'NoeudLePlusProche
