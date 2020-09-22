Attribute VB_Name = "modOperations"
'modOperations : Opérations diverses sur le Mindmap
'Par C.Dutoit, 2 Août 2000 (dutoitc@hotmail.com)
'http://www.home.ch/~spaw4758
Option Explicit



Sub InsererFils(Parent As Long, Legende As String, URL As String)
    'Redimensionner l'arbre (+1)
    ReDim Preserve Arbre(UBound(Arbre) + 1)
    
    'Créer le noeud
    Arbre(UBound(Arbre)).Legende = Legende
    Arbre(UBound(Arbre)).NbSuivants = 0
    Arbre(UBound(Arbre)).URL = URL
    
    'Ajouter le fils au parent
    If Arbre(Parent).NbSuivants = 0 Then
        Arbre(Parent).NbSuivants = 1
        ReDim Arbre(Parent).Suivants(0)
        Arbre(Parent).Suivants(0) = UBound(Arbre)
    Else
        ReDim Preserve Arbre(Parent).Suivants(UBound(Arbre(Parent).Suivants) + 1)
        Arbre(Parent).Suivants(UBound(Arbre(Parent).Suivants)) = UBound(Arbre)
        Arbre(Parent).NbSuivants = Arbre(Parent).NbSuivants + 1
    End If
    
    'Redessiner le mindmap
    DessinerAllMindMap
End Sub 'InsererFils



'Editer un noeud
Sub EditerNoeud(Index As Long)
    If Index < 0 Then Exit Sub
    
    Arbre(Index).Legende = InputBox("Entrez la légende", "Editer un noeud (1/2)", Arbre(Index).Legende)
    Arbre(Index).URL = InputBox("Entrez l'URL", "Editer un noeud (2/2)", Arbre(Index).URL)
End Sub 'EditerNoeud



'Supprimer un noeud
Sub SupprimerNoeud(Index As Long)
    'Indice correct ?
    If Index < 0 Or Index > UBound(Arbre) Then
        MsgBox "Tentative de suppression d'un noeud inexistant", vbExclamation, "Erreur..."
        Exit Sub
    End If
    
    'Tentative de suppression de la racine ?
    If Index = 0 Then
        MsgBox "impossible de supprimer le premier noeud !", vbExclamation, "Erreur..."
        Exit Sub
    End If
    
    'Supprimer de l'arbre
    Dim i, j
    For i = Index + 1 To UBound(Arbre)
        Arbre(i - 1) = Arbre(i)
    Next i
    ReDim Preserve Arbre(UBound(Arbre) - 1)
    
    'Supprimer le lien depuis le parent
    Dim k
    Dim found As Boolean
    found = False
    For i = 0 To UBound(Arbre)
        If Arbre(i).NbSuivants > 0 Then
            For j = 0 To UBound(Arbre(i).Suivants)
                If Arbre(i).Suivants(j) = Index Then 'Supprimer la référence
                    'Décaler les suivants
                    For k = j + 1 To UBound(Arbre(i).Suivants)
                        Arbre(i).Suivants(k - 1) = Arbre(i).Suivants(k)
                    Next k
                    
                    'Redimensionner l'arbre
                    ReDim Preserve Arbre(i).Suivants(UBound(Arbre(i).Suivants) - 1)
                    Arbre(i).NbSuivants = Arbre(i).NbSuivants - 1
                    found = True
                End If
                If found Then Exit For
            Next j
        End If
        If found Then Exit For
    Next i
    
    'Déplacer les liens sur les indices supérieur à l'indice du noeud à supprimer
    For i = 0 To UBound(Arbre)
        If Arbre(i).NbSuivants > 0 Then
            For j = 0 To UBound(Arbre(i).Suivants)
                If Arbre(i).Suivants(j) > Index Then Arbre(i).Suivants(j) = Arbre(i).Suivants(j) - 1
            Next j
        End If
    Next i
End Sub 'SupprimerNoeud
