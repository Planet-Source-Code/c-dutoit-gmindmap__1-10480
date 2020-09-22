VERSION 5.00
Begin VB.Form frmMap 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0FFFF&
   Caption         =   "MindMap"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   329
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   692
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmMap
Option Explicit

Private Sub Form_Click()
   DessinerAllMindMap
End Sub 'Form_Click


'Edition d'un noeud
Private Sub Form_DblClick()
    'Editer le noeud et redessiner le mindmap
    EditerNoeud (NoeudSelectionne)
    DessinerAllMindMap
    
    'Définir le titre de la fenêtre principale + ...
    If Not MyApp.Modifie Then
        MyApp.Modifie = True
        SetAppCaption
    End If
End Sub 'Form_DblClick


'Supprimer le noeud sélectionné
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        SupprimerNoeud (NoeudSelectionne)
        DessinerAllMindMap
    
        'Définir le titre de la fenêtre principale + ...
        If Not MyApp.Modifie Then
            MyApp.Modifie = True
            SetAppCaption
        End If
    Else: If KeyCode = vbKeyInsert Then frmMDI.mnuNoeudInsererFils_Click
    End If
End Sub 'Form_KeyDown


'Initialiser le mindmap
Private Sub Form_Load()
   frmMap.WindowState = vbMaximized
   DoEvents
   NouveauFichier
   DessinerAllMindMap
End Sub 'Form_Load


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   NoeudSelectionne = NoeudLePlusProche(Int(X), Int(Y))
End Sub 'Form_MouseDown


Private Sub Form_Paint()
    'Mettre à jour l'affichage
    DessinerAllMindMap
End Sub 'Form_Paint


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Cancel = 1
End Sub 'Form_QueryUnload
