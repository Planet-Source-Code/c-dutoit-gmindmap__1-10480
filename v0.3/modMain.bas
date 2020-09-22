Attribute VB_Name = "modMain"
'modMain : Module principale, divers
'Par C.Dutoit, 2 Août 2000 (dutoitc@hotmail.com)
'http://www.home.ch/~spaw4758
Option Explicit

Type TMyApp 'Données vitales pour l'application
    Fichier As String  'Nom de fichier actuel
    Modifie As Boolean 'Fichier modifié ?
End Type 'TMyApp

Global MyApp As TMyApp  'Données vitales de l'applications



'Définir le titre de la fenêtre principale
Sub SetAppCaption()
    Dim Caption As String
    Caption = App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision & " "
    
    
    If MyApp.Fichier <> "" Then Caption = Caption & "[" & MyApp.Fichier & "] "
    If MyApp.Modifie Then Caption = Caption & "*"
    
    frmMDI.Caption = Caption
End Sub 'SetAppCaption
