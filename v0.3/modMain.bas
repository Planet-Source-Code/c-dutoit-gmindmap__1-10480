Attribute VB_Name = "modMain"
'modMain : Module principale, divers
'Par C.Dutoit, 2 Ao�t 2000 (dutoitc@hotmail.com)
'http://www.home.ch/~spaw4758
Option Explicit

Type TMyApp 'Donn�es vitales pour l'application
    Fichier As String  'Nom de fichier actuel
    Modifie As Boolean 'Fichier modifi� ?
End Type 'TMyApp

Global MyApp As TMyApp  'Donn�es vitales de l'applications



'D�finir le titre de la fen�tre principale
Sub SetAppCaption()
    Dim Caption As String
    Caption = App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision & " "
    
    
    If MyApp.Fichier <> "" Then Caption = Caption & "[" & MyApp.Fichier & "] "
    If MyApp.Modifie Then Caption = Caption & "*"
    
    frmMDI.Caption = Caption
End Sub 'SetAppCaption
