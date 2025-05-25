VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CColumnFootingDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule EncapsulatePublicField
'@Folder "RSAPFootings.01_IsolatedFooting"
'//////////////////////////////////////////////////////////////////////////////////////////////////
'Data transfer object for the column footing.
'All the fields are public, so there are no getters, nor setters.
'
'Developer: Nuno Teixeira
'Email: teixeiranh@gmail.com
'//////////////////////////////////////////////////////////////////////////////////////////////////
'
Option Explicit

Public raioMaior As Double
Public angleOfRotation As Double
Public anguloDiagonal As Double
Public excentricadeX As Double
Public excentricadeY As Double
