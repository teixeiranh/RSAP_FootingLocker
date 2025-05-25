VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWallFootingDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "RSAPFootings.02_SpreadFooting"

'//////////////////////////////////////////////////////////////////////////////////////////////////
'@ModuleDescription "Data transfer class that enables the tranfer of data between procedures easily."
'To be used with the footing feature.
'All the fields are public, so there are no getters nor setters.
'
'Developer: Nuno Teixeira
'Email: teixeiranh@gmail.com
'//////////////////////////////////////////////////////////////////////////////////////////////////
'
'@IgnoreModule VariableNotUsed
'@IgnoreModule ProcedureNotUsed
'@IgnoreModule EncapsulatePublicField
'@IgnoreModule ParameterCanBeByVal
'@IgnoreModule UseMeaningfulName
Option Explicit

Private Const HIGHEST_POINT As Double = 2147483647
Private Const LOWEST_POINT As Double = -2147483648#

'@VariableDescription "zz value for the bottom edge of the panel."
Public z As Double

'@VariableDescription "xx value for the point 1 of the panel."
Public x1 As Double

'@VariableDescription "xx value for the point 2 of the panel."
Public x2 As Double

'@VariableDescription "Auxiliary varible for the determination of yy values."
Public y1Aux As Double

'@VariableDescription "Auxiliary varible for the determination of yy values."
Public y2Aux As Double

'@VariableDescription "yy value for the point 1 of the panel (lower xx value)."
Public y1 As Double

'@VariableDescription "yy value for the point 2 of the panel (higher xx value)."
Public y2 As Double

'@VariableDescription "Horizontal angle of the panel."
Public horizontalAngle As Double

'@VariableDescription "Horizontal angle of the panel."
Public complementaryhorizontalAngle As Double

'//////////////////////////////////////////////////////////////////////////////////////////////////
'@Description "Constructor that initializes the object with very big absolute numbers."
'//////////////////////////////////////////////////////////////////////////////////////////////////
'
Private Sub Class_Initialize()
    Me.z = HIGHEST_POINT
    Me.x1 = HIGHEST_POINT
    Me.x2 = LOWEST_POINT
    Me.y1Aux = HIGHEST_POINT
    Me.y2Aux = LOWEST_POINT
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'@Description "Calculate Me panel's horizontal angle in a 2D projection."
'//////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Function CalculateHorizontalAngle() As Double
    If Me.x2 = Me.x1 Then
        CalculateHorizontalAngle = MConstants.dPi / 2
    ElseIf Me.y2 > Me.y1 Then
        CalculateHorizontalAngle = Atn((Me.y2 - Me.y1) / (Me.x2 - Me.x1))
    ElseIf Me.y2 = Me.y1 Then
        CalculateHorizontalAngle = 0
    Else
        CalculateHorizontalAngle = MConstants.dPi + Atn((Me.y2 - Me.y1) / (Me.x2 - Me.x1))
    End If
    horizontalAngle = CalculateHorizontalAngle
End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////
'@Description "Verify if point 1 of Me panel is coincident with point 1 of the given argument."
'//////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Function IsPoint1Coincident(ByRef adjacentPanelDTO As CWallFootingDTO) As Boolean
    If (Me.x1 = adjacentPanelDTO.x1 And Me.y1 = adjacentPanelDTO.y1) _
    Or (Me.x1 = adjacentPanelDTO.x2 And Me.y1 = adjacentPanelDTO.y2) Then
        IsPoint1Coincident = True
    Else
        IsPoint1Coincident = False
    End If
End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////
'@Description "Verify if point 2 of Me panel is coincident with point 2 of the given argument."
'//////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Function IsPoint4Coincident(ByRef adjacentPanelDTO As CWallFootingDTO) As Boolean
    If (Me.x2 = adjacentPanelDTO.x1 And Me.y2 = adjacentPanelDTO.y1) _
    Or (Me.x2 = adjacentPanelDTO.x2 And Me.y2 = adjacentPanelDTO.y2) Then
        IsPoint4Coincident = True
    Else
        IsPoint4Coincident = False
    End If
End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////
'@Description "Verify if Me panel has the same coordinates of the argument object."
'//////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Function IsTheSamePanel(ByRef dto As CWallFootingDTO) As Boolean
    If Me.x1 = dto.x1 And Me.y1 = dto.y1 And Me.x2 = dto.x2 And Me.y2 = dto.y2 And Me.z = dto.z Then
        IsTheSamePanel = True
    Else
        IsTheSamePanel = False
    End If
End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////
'@Description "Verify if point 1 of Me panel is coincident with a point with coordinates."
'//////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Function IsPoint1CoincidentCoordinates(ByVal x1Coordinate As Double, _
ByVal y1Coordinate As Double) As Boolean
    If Me.x1 = x1Coordinate And Me.y1 = y1Coordinate Then
        IsPoint1CoincidentCoordinates = True
    Else
        IsPoint1CoincidentCoordinates = False
    End If
End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////
'@Description "Verify if point 4 of Me panel is coincident with a point with coordinates."
'//////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Function IsPoint4CoincidentCoordinates(ByVal x1Coordinate As Double, _
ByVal y1Coordinate As Double) As Boolean
    If Me.x2 = x1Coordinate And Me.y2 = y1Coordinate Then
        IsPoint4CoincidentCoordinates = True
    Else
        IsPoint4CoincidentCoordinates = False
    End If
End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////
'@Description "Finds the angle to the horizontal of the first and end points of the panels edge."
'//////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Sub FindHorizontalAngles()
    If Me.x2 = Me.x1 Then
        Me.horizontalAngle = MConstants.dPi / 2
        Me.complementaryhorizontalAngle = MConstants.dPi / 2 + MConstants.dPi
    ElseIf Me.y2 > Me.y1 Then
        Me.horizontalAngle = Atn((Me.y2 - Me.y1) / (Me.x2 - Me.x1))
        Me.complementaryhorizontalAngle = Me.horizontalAngle + MConstants.dPi
    ElseIf Me.y2 = Me.y1 Then
        Me.horizontalAngle = 0
        Me.complementaryhorizontalAngle = MConstants.dPi
    ElseIf Me.y1 > Me.y2 Then
        Me.horizontalAngle = 2 * MConstants.dPi - Atn((Me.y1 - Me.y2) / (Me.x2 - Me.x1))
        Me.complementaryhorizontalAngle = Me.horizontalAngle - MConstants.dPi
    End If
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'@Description "Verify if the current panel has a positive inclination in plan."
'//////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Function IsPositiveIncline() As Boolean
    If Me.y2 > Me.y1 Then IsPositiveIncline = True
End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////
'@Description "Verify if the current panel has a negative inclination in plan."
'//////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Function IsNegativeIncline() As Boolean
    If Me.y2 < Me.y1 Then IsNegativeIncline = True
End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////
'@Description "Verify if the current panel is vertical."
'//////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Function IsVertical() As Boolean
    If Me.x1 = Me.x2 Then IsVertical = True
End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////
'@Description "Verify if the current panel is horizontal."
'//////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Function IsHorizontal() As Boolean
    If Me.y1 = Me.y2 Then IsHorizontal = True
End Function


