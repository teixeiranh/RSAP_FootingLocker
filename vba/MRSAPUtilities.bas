Attribute VB_Name = "MRSAPUtilities"
'@Folder "RSAPFootings.00_Utilities"
'//////////////////////////////////////////////////////////////////////////////////////////////////
'@ModuleDescription "Utilities procedures to use and re-use in other Main procedures."
'RSAP = Robot Structural Analysis Professional
'
'Developer: Nuno Teixeira
'Email: teixeiranh@gmail.com
'//////////////////////////////////////////////////////////////////////////////////////////////////
'
'@IgnoreModule ParameterCanBeByVal
'@IgnoreModule UseMeaningfulName
'@IgnoreModule MoveFieldCloserToUsage
'@IgnoreModule VariableNotUsed
'@IgnoreModule VariableNotAssigned
'@IgnoreModule VariableNotAssigned
Option Explicit

'@VariableDescription "Object representing the RobotApplication."
Private robApp As RobotApplication

'@VariableDescription "Object representing the RobotNodeServer."
Private RobotNodes As RobotNodeServer

'@VariableDescription "Object representing the RobotNodeRigidLinkServer."
Private RLS As RobotNodeRigidLinkServer

'@VariableDescription "Object representing the RobotSelection."
Private NodeSel As RobotSelection

'@VariableDescription "Object representing the RobotNodeRigidLinkData."
Private RLdata As RobotNodeRigidLinkData

'@VariableDescription "Object representing the RobotLabel."
Private Label As RobotLabel

'@VariableDescription "Object representing the RobotLabel."
Private labels As RobotLabelServer

'@VariableDescription "Object representing the RobotNodeCollection."
Private AllNodesCol As RobotNodeCollection

Private I_PT_FRAME As IRobotActiveModelType

'//////////////////////////////////////////////////////////////////////////////////////////////////
'@Description "Verify if RSAP is opened."
'//////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Sub VerifyIfRobotIsOpened()
'    Dim I_PT_FRAME As IRobotActiveModelType
    Set robApp = New RobotApplication
    If Not robApp.Visible Then
        Set robApp = Nothing
        MsgBox "Start Robot and Load Model!", vbOKOnly, "Error"
        End
    Else
        '@Ignore UnassignedVariableUsage
        If (robApp.Project.Type <> I_PT_FRAME) And _
        (robApp.Project.Type <> I_PT_SHELL) And _
        (robApp.Project.Type <> I_PT_BUILDING) Then
            MsgBox "Structure type should be Frame3D or Shell or Building!", vbOKOnly, "Error"
            End
        End If
    End If
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'@Description "Verify if we have any node created in project."
'//////////////////////////////////////////////////////////////////////////////////////////////////
'
'@Ignore ProcedureNotUsed
Public Sub VerifyIfNodesWereCreated()
    Set robApp = New RobotApplication
    Set AllNodesCol = robApp.Project.structure.Nodes.GetAll
    If AllNodesCol.Count = 0 Then
        MsgBox "Please create nodes in Robot!", vbOKOnly, "Error!"
        End
    End If
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'@Description "Verify if there are any nodes currently selected."
'//////////////////////////////////////////////////////////////////////////////////////////////////
'
'@Ignore ProcedureNotUsed
Public Sub VerifyIfNodesAreSelected()
    Set robApp = New RobotApplication
    Set NodeSel = robApp.Project.structure.Selections.Get(I_OT_NODE)
    If NodeSel.Count = 0 Then
        MsgBox "Please select nodes in Robot!", vbOKOnly, "Error!"
        End
    End If
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'@Description "Clears all the references."
'//////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Sub CleanVariables()
    Set robApp = Nothing
    Set AllNodesCol = Nothing
    Set RLS = Nothing
    Set RLdata = Nothing
    Set RobotNodes = Nothing
    Set NodeSel = Nothing
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'@Description "Set the active states for RSAP."
'//////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Sub SetActiveStates()
    Set robApp = New RobotApplication
    robApp.Visible = True
    robApp.Interactive = 1
    robApp.UserControl = True
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'@Description "Creates the thickness label in RSAP."
'//////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Sub AssignThickness(ByRef labels As RobotLabelServer, ByRef slabSectionName As String, _
ByVal foundationThickness As Double)
    
    Dim Label As RobotLabel
    Set Label = labels.Create(I_LT_PANEL_THICKNESS, slabSectionName)
    
    Dim thickness As RobotThicknessData
    Set thickness = Label.Data
    thickness.ThicknessType = I_TT_HOMOGENEOUS
    
    Dim thicknessData As RobotThicknessHomoData
    Set thicknessData = thickness.Data
    thicknessData.ThickConst = foundationThickness
    
    labels.Store Label
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'@Description "Creates the panel back to the RSAP project."
'//////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Sub CreatePanel(ByRef points As RobotPointsArray, ByRef slabSectionName As String)
    
    Dim ObjNumber As Long
    Dim structure As RobotStructure
    Set structure = robApp.Project.structure
    '@Ignore UndeclaredVariable, UnassignedVariableUsage
    ObjNumber = structure.Objects.FreeNumber
    '@Ignore UnassignedVariableUsage
    structure.Objects.CreateContour ObjNumber, points
    '@Ignore UnassignedVariableUsage
    Dim slab As RobotObjObject
    Set slab = structure.Objects.Get(ObjNumber)
    slab.Main.Attribs.Meshed = True
    slab.SetLabel I_LT_PANEL_THICKNESS, slabSectionName
    slab.Initialize
    slab.Update
            
End Sub


