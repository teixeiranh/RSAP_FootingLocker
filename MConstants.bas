Attribute VB_Name = "MConstants"
'@Folder "RSAPFootings.00_Utilities"
'@IgnoreModule MoveFieldCloserToUsage
'@IgnoreModule VariableNotUsed
'@IgnoreModule ProcedureNotUsed
'@IgnoreModule ConstantNotUsed
'//////////////////////////////////////////////////////////////////////////////////////////////////
'@ModuleDescription "List of constants to use in the code."
'
'Developer: Nuno Teixeira
'Email: teixeiranh@gmail.com
'//////////////////////////////////////////////////////////////////////////////////////////////////
'
Option Explicit

'@VariableDescription "PI value taken from google.com."
Public Const dPi As Double = 3.14159265359

Public Const BIG_VALUE As Double = 2147483647
Public Const SMALL_VALUE As Double = -2147483647

'Tolerance to round values from RSAP, where 3 is for millimeters.
'Minimum value recommended is 3, maximum depends on the geometry of the structure and its precision.
Public Const TOLERANCE As Integer = 4

Public Const ERROR_MESSAGE As String = "Do not lose your Robot selection!"
    
