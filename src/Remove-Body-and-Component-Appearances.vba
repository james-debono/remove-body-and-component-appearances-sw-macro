'==============================================================================
' Remove Body and Component Appearances
'
' Clears the appearances that Apply Unique Colours writes, and nothing else.
'
' In a part, appearances applied at body level are removed. Colours applied to
' individual faces, to features, or to the part as a whole are left alone.
'
' In an assembly, the component-level appearances held by the assembly are
' removed. The part files the assembly references are never modified - whatever
' colours live inside them stay exactly as they are.
'
' Only the active display state is affected, so a coloured display state can sit
' alongside a cleared one in the same document.
'
' To use, open a part or assembly document and run the macro.
'
'   Version   0.1.0
'   Date      2026-08-07
'   Author    James Debono
'==============================================================================

Option Explicit

'--- Notes for maintenance ----------------------------------------------------
'
' There are no user settings.
'
' Nothing here may touch a referenced part file. In an assembly that means
' component-level removal only: IComponent2::RemoveMaterialProperty clears the
' override the assembly holds, never the part itself. Do not extend the assembly
' path to bodies or faces - that would edit every part file in the model,
' including ones other assemblies share, and dirty documents the user never
' opened.
'
' Removal is scoped with swInConfigurationOpts_e.swThisConfiguration. Display
' states are held against configurations, so this reaches the active display
' state, but that is an inference rather than a documented guarantee. It is worth
' re-testing after a SOLIDWORKS upgrade: colour two display states, clear one,
' and confirm the other still holds its colours.
'
' Turning SHOW_DIAGNOSTICS on reports how many entities were cleared against how
' many were tried, and names any that refused. A gap between the two is the first
' sign that an API is behaving differently to what is assumed above.
Const SHOW_DIAGNOSTICS As Boolean = False

Const MACRO_VERSION As String = "0.1.0"

Dim swApp As SldWorks.SldWorks

'--- Entry point --------------------------------------------------------------

Sub main()
    On Error GoTo mainError

    Set swApp = Application.SldWorks

    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "Please open a part or assembly document.", vbCritical
        Exit Sub
    End If

    Dim docType As Long
    docType = swModel.GetType()

    If docType = swDocumentTypes_e.swDocPART Then
        ClearPartBodies swModel
    ElseIf docType = swDocumentTypes_e.swDocASSEMBLY Then
        ClearAssemblyComponents swModel
    Else
        MsgBox "This macro only works on part or assembly documents.", vbCritical
        Exit Sub
    End If

    swModel.GraphicsRedraw2
    Exit Sub

mainError:
    MsgBox "An error occurred in main(): " & Err.Description & _
           " (Error " & Err.Number & ")", vbCritical
End Sub

'--- Part: body-level appearances ---------------------------------------------

Sub ClearPartBodies(swModel As SldWorks.ModelDoc2)

    Dim currentStep As String
    Dim swView As SldWorks.ModelView
    On Error GoTo ErrorHandler

    currentStep = "Loading bodies"
    Dim swPart As SldWorks.PartDoc
    Set swPart = swModel

    Dim vBodies As Variant
    vBodies = swPart.GetBodies2(swBodyType_e.swAllBodies, False)

    If IsEmpty(vBodies) Then
        MsgBox "No bodies found in the active part.", vbInformation
        Exit Sub
    End If

    Dim totalBodies As Long
    totalBodies = UBound(vBodies) + 1

    Set swView = swModel.ActiveView
    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = False

    currentStep = "Removing body appearances"
    Dim cleared As Long
    Dim i As Long
    Dim swBody As SldWorks.Body2

    cleared = 0
    For i = 0 To totalBodies - 1
        Set swBody = vBodies(i)

        On Error Resume Next
        swBody.RemoveMaterialProperty swInConfigurationOpts_e.swThisConfiguration, Empty
        If Err.Number = 0 Then
            cleared = cleared + 1
        ElseIf SHOW_DIAGNOSTICS Then
            Debug.Print "refused: " & swBody.Name & "  (" & Err.Description & ")"
        End If
        Err.Clear
        On Error GoTo ErrorHandler
    Next i

    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = True
    swModel.GraphicsRedraw2

    Dim msg As String
    msg = "Removed body appearances from the active display state" & vbCrLf & _
          "Bodies Cleared: " & cleared & " of " & totalBodies & vbCrLf & vbCrLf & _
          "Face, feature and part appearances were left alone." & vbCrLf & vbCrLf & _
          "Macro Version: " & MACRO_VERSION

    MsgBox msg, vbInformation
    Exit Sub

ErrorHandler:
    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = True
    MsgBox "ERROR in ClearPartBodies at step [" & currentStep & "]" & vbCrLf & _
           "Description: " & Err.Description & vbCrLf & _
           "Error " & Err.Number, vbCritical
End Sub

'--- Assembly: component-level appearances only -------------------------------

Sub ClearAssemblyComponents(swModel As SldWorks.ModelDoc2)

    Dim currentStep As String
    Dim swView As SldWorks.ModelView
    On Error GoTo ErrorHandler

    currentStep = "Loading components"
    Dim swAssy As SldWorks.AssemblyDoc
    Set swAssy = swModel

    Dim vComps As Variant
    vComps = swAssy.GetComponents(False)

    If IsEmpty(vComps) Then
        MsgBox "No components found in the active assembly.", vbInformation
        Exit Sub
    End If

    Dim totalComps As Long
    totalComps = UBound(vComps) + 1

    Set swView = swModel.ActiveView
    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = False

    ' Component level only. This clears the override the assembly holds against
    ' each component; the part files themselves are not opened, modified or
    ' marked dirty.
    currentStep = "Removing component appearances"
    Dim cleared As Long
    Dim refused As Long
    Dim i As Long
    Dim swComp As SldWorks.Component2

    cleared = 0
    refused = 0
    For i = 0 To totalComps - 1
        Set swComp = vComps(i)

        On Error Resume Next
        swComp.RemoveMaterialProperty swInConfigurationOpts_e.swThisConfiguration, Empty
        If Err.Number = 0 Then
            cleared = cleared + 1
        Else
            refused = refused + 1
            If SHOW_DIAGNOSTICS Then
                Debug.Print "refused: " & swComp.Name2 & "  (" & Err.Description & ")"
            End If
        End If
        Err.Clear
        On Error GoTo ErrorHandler
    Next i

    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = True
    swModel.GraphicsRedraw2

    Dim msg As String
    msg = "Removed component appearances from the active display state" & vbCrLf & _
          "Components Cleared: " & cleared & " of " & totalComps & vbCrLf & vbCrLf & _
          "The referenced part files were not modified." & vbCrLf & vbCrLf & _
          "Macro Version: " & MACRO_VERSION

    If refused > 0 And SHOW_DIAGNOSTICS Then
        msg = msg & vbCrLf & vbCrLf & "Refused: " & refused & _
              " - see the Immediate window (Ctrl+G)"
    End If

    MsgBox msg, vbInformation
    Exit Sub

ErrorHandler:
    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = True
    MsgBox "ERROR in ClearAssemblyComponents at step [" & currentStep & "]" & vbCrLf & _
           "Description: " & Err.Description & vbCrLf & _
           "Error " & Err.Number, vbCritical
End Sub
