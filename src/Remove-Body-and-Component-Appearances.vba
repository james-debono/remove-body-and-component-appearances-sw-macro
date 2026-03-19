'==============================================================================
' Remove Body and Component Appearances
'
' Clears the appearances that Apply Unique Colours writes, and nothing else.
'
' In a part, appearances applied to bodies are removed, along with any faces
' carrying a copy of a body's colour. A colour applied to a face or a feature on
' its own is left alone.
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
'   Version   0.5.2
'   Date      2026-08-13
'   Author    James Debono
'   Licence   MIT - full text below
'   Source    https://github.com/james-debono/solidworks-apply-colours
'
'------------------------------------------------------------------------------
' CHANGELOG (summary - see CHANGELOG.md for the full history)
'
'   0.5.2   Version reported on completion corrected.
'   0.5.1   Licence and header.
'   0.5.0   An appearance touching a body is cleared whole, faces included.
'   0.4.0   Faces carrying a copy of a body's colour go with their body.
'   0.3.0   Per-appearance diagnostics, which located the propagation problem.
'   0.2.0   Removal scoped to the active display state rather than the
'           configuration.
'   0.1.0   Initial release.
'
'------------------------------------------------------------------------------
' MIT Licence
' SPDX-License-Identifier: MIT
'
' Copyright (c) 2026 James Debono
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'==============================================================================

Option Explicit

'--- Notes for maintenance ----------------------------------------------------
'
' There are no user settings.
'
' Removal works through the render material API, not RemoveMaterialProperty.
' That matters: IBody2::RemoveMaterialProperty is scoped by *configuration*, and
' since several display states live under one configuration, 0.1.0 used it and
' wiped the colour out of every display state at once. GetRenderMaterials2 and
' AddDisplayStateSpecificRenderMaterial take a swDisplayStateOpts_e instead, so
' they can be pointed at the active display state alone.
'
' An appearance is a render material plus the list of entities it is attached to.
' There is no call to detach one entity, so the sequence is: read the entity
' list, drop all of them, add back the ones being kept, and write the result to
' this display state. An appearance left with no entities is gone.
'
' In a part the test is per appearance, not per entity: an appearance that
' touches any body is a body colour and the whole thing goes. That is because
' pattern and mirror features with "Propagate visual properties" on stamp the
' seed body's colour onto the derived bodies' FACES, and those faces end up in
' the same appearance as the seed body. Two narrower rules were tried and both
' failed - 0.3.0 kept every face, which left the affected bodies still coloured
' and took 161 seconds re-attaching 5,359 of them; 0.4.0 dropped a face only when
' its own body was in the same appearance, which never matches, because the
' stamped faces belong to a different body than the seed. An appearance touching
' no body was applied to faces or features deliberately and is still left alone.
'
' EditRebuild3 at the end is not optional. Without it the viewport keeps drawing
' the old colours until the display state is switched away and back.
'
' Nothing here may touch a referenced part file. In an assembly that means
' component entities only - and there the test stays per entity, because anything
' else in an assembly's appearance belongs to a part document. Do not extend the
' whole-appearance rule to assemblies: it would edit files the user never opened,
' including ones shared with other assemblies.
'
' Turning SHOW_DIAGNOSTICS on writes one line per appearance to the VBA Immediate
' window (Ctrl+G in the editor) giving what that appearance was attached to, what
' was cleared and what was kept, plus a timing breakdown. Every fault found while
' building this macro was diagnosed from that output.
Const SHOW_DIAGNOSTICS As Boolean = False

' Must match the Version line in the header block above. build-library.ps1 checks
' that they agree and fails the build if they drift.
Const MACRO_VERSION As String = "0.5.2"

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
        StripAppearances swModel, False
    ElseIf docType = swDocumentTypes_e.swDocASSEMBLY Then
        StripAppearances swModel, True
    Else
        MsgBox "This macro only works on part or assembly documents.", vbCritical
        Exit Sub
    End If

    Exit Sub

mainError:
    MsgBox "An error occurred in main(): " & Err.Description & _
           " (Error " & Err.Number & ")", vbCritical
End Sub

'--- Core ---------------------------------------------------------------------

Sub StripAppearances(swModel As SldWorks.ModelDoc2, ByVal isAssembly As Boolean)

    Dim currentStep As String
    Dim swView As SldWorks.ModelView
    On Error GoTo ErrorHandler

    Dim tStart As Single, tScanned As Single, tStripped As Single, tRefreshed As Single
    tStart = Timer

    currentStep = "Reading appearances for the active display state"
    Dim vAppearances As Variant
    vAppearances = swModel.Extension.GetRenderMaterials2( _
        swDisplayStateOpts_e.swThisDisplayState, Empty)

    tScanned = Timer

    If IsEmpty(vAppearances) Then
        MsgBox "There are no appearances in the active display state.", vbInformation
        Exit Sub
    End If

    Set swView = swModel.ActiveView
    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = False

    Dim examined As Long, strippedTotal As Long, keptTotal As Long
    Dim ownersCleared As Long
    Dim emptied As Long, partial As Long, untouched As Long
    Dim reattachFails As Long, writeBackFails As Long
    examined = 0: strippedTotal = 0: keptTotal = 0: ownersCleared = 0
    emptied = 0: partial = 0: untouched = 0
    reattachFails = 0: writeBackFails = 0

    Dim i As Long, j As Long
    Dim swAppearance As SldWorks.RenderMaterial
    Dim vEnts As Variant

    If SHOW_DIAGNOSTICS Then
        Debug.Print "=== Remove Body and Component Appearances " & MACRO_VERSION & " ==="
        Debug.Print "appearances in this display state: " & (UBound(vAppearances) + 1)
    End If

    currentStep = "Detaching entities"
    For i = 0 To UBound(vAppearances)
        Set swAppearance = vAppearances(i)
        examined = examined + 1

        vEnts = swAppearance.GetEntities

        If SHOW_DIAGNOSTICS Then
            Debug.Print "appearance " & i & ": " & DescribeEntities(vEnts)
        End If

        If Not IsEmpty(vEnts) Then

            ' Does this appearance belong to us at all? One body in a part, or
            ' one component in an assembly, is enough to claim the whole thing.
            Dim isOurs As Boolean
            isOurs = False
            For j = 0 To UBound(vEnts)
                If ShouldStrip(vEnts(j), isAssembly) Then
                    isOurs = True
                    Exit For
                End If
            Next j

            Dim keep As Collection
            Set keep = New Collection
            Dim strippedHere As Long
            strippedHere = 0

            If isOurs Then
                For j = 0 To UBound(vEnts)
                    If isAssembly Then
                        ' Components only. Anything else an assembly appearance
                        ' points at lives in a referenced part file.
                        If ShouldStrip(vEnts(j), True) Then
                            strippedHere = strippedHere + 1
                            ownersCleared = ownersCleared + 1
                        Else
                            keep.Add vEnts(j)
                        End If
                    Else
                        ' The whole appearance goes: the bodies, and the faces
                        ' carrying a stamped copy of a body's colour.
                        strippedHere = strippedHere + 1
                        If ShouldStrip(vEnts(j), False) Then
                            ownersCleared = ownersCleared + 1
                        End If
                    End If
                Next j
            End If

            If strippedHere > 0 Then
                strippedTotal = strippedTotal + strippedHere
                keptTotal = keptTotal + keep.Count

                If keep.Count = 0 Then
                    emptied = emptied + 1
                Else
                    partial = partial + 1
                End If

                If SHOW_DIAGNOSTICS Then
                    Debug.Print "   clearing " & strippedHere & ", keeping " & keep.Count
                End If

                swAppearance.RemoveAllEntities

                Dim k As Long
                For k = 1 To keep.Count
                    If swAppearance.AddEntity(keep(k)) = False Then
                        reattachFails = reattachFails + 1
                        If SHOW_DIAGNOSTICS Then
                            Debug.Print "   FAILED to re-attach entity " & k
                        End If
                    End If
                Next k

                If keep.Count > 0 Then
                    Dim matId1 As Long, matId2 As Long
                    swAppearance.GetMaterialIds matId1, matId2
                    If swModel.Extension.AddDisplayStateSpecificRenderMaterial( _
                        swAppearance, swDisplayStateOpts_e.swThisDisplayState, _
                        Empty, matId1, matId2) = False Then
                        writeBackFails = writeBackFails + 1
                        If SHOW_DIAGNOSTICS Then
                            Debug.Print "   FAILED to write back"
                        End If
                    End If
                End If
            Else
                untouched = untouched + 1
            End If
        Else
            untouched = untouched + 1
        End If
    Next i

    tStripped = Timer

    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = True

    currentStep = "Rebuilding"
    swModel.EditRebuild3
    tRefreshed = Timer

    Dim msg As String
    If isAssembly Then
        msg = "Removed component appearances from the active display state" & vbCrLf & _
              "Components Cleared: " & ownersCleared & vbCrLf & vbCrLf & _
              "The referenced part files were not modified."
    Else
        msg = "Removed body appearances from the active display state" & vbCrLf & _
              "Bodies Cleared: " & ownersCleared & vbCrLf & vbCrLf & _
              "Colours applied to faces or features on their own" & vbCrLf & _
              "were left alone."
    End If

    msg = msg & vbCrLf & vbCrLf & "Macro Version: " & MACRO_VERSION

    If SHOW_DIAGNOSTICS Then
        msg = msg & vbCrLf & vbCrLf & _
              "Appearances examined: " & examined & vbCrLf & _
              "  emptied completely: " & emptied & vbCrLf & _
              "  partly kept: " & partial & vbCrLf & _
              "  left alone: " & untouched & vbCrLf & _
              "Entities cleared: " & strippedTotal & vbCrLf & _
              "Entities kept: " & keptTotal & vbCrLf & _
              "Re-attach failures: " & reattachFails & vbCrLf & _
              "Write-back failures: " & writeBackFails & vbCrLf & vbCrLf & _
              "Scan:    " & Format(tScanned - tStart, "0.00") & " s" & vbCrLf & _
              "Strip:   " & Format(tStripped - tScanned, "0.00") & " s" & vbCrLf & _
              "Rebuild: " & Format(tRefreshed - tStripped, "0.00") & " s" & vbCrLf & _
              "Total:   " & Format(tRefreshed - tStart, "0.00") & " s"

        Debug.Print "emptied=" & emptied & "  partial=" & partial & _
                    "  untouched=" & untouched & _
                    "  cleared=" & strippedTotal & _
                    "  owners=" & ownersCleared & _
                    "  kept=" & keptTotal & _
                    "  reattachFails=" & reattachFails & _
                    "  writeBackFails=" & writeBackFails
        Debug.Print "=== end ==="
    End If

    MsgBox msg, vbInformation
    Exit Sub

ErrorHandler:
    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = True
    MsgBox "ERROR in StripAppearances at step [" & currentStep & "]" & vbCrLf & _
           "Description: " & Err.Description & vbCrLf & _
           "Error " & Err.Number, vbCritical
End Sub

' What makes an appearance ours: a body in a part, a component in an assembly.
'
' In a part this is only used to decide whether the appearance belongs to us at
' all - once it does, everything attached to it goes, because the faces carrying
' a stamped copy of a body's colour sit in the same appearance. In an assembly it
' is applied per entity, since the rest of an assembly's appearance belongs to
' the part files it references.
Function ShouldStrip(ByVal ent As Object, ByVal isAssembly As Boolean) As Boolean
    On Error GoTo NotOurs

    If isAssembly Then
        ShouldStrip = (TypeOf ent Is SldWorks.Component2)
    Else
        ShouldStrip = (TypeOf ent Is SldWorks.Body2)
    End If
    Exit Function

NotOurs:
    ShouldStrip = False
End Function

' A readable summary of what an appearance is attached to, for diagnostics.
Function DescribeEntities(ByVal vEnts As Variant) As String
    On Error GoTo Failed

    If IsEmpty(vEnts) Then
        DescribeEntities = "no entities"
        Exit Function
    End If

    Dim bodies As Long, faces As Long, comps As Long
    Dim feats As Long, others As Long
    Dim i As Long

    For i = 0 To UBound(vEnts)
        If TypeOf vEnts(i) Is SldWorks.Body2 Then
            bodies = bodies + 1
        ElseIf TypeOf vEnts(i) Is SldWorks.Face2 Then
            faces = faces + 1
        ElseIf TypeOf vEnts(i) Is SldWorks.Component2 Then
            comps = comps + 1
        ElseIf TypeOf vEnts(i) Is SldWorks.Feature Then
            feats = feats + 1
        Else
            others = others + 1
        End If
    Next i

    DescribeEntities = (UBound(vEnts) + 1) & " entities [" & _
                       "bodies=" & bodies & _
                       " faces=" & faces & _
                       " components=" & comps & _
                       " features=" & feats & _
                       " other=" & others & "]"
    Exit Function

Failed:
    DescribeEntities = "could not be read (" & Err.Description & ")"
End Function
