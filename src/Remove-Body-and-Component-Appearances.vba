'==============================================================================
' Remove Body and Component Appearances
'
' Clears the appearances that Apply Unique Colours writes, and nothing else.
'
' In a part, appearances applied at body level are removed, along with the faces
' that were only carrying the body's colour. A colour applied to a face on its
' own is left alone.
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
'   Version   0.4.0
'   Date      2026-08-07
'   Author    James Debono
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
' A body's appearance is listed against the body AND against every face that
' inherits it. Those faces have to go with the body. 0.3.0 kept them - on a
' 56 body weldment that meant putting back 5,359 faces one AddEntity call at a
' time, which took 161 seconds and left every affected body still coloured
' through its own faces. A face the user coloured deliberately sits in its own
' appearance, so it never shares an entity list with a body being stripped and
' is never touched by this rule.
'
' Nothing here may touch a referenced part file. In an assembly that means
' component entities only: those are held by the assembly. Do not extend this to
' faces or bodies in an assembly - those belong to the part documents, and
' editing them would dirty files the user never opened, including ones shared
' with other assemblies.
'
' On while this macro is still being proven. It writes one line per appearance to
' the VBA Immediate window (Ctrl+G in the editor) giving what that appearance was
' attached to, what was cleared, what was kept, and whether writing it back
' succeeded - plus a timing breakdown.
Const SHOW_DIAGNOSTICS As Boolean = True

Const MACRO_VERSION As String = "0.4.0"

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
    Dim facesFollowed As Long
    Dim emptied As Long, partial As Long, untouched As Long
    Dim reattachFails As Long, writeBackFails As Long
    examined = 0: strippedTotal = 0: keptTotal = 0: facesFollowed = 0
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

            ' First pass: which bodies is this appearance losing? Needed before
            ' the faces can be judged, since a face is only dropped when the body
            ' it belongs to is being dropped from this same appearance.
            Dim losing As Object
            Set losing = CreateObject("Scripting.Dictionary")

            If Not isAssembly Then
                For j = 0 To UBound(vEnts)
                    If TypeOf vEnts(j) Is SldWorks.Body2 Then
                        Dim swOwnedBody As SldWorks.Body2
                        Set swOwnedBody = vEnts(j)
                        If Not losing.Exists(swOwnedBody.Name) Then
                            losing.Add swOwnedBody.Name, True
                        End If
                    End If
                Next j
            End If

            ' Second pass: sort the entities into those going and those staying.
            Dim keep As Collection
            Set keep = New Collection
            Dim strippedHere As Long
            strippedHere = 0

            For j = 0 To UBound(vEnts)
                If ShouldStrip(vEnts(j), isAssembly) Then
                    strippedHere = strippedHere + 1
                ElseIf FaceBelongsToStrippedBody(vEnts(j), losing, isAssembly) Then
                    strippedHere = strippedHere + 1
                    facesFollowed = facesFollowed + 1
                Else
                    keep.Add vEnts(j)
                End If
            Next j

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

                ' No API detaches a single entity, so drop the lot and put back
                ' the survivors.
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

                ' Written back against this display state only. An appearance
                ' with nothing left attached simply ceases to exist.
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
              "Components Cleared: " & strippedTotal & vbCrLf & vbCrLf & _
              "The referenced part files were not modified."
    Else
        msg = "Removed body appearances from the active display state" & vbCrLf & _
              "Bodies Cleared: " & (strippedTotal - facesFollowed) & vbCrLf & vbCrLf & _
              "Face, feature and part appearances were left alone."
    End If

    msg = msg & vbCrLf & vbCrLf & "Macro Version: " & MACRO_VERSION

    If SHOW_DIAGNOSTICS Then
        msg = msg & vbCrLf & vbCrLf & _
              "Appearances examined: " & examined & vbCrLf & _
              "  emptied completely: " & emptied & vbCrLf & _
              "  partly kept: " & partial & vbCrLf & _
              "  left alone: " & untouched & vbCrLf & _
              "Faces dropped with their body: " & facesFollowed & vbCrLf & _
              "Entities kept: " & keptTotal & vbCrLf & _
              "Re-attach failures: " & reattachFails & vbCrLf & _
              "Write-back failures: " & writeBackFails & vbCrLf & vbCrLf & _
              "Scan:    " & Format(tScanned - tStart, "0.00") & " s" & vbCrLf & _
              "Strip:   " & Format(tStripped - tScanned, "0.00") & " s" & vbCrLf & _
              "Rebuild: " & Format(tRefreshed - tStripped, "0.00") & " s" & vbCrLf & _
              "Total:   " & Format(tRefreshed - tStart, "0.00") & " s"

        Debug.Print "emptied=" & emptied & "  partial=" & partial & _
                    "  untouched=" & untouched & _
                    "  facesFollowed=" & facesFollowed & _
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

' Which entities this macro is responsible for clearing.
'
' A part's appearances may be attached to faces, features, bodies or the part
' itself; only bodies are ours. An assembly's may be attached to components or
' to the assembly document; only components are ours, and those are the only
' ones the assembly owns rather than the part files it references.
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

' True when this entity is a face that is only carrying the colour of a body
' being cleared from the same appearance.
'
' Removing the body alone is not enough: the colour stays visible through the
' faces, because a face-level appearance wins over a body-level one. A face the
' user coloured deliberately belongs to a different appearance and so is never
' in this entity list to begin with.
Function FaceBelongsToStrippedBody(ByVal ent As Object, ByVal losing As Object, _
                                   ByVal isAssembly As Boolean) As Boolean
    On Error GoTo NotOurs

    FaceBelongsToStrippedBody = False
    If isAssembly Then Exit Function
    If losing Is Nothing Then Exit Function
    If losing.Count = 0 Then Exit Function
    If Not (TypeOf ent Is SldWorks.Face2) Then Exit Function

    Dim swFace As SldWorks.Face2
    Set swFace = ent

    Dim swOwner As SldWorks.Body2
    Set swOwner = swFace.GetBody()
    If swOwner Is Nothing Then Exit Function

    FaceBelongsToStrippedBody = losing.Exists(swOwner.Name)
    Exit Function

NotOurs:
    FaceBelongsToStrippedBody = False
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
