Attribute VB_Name = "PWDMathType133Check"
Option Explicit
' Stable 1.3.3: preserve unrelated fields across MathType's whole-document updates.
Private gridCalibrated As Boolean
Private gridOffSetting As Long

Private Const MT_UI As String = "MathTypeCommands.UILib."
Private busy As Boolean
Private lastConverted As InlineShape


Public Sub PWD_Run133_Check()
    Dim doc As Document, choice As VbMsgBoxResult, backupPath As String
    Dim processingStarted As Boolean, message As String, stamp As String
    Dim count As Long, b As Bookmark, answer As VbMsgBoxResult
    Dim phaseStart As Single, processingSeconds As Double
    Dim protectedFields As New Collection, afterText As String
    Dim fresh As Collection, nativeBefore As Long, converted As Long
    Dim names As New Collection, originalText As String, otherCount As Long
    Dim tableCount As Long, floatCount As Long, warnings As String, completion As String
    On Error GoTo Failed
    If Documents.Count = 0 Then Err.Raise vbObjectError + 8100, , Zh("8BF7 5148 6253 5F00 9700 8981 5904 7406 7684 0057 006F 0072 0064 6587 6863 3002")
    Set doc = ActiveDocument
    If doc.ReadOnly Or doc.ProtectionType <> wdNoProtection Or doc.TrackRevisions Then _
        Err.Raise vbObjectError + 8101, , Zh("8BF7 4F7F 7528 53EF 7F16 8F91 6587 6863 FF0C 5E76 5173 95ED 4FEE 8BA2 6A21 5F0F 3002")
    phaseStart = Timer
    ValidateExistingInputs doc
    Set fresh = DiscoverNewInputs(doc)
    For Each b In doc.Bookmarks
        If Left$(b.Name, 7) = "PWD_EQ_" And b.Range.StoryType = wdMainTextStory Then count = count + 1
    Next b
    If count = 0 And fresh.Count = 0 Then Err.Raise vbObjectError + 8102, , "No recognized formula inputs or markers."
    processingSeconds = ElapsedSeconds(phaseStart)
    choice = MsgBox(Zh("5F00 59CB 524D 662F 5426 5907 4EFD FF1F") & vbCrLf & Zh("201C 662F 201D FF1A 5148 4FDD 5B58 5F53 524D 5185 5BB9 FF0C 5E76 5907 4EFD 5230 201C 516C 5F0F 5907 4EFD 201D 6587 4EF6 5939 3002") & vbCrLf & Zh("201C 5426 201D FF1A 4E0D 5907 4EFD FF0C 76F4 63A5 5904 7406 FF1B 201C 53D6 6D88 201D FF1A 9000 51FA 3002"), vbYesNoCancel + vbQuestion + vbDefaultButton1, "PWD MathType 1.3.3 CHECK")
    If choice = vbCancel Then Exit Sub
    If choice = vbYes Then
        If Not SaveCurrentForBackup(doc) Then Exit Sub
        backupPath = CreateActualBackup(doc)
        Application.StatusBar = Zh("5907 4EFD 5DF2 5B8C 6210 FF1A") & backupPath
    End If
    ' From here, all changes belong to the current document; do not SaveAs a result copy.
    processingStarted = True
    phaseStart = Timer
    DropEmptyProtocolMarks doc
    RegisterNewInputs doc, fresh
    For Each b In doc.Bookmarks
        If Left$(b.Name, 7) = "PWD_EQ_" And b.Range.StoryType = wdMainTextStory Then names.Add b.Name
    Next b
    nativeBefore = MathTypeCount(doc)

    originalText = ProseOutsideEquations(doc)
    otherCount = OtherInlineCount(doc)
    tableCount = doc.Tables.Count
    floatCount = doc.Shapes.Count
    processingStarted = True
    doc.Activate
    doc.Saved = False
    ProtectOtherFields doc, protectedFields
    processingSeconds = processingSeconds + ElapsedSeconds(phaseStart)
    Application.Run MT_UI & "MTCommand_FormatEqnNum"
    phaseStart = Timer
    ConvertDocumentCoreV8
    RestoreOtherFields protectedFields
    converted = MathTypeCount(doc) - nativeBefore
    ValidateSuiteBookmarks doc, names
    afterText = ProseOutsideEquations(doc)
    If afterText <> originalText Then Err.Raise vbObjectError + 8110, , Zh("516C 5F0F 4EE5 5916 7684 6587 5B57 53D1 751F 53D8 5316 FF0C 8BF7 68C0 67E5 540E 518D 4FDD 5B58 3002") & vbCrLf & ProseDifference(originalText, afterText)
    If OtherInlineCount(doc) <> otherCount Or doc.Tables.Count <> tableCount Or doc.Shapes.Count <> floatCount Then _
        Err.Raise vbObjectError + 8111, , Zh("56FE 7247 7B49 975E 516C 5F0F 5BF9 8C61 6216 8868 683C 6570 91CF 53D1 751F 53D8 5316 FF0C 8BF7 68C0 67E5 540E 518D 4FDD 5B58 3002")
    warnings = InspectFormulaObjects(doc)
    ShowEquationResults doc
    doc.Saved = False
    completion = Zh("5904 7406 5B8C 6210 3002 73B0 5728 4FDD 5B58 5230 5F53 524D 6587 4EF6 5417 FF1F")
    If warnings <> "" Then completion = Zh("516C 5F0F 5904 7406 5B8C 6210 FF0C 4EE5 4E0B 9879 76EE 9700 8981 770B 4E00 4E0B FF1A") & vbCrLf & warnings & vbCrLf & Zh("73B0 5728 4FDD 5B58 5230 5F53 524D 6587 4EF6 5417 FF1F")
    processingSeconds = processingSeconds + ElapsedSeconds(phaseStart)
    Debug.Print "PWD Stable 1.3.3 seconds=" & Format$(processingSeconds, "0.0") & "; converted=" & converted & "; existing=" & nativeBefore
    completion = Zh("672C 6B21 6210 529F 8F6C 6362 FF1A") & converted & Zh("0020 6761 FF1B 5168 6587 0020 004D 0061 0074 0068 0054 0079 0070 0065 0020 516C 5F0F FF1A") & (nativeBefore + converted) & Zh("0020 6761") & vbCrLf & _
        Zh("5904 7406 8017 65F6 FF1A") & Format$(processingSeconds, "0.0") & Zh("0020 79D2") & vbCrLf & completion
    answer = MsgBox(completion, vbYesNo + vbQuestion + vbDefaultButton1, "PWD MathType 1.3.3 CHECK")
    If answer = vbYes Then
        If Not SaveCurrentResult(doc) Then doc.Saved = False
    Else
        doc.Saved = False
    End If
    Application.StatusBar = ""
    Exit Sub
Failed:
    message = Err.Description
    On Error Resume Next
    Err.Clear
    RestoreOtherFields protectedFields
    If Err.Number <> 0 Then message = message & vbCrLf & "Field lock restore: " & Err.Description
    On Error GoTo 0
    Application.StatusBar = ""
    If processingStarted Then
        On Error Resume Next
        Err.Clear
        ShowEquationResults doc
        If Err.Number <> 0 Then message = message & vbCrLf & "Display: " & Err.Description
        Err.Clear
        doc.Saved = False
        If Err.Number <> 0 Then message = message & vbCrLf & "Unsaved state: " & Err.Description
        On Error GoTo 0
        MsgBox Zh("5904 7406 5DF2 505C 6B62 FF0C 5F53 524D 66F4 6539 5C1A 672A 4FDD 5B58 3002 8BF7 68C0 67E5 540E 51B3 5B9A 662F 5426 4FDD 5B58 3002") & vbCrLf & message, vbExclamation, "PWD MathType 1.3.3 CHECK"
    Else
        MsgBox Zh("5C1A 672A 5F00 59CB 8F6C 6362 3002") & vbCrLf & message, vbExclamation, "PWD MathType 1.3.3 CHECK"
    End If
End Sub

Private Function SaveCurrentForBackup(ByVal doc As Document) As Boolean
    doc.Activate
    If doc.Path = "" Then
        If Application.Dialogs(wdDialogFileSaveAs).Show <> -1 Then Exit Function
    Else
        doc.Save
    End If
    If doc.Path = "" Or Not doc.Saved Then Err.Raise vbObjectError + 8103, , Zh("5F53 524D 5185 5BB9 6CA1 6709 4FDD 5B58 6210 529F FF0C 672A 5F00 59CB 8F6C 6362 3002")
    SaveCurrentForBackup = True
End Function

Private Function CreateActualBackup(ByVal doc As Document) As String
    Dim fs As Object, folder As String, stem As String, extension As String
    Dim candidate As String, base As String, index As Long
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FileExists(doc.FullName) Then Err.Raise vbObjectError + 8104, , Zh("65E0 6CD5 8BFB 53D6 5F53 524D 6587 4EF6 8FDB 884C 5907 4EFD 3002 8BF7 5148 4FDD 5B58 5230 672C 5730 6587 4EF6 5939 3002")
    folder = fs.BuildPath(doc.Path, Zh("516C 5F0F 5907 4EFD"))
    If Not fs.FolderExists(folder) Then fs.CreateFolder folder
    stem = fs.GetBaseName(doc.Name)
    extension = fs.GetExtensionName(doc.Name)
    base = stem & "_" & Format$(Now, "yyyymmdd_hhnnss")
    candidate = fs.BuildPath(folder, base & "." & extension)
    Do While fs.FileExists(candidate)
        index = index + 1
        candidate = fs.BuildPath(folder, base & "_" & CStr(index) & "." & extension)
    Loop
    fs.CopyFile doc.FullName, candidate, False
    If Not fs.FileExists(candidate) Then Err.Raise vbObjectError + 8105, , Zh("5907 4EFD 672A 80FD 786E 8BA4 6210 529F FF0C 672A 5F00 59CB 8F6C 6362 3002")
    If fs.GetFile(candidate).Size = 0 Or fs.GetFile(candidate).Size <> fs.GetFile(doc.FullName).Size Then _
        Err.Raise vbObjectError + 8106, , Zh("5907 4EFD 672A 80FD 786E 8BA4 6210 529F FF0C 672A 5F00 59CB 8F6C 6362 3002")
    CreateActualBackup = candidate
End Function

Private Function SaveCurrentResult(ByVal doc As Document) As Boolean
    doc.Activate
    If doc.Path = "" Then
        If Application.Dialogs(wdDialogFileSaveAs).Show <> -1 Then Exit Function
    Else
        doc.Save
    End If
    SaveCurrentResult = doc.Saved
End Function


Private Function InspectFormulaObjects(ByVal doc As Document) As String
    Dim eq As InlineShape, i As Long, width As Single, note As String
    Dim state As Long, right As Long, count As Long
    For Each eq In doc.InlineShapes
        If IsMathType(eq) Then
            i = i + 1: note = ""
            state = SnapGridState(eq.Range)
            right = eq.Range.ParagraphFormat.AutoAdjustRightIndent
            If state <> 0 Or right <> False Then note = note & Zh("7F51 683C 8BBE 7F6E 9700 590D 6838 FF1B")
            If Not eq.Range.Information(wdWithInTable) Then
                If eq.Range.Sections(1).PageSetup.TextColumns.Count = 1 Then
                    With eq.Range.Sections(1).PageSetup
                        width = .PageWidth - .LeftMargin - .RightMargin - .Gutter
                    End With
                    If eq.Width > width Then note = note & Zh("5BBD 5EA6 8D85 8FC7 5355 680F 6B63 6587 FF1B")
                End If
            End If
            If note <> "" Then
                count = count + 1
                If count <= 6 Then InspectFormulaObjects = InspectFormulaObjects & "#" & i & " " & note & vbCrLf
            End If
        End If
    Next eq
    If count > 6 Then InspectFormulaObjects = InspectFormulaObjects & Zh("53E6 5916 8FD8 6709 63D0 793A 9879 FF1A") & (count - 6) & vbCrLf
End Function

Private Sub PWD_ConvertU8()
    Dim protectedFields As New Collection, errorText As String
    On Error GoTo Failed
    If Documents.Count = 0 Then Err.Raise vbObjectError + 7301, , "Open a Word document first."
    If ActiveDocument.ReadOnly Or ActiveDocument.ProtectionType <> wdNoProtection Or ActiveDocument.TrackRevisions Then _
        Err.Raise vbObjectError + 8223, , "Use an editable document with Track Changes off."
    ProtectOtherFields ActiveDocument, protectedFields
    PWD_ConvertStrictU8
    RestoreOtherFields protectedFields
    Exit Sub
Failed:
    errorText = Err.Description
    On Error Resume Next
    Err.Clear
    RestoreOtherFields protectedFields
    If Err.Number <> 0 Then errorText = errorText & vbCrLf & "Field lock restore: " & Err.Description
    On Error GoTo 0
    MsgBox errorText, vbExclamation, "PWD MathType 1.3.3 CHECK"
End Sub

Private Sub PWD_ConvertStrictU8(Optional ByVal restoreBookmark As String = "")
    Dim doc As Document, target As Range, p As Range, eq As InlineShape
    Dim kind As Long, tex As String, why As String, beforeObjects As Long, ordinal As Long
    Dim existingNumbers As Long, recordOpen As Boolean, beforeText As String
    Dim oldScreen As Boolean, oldTracking As Boolean, propertyAdded As Boolean
    Dim errorText As String, undoOK As Boolean, reverse As Boolean, stepName As String
    Dim paragraphsBefore As Long, inputFormat As ParagraphFormat
    Set lastConverted = Nothing
    stepName = "Check input"
    On Error GoTo PreflightFailed
    If busy Then Err.Raise vbObjectError + 7300, , "An equation conversion is already running."
    If Documents.Count = 0 Then Err.Raise vbObjectError + 7301, , "Open a Word document first."
    Set doc = ActiveDocument
    If doc.ReadOnly Or doc.ProtectionType <> wdNoProtection Then Err.Raise vbObjectError + 7302, , "Use an editable document."
    If doc.TrackRevisions Then Err.Raise vbObjectError + 7303, , "Turn off Track Changes before converting equations."
    If Selection.StoryType <> wdMainTextStory Then Err.Raise vbObjectError + 7304, , "This command handles equations in the main document."

    Set eq = SelectedEquation()
    reverse = Not eq Is Nothing
    If reverse Then
        Set target = eq.Range.Duplicate
    Else
        stepName = "Locate TeX input"
        Set target = InputRange()
        If target Is Nothing Then Err.Raise vbObjectError + 7305, , "Place the cursor inside the formula code, or select that one formula."
        If Not PWD_ParseU8(target.Text, kind, tex, why) Then Err.Raise vbObjectError + 7306, , why
        If target.Fields.Count > 0 Or target.InlineShapes.Count > 0 Then Err.Raise vbObjectError + 7307, , "Select formula code without other objects or fields."
        stepName = "Check display and existing number boundaries"
        If kind <> 1 Then ValidateDisplay target
    End If
    Set p = FullParagraph(target)
    Set inputFormat = p.ParagraphFormat.Duplicate
    existingNumbers = PWD_NumberCountU8(p)
    EnsureGridCalibration
    ShowLocalEquationResults doc, target
    beforeText = StableDocumentText(doc)
    beforeObjects = doc.InlineShapes.Count
    paragraphsBefore = doc.Paragraphs.Count
    ordinal = doc.Range(0, target.Start).InlineShapes.Count + 1
    oldScreen = Application.ScreenUpdating
    oldTracking = doc.TrackRevisions
    busy = True
    On Error GoTo Failed
    stepName = "Begin conversion"
    Application.UndoRecord.StartCustomRecord "PWD MathType one-key"
    recordOpen = True
    If reverse Then
        stepName = "MathType object to TeX"
        target.Select
        Application.Run MT_UI & "MTCommand_TeXToggle"
        ShowLocalEquationResults doc, target
        If doc.InlineShapes.Count <> beforeObjects - 1 Then Err.Raise vbObjectError + 7308, , "MathType did not return TeX."
    Else
        If kind = 3 And existingNumbers = 0 Then propertyAdded = EnsureNumberFormat(doc)
        stepName = "TeX to MathType object"
        target.Text = tex
        target.Select
        Application.Run MT_UI & "MTCommand_TeXToggle"
        ShowLocalEquationResults doc, target
        If doc.InlineShapes.Count <> beforeObjects + 1 Then Err.Raise vbObjectError + 7309, , "MathType did not create an equation. Check its TeX input message."
        stepName = "Locate converted object"
        Set eq = doc.InlineShapes(ordinal)
        If Not IsMathType(eq) Then Err.Raise vbObjectError + 7310, , "The new object is not a MathType equation."
        If kind = 3 And existingNumbers = 0 Then
            stepName = "Insert native number"
            eq.Range.Select
            Selection.Collapse wdCollapseEnd
            Selection.TypeText vbTab
            Application.Run MT_UI & "MTCommand_InsertEqnNum"
            If PWD_NumberCountU8(eq.Range.Paragraphs(1).Range) <> 1 Then _
                Err.Raise vbObjectError + 7311, , "MathType did not insert its native equation number."
        End If
        Set lastConverted = eq
        stepName = "Check preserved numbering"
        If existingNumbers > 0 Then
            If PWD_NumberCountU8(eq.Range.Paragraphs(1).Range) <> existingNumbers Then _
                Err.Raise vbObjectError + 7312, , "The existing number was not preserved."
        End If
        If kind <> 1 Then
            stepName = "Normalize display paragraph layout"
            eq.Range.ParagraphFormat = inputFormat
            ApplyDisplayLayout eq
            RemoveAddedEmptyParagraph eq, paragraphsBefore
        End If
        If restoreBookmark <> "" Then doc.Bookmarks.Add Name:=restoreBookmark, Range:=eq.Range
        stepName = "Normalize formula paragraph grid"
        ReleaseFormulaGrid eq.Range
        If kind <> 1 Then
            Set p = FullParagraph(eq.Range)
            doc.Range(p.End - 1, p.End - 1).Select
        Else
            eq.Range.Select
            Selection.Collapse wdCollapseEnd
        End If
    End If
    Application.UndoRecord.EndCustomRecord
    recordOpen = False
    ShowLocalEquationResults doc, target
    Application.ScreenUpdating = oldScreen
    doc.TrackRevisions = oldTracking
    busy = False
    Exit Sub
PreflightFailed:
    Err.Raise vbObjectError + 7314, "PWD MathType 1.3.3 CHECK", stepName & " / " & Err.Number & ": " & Err.Description
Failed:
    errorText = stepName & " / " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If recordOpen Then Application.UndoRecord.EndCustomRecord
    undoOK = True
    If recordOpen And (StableDocumentText(doc) <> beforeText Or doc.InlineShapes.Count <> beforeObjects) Then
        undoOK = False
        undoOK = doc.Undo(1)
        undoOK = (StableDocumentText(doc) = beforeText And doc.InlineShapes.Count = beforeObjects)
        If restoreBookmark <> "" Then
            If Not doc.Bookmarks.Exists(restoreBookmark) Then undoOK = False
        End If
    End If
    Err.Clear
    ShowLocalEquationResults doc, target
    If Err.Number <> 0 Then errorText = errorText & " / Could not restore equation display: " & Err.Description
    If propertyAdded Then doc.CustomDocumentProperties("MTEquationNumber2").Delete
    Application.ScreenUpdating = oldScreen
    doc.TrackRevisions = oldTracking
    busy = False
    On Error GoTo 0
    If Not undoOK Then errorText = errorText & vbCrLf & "Rollback could not be verified. Keep this test copy for diagnosis and restart from the original input."
    Err.Raise vbObjectError + 7313, "PWD MathType 1.3.3 CHECK", errorText
End Sub

Private Function EnsureNumberFormat(ByVal doc As Document) As Boolean
    Dim prop As Object
    For Each prop In doc.CustomDocumentProperties
        If prop.Name = "MTEquationNumber2" Then Exit Function
        If prop.Name = "MTEquationNumber" Then Exit Function
    Next prop
    doc.CustomDocumentProperties.Add Name:="MTEquationNumber2", LinkToContent:=False, _
        Type:=msoPropertyTypeString, Value:="(#E1)"
    EnsureNumberFormat = True
End Function

Private Function SelectedEquation() As InlineShape
    Dim r As Range, item As InlineShape
    Set r = Selection.Range
    If r.InlineShapes.Count = 1 Then
        Set item = r.InlineShapes(1)
        If r.Start = item.Range.Start And r.End = item.Range.End And IsMathType(item) Then
            Set SelectedEquation = item
        End If
    ElseIf r.Start = r.End Then
        For Each item In r.Paragraphs(1).Range.InlineShapes
            If (r.Start = item.Range.Start Or r.Start = item.Range.End) And IsMathType(item) Then
                Set SelectedEquation = item: Exit Function
            End If
        Next item
    End If
End Function

Private Function IsMathType(ByVal item As InlineShape) As Boolean
    On Error GoTo NotEquation
    If item.Type <> wdInlineShapeEmbeddedOLEObject Then Exit Function
    IsMathType = (LCase$(item.OLEFormat.ProgID) = "equation.dsmt4")
NotEquation:
End Function

Private Function InputRange() As Range
    Dim r As Range, scope As Range, opening As Range, closing As Range, candidate As Range
    Dim opens As Variant, closes As Variant, k As Long, cursor As Long, pos As Long
    Set r = Selection.Range.Duplicate
    If r.Start <> r.End Then
        Do While r.End > r.Start
            If Right$(r.Text, 1) <> vbCr Then Exit Do
            r.End = r.End - 1
        Loop
        TrimRange r
        Set InputRange = r
        Exit Function
    End If
    cursor = r.Start
    Set scope = r.Paragraphs(1).Range.Duplicate
    scope.End = scope.End - 1
    opens = Array("\begin{equation}", "\begin{equation*}", "\[", "$")
    closes = Array("\end{equation}", "\end{equation*}", "\]", "$")
    For k = 0 To 3
        pos = scope.Start
        Do While pos < scope.End
            Set opening = scope.Duplicate
            opening.Start = pos
            If Not FindLiteral(opening, CStr(opens(k))) Then Exit Do
            pos = opening.End
            If Not RangeEscaped(opening) Then
                Set closing = scope.Duplicate
                closing.Start = opening.End
                Do While closing.Start < scope.End
                    If Not FindLiteral(closing, CStr(closes(k))) Then Exit Do
                    If Not RangeEscaped(closing) Then Exit Do
                    closing.SetRange closing.End, scope.End
                Loop
                If closing.Text = CStr(closes(k)) Then
                    If cursor >= opening.Start And cursor <= closing.End Then
                        Set candidate = scope.Duplicate
                        candidate.SetRange opening.Start, closing.End
                        Set InputRange = candidate: Exit Function
                    End If
                    pos = closing.End
                End If
            End If
        Loop
    Next k
End Function

Private Function FindLiteral(ByVal r As Range, ByVal value As String) As Boolean
    Dim scopeRange As Range, characterRange As Range, probe As Range
    Dim lowerBound As Long, upperBound As Long, firstCharacter As String
    If Len(value) = 0 Then Exit Function
    Set scopeRange = r.Duplicate
    lowerBound = scopeRange.Start: upperBound = scopeRange.End
    If lowerBound >= upperBound Then Exit Function
    ' Cheap absence check; never turn string offsets into Word coordinates.
    If InStr(1, scopeRange.Text, value, vbBinaryCompare) = 0 Then Exit Function
    firstCharacter = Left$(value, 1)
    For Each characterRange In scopeRange.Characters
        If characterRange.Start >= upperBound Then Exit For
        If characterRange.Start >= lowerBound And characterRange.End <= upperBound Then
            If characterRange.Text = firstCharacter Then
                Set probe = scopeRange.Duplicate
                probe.SetRange characterRange.Start, characterRange.End
                ' All supported outer delimiters are ASCII; expand by native Word characters.
                If Len(value) > 1 Then probe.MoveEnd Unit:=wdCharacter, Count:=Len(value) - 1
                If probe.Start >= lowerBound And probe.End <= upperBound Then
                    If probe.Text = value Then
                        r.SetRange probe.Start, probe.End
                        FindLiteral = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next characterRange
End Function

Private Function RangeEscaped(ByVal r As Range) As Boolean
    Dim prior As Range, count As Long, pos As Long
    pos = r.Start
    Do While pos > r.Paragraphs(1).Range.Start
        Set prior = r.Document.Range(pos - 1, pos)
        If prior.Text <> "\" Then Exit Do
        count = count + 1: pos = pos - 1
    Loop
    RangeEscaped = (count Mod 2 = 1)
End Function

Private Sub TrimRange(ByVal r As Range)
    Do While r.End > r.Start
        If PWD_TrimU8(Left$(r.Text, 1)) <> "" Then Exit Do
        r.Start = r.Start + 1
    Loop
    Do While r.End > r.Start
        If PWD_TrimU8(Right$(r.Text, 1)) <> "" Then Exit Do
        r.End = r.End - 1
    Loop
End Sub

Private Sub ValidateDisplay(ByVal r As Range)
    Dim p As Range, before As Range, after As Range, f As Field, number As Field
    Dim numberEnd As Long, paragraphEnd As Long
    If r.Paragraphs.Count <> 1 Then Err.Raise vbObjectError + 7320, , "Keep each display formula in one Word paragraph."
    If r.Information(wdWithInTable) Then Err.Raise vbObjectError + 7321, , "Place numbered/display input outside tables in this version."
    Set p = FullParagraph(r)
    Set before = r.Document.Range(p.Start, r.Start)
    If PWD_TrimU8(before.Text) <> "" Then Err.Raise vbObjectError + 7322, , "A display formula needs its own paragraph."
    If PWD_NumberCountU8(p) > 1 Then Err.Raise vbObjectError + 7323, , "This paragraph contains more than one equation number."
    For Each f In p.Fields
        If PWD_IsNativeNumberU8(f) Then Set number = f
    Next f
    If number Is Nothing Then
        Set after = r.Document.Range(r.End, p.End - 1)
    Else
        If number.Code.Start - 1 < r.End Then Err.Raise vbObjectError + 7324, , "The existing number must follow the formula."
        numberEnd = NativeFieldEnd(number)
        paragraphEnd = p.End - 1
        If numberEnd < r.End Or numberEnd > paragraphEnd Then
            Err.Raise vbObjectError + 7327, , "Number boundary: end=" & numberEnd & _
                "; paragraph end=" & paragraphEnd & "; TeX end=" & r.End
        End If
        Set after = r.Document.Range(numberEnd, paragraphEnd)
        If PWD_TrimU8(after.Text) <> "" Then Err.Raise vbObjectError + 7325, , "Unexpected text after the equation number."
        Set after = r.Document.Range(r.End, number.Code.Start - 1)
    End If
    If PWD_TrimU8(after.Text) <> "" Then Err.Raise vbObjectError + 7326, , "Select only the formula; keep prose in another paragraph."
End Sub

Private Function FullParagraph(ByVal r As Range) As Range
    Dim p As Range
    Set p = r.Paragraphs(1).Range.Duplicate
    Set FullParagraph = p
End Function

Private Function NativeFieldEnd(ByVal f As Field) As Long
    Dim pos As Long, doc As Document
    Set doc = f.Code.Document
    If f.Type = wdFieldMacroButton Or f.Type = wdFieldGoToButton Then
        pos = f.Code.End + 1
    Else
        pos = f.Result.End + 1
    End If
    If pos < doc.Content.Start Or pos > doc.Content.End Then
        Err.Raise vbObjectError + 7328, , "Invalid field endpoint: " & pos
    End If
    NativeFieldEnd = pos
End Function

Private Function PWD_IsNativeNumberU8(ByVal f As Field) As Boolean
    If f.Type = wdFieldMacroButton Then
        PWD_IsNativeNumberU8 = (InStr(1, f.Code.Text, "MTPlaceRef", vbTextCompare) > 0)
    End If
End Function

Private Function PWD_NumberCountU8(ByVal r As Range) As Long
    Dim f As Field
    For Each f In r.Fields
        If PWD_IsNativeNumberU8(f) Then PWD_NumberCountU8 = PWD_NumberCountU8 + 1
    Next f
End Function

Private Sub PWD_UpdateNumbersStrictU8()
    Application.Run MT_UI & "MTCommand_UpdateEqns"
End Sub

Private Function PWD_ParseU8(ByVal value As String, ByRef kind As Long, ByRef tex As String, ByRef why As String) As Boolean
    Dim body As String, i As Long, ch As String, depth As Long
    kind = 0: tex = "": why = ""
    value = PWD_TrimU8(value)
    If InStr(value, vbCr) > 0 Or InStr(value, vbLf) > 0 Then
        why = "Use one Word paragraph per formula; Shift+Enter is allowed inside it.": Exit Function
    End If
    If Len(value) >= 32 And Left$(value, 17) = "\begin{equation*}" Then
        If Right$(value, 15) <> "\end{equation*}" Then GoTo Invalid
        kind = 2: body = Mid$(value, 18, Len(value) - 32)
    ElseIf Len(value) >= 30 And Left$(value, 16) = "\begin{equation}" Then
        If Right$(value, 14) <> "\end{equation}" Then GoTo Invalid
        kind = 3: body = Mid$(value, 17, Len(value) - 30)
    ElseIf Len(value) >= 4 And Left$(value, 2) = "\[" And Right$(value, 2) = "\]" Then
        kind = 2: body = Mid$(value, 3, Len(value) - 4)
    ElseIf Len(value) >= 2 And Left$(value, 1) = "$" And Right$(value, 1) = "$" Then
        If PWD_EscapedU8(value, Len(value)) Then GoTo Invalid
        kind = 1: body = Mid$(value, 2, Len(value) - 2)
    Else
        GoTo Invalid
    End If
    body = PWD_TrimU8(Replace(body, Chr(11), " "))
    If body = "" Then why = "The formula is empty.": Exit Function
    If InStr(body, "\begin{equation") > 0 Or InStr(body, "\end{equation") > 0 Or _
       InStr(body, "\[") > 0 Or InStr(body, "\]") > 0 Then
        why = "Select one formula, with one outer delimiter pair.": Exit Function
    End If
    If HasCommand(body, "tag") Or HasCommand(body, "label") Or HasCommand(body, "nonumber") Or HasCommand(body, "notag") Then
        why = "Leave numbering to MathType. Remove tag/label/nonumber/notag from this input.": Exit Function
    End If
    For i = 1 To Len(body)
        ch = Mid$(body, i, 1)
        If (AscW(ch) And &HFFFF&) < 32 And ch <> vbTab Then why = "Unsupported control character in formula.": Exit Function
        If Not PWD_EscapedU8(body, i) Then
            If ch = "$" Then why = "Select one formula; $$ delimiters are not supported.": Exit Function
            If ch = "%" Then why = "Remove TeX comments; use \% for a literal percent sign.": Exit Function
            If ch = "{" Then depth = depth + 1
            If ch = "}" Then depth = depth - 1
            If depth < 0 Then why = "Unbalanced braces.": Exit Function
        End If
    Next i
    If depth <> 0 Then why = "Unbalanced braces.": Exit Function
    If kind = 1 Then tex = "$" & body & "$" Else tex = "\[" & body & "\]"
    PWD_ParseU8 = True
    Exit Function
Invalid:
    why = "Use $...$, \[...\], or \begin{equation}...\end{equation}."
End Function

Private Function PWD_TrimU8(ByVal value As String) As String
    Do While Len(value) > 0
        If Not IsWhite(Left$(value, 1)) Then Exit Do
        value = Mid$(value, 2)
    Loop
    Do While Len(value) > 0
        If Not IsWhite(Right$(value, 1)) Then Exit Do
        value = Left$(value, Len(value) - 1)
    Loop
    PWD_TrimU8 = value
End Function

Private Function IsWhite(ByVal ch As String) As Boolean
    IsWhite = (ch = " " Or ch = vbTab Or ch = Chr(11))
End Function

Private Function PWD_EscapedU8(ByVal value As String, ByVal position As Long) As Boolean
    Dim count As Long
    position = position - 1
    Do While position > 0
        If Mid$(value, position, 1) <> "\" Then Exit Do
        count = count + 1: position = position - 1
    Loop
    PWD_EscapedU8 = (count Mod 2 = 1)
End Function

Private Function HasCommand(ByVal value As String, ByVal name As String) As Boolean
    Dim pos As Long, after As String
    pos = InStr(1, value, "\" & name, vbBinaryCompare)
    Do While pos > 0
        after = Mid$(value, pos + Len(name) + 1, 1)
        If Not PWD_EscapedU8(value, pos) Then
            If after = "" Or Not (after Like "[A-Za-z]") Then HasCommand = True: Exit Function
        End If
        pos = InStr(pos + 1, value, "\" & name, vbBinaryCompare)
    Loop
End Function

Private Function NumberCount(ByVal doc As Document) As Long
    Dim f As Field
    For Each f In doc.Content.Fields
        If PWD_IsNativeNumberU8(f) Then NumberCount = NumberCount + 1
    Next f
End Function

Private Sub ApplyDisplayLayout(ByVal eq As InlineShape)
    Dim doc As Document, p As Range, prefix As Range, gap As Range, f As Field, number As Field
    Dim numbered As Boolean, beforeSpace As Single, afterSpace As Single
    Dim lineRule As Long, lineSpace As Single, width As Single, styleName As String, st As Style, candidate As Style
    Set doc = eq.Range.Document
    Set p = FullParagraph(eq.Range)
    beforeSpace = p.ParagraphFormat.SpaceBefore
    afterSpace = p.ParagraphFormat.SpaceAfter
    lineRule = p.ParagraphFormat.LineSpacingRule
    lineSpace = p.ParagraphFormat.LineSpacing
    For Each f In p.Fields
        If PWD_IsNativeNumberU8(f) Then Set number = f
    Next f
    numbered = Not number Is Nothing
    Set prefix = doc.Range(p.Start, eq.Range.Start)
    If PWD_TrimU8(prefix.Text) <> "" Then Err.Raise vbObjectError + 7601, , "Unexpected text before display equation."
    If numbered Then
        Set gap = doc.Range(eq.Range.End, number.Code.Start - 1)
        If PWD_TrimU8(gap.Text) <> "" Then Err.Raise vbObjectError + 7602, , "Unexpected text between equation and number."
        gap.Text = vbTab
        prefix.Text = vbTab
        styleName = "PWDDisplayNumbered"
    Else
        prefix.Text = ""
        styleName = "PWDDisplayCentered"
    End If
    Set p = FullParagraph(eq.Range)
    For Each candidate In doc.Styles
        If candidate.NameLocal = styleName Then Set st = candidate: Exit For
    Next candidate
    If st Is Nothing Then Set st = doc.Styles.Add(Name:=styleName, Type:=wdStyleTypeParagraph)
    st.BaseStyle = wdStyleNormal
    st.NextParagraphStyle = wdStyleNormal
    st.ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText
    p.Style = st
    With p.ParagraphFormat
        .OutlineLevel = wdOutlineLevelBodyText
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LeftIndent = 0
        .RightIndent = 0
        .FirstLineIndent = 0
        .SpaceBefore = beforeSpace
        .SpaceAfter = afterSpace
        .LineSpacingRule = lineRule
        If lineRule <> wdLineSpaceSingle And lineRule <> wdLineSpace1pt5 And lineRule <> wdLineSpaceDouble Then .LineSpacing = lineSpace
        .TabStops.ClearAll
        If numbered Then
            .Alignment = wdAlignParagraphLeft
            If p.Sections(1).PageSetup.TextColumns.Count = 1 Then
                With p.Sections(1).PageSetup
                    width = .PageWidth - .LeftMargin - .RightMargin - .Gutter
                End With
            Else
                eq.Range.Select
                width = CSng(Application.Run("MathTypeCommands.MTInsertEquation.GetDocWidth", doc))
            End If
            .TabStops.Add Position:=width / 2, Alignment:=wdAlignTabCenter
            .TabStops.Add Position:=width, Alignment:=wdAlignTabRight
        Else
            .Alignment = wdAlignParagraphCenter
        End If
    End With
End Sub

Private Sub RemoveAddedEmptyParagraph(ByVal eq As InlineShape, ByVal countBefore As Long)
    Dim doc As Document, p As Range, nextPara As Range
    Set doc = eq.Range.Document
    If doc.Paragraphs.Count <= countBefore Then Exit Sub
    Set p = FullParagraph(eq.Range)
    If p.End >= doc.Content.End Then Exit Sub
    Set nextPara = doc.Range(p.End, p.End)
    nextPara.Expand wdParagraph
    If nextPara.Text <> vbCr Then Exit Sub
    If nextPara.Fields.Count > 0 Or nextPara.InlineShapes.Count > 0 Or nextPara.Bookmarks.Count > 0 Then Exit Sub
    nextPara.Delete
End Sub

Private Sub HideCodes(ByVal r As Range)
    Dim f As Field
    For Each f In r.Fields
        If IsMathTypeWrapper(f) Then
            f.ShowCodes = False
            If f.Code.Fields.Count > 0 Then HideCodes f.Code
        ElseIf IsMathTypeSequence(f) Then
            f.ShowCodes = False
        End If
    Next f
End Sub

Private Sub ShowLocalEquationResults(ByVal doc As Document, ByVal scopeRange As Range)
    Dim win As Window
    For Each win In doc.Windows
        win.View.ShowFieldCodes = False
    Next win
    HideCodes FullParagraph(scopeRange)
End Sub

Private Function Zh(ByVal hexText As String) As String
    Dim codes As Variant, value As Variant
    codes = Split(hexText, " ")
    For Each value In codes
        Zh = Zh & ChrW$(CLng("&H" & CStr(value)))
    Next value
End Function

Private Sub ConvertDocumentCoreV8()
    Dim doc As Document, b As Bookmark, r As Range, names() As String, positions() As Long
    Dim n As Long, i As Long, j As Long, temp As String, pos As Long
    Dim todo As Long, skipped As Long, done As Long, kind As Long
    Dim tex As String, why As String, currentName As String, errorText As String
    Dim oldCancel As Long, numberBefore As Long, expectedNumbers As Long
    Dim beforeObjects As Long, previousEnd As Long
    On Error GoTo Failed
    oldCancel = Application.EnableCancelKey
    If Documents.Count = 0 Then Err.Raise vbObjectError + 7800, , Zh("8BF7 5148 6253 5F00 9700 8981 8F6C 6362 7684 0020 0057 006F 0072 0064 0020 6587 6863 3002")
    Set doc = ActiveDocument
    If doc.ReadOnly Or doc.ProtectionType <> wdNoProtection Or doc.TrackRevisions Then _
        Err.Raise vbObjectError + 7801, , Zh("8BF7 4F7F 7528 53EF 7F16 8F91 6587 6863 FF0C 5E76 5148 5173 95ED 4FEE 8BA2 6A21 5F0F 3002")
    For Each b In doc.Bookmarks
        If Left$(b.Name, 7) = "PWD_EQ_" And b.Range.StoryType = wdMainTextStory Then
            n = n + 1
            ReDim Preserve names(1 To n)
            ReDim Preserve positions(1 To n)
            names(n) = b.Name
            positions(n) = b.Range.Start
        End If
    Next b
    If n = 0 Then Err.Raise vbObjectError + 7802, , Zh("6CA1 6709 627E 5230 5F85 5904 7406 516C 5F0F 6807 8BB0 3002 8BF7 4F7F 7528 6309 901A 7528 89C4 5219 5BFC 51FA 7684 0020 0057 006F 0072 0064 FF1B 5355 4E2A 624B 5199 516C 5F0F 53EF 8FD0 884C 0020 0050 0057 0044 005F 0043 006F 006E 0076 0065 0072 0074 004F 006E 0065 3002")
    For i = 2 To n
        temp = names(i): pos = positions(i): j = i - 1
        Do While j >= 1
            If positions(j) <= pos Then Exit Do
            names(j + 1) = names(j): positions(j + 1) = positions(j): j = j - 1
        Loop
        names(j + 1) = temp: positions(j + 1) = pos
    Next i
    ' Validate all inputs before making changes.
    numberBefore = NumberCount(doc)
    beforeObjects = MathTypeCount(doc)
    For i = 1 To n
        currentName = names(i)
        Set r = doc.Bookmarks(currentName).Range
        If i > 1 And r.Start < previousEnd Then Err.Raise vbObjectError + 7809, , "Overlapping formula bookmarks."
        previousEnd = r.End
        If CompletedInput(r) Then
            skipped = skipped + 1
        Else
            If r.InlineShapes.Count > 0 Or r.Fields.Count > 0 Then _
                Err.Raise vbObjectError + 7803, , Zh("8FD9 4E2A 516C 5F0F 6807 8BB0 91CC 6DF7 5165 4E86 5176 4ED6 5BF9 8C61 6216 57DF FF0C 8BF7 68C0 67E5 3002")
            If Not PWD_ParseU8(r.Text, kind, tex, why) Then _
                Err.Raise vbObjectError + 7804, , why
            If kind <> 1 Then ValidateDisplay r
            If kind = 3 Then
                If PWD_NumberCountU8(FullParagraph(r)) = 0 Then expectedNumbers = expectedNumbers + 1
            End If
            todo = todo + 1
        End If
    Next i
    EnsureGridCalibration
    ShowEquationResults doc
    SyncChaptersStrict
    ShowEquationResults doc
    Application.EnableCancelKey = wdCancelInterrupt
    For i = 1 To n
        currentName = names(i)
        Application.StatusBar = "MathType " & i & " / " & n & "   " & currentName
        If Not doc.Bookmarks.Exists(currentName) Then _
            Err.Raise vbObjectError + 7805, , Zh("516C 5F0F 5B9A 4F4D 6807 8BB0 4E22 5931 FF0C 8BF7 4FDD 7559 5F53 524D 6587 4EF6 5E76 628A 63D0 793A 53D1 7ED9 6211 3002")
        Set r = doc.Bookmarks(currentName).Range
        If Not CompletedInput(r) Then
            r.Select
            PWD_ConvertStrictU8 currentName
            If lastConverted Is Nothing Then Err.Raise vbObjectError + 7806, , Zh("6CA1 6709 627E 5230 521A 521A 8F6C 6362 7684 0020 004D 0061 0074 0068 0054 0079 0070 0065 0020 516C 5F0F 3002")
            done = done + 1
        Else
            ReleaseFormulaGrid r
        End If
    Next i
    If MathTypeCount(doc) <> beforeObjects + todo Then _
        Err.Raise vbObjectError + 7807, , Zh("65B0 589E 7684 0020 004D 0061 0074 0068 0054 0079 0070 0065 0020 516C 5F0F 6570 91CF 4E0E 5F85 8F6C 6362 6570 91CF 4E0D 4E00 81F4 3002")
    If NumberCount(doc) <> numberBefore + expectedNumbers Then _
        Err.Raise vbObjectError + 7808, , Zh("81EA 52A8 7F16 53F7 7684 6570 91CF 4E0E 5F85 7F16 53F7 516C 5F0F 6570 91CF 4E0D 4E00 81F4 3002")
    If NumberCount(doc) > 0 Then PWD_UpdateNumbersStrictU8
    HideCodes doc.Content
    doc.Repaginate
    Application.StatusBar = ""
    Application.EnableCancelKey = oldCancel
    Exit Sub
Failed:
    errorText = Err.Description
    Application.StatusBar = ""
    Application.EnableCancelKey = oldCancel
    On Error Resume Next
    Err.Clear
    ShowEquationResults doc
    If Err.Number <> 0 Then errorText = errorText & " / Display cleanup: " & Err.Description
    On Error GoTo 0
    Err.Raise vbObjectError + 7899, , currentName & ": " & errorText
End Sub

Public Sub PWD_One133_Check()
    PWD_ConvertU8
End Sub

Private Function CompletedInput(ByVal r As Range) As Boolean
    Dim eq As InlineShape
    If r.InlineShapes.Count <> 1 Then Exit Function
    Set eq = r.InlineShapes(1)
    If Not IsMathType(eq) Then Exit Function
    CompletedInput = (r.Start = eq.Range.Start And r.End = eq.Range.End)
End Function

Private Function MathTypeCount(ByVal doc As Document) As Long
    Dim eq As InlineShape
    For Each eq In doc.InlineShapes
        If IsMathType(eq) Then MathTypeCount = MathTypeCount + 1
    Next eq
End Function

Public Sub PWD_Chapters133_Check()
    Dim protectedFields As New Collection, errorText As String
    On Error GoTo Failed
    If Documents.Count = 0 Then Err.Raise vbObjectError + 7900, , "Open a document first."
    If ActiveDocument.ReadOnly Or ActiveDocument.ProtectionType <> wdNoProtection Or ActiveDocument.TrackRevisions Then _
        Err.Raise vbObjectError + 8223, , "Use an editable document with Track Changes off."
    ProtectOtherFields ActiveDocument, protectedFields
    SyncChaptersStrict
    RestoreOtherFields protectedFields
    Application.StatusBar = ""
    Exit Sub
Failed:
    errorText = Err.Description
    On Error Resume Next
    Err.Clear
    RestoreOtherFields protectedFields
    If Err.Number <> 0 Then errorText = errorText & vbCrLf & "Field lock restore: " & Err.Description
    On Error GoTo 0
    MsgBox errorText, vbExclamation, "PWD MathType 1.3.3 CHECK"
End Sub

Private Sub SyncChaptersStrict()
    Dim doc As Document, p As Paragraph, titles As New Collection, nums As New Collection
    Dim r As Range, f As Field, owned As Field, i As Long, chapter As Long
    Dim mode As String, byChapter As Boolean, sectionOnly As Boolean, prop As Object
    Dim undoOpen As Boolean, dirty As Boolean, errorText As String, seen As Object
    Dim st As Style, oldSelection As Range, previousChapter As Long
    If Documents.Count = 0 Then Err.Raise vbObjectError + 7900, , "Open a document first."
    Set doc = ActiveDocument
    If doc.ReadOnly Or doc.ProtectionType <> wdNoProtection Or doc.TrackRevisions Then _
        Err.Raise vbObjectError + 7901, , "Use an editable document with Track Changes off."
    Set oldSelection = Selection.Range.Duplicate
    For Each prop In doc.CustomDocumentProperties
        If prop.Name = "MTEquationNumber2" Then mode = CStr(prop.Value): Exit For
    Next prop
    If mode = "" Then
        Application.Run MT_UI & "MTCommand_FormatEqnNum"
        For Each prop In doc.CustomDocumentProperties
            If prop.Name = "MTEquationNumber2" Then mode = CStr(prop.Value): Exit For
        Next prop
        If mode = "" Then Err.Raise vbObjectError + 7902, , "Set the MathType equation number format first."
    End If
    byChapter = InStr(1, mode, "#C", vbTextCompare) > 0 Or InStr(1, mode, "#S", vbTextCompare) > 0
    sectionOnly = InStr(1, mode, "#C", vbTextCompare) = 0 And InStr(1, mode, "#S", vbTextCompare) > 0
    Set seen = CreateObject("Scripting.Dictionary")
    ' Read real level-one headings; ignore abstract, references and subsections.
    For Each p In doc.Paragraphs
        If p.OutlineLevel = wdOutlineLevel1 And Not p.Range.Information(wdWithInTable) Then
            Set r = p.Range.Duplicate
            r.TextRetrievalMode.IncludeHiddenText = False
            r.TextRetrievalMode.IncludeFieldCodes = False
            chapter = ChapterNumber(Trim$(Replace(r.Text, vbCr, "")))
            If chapter = 0 And p.Range.ListFormat.ListType <> wdListNoNumbering Then _
                chapter = ChapterNumber(p.Range.ListFormat.ListString & " ")
            If chapter > 0 Then
                If seen.Exists(CStr(chapter)) Or chapter <= previousChapter Then _
                    Err.Raise vbObjectError + 7903, , "Chapter numbers must be distinct and increasing: " & chapter
                seen.Add CStr(chapter), True
                titles.Add p.Range.Duplicate
                nums.Add chapter
                previousChapter = chapter
            End If
        End If
    Next p
    If byChapter And titles.Count = 0 Then Err.Raise vbObjectError + 7904, , _
        ZC("6CA1 6709 627E 5230 5E26 7AE0 53F7 7684 4E00 7EA7 6807 9898 3002")
    ' Manual breaks inside chapter content may override chapter counters; do not silently remove them.
    If byChapter And titles.Count > 0 Then
        Set r = titles(1)
        For Each f In doc.Fields
            If IsChapterBreak(f) And Not IsOwnedChapter(f) Then
                If f.Code.Start >= r.Start Then Err.Raise vbObjectError + 7905, , _
                    ZC("6B63 6587 5185 5DF2 6709 624B 52A8 8BBE 7F6E 7684 004D 0061 0074 0068 0054 0079 0070 0065 7AE0 8282 6807 8BB0 FF0C 8BF7 5148 68C0 67E5 5176 4F4D 7F6E 3002")
            End If
        Next f
    End If
    On Error GoTo Failed
    Application.UndoRecord.StartCustomRecord "PWD sync MathType chapters"
    undoOpen = True
    ' Changing format back to continuous must neutralize our old chapter resets.
    For Each f In doc.Fields
        If IsOwnedChapter(f) Then
            dirty = True
            SetSeq f, "MTEqn", "\c \h"
            SetSeq f, "MTChap", "\c \h"
            SetSeq f, "MTSec", "\c \h"
        End If
    Next f
    If byChapter Then
        For i = titles.Count To 1 Step -1
            Set r = titles(i)
            chapter = CLng(nums(i))
            Set owned = Nothing
            For Each f In r.Fields
                If IsOwnedChapter(f) Then
                    If Not owned Is Nothing Then Err.Raise vbObjectError + 7906, , "Duplicate chapter markers."
                    Set owned = f
                End If
            Next f
            dirty = True
            If owned Is Nothing Then
                r.Collapse wdCollapseStart
                Set owned = doc.Fields.Add(r, wdFieldMacroButton, _
                    "MTEditEquationSection3 PWDChapterSync ", False)
                AddSeq owned, "MTEqn", "\r 0 \h"
                AddSeq owned, "MTSec", "\r 1 \h"
                AddSeq owned, "MTChap", "\r " & CStr(chapter) & " \h"
            End If
            SetSeq owned, "MTEqn", "\r 0 \h"
            If sectionOnly Then
                SetSeq owned, "MTSec", "\r " & CStr(chapter) & " \h"
            Else
                SetSeq owned, "MTSec", "\r 1 \h"
            End If
            SetSeq owned, "MTChap", "\r " & CStr(chapter) & " \h"
            Set r = doc.Range(owned.Code.Start - 1, owned.Code.End + 1)
            r.Font.Hidden = True
            owned.ShowCodes = False
        Next i
    End If
    ' Update in document order, including hidden sequence fields nested in macrobuttons.
    UpdateNested doc.Content
    If NumberCount(doc) > 0 Then Application.Run MT_UI & "MTCommand_UpdateEqns"
    HideCodes doc.Content
    Application.UndoRecord.EndCustomRecord
    undoOpen = False
    oldSelection.Select
    Exit Sub
Failed:
    errorText = Err.Description
    On Error Resume Next
    If undoOpen Then Application.UndoRecord.EndCustomRecord
    If undoOpen And dirty Then
        If Not doc.Undo(1) Then errorText = errorText & " / Undo failed; inspect the document."
    End If
    oldSelection.Select
    On Error GoTo 0
    Err.Raise vbObjectError + 7907, , errorText
End Sub

Private Function IsChapterBreak(ByVal f As Field) As Boolean
    If f.Type = wdFieldMacroButton Then _
        IsChapterBreak = InStr(1, f.Code.Text, "MTEditEquationSection", vbTextCompare) > 0
End Function

Private Function IsOwnedChapter(ByVal f As Field) As Boolean
    IsOwnedChapter = IsChapterBreak(f) And InStr(1, f.Code.Text, "PWDChapterSync", vbBinaryCompare) > 0
End Function

Private Sub AddSeq(ByVal parent As Field, ByVal name As String, ByVal switches As String)
    Dim r As Range, f As Field
    Set r = parent.Code.Duplicate
    r.Collapse wdCollapseEnd
    Set f = r.Document.Fields.Add(r, wdFieldSequence, name & " " & switches, False)
End Sub

Private Sub SetSeq(ByVal parent As Field, ByVal name As String, ByVal switches As String)
    Dim f As Field, text As String
    For Each f In parent.Code.Fields
        text = Trim$(f.Code.Text)
        If f.Type = wdFieldSequence And InStr(1, text, name & " ", vbTextCompare) > 0 Then
            f.Code.Text = " SEQ " & name & " " & switches & " "
            Exit Sub
        End If
    Next f
    Err.Raise vbObjectError + 7908, , "Missing chapter counter: " & name
End Sub

Private Sub UpdateNested(ByVal r As Range)
    Dim f As Field
    For Each f In r.Fields
        If IsMathTypeWrapper(f) Then
            If f.Code.Fields.Count > 0 Then UpdateNested f.Code
            f.Update
        ElseIf IsMathTypeSequence(f) Then
            f.Update
        End If
    Next f
End Sub

Private Function FieldWord(ByVal code As String, ByVal ordinal As Long) As String
    Dim words As Variant, word As Variant, n As Long, value As String
    code = Replace(Replace(Replace(code, vbTab, " "), vbCr, " "), vbLf, " ")
    words = Split(Trim$(code), " ")
    For Each word In words
        value = CStr(word)
        If value <> "" Then
            n = n + 1
            If n = ordinal Then
                If Len(value) >= 2 Then
                    If Left$(value, 1) = Chr$(34) And Right$(value, 1) = Chr$(34) Then value = Mid$(value, 2, Len(value) - 2)
                End If
                FieldWord = LCase$(value)
                Exit Function
            End If
        End If
    Next word
End Function

Private Function IsMathTypeWrapper(ByVal f As Field) As Boolean
    Dim code As String, commandName As String, suffix As String, prefix As String
    If f.Type <> wdFieldMacroButton Then Exit Function
    code = f.Code.Text
    If FieldWord(code, 1) <> "macrobutton" Then Exit Function
    commandName = FieldWord(code, 2)
    If commandName = "mtplaceref" Then IsMathTypeWrapper = True: Exit Function
    prefix = "mteditequationsection"
    If Left$(commandName, Len(prefix)) <> prefix Then Exit Function
    suffix = Mid$(commandName, Len(prefix) + 1)
    IsMathTypeWrapper = (suffix = "" Or Not (suffix Like "*[!0-9]*"))
End Function

Private Function IsMathTypeSequence(ByVal f As Field) As Boolean
    Dim code As String, counter As String
    If f.Type <> wdFieldSequence Then Exit Function
    code = f.Code.Text
    If FieldWord(code, 1) <> "seq" Then Exit Function
    counter = FieldWord(code, 2)
    IsMathTypeSequence = (counter = "mteqn" Or counter = "mtchap" Or counter = "mtsec")
End Function

Private Function ChapterNumber(ByVal text As String) As Long
    Dim rx As Object, hit As Object, s As String
    s = Trim$(text)
    Set rx = CreateObject("VBScript.RegExp")
    rx.IgnoreCase = True
    rx.Pattern = "^(Chapter\s+)?([0-9]+)(\.|\s|" & ChrW(&H3001) & "|" & ChrW(&HFF0E) & "|$)"
    If rx.Test(s) Then
        Set hit = rx.Execute(s)(0)
        ' Reject subsection prefixes such as 3.1 even if incorrectly styled Heading 1.
        If Len(s) > hit.Length Then
            If Right$(hit.Value, 1) = "." And Mid$(s, hit.Length + 1, 1) Like "#" Then Exit Function
        End If
        If Len(hit.SubMatches(1)) < 5 Then ChapterNumber = CLng(hit.SubMatches(1))
        Exit Function
    End If
    rx.Pattern = "^" & ChrW(&H7B2C) & "\s*([0-9]+)\s*" & ChrW(&H7AE0)
    If rx.Test(s) Then
        Set hit = rx.Execute(s)(0)
        If Len(hit.SubMatches(0)) < 5 Then ChapterNumber = CLng(hit.SubMatches(0))
        Exit Function
    End If
    rx.Pattern = "^" & ChrW(&H7B2C) & "([^ ]{1,8})" & ChrW(&H7AE0)
    If rx.Test(s) Then
        Set hit = rx.Execute(s)(0)
        ChapterNumber = ChineseChapter(hit.SubMatches(0))
    End If
End Function

Private Function ChineseChapter(ByVal s As String) As Long
    Dim i As Long, n As Long, total As Long, digit As Long, unit As Long, ch As Long
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1)) And &HFFFF&
        digit = -1: unit = 0
        Select Case ch
            Case &H96F6, &H3007: digit = 0
            Case &H4E00: digit = 1
            Case &H4E8C, &H4E24: digit = 2
            Case &H4E09: digit = 3
            Case &H56DB: digit = 4
            Case &H4E94: digit = 5
            Case &H516D: digit = 6
            Case &H4E03: digit = 7
            Case &H516B: digit = 8
            Case &H4E5D: digit = 9
            Case &H5341: unit = 10
            Case &H767E: unit = 100
            Case &H5343: unit = 1000
            Case Else: Exit Function
        End Select
        If digit >= 0 Then
            n = digit
        Else
            If n = 0 Then n = 1
            total = total + n * unit: n = 0
        End If
    Next i
    ChineseChapter = total + n
End Function

Private Function ZC(ByVal value As String) As String
    ZC = Zh(value)
End Function

Private Sub ReleaseFormulaGrid(ByVal formulaRange As Range)
    Dim p As Range, oldRight As Long, oldGrid As Long, state As Long
    Dim touched As Boolean, errorText As String, restored As Boolean
    Set p = FullParagraph(formulaRange)
    oldRight = p.ParagraphFormat.AutoAdjustRightIndent
    oldGrid = p.ParagraphFormat.DisableLineHeightGrid
    If oldRight = wdUndefined Or oldGrid = wdUndefined Then _
        Err.Raise vbObjectError + 7950, , "Select one paragraph with consistent grid settings."
    On Error GoTo Failed
    EnsureGridCalibration
    touched = True
    p.ParagraphFormat.AutoAdjustRightIndent = False
    p.ParagraphFormat.DisableLineHeightGrid = gridOffSetting
    state = SnapGridState(p)
    If state <> 0 Then Err.Raise vbObjectError + 7951, , "Could not confirm paragraph snapToGrid is off."
    If p.ParagraphFormat.AutoAdjustRightIndent <> False Then _
        Err.Raise vbObjectError + 7952, , "Could not disable automatic right indentation."
    Exit Sub
Failed:
    errorText = Err.Description
    On Error Resume Next
    If touched Then
        p.ParagraphFormat.AutoAdjustRightIndent = oldRight
        p.ParagraphFormat.DisableLineHeightGrid = oldGrid
        restored = (p.ParagraphFormat.AutoAdjustRightIndent = oldRight And _
                    p.ParagraphFormat.DisableLineHeightGrid = oldGrid)
    Else
        restored = True
    End If
    On Error GoTo 0
    If Not restored Then errorText = errorText & " / Restore failed; inspect the paragraph."
    Err.Raise vbObjectError + 7953, , errorText
End Sub

Private Function ProbeGridState(ByVal r As Range) As Long
    Dim xml As Object, ps As Object, p As Object, node As Object, st As Object, styles As Object
    Dim attr As Object, id As String, i As Long
    ProbeGridState = -1
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    xml.async = False: xml.validateOnParse = False: xml.resolveExternals = False
    xml.setProperty "ProhibitDTD", True
    If Not xml.LoadXML(r.Document.Content.WordOpenXML) Then Exit Function
    xml.setProperty "SelectionNamespaces", "xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'"
    Set ps = xml.selectNodes("//w:document/w:body/w:p")
    If ps.Length <> 1 Then Exit Function
    Set p = ps.Item(0)
    Set node = p.selectSingleNode("w:pPr/w:snapToGrid")
    If Not node Is Nothing Then ProbeGridState = OnOffValue(node): Exit Function
    ' Table conditional styles require a separate resolver; never guess when direct flags are absent.
    If Not p.selectSingleNode("ancestor::w:tc") Is Nothing Then Exit Function
    Set styles = xml.selectSingleNode("//w:styles")
    If styles Is Nothing Then Exit Function
    Set node = p.selectSingleNode("w:pPr/w:pStyle")
    If Not node Is Nothing Then
        Set attr = node.Attributes.getNamedItem("w:val")
        If attr Is Nothing Then Exit Function
        id = attr.Text
    Else
        For Each st In styles.selectNodes("w:style[@w:type='paragraph']")
            Set attr = st.Attributes.getNamedItem("w:default")
            If Not attr Is Nothing Then
                If attr.Text = "1" Or attr.Text = "true" Or attr.Text = "on" Then
                    id = st.Attributes.getNamedItem("w:styleId").Text
                    Exit For
                End If
            End If
        Next st
    End If
    For i = 1 To 50
        If id = "" Then Exit For
        Set node = Nothing
        For Each st In styles.selectNodes("w:style[@w:type='paragraph']")
            If st.Attributes.getNamedItem("w:styleId").Text = id Then Set node = st: Exit For
        Next st
        If node Is Nothing Then Exit Function
        Set st = node
        Set node = st.selectSingleNode("w:pPr/w:snapToGrid")
        If Not node Is Nothing Then ProbeGridState = OnOffValue(node): Exit Function
        id = ""
        Set node = st.selectSingleNode("w:basedOn")
        If Not node Is Nothing Then
            Set attr = node.Attributes.getNamedItem("w:val")
            If attr Is Nothing Then Exit Function
            id = attr.Text
        End If
    Next i
    If id <> "" Then Exit Function
    Set node = styles.selectSingleNode("w:docDefaults/w:pPrDefault/w:pPr/w:snapToGrid")
    If Not node Is Nothing Then
        ProbeGridState = OnOffValue(node)
    Else
        ' ISO/IEC 29500: if unspecified throughout the style hierarchy, snapping is on.
        ProbeGridState = 1
    End If
End Function

Private Function OnOffValue(ByVal node As Object) As Long
    Dim attr As Object
    Set attr = node.Attributes.getNamedItem("w:val")
    OnOffValue = 1
    If attr Is Nothing Then Exit Function
    Select Case LCase$(attr.Text)
        Case "0", "false", "off": OnOffValue = 0
        Case "1", "true", "on": OnOffValue = 1
        Case Else: OnOffValue = -1
    End Select
End Function

Private Function ParagraphHasFormula(ByVal r As Range) As Boolean
    Dim eq As InlineShape, b As Bookmark
    For Each eq In r.InlineShapes
        If IsMathType(eq) Then ParagraphHasFormula = True: Exit Function
    Next eq
    If r.OMaths.Count > 0 Then ParagraphHasFormula = True: Exit Function
    For Each b In r.Bookmarks
        If Left$(b.Name, 7) = "PWD_EQ_" Then ParagraphHasFormula = True: Exit Function
    Next b
End Function

Public Sub PWD_Grid133_Check()
    Dim doc As Document, p As Paragraph, count As Long, dirty As Boolean
    Dim undoOpen As Boolean, errorText As String, saved As Range
    On Error GoTo Failed
    If Documents.Count = 0 Then Err.Raise vbObjectError + 7953, , "Open a Word document first."
    Set doc = ActiveDocument
    If doc.ReadOnly Or doc.ProtectionType <> wdNoProtection Or doc.TrackRevisions Then _
        Err.Raise vbObjectError + 7954, , "Use an editable document with Track Changes off."
    Set saved = Selection.Range.Duplicate
    EnsureGridCalibration
    Application.UndoRecord.StartCustomRecord "PWD release formula paragraph grid"
    undoOpen = True
    For Each p In doc.Paragraphs
        If ParagraphHasFormula(p.Range) Then
            dirty = True
            ReleaseFormulaGrid p.Range
            count = count + 1
        End If
    Next p
    Application.UndoRecord.EndCustomRecord
    undoOpen = False
    saved.Select
    Application.StatusBar = ""
    Exit Sub
Failed:
    errorText = Err.Description
    On Error Resume Next
    If undoOpen Then Application.UndoRecord.EndCustomRecord
    If undoOpen And dirty Then
        If Not doc.Undo(1) Then errorText = errorText & " / Undo failed; inspect the document."
    End If
    If Not saved Is Nothing Then saved.Select
    On Error GoTo 0
    MsgBox errorText, vbExclamation, "PWD MathType 1.3.3 CHECK"
End Sub

Private Function OtherInlineCount(ByVal doc As Document) As Long
    Dim eq As InlineShape
    For Each eq In doc.InlineShapes
        If Not IsMathType(eq) Then OtherInlineCount = OtherInlineCount + 1
    Next eq
End Function

Private Function ProseOutsideEquations(ByVal doc As Document) As String
    Dim starts() As Long, ends() As Long, n As Long, i As Long, j As Long
    Dim b As Bookmark, f As Field, eq As InlineShape, a As Long, z As Long, cursor As Long, text As String
    For Each b In doc.Bookmarks
        If Left$(b.Name, 7) = "PWD_EQ_" And b.Range.StoryType = wdMainTextStory Then AddExcluded starts, ends, n, b.Range.Start, b.Range.End
    Next b
    For Each eq In doc.InlineShapes
        If IsMathType(eq) Then AddExcluded starts, ends, n, eq.Range.Start, eq.Range.End
    Next eq
    For Each f In doc.Fields
        If PWD_IsNativeNumberU8(f) Or IsChapterBreak(f) Or IsMathTypeReference(f) Then _
            AddExcluded starts, ends, n, f.Code.Start - 1, NativeFieldEnd(f)
    Next f
    For i = 2 To n
        a = starts(i): z = ends(i): j = i - 1
        Do While j >= 1
            If starts(j) <= a Then Exit Do
            starts(j + 1) = starts(j): ends(j + 1) = ends(j): j = j - 1
        Loop
        starts(j + 1) = a: ends(j + 1) = z
    Next i
    cursor = doc.Content.Start
    For i = 1 To n
        If starts(i) > cursor Then text = text & doc.Range(cursor, starts(i)).Text
        If ends(i) > cursor Then cursor = ends(i)
    Next i
    If cursor < doc.Content.End Then text = text & doc.Range(cursor, doc.Content.End).Text
    text = Replace(text, vbTab, "")
    text = Replace(text, vbCr, "")
    text = Replace(text, Chr(7), "")
    ProseOutsideEquations = text
End Function

Private Sub AddExcluded(ByRef starts() As Long, ByRef ends() As Long, ByRef n As Long, ByVal a As Long, ByVal z As Long)
    n = n + 1: ReDim Preserve starts(1 To n): ReDim Preserve ends(1 To n)
    starts(n) = a: ends(n) = z
End Sub

Private Sub ValidateSuiteBookmarks(ByVal doc As Document, ByVal expected As Collection)
    Dim name As Variant, b As Bookmark, count As Long
    For Each name In expected
        If Not doc.Bookmarks.Exists(CStr(name)) Then Err.Raise vbObjectError + 8016, , Zh("8BE5 516C 5F0F 672A 5B8C 6210 6216 6807 8BB0 4E22 5931 FF1A") & CStr(name)
        If Not CompletedInput(doc.Bookmarks(CStr(name)).Range) Then Err.Raise vbObjectError + 8017, , Zh("8BE5 516C 5F0F 672A 5B8C 6210 6216 6807 8BB0 4E22 5931 FF1A") & CStr(name)
    Next name
    For Each b In doc.Bookmarks
        If Left$(b.Name, 7) = "PWD_EQ_" And b.Range.StoryType = wdMainTextStory Then count = count + 1
    Next b
    If count <> expected.Count Then Err.Raise vbObjectError + 8018, , "Formula marker count changed."
End Sub

Private Function StableDocumentText(ByVal doc As Document) As String
    Dim r As Range
    Set r = doc.Content.Duplicate
    r.TextRetrievalMode.IncludeHiddenText = True
    r.TextRetrievalMode.IncludeFieldCodes = True
    StableDocumentText = r.Text
End Function

Private Sub ShowEquationResults(ByVal doc As Document)
    Dim win As Window
    For Each win In doc.Windows
        win.View.ShowFieldCodes = False
    Next win
    HideCodes doc.Content
End Sub

Private Sub EnsureGridCalibration()
    Dim original As Document, scratch As Document, r As Range
    Dim falseState As Long, trueState As Long, restoreState As Long
    Dim errorText As String
    If gridCalibrated Then Exit Sub
    On Error GoTo Failed
    Set original = ActiveDocument
    Set scratch = Documents.Add(Visible:=False)
    scratch.Content.Text = "Paragraph grid calibration."
    Set r = scratch.Paragraphs(1).Range
    r.ParagraphFormat.DisableLineHeightGrid = False
    falseState = ProbeGridState(r)
    r.ParagraphFormat.DisableLineHeightGrid = True
    trueState = ProbeGridState(r)
    If falseState = 0 And trueState = 1 Then
        gridOffSetting = False
    ElseIf falseState = 1 And trueState = 0 Then
        gridOffSetting = True
    Else
        Err.Raise vbObjectError + 7970, , "Grid calibration failed: False=" & falseState & ", True=" & trueState
    End If
    r.ParagraphFormat.DisableLineHeightGrid = gridOffSetting
    restoreState = ProbeGridState(r)
    If restoreState <> 0 Then Err.Raise vbObjectError + 7971, , "Grid calibration did not verify."
    scratch.Close SaveChanges:=wdDoNotSaveChanges
    Set scratch = Nothing
    original.Activate
    gridCalibrated = True
    Exit Sub
Failed:
    errorText = Err.Description
    gridCalibrated = False
    On Error Resume Next
    If Not scratch Is Nothing Then
        Err.Clear
        scratch.Close SaveChanges:=wdDoNotSaveChanges
        If Err.Number <> 0 Then errorText = errorText & " / Scratch cleanup failed: " & Err.Description
    End If
    Err.Clear
    If Not original Is Nothing Then original.Activate
    If Err.Number <> 0 Then errorText = errorText & " / Could not restore original document: " & Err.Description
    On Error GoTo 0
    Err.Raise vbObjectError + 7972, , errorText
End Sub

Private Function SnapGridState(ByVal r As Range) As Long
    Dim p As Range, value As Long
    EnsureGridCalibration
    Set p = r.Paragraphs(1).Range.Duplicate
    value = p.ParagraphFormat.DisableLineHeightGrid
    If value = wdUndefined Then
        SnapGridState = -1
    ElseIf value = gridOffSetting Then
        SnapGridState = 0
    Else
        SnapGridState = 1
    End If
End Function

' Discover raw math without converting old MathType objects or scanning by rendered-text offsets.
Private Function DiscoverNewInputs(ByVal doc As Document) As Collection
    Dim found As New Collection, para As Paragraph, scope As Range, hit As Range, closing As Range, candidate As Range
    Dim opens As Variant, closes As Variant, k As Long, best As Long, pos As Long, closePos As Long
    Dim openKind As Long, code As String, normalized As String, why As String, kind As Long
    Dim openingEnd As Long, coveredEnd As Long
    Dim skip As Boolean, b As Bookmark, styleName As String, paragraphText As String
    Dim markStarts() As Long, markEnds() As Long, markNames() As String
    Dim markCount As Long, markIndex As Long, marked As Range
    If doc.Bookmarks.Count > 0 Then
        ReDim markStarts(1 To doc.Bookmarks.Count)
        ReDim markEnds(1 To doc.Bookmarks.Count)
        ReDim markNames(1 To doc.Bookmarks.Count)
    End If
    For Each b In doc.Bookmarks
        If Left$(b.Name, 7) = "PWD_EQ_" And b.Range.StoryType = wdMainTextStory Then
            Set marked = b.Range
            If Not EmptyProtocolRange(marked) Then
                markCount = markCount + 1
                markStarts(markCount) = marked.Start
                markEnds(markCount) = marked.End
                markNames(markCount) = b.Name
            End If
        End If
    Next b
    opens = Array("\begin{equation}", "\begin{equation*}", "\[", "\(", "$")
    closes = Array("\end{equation}", "\end{equation*}", "\]", "\)", "$")
    For Each para In doc.Paragraphs
        Set scope = para.Range.Duplicate
        paragraphText = scope.Text
        If InStr(1, paragraphText, "\", vbBinaryCompare) = 0 And InStr(1, paragraphText, "$", vbBinaryCompare) = 0 Then GoTo NextInputParagraph
        styleName = LCase$(scope.Style.NameLocal)
        If InStr(styleName, "code") = 0 And InStr(styleName, "verbatim") = 0 Then
            pos = scope.Start
            Do While pos < scope.End
                best = scope.End: openKind = -1
                For k = 0 To 4
                    Set hit = scope.Duplicate
                    hit.Start = pos
                    Do While hit.Start < hit.End
                        If Not FindLiteral(hit, CStr(opens(k))) Then Exit Do
                        If Not RangeEscaped(hit) Then
                            If hit.Start < best Then
                                best = hit.Start
                                openingEnd = hit.End
                                openKind = k
                            End If
                            Exit Do
                        End If
                        If hit.End >= scope.End Then Exit Do
                        hit.SetRange hit.End, scope.End
                    Loop
                Next k
                If openKind < 0 Then Exit Do
                ' ValidateExistingInputs already checked these marked ranges.
                ' Never rediscover their individual delimiters as new formulas.
                coveredEnd = best
                For markIndex = 1 To markCount
                    If markStarts(markIndex) <= best And best < markEnds(markIndex) Then
                        If markEnds(markIndex) > coveredEnd Then coveredEnd = markEnds(markIndex)
                    End If
                Next markIndex
                If coveredEnd > best Then
                    pos = coveredEnd
                    GoTo NextCandidate
                End If
                Set closing = scope.Duplicate
                closing.SetRange openingEnd, scope.End
                closePos = -1
                Do While closing.Start < scope.End
                    If Not FindLiteral(closing, CStr(closes(openKind))) Then Exit Do
                    If Not RangeEscaped(closing) Then
                        If closing.Start < openingEnd Or closing.End <= openingEnd Then _
                            Err.Raise vbObjectError + 8211, , "Formula closing delimiter did not advance past its opening delimiter."
                        closePos = closing.End
                        Exit Do
                    End If
                    If closing.End >= scope.End Then Exit Do
                    closing.SetRange closing.End, scope.End
                Loop
                If closePos < 0 Then Err.Raise vbObjectError + 8200, , "Unclosed formula near: " & Left$(scope.Text, 100)
                Set candidate = doc.Range(best, closePos)
                If candidate.InlineShapes.Count > 0 Or candidate.Fields.Count > 0 Then _
                    Err.Raise vbObjectError + 8201, , "Formula code overlaps an object or field."
                code = candidate.Text
                normalized = NormalizeInput(code)
                If Not PWD_ParseU8(normalized, kind, code, why) Then _
                    Err.Raise vbObjectError + 8202, , "New formula: " & why & " / " & Left$(normalized, 100)
                If kind <> 1 Then ValidateDisplay candidate
                skip = False
                For markIndex = 1 To markCount
                    If markStarts(markIndex) < closePos And markEnds(markIndex) > best Then
                        If markStarts(markIndex) = best And markEnds(markIndex) = closePos And normalized = candidate.Text Then
                            skip = True
                        Else
                            Err.Raise vbObjectError + 8203, , "An existing formula marker partly overlaps the new code: " & markNames(markIndex)
                        End If
                    End If
                Next markIndex
                If Not skip Then found.Add candidate.Duplicate
                pos = closePos
NextCandidate:
            Loop
        End If
NextInputParagraph:
    Next para
    Set DiscoverNewInputs = found
End Function

Private Function NormalizeInput(ByVal value As String) As String
    value = PWD_TrimU8(value)
    If Left$(value, 2) = "\(" And Right$(value, 2) = "\)" Then
        NormalizeInput = "$" & Mid$(value, 3, Len(value) - 4) & "$"
    Else
        NormalizeInput = value
    End If
End Function

Private Sub RegisterNewInputs(ByVal doc As Document, ByVal inputs As Collection)
    Dim i As Long, serial As Long, name As String, r As Range, normalized As String
    Dim opened As Boolean, errorText As String, undoOK As Boolean
    Dim beforeText As String, beforeMarks As Long, changed As Boolean
    If inputs.Count = 0 Then Exit Sub
    On Error GoTo Failed
    beforeText = StableDocumentText(doc): beforeMarks = doc.Bookmarks.Count
    Application.UndoRecord.StartCustomRecord "PWD register pasted formula inputs"
    opened = True
    ' Backward order keeps later edits away from earlier ranges, including adjacent formulas.
    For i = inputs.Count To 1 Step -1
        Set r = inputs(i)
        normalized = NormalizeInput(r.Text)
        If r.Text <> normalized Then r.Text = normalized
        Do
            serial = serial + 1
            name = "PWD_EQ_" & Format$(serial, "000000")
        Loop While doc.Bookmarks.Exists(name)
        doc.Bookmarks.Add Name:=name, Range:=r
    Next i
    Application.UndoRecord.EndCustomRecord
    Exit Sub
Failed:
    errorText = Err.Description
    On Error Resume Next
    If opened Then
        Application.UndoRecord.EndCustomRecord
        Err.Clear
        changed = (StableDocumentText(doc) <> beforeText Or doc.Bookmarks.Count <> beforeMarks)
        If Err.Number <> 0 Then
            errorText = errorText & " / Could not inspect registration state."
        ElseIf changed Then
            undoOK = doc.Undo(1)
            If Not undoOK Then errorText = errorText & " / Could not confirm registration undo."
        End If
    End If
    On Error GoTo 0
    Err.Raise vbObjectError + 8204, , errorText
End Sub

Private Sub ValidateExistingInputs(ByVal doc As Document)
    Dim b As Bookmark, r As Range, kind As Long, tex As String, why As String
    For Each b In doc.Bookmarks
        If Left$(b.Name, 7) = "PWD_EQ_" And b.Range.StoryType = wdMainTextStory Then
            Set r = b.Range
            If Not CompletedInput(r) And Not EmptyProtocolRange(r) Then
                If r.InlineShapes.Count > 0 Or r.Fields.Count > 0 Then Err.Raise vbObjectError + 8205, , "Mixed content in formula marker: " & b.Name
                If Not PWD_ParseU8(r.Text, kind, tex, why) Then Err.Raise vbObjectError + 8206, , b.Name & ": " & why
                If kind <> 1 Then ValidateDisplay r
            End If
        End If
    Next b
End Sub

Private Function EmptyProtocolRange(ByVal r As Range) As Boolean
    If r.InlineShapes.Count > 0 Or r.Fields.Count > 0 Then Exit Function
    EmptyProtocolRange = (PWD_TrimU8(r.Text) = "")
End Function

Private Sub DropEmptyProtocolMarks(ByVal doc As Document)
    Dim names As New Collection, b As Bookmark, name As Variant
    For Each b In doc.Bookmarks
        If Left$(b.Name, 7) = "PWD_EQ_" And b.Range.StoryType = wdMainTextStory Then
            If EmptyProtocolRange(b.Range) Then names.Add b.Name
        End If
    Next b
    For Each name In names
        doc.Bookmarks(CStr(name)).Delete
    Next name
End Sub

Private Function ElapsedSeconds(ByVal started As Single) As Double
    Dim elapsed As Double
    elapsed = CDbl(Timer) - CDbl(started)
    If elapsed < 0 Then elapsed = elapsed + 86400#
    ElapsedSeconds = elapsed
End Function

Private Function IsMathTypeReference(ByVal f As Field) As Boolean
    Dim targetName As String, targetRange As Range, nested As Field, owner As Field
    Dim hasCounter As Boolean
    If f.Type <> wdFieldRef And f.Type <> wdFieldGoToButton Then Exit Function
    targetName = FieldWord(f.Code.Text, 2)
    If targetName = "" Then Exit Function
    If Not f.Code.Document.Bookmarks.Exists(targetName) Then Exit Function
    Set targetRange = f.Code.Document.Bookmarks(targetName).Range
    If targetRange.StoryType <> wdMainTextStory Then Exit Function
    If targetRange.End <= targetRange.Start Then Exit Function
    For Each nested In targetRange.Fields
        If Not IsMathTypeSequence(nested) Then Exit Function
        hasCounter = True
    Next nested
    If Not hasCounter Then Exit Function
    ' Official MathType references bookmark a numbering fragment inside MTPlaceRef.
    ' A bookmark over mixed prose and an equation must remain protected.
    For Each owner In f.Code.Document.Content.Fields
        If owner.Type = wdFieldMacroButton Then
            If FieldWord(owner.Code.Text, 2) = "mtplaceref" Then
                If targetRange.Start >= owner.Code.Start And targetRange.End <= owner.Code.End Then
                    IsMathTypeReference = True
                    Exit Function
                End If
            End If
        End If
    Next owner
End Function

Private Sub ProtectOtherFields(ByVal doc As Document, ByVal protectedFields As Collection)
    Dim firstStory As Range, storyRange As Range, seen As Object
    For Each firstStory In doc.StoryRanges
        Set storyRange = firstStory.Duplicate
        Do While Not storyRange Is Nothing
            Set seen = CreateObject("Scripting.Dictionary")
            ProtectFieldsInRange storyRange, protectedFields, seen
            Set storyRange = storyRange.NextStoryRange
        Loop
    Next firstStory
End Sub

Private Sub ProtectFieldsInRange(ByVal scopeRange As Range, ByVal protectedFields As Collection, ByVal seen As Object)
    Dim f As Field, key As String, related As Boolean
    For Each f In scopeRange.Fields
        key = CStr(f.Code.Start) & ":" & CStr(f.Type)
        If Not seen.Exists(key) Then
            seen.Add key, True
            related = IsMathTypeWrapper(f) Or IsMathTypeSequence(f) Or IsMathTypeReference(f)
            If Not related Then
                If Not f.Locked Then
                    protectedFields.Add f
                    f.Locked = True
                    If Not f.Locked Then Err.Raise vbObjectError + 8220, , "Could not protect a non-MathType field."
                End If
            End If
            If f.Code.Fields.Count > 0 Then ProtectFieldsInRange f.Code, protectedFields, seen
            If f.Result.Fields.Count > 0 Then ProtectFieldsInRange f.Result, protectedFields, seen
        End If
    Next f
End Sub

Private Sub RestoreOtherFields(ByVal protectedFields As Collection)
    Dim i As Long, f As Field, failures As String
    For i = protectedFields.Count To 1 Step -1
        On Error Resume Next
        Err.Clear
        Set f = protectedFields(i)
        f.Locked = False
        If Err.Number = 0 Then
            If f.Locked Then Err.Raise vbObjectError + 8221, , "Field remained locked."
        End If
        If Err.Number = 0 Then
            protectedFields.Remove i
        Else
            failures = failures & " #" & CStr(i) & ": " & Err.Description
        End If
        On Error GoTo 0
    Next i
    If failures <> "" Then Err.Raise vbObjectError + 8222, , "Could not restore original field lock state:" & failures
End Sub

Private Function ProseDifference(ByVal beforeText As String, ByVal afterText As String) As String
    Dim i As Long, limit As Long, beginAt As Long
    limit = Len(beforeText)
    If Len(afterText) < limit Then limit = Len(afterText)
    For i = 1 To limit
        If Mid$(beforeText, i, 1) <> Mid$(afterText, i, 1) Then Exit For
    Next i
    beginAt = i - 25
    If beginAt < 1 Then beginAt = 1
    ProseDifference = Zh("53D8 5316 524D FF1A") & Mid$(beforeText, beginAt, 100) & vbCrLf & _
                      Zh("53D8 5316 540E FF1A") & Mid$(afterText, beginAt, 100)
End Function
