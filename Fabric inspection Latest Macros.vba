Option Explicit

' =====================================================================================
'   Buddhimaan Report Generator (Songskoron 10.7 - Churannto Comment o Logic Shoho)
'   Biboron: Ei macro'ti ek click'ei Shading, Defect, Shortage ebong Bowing/Skewing
'   porikkha kore Summary sheet'er churannto result (B49) ebong bistarito comment (F47, E12)
'   nije thekei toiri kore dey. Eti ekhon shompurno nirbhul ebong churannto.
' =====================================================================================

Sub GenerateResultAndComment()
    ' --- System'er Goti o Bhul Bebosthapona ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrorHandler

    ' --- Proyojoniyo Variable Ghoshona ---
    Dim WsSummary As Worksheet
    Dim fabricType As String, individualStdPoint As Double
    Dim isShadingFail As Boolean, isDefectFail As Boolean
    Dim shadingComment As String, defectComment As String, shortageComment As String, bowingSkewingComment As String, highDefectComment As String
    Dim gsmMoistureComment As String, e12DefectComment As String
    Dim finalComment As String, commentBuilder As String
    Dim missingInfo As String, commentCounter As Integer

    ' --- Summary Sheet Nirdharon ---
    On Error Resume Next
    Set WsSummary = ThisWorkbook.Sheets("Summary")
    On Error GoTo 0
    If WsSummary Is Nothing Then
        MsgBox "Bhul: 'Summary' sheet'ti khuje paowa jaayni!", vbCritical
        GoTo CleanUp
    End If
    
    ' ===== Shurokkha Bebostha: Shoyongkriyo Backup Toiri (Bondho Kora Hoyeche) =====
    ' ... (Backup code commented out for now) ...
    ' =======================================================

    ' ===== Kaaj Shurur Ager Porikkha: Joruri Data Jachai =====
    If IsEmpty(WsSummary.Range("B27").Value) Or Not IsNumeric(WsSummary.Range("B27").Value) Then missingInfo = missingInfo & "- 'Check Roll' (B27)" & vbLf
    If IsEmpty(WsSummary.Range("B41").Value) Or Not IsNumeric(WsSummary.Range("B41").Value) Then missingInfo = missingInfo & "- 'Average Point' (B41)" & vbLf
    If IsEmpty(WsSummary.Range("B43").Value) Or Not IsNumeric(WsSummary.Range("B43").Value) Then missingInfo = missingInfo & "- 'Standard Point' (B43)" & vbLf
    If missingInfo <> "" Then
        MsgBox "Kaaj shuru kora jachche na. Onugroho kore nicher tothogulo din:" & vbLf & vbLf & missingInfo, vbCritical, "Proyojoniyo Totho Nei"
        GoTo CleanUp
    End If
    ' =======================================================

    ' --- User theke Joruri Totho Newa ---
    fabricType = InputBox("Ei kaporti ki 'Solid' naki 'Stripe'?", "Kaporer Dhoron Nirbachon", "Solid")
    If fabricType = "" Then GoTo CleanUp
    If LCase(fabricType) <> "solid" And LCase(fabricType) <> "stripe" Then
        MsgBox "Bhul input. Onugroho kore 'Solid' othoba 'Stripe' likhun.", vbCritical
        GoTo CleanUp
    End If
    
    individualStdPoint = Application.InputBox("Onugroho kore 'Individual Standard Point' din:", "Individual STD Point", Type:=1)
    If individualStdPoint = 0 Then GoTo CleanUp
    ' ==========================================

    ' --- Shob Porikkha Chalano ---
    CheckIndividualRolls WsSummary, individualStdPoint
    highDefectComment = GetHighDefectComment(WsSummary, 3, e12DefectComment) ' Top 3 defects
    defectComment = CheckDefects(WsSummary, isDefectFail, highDefectComment)
    shadingComment = CheckShading(WsSummary, isShadingFail)
    shortageComment = CheckShortage(WsSummary)
    bowingSkewingComment = CheckBowingSkewing(fabricType)
    gsmMoistureComment = CheckGsmMoisture(WsSummary)

    ' --- Churannto Result (B49) Nirdharon ---
    If isDefectFail Or isShadingFail Then
        WsSummary.Range("B49").Value = "FAIL"
    Else
        WsSummary.Range("B49").Value = "PASS"
    End If

    ' --- Buddhimaan Comment (F47) Toiri (Notun Niyom Onujayi) ---
    commentCounter = 1
    
    If WsSummary.Range("B49").Value = "FAIL" Then
        Dim failReason As String
        If isDefectFail Then
            failReason = " " & defectComment
        End If
        
        If isShadingFail Then
            If failReason <> "" Then failReason = failReason & " & "
            failReason = failReason & shadingComment
        End If
        
        commentBuilder = "DUE TO " & failReason & "." & vbLf
        
        If Not isDefectFail And highDefectComment <> "" Then
             commentBuilder = commentBuilder & "1. " & highDefectComment & vbLf
             commentCounter = commentCounter + 1
        End If
        If shadingComment <> "" And Not isShadingFail Then
             commentBuilder = commentBuilder & commentCounter & ". " & shadingComment & vbLf
             commentCounter = commentCounter + 1
        End If

    Else ' Report PASS
        If highDefectComment <> "" Then
            commentBuilder = commentBuilder & commentCounter & ". " & highDefectComment & vbLf
            commentCounter = commentCounter + 1
        End If
        If shadingComment <> "" Then
            commentBuilder = commentBuilder & commentCounter & ". " & shadingComment & vbLf
            commentCounter = commentCounter + 1
        End If
    End If
    
    ' Onanno shadharon issue'gulo jog kora
    If shortageComment <> "" Then
        If commentBuilder <> "" And Right(commentBuilder, 1) <> vbLf Then commentBuilder = commentBuilder & vbLf
        If commentBuilder = "" Then commentCounter = 1
        commentBuilder = commentBuilder & commentCounter & ". " & shortageComment & vbLf
        commentCounter = commentCounter + 1
    End If
    If bowingSkewingComment <> "" Then
        If commentBuilder <> "" And Right(commentBuilder, 1) <> vbLf Then commentBuilder = commentBuilder & vbLf
        If commentBuilder = "" Then commentCounter = 1
        commentBuilder = commentBuilder & commentCounter & ". " & bowingSkewingComment & vbLf
        commentCounter = commentCounter + 1
    End If
    If gsmMoistureComment <> "" Then
        If commentBuilder <> "" And Right(commentBuilder, 1) <> vbLf Then commentBuilder = commentBuilder & vbLf
        If commentBuilder = "" Then commentCounter = 1
        commentBuilder = commentBuilder & commentCounter & ". " & gsmMoistureComment & vbLf
    End If
    
    finalComment = commentBuilder
    If Right(finalComment, 1) = vbLf Then finalComment = Left(finalComment, Len(finalComment) - 1)

    WsSummary.Range("F47").Value = UCase(finalComment)
    WsSummary.Range("E12").Value = e12DefectComment
    
    MsgBox "Result ebong comment shofolbhabe toiri kora hoyeche!", vbInformation

CleanUp:
    ' --- Excel'ke aager obosthay firiye ana ---
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "Ekti oprottashito truti ghoteche: " & Err.Description, vbCritical
    GoTo CleanUp
End Sub


' =============================================================================
'   Shohokari Function'gulo
' =============================================================================

Private Sub CheckIndividualRolls(WsSummary As Worksheet, individualStdPoint As Double)
    Dim Ws As Worksheet, i As Long
    Dim failedRollCount As Long, failedYardsSum As Double
    Dim avgPointCells As Variant, yardCells As Variant, rollAvgPoint As Double
    
    avgPointCells = Array("D40", "H40", "L40", "P40", "T40")
    yardCells = Array("D19", "H19", "L19", "P19", "T19")
    
    For Each Ws In ThisWorkbook.Worksheets
        If Ws.Visible = xlSheetVisible And InStr(1, Ws.Name, "Page", vbTextCompare) > 0 Then
            For i = 0 To 4
                If IsNumeric(Ws.Range(avgPointCells(i)).Value) Then
                    rollAvgPoint = Ws.Range(avgPointCells(i)).Value
                    If rollAvgPoint > individualStdPoint Then
                        failedRollCount = failedRollCount + 1
                        If IsNumeric(Ws.Range(yardCells(i)).Value) Then
                            failedYardsSum = failedYardsSum + Ws.Range(yardCells(i)).Value
                        End If
                    End If
                End If
            Next i
        End If
    Next Ws
    
    WsSummary.Range("B45").Value = failedRollCount
    WsSummary.Range("B47").Value = failedYardsSum
End Sub

Private Function GetHighDefectComment(WsSummary As Worksheet, topCount As Integer, ByRef e12Comment As String) As String
    Dim defects As Object, Ws As Worksheet, i As Long, totalPoints As Double, defectName As String
    Set defects = CreateObject("Scripting.Dictionary")
    
    For Each Ws In ThisWorkbook.Worksheets
        If Ws.Visible = xlSheetVisible And InStr(1, Ws.Name, "Page", vbTextCompare) > 0 Then
            For i = 23 To 38
                totalPoints = Application.WorksheetFunction.Sum(Ws.Range(Ws.Cells(i, "V"), Ws.Cells(i, "AO")))
                If totalPoints > 0 Then
                    defectName = Trim(CStr(Ws.Cells(i, "A").Value))
                    If defectName <> "" Then
                        If defects.Exists(defectName) Then
                            defects(defectName) = defects(defectName) + totalPoints
                        Else
                            defects(defectName) = totalPoints
                        End If
                    End If
                End If
            Next i
        End If
    Next Ws
    
    If defects.Count > 0 Then
        Dim allSortedDefects As String, key As Variant
        Dim tempDict As Object
        Set tempDict = CreateObject("Scripting.Dictionary")
        For Each key In defects.Keys
            tempDict.Add key, defects(key)
        Next key
        
        Do While tempDict.Count > 0
            Dim maxPoints As Double, maxKey As String
            maxPoints = -1
            For Each key In tempDict.Keys
                If tempDict(key) > maxPoints Then
                    maxPoints = tempDict(key)
                    maxKey = key
                End If
            Next key
            allSortedDefects = allSortedDefects & maxKey & ", "
            tempDict.Remove maxKey
        Loop
        If Len(allSortedDefects) > 0 Then allSortedDefects = Left(allSortedDefects, Len(allSortedDefects) - 2)
        e12Comment = "Found- {" & allSortedDefects & "}."

        Dim sortedDefects As String, j As Integer
        For j = 1 To Application.WorksheetFunction.Min(topCount, defects.Count)
            maxPoints = -1
            For Each key In defects.Keys
                If defects(key) > maxPoints Then
                    maxPoints = defects(key)
                    maxKey = key
                End If
            Next key
            sortedDefects = sortedDefects & maxKey & ", "
            defects.Remove maxKey
        Next j
        If Len(sortedDefects) > 0 Then sortedDefects = Left(sortedDefects, Len(sortedDefects) - 2)
        
        Dim avgPoint As Double, stdPoint As Double
        avgPoint = WsSummary.Range("B41").Value
        stdPoint = WsSummary.Range("B43").Value
        
        Dim avgPointText As String
        avgPointText = " {AVG POINT-" & Format(avgPoint, "0.00") & "}"
        
        If avgPoint <= stdPoint And avgPoint >= (stdPoint - 5) Then
            GetHighDefectComment = " " & sortedDefects & avgPointText
        Else
            GetHighDefectComment = sortedDefects & avgPointText
        End If
    End If
End Function

Private Function CheckShading(WsSummary As Worksheet, ByRef isFail As Boolean) As String
    Dim Ws As Worksheet, totalCheckedRolls As Long, criticalShadingRolls As Long
    Dim i As Long, j As Long, k As Long, lastCol As Long
    Dim shadingPercent As Double, isRollCritical As Boolean
    Dim critCSV As Boolean, critSSV As Boolean, critETE As Boolean
    Dim minorCSV As Boolean, minorSSV As Boolean, minorETE As Boolean
    Dim critRanges As Object, minorRanges As Object, cellValue As String
    Set critRanges = CreateObject("Scripting.Dictionary")
    Set minorRanges = CreateObject("Scripting.Dictionary")
    
    isFail = False
    totalCheckedRolls = WsSummary.Range("B27").Value
    If totalCheckedRolls = 0 Then Exit Function
    
    For Each Ws In ThisWorkbook.Worksheets
        If Ws.Visible = xlSheetVisible And InStr(1, Ws.Name, "Page", vbTextCompare) > 0 Then
            lastCol = Ws.Cells(11, Ws.Columns.Count).End(xlToLeft).Column
            For i = 2 To lastCol Step 4
                isRollCritical = False
                Dim hasMinorShadeInRoll As Boolean
                
                For j = i To i + 3
                    For k = 15 To 17
                        cellValue = Trim(CStr(Ws.Cells(k, j).Value))
                        If cellValue <> "" Then
                            If IsCriticalShading(cellValue) Then
                                isRollCritical = True
                                If Not critRanges.Exists(cellValue) Then critRanges.Add cellValue, 1
                                If k = 15 Then critETE = True
                                If k = 16 Then critSSV = True
                                If k = 17 Then critCSV = True
                            ElseIf IsMinorShading(cellValue) Then
                                hasMinorShadeInRoll = True
                                If Not minorRanges.Exists(cellValue) Then minorRanges.Add cellValue, 1
                                If k = 15 Then minorETE = True
                                If k = 16 Then minorSSV = True
                                If k = 17 Then minorCSV = True
                            End If
                        End If
                    Next k
                Next j
                If isRollCritical Then criticalShadingRolls = criticalShadingRolls + 1
            Next i
        End If
    Next Ws
    
    Dim critDetails As String, critRangeStr As String
    critDetails = GetShadeDetails(critCSV, critSSV, critETE)
    critRangeStr = GetRangeString(critRanges)
    
    Dim minorDetails As String, minorRangeStr As String
    minorDetails = GetShadeDetails(minorCSV, minorSSV, minorETE)
    minorRangeStr = GetRangeString(minorRanges)
    
    If criticalShadingRolls > 0 Then
        shadingPercent = (criticalShadingRolls / totalCheckedRolls) * 100
        If shadingPercent >= 20 Then
            isFail = True
            CheckShading = " " & critDetails & "-" & critRangeStr & ""
        Else
            ' Pass, kintu gurutoro shading aache (<20%)
            CheckShading = " " & critDetails & "-" & critRangeStr & ""
        End If
    ElseIf minorRanges.Count > 0 Then
        ' Shudhu tokhon'i minor dekhano hobe jokhon kono gurutoro shading nei
        CheckShading = " " & minorDetails & " SHADE RANGE-" & minorRangeStr & ""
    End If
End Function

Private Function GetShadeDetails(hasCSV As Boolean, hasSSV As Boolean, hasETE As Boolean) As String
    Dim details As String
    If hasCSV Then details = details & "CSV, "
    If hasSSV Then details = details & "SSV, "
    If hasETE Then details = details & "END TO END, "
    If Len(details) > 0 Then
        details = Left(details, Len(details) - 2)
        If InStrRev(details, ",") > 0 Then
            details = Left(details, InStrRev(details, ",") - 1) & " &" & Mid(details, InStrRev(details, ",") + 1)
        End If
    End If
    GetShadeDetails = details
End Function

Private Function GetRangeString(foundRanges As Object) As String
    Dim sortedRanges As Object, rangeKey As Variant
    Dim minRange As String, maxRange As String
    Set sortedRanges = CreateObject("System.Collections.ArrayList")
    
    For Each rangeKey In foundRanges.Keys
        If CStr(rangeKey) <> "" Then
            sortedRanges.Add CStr(rangeKey)
        End If
    Next
    
    If sortedRanges.Count > 0 Then
        sortedRanges.Sort
        minRange = sortedRanges(0)
        maxRange = sortedRanges(sortedRanges.Count - 1)
        If minRange = maxRange Then
            GetRangeString = minRange
        Else
            GetRangeString = minRange & " TO " & maxRange
        End If
    End If
End Function

Private Function IsCriticalShading(val As Variant) As Boolean
    Dim strVal As String, firstPart As String
    IsCriticalShading = False
    If IsEmpty(val) Then Exit Function
    strVal = Trim(CStr(val))
    If strVal = "" Then Exit Function
    
    If IsNumeric(strVal) Then
        If CDbl(strVal) <= 4 Then IsCriticalShading = True
    ElseIf InStr(1, strVal, "/") > 0 Then
        firstPart = Split(strVal, "/")(0)
        If IsNumeric(firstPart) Then
            If CDbl(firstPart) < 4 Then IsCriticalShading = True
        End If
    End If
End Function

Private Function IsMinorShading(val As Variant) As Boolean
    Dim strVal As String
    IsMinorShading = False
    If IsEmpty(val) Then Exit Function
    strVal = Trim(CStr(val))
    If strVal = "" Then Exit Function
    
    If strVal = "4/5" Then IsMinorShading = True
End Function

Private Function CheckDefects(WsSummary As Worksheet, ByRef isFail As Boolean, highDefectComment As String) As String
    Dim avgPoint As Double, stdPoint As Double
    isFail = False
    avgPoint = WsSummary.Range("B41").Value
    stdPoint = WsSummary.Range("B43").Value
    
    If avgPoint > stdPoint Then
        isFail = True
        CheckDefects = highDefectComment
    End If
End Function

Private Function CheckShortage(WsSummary As Worksheet) As String
    Dim orderWidth As Double, actualWidth As Double, lengthShortPercent As Double, afterRelaxationShort As Double
    Dim widthShortageVal As Double, shortageComment As String
    
    If IsNumeric(WsSummary.Range("B15").Value) Then orderWidth = WsSummary.Range("B15").Value
    If IsNumeric(WsSummary.Range("B19").Value) Then actualWidth = WsSummary.Range("B19").Value
    If IsNumeric(WsSummary.Range("I37").Value) Then afterRelaxationShort = WsSummary.Range("I37").Value
    
    Dim lengthPart As String, widthPart As String, relaxationPart As String
    
    If IsNumeric(WsSummary.Range("B33").Value) And WsSummary.Range("B33").Value < 0 Then
        If IsNumeric(WsSummary.Range("B31").Value) And WsSummary.Range("B31").Value > 0 Then
            lengthShortPercent = (Abs(WsSummary.Range("B33").Value) / WsSummary.Range("B31").Value) * 100
            lengthPart = "LENGTH " & Format(lengthShortPercent, "0.00") & "% SHORT"
        End If
    End If
    
    If actualWidth > 0 And orderWidth > 0 And actualWidth < orderWidth Then
        widthShortageVal = orderWidth - actualWidth
        widthPart = "CUTTABLE WIDTH " & Format(widthShortageVal, "0.0") & """ SHORT"
    End If

    If afterRelaxationShort > 0 Then
        relaxationPart = "AFTER RELAXATION LENGTH " & Format(afterRelaxationShort, "0.00") & "% SHORT"
    End If
    
    If lengthPart <> "" Then shortageComment = shortageComment & lengthPart & " & "
    If widthPart <> "" Then shortageComment = shortageComment & widthPart & " & "
    If relaxationPart <> "" Then shortageComment = shortageComment & relaxationPart & " & "
    
    If shortageComment <> "" Then
        shortageComment = Left(shortageComment, Len(shortageComment) - 3)
        CheckShortage = " " & shortageComment
    End If
End Function

Private Function CheckBowingSkewing(fabricType As String) As String
    Dim Ws As Worksheet, resultComment As String, lastRow As Long, i As Long
    Dim actualWidth As Double, point As Double, percent As Double
    Dim minPercent As Double, maxPercent As Double, failCount As Long
    
    For Each Ws In ThisWorkbook.Worksheets
        If InStr(1, UCase(Ws.Name), "BOWING", vbTextCompare) > 0 And InStr(1, UCase(Ws.Name), "SKEW", vbTextCompare) > 0 Then
            Exit For
        End If
        Set Ws = Nothing
    Next Ws
    
    If Ws Is Nothing Or Ws.Visible = xlSheetHidden Then Exit Function
    
    lastRow = Ws.Cells(Ws.Rows.Count, "A").End(xlUp).Row
    If lastRow >= 10 Then
        failCount = 0
        For i = 10 To lastRow
            actualWidth = 0
            point = 0
            If IsNumeric(Ws.Cells(i, "C").Value) Then actualWidth = Ws.Cells(i, "C").Value
            If IsNumeric(Ws.Cells(i, "D").Value) Then point = Ws.Cells(i, "D").Value
            
            If actualWidth > 0 Then
                percent = (point * 100) / actualWidth
                
                Dim isRollFail As Boolean
                isRollFail = False
                
                If LCase(fabricType) = "stripe" Then
                    If percent > 2 Then isRollFail = True
                Else ' Solid
                    If percent > 3 Then isRollFail = True
                End If
                
                Ws.Cells(i, "F").Value = IIf(isRollFail, "FAIL", "PASS")
                
                If isRollFail Then
                    failCount = failCount + 1
                    If failCount = 1 Then
                        minPercent = percent
                        maxPercent = percent
                    Else
                        If percent < minPercent Then minPercent = percent
                        If percent > maxPercent Then maxPercent = percent
                    End If
                End If
            End If
        Next i
        
        If failCount > 0 Then
            Dim issueType As String, issueTitle As String
            If LCase(fabricType) = "stripe" Then
                issueType = "BOWING"
                issueTitle = " "
            Else
                issueType = "SKEWING"
                issueTitle = " "
            End If
            
            If minPercent = maxPercent Then
                resultComment = issueTitle & issueType & " FOUND " & Format(maxPercent, "0.00") & "%."
            Else
                resultComment = issueTitle & issueType & " FOUND " & Format(minPercent, "0.00") & "% TO " & Format(maxPercent, "0.00") & "%."
            End If
            CheckBowingSkewing = resultComment
        End If
    End If
End Function

Private Function CheckGsmMoisture(WsSummary As Worksheet) As String
    Dim reqGsm As Double, stdMoisture As Double
    Dim minGsm As Double, maxGsm As Double, minMoisture As Double, maxMoisture As Double
    Dim gsmIssue As Boolean, moistureIssue As Boolean
    Dim Ws As Worksheet, i As Long, j As Long, lastCol As Long, gsmVal As Double, moistureVal As Double
    Dim gsmCount As Long, moistureCount As Long ' ???? ????? ???? ??????????

    If IsNumeric(WsSummary.Range("H20").Value) Then reqGsm = WsSummary.Range("H20").Value
    If IsNumeric(WsSummary.Range("H16").Value) Then stdMoisture = WsSummary.Range("H16").Value
    minGsm = 9999
    minMoisture = 9999

    For Each Ws In ThisWorkbook.Worksheets
        If Ws.Visible = xlSheetVisible And InStr(1, Ws.Name, "Page", vbTextCompare) > 0 Then
            lastCol = Ws.Cells(11, Ws.Columns.Count).End(xlToLeft).Column
            For i = 2 To lastCol
                For j = 42 To 43
                    If IsNumeric(Ws.Cells(j, i).Value) And Ws.Cells(j, i).Value > 0 Then
                        If j = 42 Then ' GSM
                            gsmVal = Ws.Cells(j, i).Value
                            If gsmVal < minGsm Then minGsm = gsmVal
                            If gsmVal > maxGsm Then maxGsm = gsmVal
                            gsmCount = gsmCount + 1
                        Else ' Moisture
                            moistureVal = Ws.Cells(j, i).Value
                            If moistureVal < minMoisture Then minMoisture = moistureVal
                            If moistureVal > maxMoisture Then maxMoisture = moistureVal
                            moistureCount = moistureCount + 1
                        End If
                    End If
                Next j
            Next i
        End If
    Next Ws

    ' --- Summary Sheet ? ?????? ??????? ???? ??? ??????? ?????? ---
    If gsmCount > 0 Then
        If gsmCount = 1 Then ' <-- ???? ???? ?????? ? ???? ??????? ??? ??????
            WsSummary.Range("H21").Value = minGsm
        Else ' <-- ???? ?????? ? ?? ???? ???? ????? ??????
            WsSummary.Range("H21").Value = minGsm & " TO " & maxGsm
        End If
    End If

    If moistureCount > 0 Then
        If moistureCount = 1 Then ' <-- ???? ???? ?????? ? ???? ??????? ??? ??????
            WsSummary.Range("H17").Value = minMoisture & "%"
        Else ' <-- ???? ?????? ? ?? ???? ???? ????? ??????
            WsSummary.Range("H17").Value = minMoisture & "% TO " & maxMoisture & "%"
        End If
    End If
    ' ------------------------------------------------------------------

    Dim gsmMoistureComment As String, issueTitle As String

    If gsmCount > 0 And (maxGsm > reqGsm * 1.05 Or minGsm < reqGsm * 0.95) Then gsmIssue = True
    If moistureCount > 0 And maxMoisture > stdMoisture Then moistureIssue = True

    If gsmIssue And moistureIssue Then
        issueTitle = " "
    ElseIf gsmIssue Then
        issueTitle = " "
    ElseIf moistureIssue Then
        issueTitle = " "
    End If

    ' --- ?????? ????? ???? ??? ??????? ?????? ---
    If gsmIssue Then
        If gsmCount = 1 Then
             gsmMoistureComment = "GSM FOUND " & minGsm
        Else
             gsmMoistureComment = "GSM FOUND " & minGsm & " TO " & maxGsm
        End If
    End If

    If moistureIssue Then
        If gsmMoistureComment <> "" Then gsmMoistureComment = gsmMoistureComment & " & "
        
        If moistureCount = 1 Then
            gsmMoistureComment = gsmMoistureComment & "MOISTURE FOUND " & minMoisture & "%"
        Else
            gsmMoistureComment = gsmMoistureComment & "MOISTURE FOUND " & minMoisture & "% TO " & maxMoisture & "%"
        End If
    End If
    ' -------------------------------------------

    If gsmMoistureComment <> "" Then
        CheckGsmMoisture = issueTitle & gsmMoistureComment & "."
    End If
End Function
