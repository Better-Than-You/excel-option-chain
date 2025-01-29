#?
#####? A Public link is created of chat with ChatGpt "https://chatgpt.com/share/6795f6f3-efb0-8012-807c-af6f2c8a4225" #####


# Private Sub Worksheet_Change(ByVal Target As Range)
#     Dim nearestCell As Range
#     Dim lookupValue As Double
#     Dim cell As Range
#     Dim specificYellow As Long
#     Dim lookupRange As Range
#     Dim currentDiff As Double
#     Dim minDiff As Double
    
#     ' Define the specific yellow color (RGB: 255, 255, 0)
#     specificYellow = RGB(255, 255, 0)

#     ' Check if Q10 value changes
#     If Not Intersect(Target, Me.Range("Q10")) Is Nothing Then
#         ' Get the changing value from Q10
#         lookupValue = Me.Range("Q10").Value
        
#         ' Set the lookup range
#         Set lookupRange = Me.Range("Q11:Q52")
        
#         ' Initialize variables
#         Set nearestCell = Nothing
#         minDiff = WorksheetFunction.Max(lookupRange) - WorksheetFunction.Min(lookupRange)
        
#         ' Loop through each cell in the range
#         For Each cell In lookupRange
#             If IsNumeric(cell.Value) Then
#                 currentDiff = Abs(cell.Value - lookupValue)
#                 If currentDiff <= minDiff Then
#                     minDiff = currentDiff
#                     Set nearestCell = cell
#                 End If
#             End If
#         Next cell
        
#         ' Clear previous yellow highlights
#         lookupRange.Interior.ColorIndex = xlNone

#         ' Highlight the nearest cell in yellow
#         If Not nearestCell Is Nothing Then
#             nearestCell.Interior.Color = specificYellow
#         End If
#     End If
# End Sub

#? better to the above color whole row 

# Private Sub Worksheet_Change(ByVal Target As Range)
#     Dim nearestCell As Range
#     Dim lookupValue As Double
#     Dim cell As Range
#     Dim specificYellow As Long
#     Dim lookupRange As Range
#     Dim currentDiff As Double
#     Dim minDiff As Double
#     Dim rowRange As Range
    
#     ' Define the specific yellow color (RGB: 255, 255, 0)
#     specificYellow = RGB(255, 255, 0)
    
#     ' Check if Q10 value changes
#     If Not Intersect(Target, Me.Range("Q10")) Is Nothing Then
#         ' Get the changing value from Q10
#         lookupValue = Me.Range("Q10").Value
        
#         ' Set the lookup range
#         Set lookupRange = Me.Range("Q12:Q52")
        
#         ' Initialize variables
#         Set nearestCell = Nothing
#         minDiff = WorksheetFunction.Max(lookupRange) - WorksheetFunction.Min(lookupRange)
        
#         ' Loop through each cell in the range
#         For Each cell In lookupRange
#             If IsNumeric(cell.Value) Then
#                 currentDiff = Abs(cell.Value - lookupValue)
#                 If currentDiff <= minDiff Then
#                     minDiff = currentDiff
#                     Set nearestCell = cell
#                 End If
#             End If
#         Next cell
        
#         ' Clear previous yellow highlights
#         lookupRange.EntireRow.Interior.ColorIndex = xlNone

#         ' Highlight the entire row containing the nearest cell in yellow
#         If Not nearestCell Is Nothing Then
#             Set rowRange = nearestCell.EntireRow
#             rowRange.Interior.Color = specificYellow
#         End If
#     End If
# End Sub

#? now more better colouring above and below of rows 

# Private Sub Worksheet_Change(ByVal Target As Range)
#     Dim nearestCell As Range
#     Dim lookupValue As Double
#     Dim cell As Range
#     Dim specificYellow As Long
#     Dim aboveColor As Long
#     Dim belowColor As Long
#     Dim lookupRange As Range
#     Dim currentDiff As Double
#     Dim minDiff As Double
#     Dim rowRange As Range
#     Dim ws As Worksheet
#     Dim i As Long
    
#     ' Define specific colors
#     specificYellow = RGB(255, 255, 0) ' Yellow for the nearest row
#     aboveColor = RGB(200, 200, 255)  ' Light blue for rows above
#     belowColor = RGB(255, 200, 200) ' Light red for rows below
    
#     Set ws = Me
    
#     ' Check if Q10 value changes
#     If Not Intersect(Target, ws.Range("Q10")) Is Nothing Then
#         ' Get the changing value from Q10
#         lookupValue = ws.Range("Q10").Value
        
#         ' Set the lookup range
#         Set lookupRange = ws.Range("Q12:Q52")
        
#         ' Initialize variables
#         Set nearestCell = Nothing
#         minDiff = WorksheetFunction.Max(lookupRange) - WorksheetFunction.Min(lookupRange)
        
#         ' Loop through each cell in the range
#         For Each cell In lookupRange
#             If IsNumeric(cell.Value) Then
#                 currentDiff = Abs(cell.Value - lookupValue)
#                 If currentDiff <= minDiff Then
#                     minDiff = currentDiff
#                     Set nearestCell = cell
#                 End If
#             End If
#         Next cell
        
#         ' Clear all previous highlights
#         ws.Rows("12:52").Interior.ColorIndex = xlNone
        
#         ' Highlight the nearest row in yellow
#         If Not nearestCell Is Nothing Then
#             Set rowRange = nearestCell.EntireRow
#             rowRange.Interior.Color = specificYellow
            
#             ' Highlight rows above in light blue
#             For i = 12 To nearestCell.Row - 1
#                 ws.Rows(i).Interior.Color = aboveColor
#             Next i
            
#             ' Highlight rows below in light red
#             For i = nearestCell.Row + 1 To 52
#                 ws.Rows(i).Interior.Color = belowColor
#             Next i
#         End If
#     End If
# End Sub

#? colouring above and below of yellow row with all strike prices green color 

# Private Sub Worksheet_Change(ByVal Target As Range)
#     Dim nearestCell As Range
#     Dim lookupValue As Double
#     Dim cell As Range
#     Dim specificYellow As Long
#     Dim aboveColor As Long
#     Dim belowColor As Long
#     Dim greenColor As Long
#     Dim lookupRange As Range
#     Dim currentDiff As Double
#     Dim minDiff As Double
#     Dim rowRange As Range
#     Dim ws As Worksheet
#     Dim i As Long
    
#     ' Define specific colors
#     specificYellow = RGB(255, 255, 0) ' Yellow for the nearest row
#     aboveColor = RGB(200, 200, 255)  ' Light blue for rows above
#     belowColor = RGB(255, 200, 200)  ' Light red for rows below
#     greenColor = RGB(246, 235, 97) ' Light green for column Q
    
#     Set ws = Me
    
#     ' Check if Q10 value changes
#     If Not Intersect(Target, ws.Range("Q10")) Is Nothing Then
#         ' Get the changing value from Q10
#         lookupValue = ws.Range("Q10").Value
        
#         ' Set the lookup range
#         Set lookupRange = ws.Range("Q12:Q52")
        
#         ' Initialize variables
#         Set nearestCell = Nothing
#         minDiff = WorksheetFunction.Max(lookupRange) - WorksheetFunction.Min(lookupRange)
        
#         ' Loop through each cell in the range
#         For Each cell In lookupRange
#             If IsNumeric(cell.Value) Then
#                 currentDiff = Abs(cell.Value - lookupValue)
#                 If currentDiff <= minDiff Then
#                     minDiff = currentDiff
#                     Set nearestCell = cell
#                 End If
#             End If
#         Next cell
        
#         ' Clear all previous highlights
#         ws.Rows("12:52").Interior.ColorIndex = xlNone
        
#         ' Set column Q to green
#         lookupRange.Interior.Color = greenColor
        
#         ' Highlight the nearest row in yellow
#         If Not nearestCell Is Nothing Then
#             Set rowRange = nearestCell.EntireRow
#             rowRange.Interior.Color = specificYellow
            
#             ' Ensure column Q in the nearest row remains green
#             nearestCell.Interior.Color = specificYellow
            
#             ' Highlight rows above in light blue
#             For i = 12 To nearestCell.Row - 1
#                 ws.Rows(i).Interior.Color = aboveColor
#                 ws.Cells(i, "Q").Interior.Color = greenColor
#             Next i
            
#             ' Highlight rows below in light red
#             For i = nearestCell.Row + 1 To 52
#                 ws.Rows(i).Interior.Color = belowColor
#                 ws.Cells(i, "Q").Interior.Color = greenColor
#             Next i
#         End If
#     End If
# End Sub






##############? Macro for applying highest in volume and oi ##################

#? For specific column
# Sub ApplyConditionalFormatting()
#     Dim ws As Worksheet
#     Dim dataRange As Range
#     Dim highestValue As Double

#     ' Set the worksheet to "nifty"
#     Set ws = ThisWorkbook.Sheets("nifty") ' Ensure the sheet name matches exactly
#     Set dataRange = ws.Range("T12:T52") ' Define the range for conditional formatting

#     ' Find the highest value in the range
#     highestValue = WorksheetFunction.Max(dataRange)

#     ' Clear existing conditional formatting
#     dataRange.FormatConditions.Delete

#     ' Apply yellow color for the highest value
#     With dataRange
#         .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=" & highestValue
#         .FormatConditions(1).Interior.Color = RGB(255, 255, 0) ' Yellow
#     End With

#     ' Apply green color for values >= 75% and < 85% of the highest value
#     With dataRange
#         .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=" & highestValue * 0.75
#         .FormatConditions(2).Interior.Color = RGB(144, 238, 144) ' Light green
#         .FormatConditions(2).StopIfTrue = True
#     End With

#     ' Apply dark green color for values >= 85% and < 90% of the highest value
#     With dataRange
#         .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=" & highestValue * 0.85
#         .FormatConditions(3).Interior.Color = RGB(0, 100, 0) ' Dark green
#         .FormatConditions(3).StopIfTrue = True
#     End With

#     ' Apply red color for values > 90% of the highest value
#     With dataRange
#         .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=" & highestValue * 0.9
#         .FormatConditions(4).Interior.Color = RGB(255, 0, 0) ' Red
#         .FormatConditions(4).StopIfTrue = True
#     End With

#     MsgBox "Conditional formatting applied successfully to the 'nifty' sheet!", vbInformation
# End Sub
 

#? absolutely a general purpose (applying to anyone by calling this function:- ApplyConditionalFormattingToColumn)

# Sub ApplyConditionalFormattingToColumn(columnLetter As String, startRow As Long, endRow As Long)
#     Dim ws As Worksheet
#     Dim dataRange As Range
#     Dim highestValue As Double

#     ' Set the worksheet to "nifty"
#     Set ws = ThisWorkbook.Sheets("nifty") ' Ensure the sheet name matches exactly
#     Set dataRange = ws.Range(columnLetter & startRow & ":" & columnLetter & endRow) ' Define the range dynamically

#     ' Find the highest value in the range
#     highestValue = WorksheetFunction.Max(dataRange)

#     ' Clear existing conditional formatting
#     dataRange.FormatConditions.Delete

#     ' Apply yellow color for the highest value
#     With dataRange
#         .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=" & highestValue
#         .FormatConditions(1).Interior.Color = RGB(255, 255, 0) ' Yellow
#     End With

#     ' Apply green color for values >= 75% and < 85% of the highest value
#     With dataRange
#         .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=" & highestValue * 0.75
#         .FormatConditions(2).Interior.Color = RGB(144, 238, 144) ' Light green
#         .FormatConditions(2).StopIfTrue = True
#     End With

#     ' Apply dark green color for values >= 85% and < 90% of the highest value
#     With dataRange
#         .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=" & highestValue * 0.85
#         .FormatConditions(3).Interior.Color = RGB(0, 100, 0) ' Dark green
#         .FormatConditions(3).StopIfTrue = True
#     End With

#     ' Apply red color for values > 90% of the highest value
#     With dataRange
#         .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=" & highestValue * 0.9
#         .FormatConditions(4).Interior.Color = RGB(255, 0, 0) ' Red
#         .FormatConditions(4).StopIfTrue = True
#     End With

#     MsgBox "Conditional formatting applied successfully to column " & columnLetter & "!", vbInformation
# End Sub

# Sub ApplyFormattingForSpecificColumns()
#     ' Apply formatting to column T
#     Call ApplyConditionalFormattingToColumn("T", 12, 52)
    
#     ' Apply formatting to column U
#     Call ApplyConditionalFormattingToColumn("U", 12, 52)
# End Sub


#? automation logic of upper vba (Note:- you can directly paste this code in module of above code or you can make a different module)

# ' Module 2: Automation Logic

# Dim nextRunTime As Date

# Sub AutoRunFormatting()
#     ' Run the ApplyFormattingForSpecificColumns macro
#     Call ApplyFormattingForSpecificColumns
    
#     ' Schedule the next run after 5 seconds
#     nextRunTime = Now + TimeValue("00:00:05")
#     Application.OnTime nextRunTime, "AutoRunFormatting"
# End Sub

# Sub StopAutoRun()
#     ' Stop the automatic execution
#     On Error Resume Next
#     Application.OnTime nextRunTime, "AutoRunFormatting", Schedule:=False
#     On Error GoTo 0
#     MsgBox "Automatic execution stopped.", vbInformation
# End Sub





















# https://chatgpt.com/share/6795f6f3-efb0-8012-807c-af6f2c8a4225