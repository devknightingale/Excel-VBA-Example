' This code must be placed into the Worksheet SelectionChange area of the Developer code area for the sheet it is intended to work for. At the time of using this coded worksheet, I also created a module to autofit column width, but that code was simply available online and so is not included here. Everything below is written by myself and myself only. It may be a bit messy as I would add and remove things or comment out things that were no longer being used etc. This was in use for about 2 years with occasional updates to add new categories. 




Option Compare Text





Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    'This is just to declare "cell" as a range, do not delete this, it is necessary to make the automatic tagging of [COUNTY]/[STATE] work in the For Each/Next loop below.
    Dim cell As Range
    
    'This declares the variables that count the different categories. Do not change anything here.
    'If you are adding a category to be counted, it must be added as an Integer so that it will be counted as a whole number.
    Dim intIssuesCount, intDuplicate, intOutbreak, intClose, intCaseStatus, intSymptomOnset As Integer
    Dim intDate, intIds, intShell, intLab As Integer
    
    
    
    
    'For each cell in the determined range, checks to see if 1. The value in the cell is Numeric (no letters), and 2. The cell is empty.
    'If BOTH conditions are true, adds the appropriate prefix to the numbers in the cell.
    'You DO NOT need to type [COUNTY] or [STATE] at the beginning of each number in columns A or B! The prefix will be added automatically when you select or tab into a different cell.
    
    'Again, if you need to expand the range for this function, simply change ("A3:A500") to ("A3:A1000") or whatever. Make sure to include the quotes. Same goes for any column such as B.
    'Column A is for the [COUNTY] number. Column B is for the [STATE] number.
    
    Dim intRCCount, intShellReviewCount, intClosed, intNotRCCount, intAntibody, intDataProjects, intReinfection, intNDrive, intCompleteCount, intBreakthroughs, intOpenDRRCount, intCaseReports, intSymptomResolution As Integer
    intShellReviewCount = 0
    intBreakthroughs = 0
    intNDrive = 0
    intTS = 0
    intOpenDRRCount = 0
    intCaseReports = 0
    intReinfection = 0
    intCompleteCount = 0
    intClosed = 0
    intDataProjects = 0
    
    Dim lngHoursWorked As Long
 
    
    
    
    For Each cell In Range("A3:A200")
        'Checks to see if the cell has a value in it, and if that value is Numeric (numbers only)
        'If it is only numbers, then it checks to see how many numbers are entered.
        'Based on the number entered, the prefix will change to follow [COUNTY] numbering conventions.
        If IsNumeric(cell.Value) = True And IsEmpty(cell.Value) = False Then
            If Len(cell.Value) >= 5 Then
                cell.Value = "[COUNTY]" & cell.Value
            End If
            If Len(cell.Value) = 4 Then
                cell.Value = "[COUNTY]0" & cell.Value
            End If
            If Len(cell.Value) = 3 Then
                cell.Value = "[COUNTY]00" & cell.Value
            End If
            If Len(cell.Value) = 2 Then
                cell.Value = "[COUNTY]000" & cell.Value
            End If
            If Len(cell.Value) = 1 Then
                cell.Value = "[COUNTY]0000" & cell.Value
            End If
         End If
    Next
    For Each cell In Range("B3:B200")
        If IsNumeric(cell.Value) = True And IsEmpty(cell.Value) = False Then
            cell.Value = "[STATE]" & cell.Value
        End If
    Next
    ' There is probably a more efficient way to do something like the below, but this worked for the purpose of the worksheet. 
    For Each cell In Range("C3:C200")
        If InStr(cell.Value, "Open Investigations") Or InStr(cell.Value, "DRR") Or InStr(cell.Value, "OI") Then
            intOpenDRRCount = intOpenDRRCount + 1
        End If
        If InStr(cell.Value, "Breakthrough") Or InStr(cell.Value, "BT") Then
            intBreakthroughs = intBreakthroughs + 1
        End If
        If InStr(cell.Value, "Shell") Then
            intShellReviewCount = intShellReviewCount + 1
        End If
        If InStr(cell.Value, "Incomplete") Or InStr(cell.Value, "Complete") Then
            intCompleteCount = intCompleteCount + 1
        End If
        If InStr(cell.Value, "Closed") Then
            intClosed = intClosed + 1
        End If
        If InStr(cell.Value, "Missing") Or InStr(cell.Value, "Case Info") Or InStr(cell.Value, "Patient Info") Then
            intDataProjects = intDataProjects + 1
        End If

    Next
    For Each cell In Range("E3:E200")
        If InStr(cell.Value, "Troubleshooting") Or InStr(cell.Value, "Out") Or InStr(cell.Value, "Duplicate") Or InStr(cell.Value, "Dup") Or InStr(cell.Value, "OOJ") Or InStr(cell.Value, "TS") Then
        intTS = intTS + 1
        End If

    Next
    
    
    
' Calculate time difference for hours worked
Dim HourStartTime, HourEndTime, HourEndTimeRaw, HourStartTimeRaw As Date
Dim strHours As String

strHours = " HOURS"



HourStartTimeRaw = Format((Range("J4").Value), "hh:mm")
HourStartTime = Hour(HourStartTimeRaw) * 60 + Minute(HourStartTimeRaw)
HourEndTimeRaw = Format((Range("J7").Value), "hh:mm")
HourEndTime = Hour(HourEndTimeRaw) * 60 + Minute(HourEndTimeRaw)
HourLunchTimeOutRaw = Format((Range("J5").Value), "hh:mm")
HourLunchTimeInRaw = Format((Range("J6").Value), "hh:mm")
HourLunchTimeOut = Hour(HourLunchTimeOutRaw) * 60 + Minute(HourLunchTimeOutRaw)
HourLunchTimeIn = Hour(HourLunchTimeInRaw) * 60 + Minute(HourLunchTimeInRaw)

ShiftHoursWithLunch = (HourEndTime - HourStartTime) / 60
LunchTimeCalculated = (HourLunchTimeIn - HourLunchTimeOut) / 60
ShiftHoursFinal = ShiftHoursWithLunch - LunchTimeCalculated



'***** LABELED TOTALS FOR CATEGORIES *****
    Range("H11").Value = intShellReviewCount
    Range("H12").Value = intOpenDRRCount
    Range("H13").Value = intBreakthroughs
    Range("H14").Value = intDataProjects
    Range("H15").Value = intCompleteCount
    Range("H16").Value = intTS
    Range("J8").Value = ShiftHoursFinal & strHours
    
'****** TESTING SHIT ************
'This was just used to test things out when I was trying to figure out how to auto-calculate shifts. It was more complicated than I expected it to be. 
    'Range("J23").Value = lngLunch
    'Range("J24").Value = lngHoursWorked
    'Range("J25").Value = lngMinutesWorked
    'Range("J26").Value = lngShift
    'Range("J27").Value = lngShiftCalc
    'Range("J28").Value = HourStartTimeRaw
    'Range("J29").Value = HourEndTimeRaw
    'Range("J30").Value = HourStartTime
    'Range("J31").Value = HourEndTime
    'Range("J32").Value = ShiftHoursWithLunch
    'Range("J33").Value = ShiftHoursFinal
    
    
    
    
    
    
'********** END LABELED TOTALS ***********


'******* TOTALS BREAKDOWN SECTION ********'

    Range("H17").Value = intShellReviewCount + intOpenDRRCount + intBreakthroughs + intCaseReports + intCompleteCount + intClosed
    

   
'******* END OF BREAKDOWN SECTION ********'
            
    
'******* OLD CODE - OBSOLETE ********'
                
'   Range("G7").Value = intAntibody
'   Range("G11").Value = intReinfection
'   Range("G12").Value = intShellCount
'   Range("G13").Value = intCompleteCount

'******* OLD CODE - OBSOLETE ********'
    
End Sub













