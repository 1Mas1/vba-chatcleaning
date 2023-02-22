Attribute VB_Name = "telegram_cleaning"

Sub telegramMain()

'removes all filters on activesheet
Call clearFilter

'applies a filter to Source (k) equal to Facebook Messenger
ActiveSheet.Range("a1").AutoFilter field:=11, Criteria1:="Telegram"

Call identifyRecipient

Call tgSplitFromColumn

'applies a filter to to To Attributed column equal to blank. This will remove the Groups and prevent them from beign split
ActiveSheet.Range("a1").AutoFilter field:=9, Criteria1:=(Blanks)

Call tgSplitToColumn

End Sub

Sub tgSplitToColumn()
'splits identifier into 'to' and 'to Attributed'

Dim arrFullID() As String

'this will find the last row using the column #
Dim lastrow As Long
Dim col As Range
Set col = Rows(1).Find("#", LookIn:=xlValues, lookat:=xlWhole)
lastrow = Cells(Rows.Count, col.column).End(xlUp).Row

With ActiveSheet
If .AutoFilterMode Then
    For i = 2 To lastrow
        If Not .Rows(i).Hidden Then
        
        Dim fullId As String
    
        'identifies the full identifier
        fullId = .Cells(i, 8).value
        
        'if there is no second part of the identifer use the first (ie no string with " " delimiter)
        If InStr(fullId, " ") Then
            'splits into array by " "
            arrFullID = Split(fullId, " ", 2)
            On Error GoTo ErrorHandler
        
                Cells(i, 8) = arrFullID(0)
                Cells(i, 9) = arrFullID(1)
        Else
            Cells(i, 9) = fullId
            On Error GoTo ErrorHandler
        End If
        
ErrorHandler:
Resume Next

End If
Next i
End If

End With
End Sub


Sub tgSplitFromColumn()

'splits identifier into 'From' and 'From Attributed'

Dim arrFullID() As String

'this will find the last row using the column #
Dim lastrow As Long
Dim col As Range
Set col = Rows(1).Find("#", LookIn:=xlValues, lookat:=xlWhole)
lastrow = Cells(Rows.Count, col.column).End(xlUp).Row

On Error GoTo nextiteration

With ActiveSheet
If .AutoFilterMode Then
    For i = 57 To lastrow
        If Not .Rows(i).Hidden Then
        
        Dim fullId As String
    
        'identifies the full identifier
        fullId = .Cells(i, 6).value
        
        'identifies "System Messages" and splits accordingly
        If InStr(fullId, "System Message System Message") Then
            Cells(i, 6).value = "System Message"
            Cells(i, 7).value = "System Message"
            GoTo nextiteration
        End If
        
        'if there is no second part of the identifer use the first (ie no string with " " delimiter)
        If InStr(fullId, " ") Then
        
            'splits into array by " "
            arrFullID = Split(fullId, " ", 2)
                
                Cells(i, 6) = arrFullID(0)
                Cells(i, 7) = arrFullID(1)
        Else
            Cells(i, 7) = fullId
            
        End If
        
nextiteration:
Resume Next

End If
Next i
End If

End With

End Sub


Sub clearFilter()
'removes all filters on activesheet

On Error Resume Next
ActiveSheet.ShowAllData


End Sub

Function RemoveStringFromArray(arr As Variant, strRemove As String)
'takes two arguemnts, array and string. loops through array and removes elements matching the string

Dim i As Integer
Dim j As Integer
Dim newArray() As Variant

j = 0
For i = LBound(arr) To UBound(arr)
    'checks if the current element matches the string to remove
    If arr(i) <> strRemove Then
        'resizes the array to include this element
        ReDim Preserve newArray(i)
        'sets the new array element = arr(i)
        newArray(j) = arr(i)
        j = j + 1
    End If
Next i

RemoveStringFromArray = newArray

End Function


Function CleanString(ByVal str As Variant) As String
'takes a string and removes the three types of carriage return and trims

str = CStr(str)

str = Replace(str, vbLf, "")
str = Replace(str, vbCr, "")
str = Replace(str, Chr(10), "")
str = Trim(str)

CleanString = str
            
End Function

Sub identifyRecipient()
'identifies the to column based upon the participants, name (L) column and from column

'sets the currentsheet as 'MainSheet'
Dim mainSheet As Worksheet
Set mainSheet = Worksheets("MainSheet")
mainSheet.Activate

'this will find the last row using the column #
Dim lastrow As Long
Dim col As Range
Set col = Rows(1).Find("#", LookIn:=xlValues, lookat:=xlWhole)
lastrow = Cells(Rows.Count, col.column).End(xlUp).Row

Dim filteredRange As Range
Set filteredRange = Intersect(Range("j2:j" & lastrow), ActiveSheet.AutoFilter.Range)

Dim visibleCells As Range
Set visibleCells = filteredRange.SpecialCells(xlCellTypeVisible)

Dim dict As Object
Set dict = CreateObject("scripting.Dictionary")

Dim cellValue As String
Dim counter As Integer
counter = 1

'creates a flag which is raised when the sender is "system message system message"
Dim systemFlag As Boolean
systemFlag = False

For Each cell In visibleCells
    cellValue = cell.value
    
    'search for the name column and assign to groupName
    Dim groupName As String
    groupName = cell.Offset(0, findName)
    
    'identifies the sender
    Dim sender As String
    sender = cell.Offset(0, -4).value
    
    'trims whitespace from sender
    sender = CleanString(sender)
        
    'if sender is system message, raise flag
    If InStr(sender, "System Message System Message") Then
        systemFlag = True
    End If
    
    'checks whether the participant cell is blank. if so skip
    If cellValue = "" Then
        GoTo nextiteration
    End If
    
    'checks whether column Name is not blank. if not, use value as group name
    If Not groupName = "" Then
        cell.Offset(0, -1).value = CStr(cell.Offset(0, 1).value) & " Group " & groupName
        cell.Offset(0, -2).value = CStr(cell.Offset(0, 1).value) & " Group " & groupName
        GoTo nextiteration
    End If
   
    'splits participants by hard return
    Dim cellValueSplit() As String 'temp array to store split participants
    cellValueSplit = Split(cellValue, vbLf)
        
    'function to remove blank elements from array
    cleanedArray = RemoveStringFromArray(cellValueSplit, "")
    
    'loops through each part of array containing the participants
    For k = LBound(cleanedArray) To UBound(cleanedArray)
                
        'each part of array saved to individualPar variable
        individualpar = cleanedArray(k)
        
        'checks if there is (owner) tag on individualPar and removes if present
        If InStr(individualpar, "(owner)") Then
            individualpar = Replace(individualpar, "(owner)", "")
                    
        'if there is no (owner) and the system message flag is raised, then clean the participant and use that as To. Go to next iteration
        ElseIf systemFlag = True Then
            individualpar = CleanString(individualpar)
            cell.Offset(0, -2).value = individualpar
            systemFlag = False
            GoTo nextiteration
        End If
        
        'function to clean the individual participant. Removes hard returns and trims
        cleanedstring = CleanString(individualpar)
                    
        If Not IsEmpty(cell.Offset(0, -4).value) Then
                
            'checks individualPar is not the same as sender
            If cleanedstring <> sender Then
         
                'if not the same, pindividualPar set as value of 8th column (recipient)
                cell.Offset(0, -2).value = cleanedstring
            Else
                'if the userflag is true, ie it is the subject account AND they are the sender, store details in userDict
                If userFlag = True Then
                    userDict.Add cell.Row, cleanedstring
                End If
            End If
        End If
        
        
   Next k
   
nextiteration:
Next cell
End Sub


Function findName()
'find and returns the column number of the NAME column for telegram

Dim colNumber As Variant
colNumber = WorksheetFunction.Match("Name", ActiveWorkbook.Sheets("MainSheet").Range("1:1"), 0)

'deducts column J from the value of Name as J will always be column 10
findName = colNumber - 10

End Function
