Attribute VB_Name = "snapchat_cleaning"

Sub SnapchatMain()

'removes all filters on activesheet
Call clearFilter

'applies a filter to Source (k) equal to Instagram
ActiveSheet.Range("a1").AutoFilter field:=11, Criteria1:="Snapchat"

Call identifyRecipient

Call scSplitFromColumn

'applies a filter to to To Attributed column equal to blank. This will remove the Groups and prevent them from beign split
ActiveSheet.Range("a1").AutoFilter field:=9, Criteria1:=(Blanks)

Call scSplitToColumn


End Sub


Sub identifyRecipient()
'Loops through all cells in participants column and splits by hard return.
'If less than 2: compares both participants to the from column to identify the recipient
'If more than 2: checks whether value in group dictionary and returns key if so. If not add value to dict, and increase group counter by 1

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

'creates  flag which is raised when the subject account (owner) is the sender
Dim userFlag As Boolean
userFlag = False

'creates a dict to store the subject user name + row nubmber when the subject is the sender
Dim userDict As Object
Set userDict = CreateObject("scripting.Dictionary")


For Each cell In visibleCells
    cellValue = cell.value
    
    'identifies the sender
    Dim sender As String
    sender = cell.Offset(0, -4).value
    'trims whitespace from sender
    sender = CleanString(sender)
        
    'checks whether the participant cell is blank. if so skip
    If cellValue = "" Then
        GoTo nextiteration
    End If
    
    'checks whether the value is already in the group dict. If so, sets att/user columns
    If dict.Exists(cellValue) Then
        cell.Offset(0, -1).value = dict(cellValue) 'to attributed
        cell.Offset(0, -2).value = dict(cellValue) 'to user
    Else
        'if not, splits participants by hard return
        Dim cellValueSplit() As String 'temp array to store split participants
        cellValueSplit = Split(cellValue, vbLf)
        
        'function to remove blank elements from array
        cleanedArray = RemoveStringFromArray(cellValueSplit, "")
        
        'checks whether the cleaned array contains more than 2 participants split by hard return
        If UBound(cleanedArray) > 1 Then
            dict(cellValue) = CStr(cell.Offset(0, 1).value) & " Group " & counter
            cell.Offset(0, -1).value = CStr(cell.Offset(0, 1).value) & " Group " & counter
            cell.Offset(0, -2).value = CStr(cell.Offset(0, 1).value) & " Gourp " & counter
            counter = counter + 1
        Else
            'if there are only two participants, loops through each part of array containing the participants
            For k = LBound(cleanedArray) To UBound(cleanedArray)
                
                'each part of array saved to individualPar variable
                individualpar = cleanedArray(k)
                    
                'checks if there is (owner) tag on individualPar and removes if present
                If InStr(individualpar, "(owner)") Then
                    individualpar = Replace(individualpar, "(owner)", "")
                    userFlag = True
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
                    
                userFlag = False
                
                End If
             End If
            Next k
        End If
    End If
nextiteration:
Next cell



End Sub


Sub scSplitToColumn()
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


Sub scSplitFromColumn()

'splits identifier into 'From' and 'From Attributed'

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
        fullId = .Cells(i, 6).value
        
        'if there is no second part of the identifer use the first (ie no string with " " delimiter)
        If InStr(fullId, " ") Then
        
            'splits into array by " "
            arrFullID = Split(fullId, " ", 2)
         
            On Error GoTo ErrorHandler
        
                Cells(i, 6) = arrFullID(0)
                Cells(i, 7) = arrFullID(1)
        Else
            Cells(i, 7) = fullId
            On Error GoTo ErrorHandler
            
        End If
        
ErrorHandler:
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



