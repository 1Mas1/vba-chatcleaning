Attribute VB_Name = "whatsapp_cleaning"

Sub WhatsappMain()

'removes all filters on activesheet
Call clearFilter

'applies a filter to Source (k) equal to Instagram
ActiveSheet.Range("a1").AutoFilter field:=11, Criteria1:="WhatsApp"

Call identifyRecipient

Call waSplitFromColumn

'applies a filter to to To Attributed column equal to blank. This will remove the Groups and prevent them from beign split
ActiveSheet.Range("a1").AutoFilter field:=9, Criteria1:=(Blanks)

Call waSplitToColumn


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

'creates a flag which is raised when the sender is "system message system message"
Dim systemFlag As Boolean
systemFlag = False

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
    
    'if sender is system message . . .
    If InStr(sender, "System Message System Message") Then
        systemFlag = True
    End If
    
    'checks whether the participant cell is blank. if so skip
    If cellValue = "" Then
        systemFlag = True
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
            On Error GoTo nextiteration
            dict(cellValue) = CStr(cell.Offset(0, 1).value) & " Group " & counter
            cell.Offset(0, -1).value = CStr(cell.Offset(0, 1).value) & " Group " & counter
            cell.Offset(0, -2).value = CStr(cell.Offset(0, 1).value) & " Group " & counter
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
                    
                userFlag = False
                
                End If
             End If
            Next k
        End If
    End If

nextiteration:
Next cell


End Sub


Sub waSplitToColumn()
'splits identifier into 'to' and 'to Attributed'

Dim arrFullID() As String

'this will find the last row using the column #
Dim lastrow As Long
Dim col As Range
Set col = Rows(1).Find("#", LookIn:=xlValues, lookat:=xlWhole)
lastrow = Cells(Rows.Count, col.column).End(xlUp).Row

On Error GoTo nextiteration

With ActiveSheet
If .AutoFilterMode Then
    For i = 2 To lastrow
        If Not .Rows(i).Hidden Then
        
        Dim fullId As String
    
        'identifies the full identifier
        fullId = .Cells(i, 8).value
        
        'skips row if no data
        If fullId = "" Then
        GoTo nextiteration
        End If
        
        'identifies "System Messages" and splits accordingly
        If InStr(fullId, "System Message System Message") Then
            Cells(i, 8).value = "System Message"
            Cells(i, 9).value = "System Message"
            GoTo nextiteration
        End If
        

        'splits into array by "@"
        arrFullID = Split(fullId, "@", 2)
               
        'stores number in from user column
        'cleans number removes 44 and replace with 0
        arrFullID(0) = msisdn44to0(arrFullID(0))
        
        Cells(i, 8) = arrFullID(0)
               
        'splits string after " " to get saved name if present
        Dim waUserName As String
        waUserName = Mid(arrFullID(1), InStr(arrFullID(1), " ") + 1)
        'waUserName = Split(arrFullID(1), " ")
               
        'if no username stored, use number
        Dim noUsername As String
        noUsername = arrFullID(1)
        
        If Len(noUsername) = 14 Or Len(noUsername) = 15 Then
            Cells(i, 9).value = arrFullID(0)
        Else
        
            If waUserName = "" Then
                Cells(i, 8).value = arrFullID(0)
            Else
                Cells(i, 9) = waUserName
            End If
        End If
            
End If
nextiteration:
Resume Next
Next i
End If

End With

End Sub


Sub waSplitFromColumn()
'splits identifier into 'from' and 'from Attributed'

Dim arrFullID() As String

'this will find the last row using the column #
Dim lastrow As Long
Dim col As Range
Set col = Rows(1).Find("#", LookIn:=xlValues, lookat:=xlWhole)
lastrow = Cells(Rows.Count, col.column).End(xlUp).Row

On Error GoTo nextiteration

With ActiveSheet
If .AutoFilterMode Then
    For i = 2 To lastrow
        If Not .Rows(i).Hidden Then
        
        Dim fullId As String
    
        'identifies the full identifier
        fullId = .Cells(i, 6).value
        
        'skips row if no data
        If fullId = "" Then
        GoTo nextiteration
        End If
        
        'identifies "System Messages" and splits accordingly
        If InStr(fullId, "System Message System Message") Then
            Cells(i, 6).value = "System Message"
            Cells(i, 7).value = "System Message"
            GoTo nextiteration
        End If
        

        'splits into array by " "
        arrFullID = Split(fullId, "@", 2)
               
        'stores number in from user column
        'cleans number removes 44 and replace with 0
        arrFullID(0) = msisdn44to0(arrFullID(0))
        
        Cells(i, 6) = arrFullID(0)
               
        'splits string after " " to get saved name if present
        Dim waUserName As String
        waUserName = Mid(arrFullID(1), InStr(arrFullID(1), " ") + 1)
        'waUserName = Split(arrFullID(1), " ")
               
        'if no username stored, use number
        Dim noUsername As String
        noUsername = arrFullID(1)
        
        If Len(noUsername) = 14 Or Len(noUsername) = 15 Then
            Cells(i, 7).value = arrFullID(0)
        Else
        
            If waUserName = "" Then
                Cells(i, 6).value = arrFullID(0)
            Else
                Cells(i, 7) = waUserName
            End If
        End If
            
End If
nextiteration:
Resume Next
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


Function msisdn44to0(ByVal str As Variant) As String
'takes a string and removes the first 2x digits (44) and replaces with 0

Dim subString  As String

subString = Right(str, Len(str) - 2)
str = "0" + subString

msisdn44to0 = str

End Function


