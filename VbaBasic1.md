# VBA Interacting Essentials
## DtaTypes
Short integer (whole number). -32,768 to 32,767ong integer (whole numbers). -2,147,483,648   to  2,147,483,647
Single/Double Used to hold values with decim
Short integer (whole number). -32,768 to 32,767
Date Holds date data types. 1/1/100 to 12/31/9999
## VBA Common Operations (Required syntax in bold)
if Statement
```
If numGrade > 90 Then 
	letterGrade = “A”
ElseIf numGrade > 80 Then
	letterGrade = “B”
Else
	letterGrade = “F”
End If
```
for
```
For x=0 to 49
	‘Loop Over Code
Next x
```
For Each … Next Loop
```
For Each Item In Selection
	Item.Offset(0, 1) = Item * 2
Next
```
Do … Loop While
```
Do
	.Range(“A1”).Offset(Item,0) = Item
Loop While myBool = True
```
Do While 
```
Do While myBool=True
	.Range(“A1”).Offset(Item,0) = Item
Loop
```
### Comparision & Interactions 
User input
```
usrInput = InputBox(“Please Enter Your Name”)
Msgbox “Hello world”
Not Equal: <>
```

### Referencing Workbooks/Worksheets/Ranges
Workbook
```
Workbook that contains code: ThisWorkbook
Using the Active Workbook: Active Workbook
Using Numbered Index: Workbooks(1)
Using Workbook Name: Workbooks(“myWkbk”)
```
Worksheet
```
Using the Active Worksheet:
ActiveSheet
Using the Selected Worksheet:
Windows.SelectedSheets
Using Numbered Index:
Worksheets(1)
Using Worksheet Name:
Worksheets(“myWksht”)
```
Range
```
Reference Single Cell:
Range(“A1”)
Refernce Multiple Adjacent Cells:
Range(“A1:C5”)
Reference Multiple Non Adjacent Cells
Range(“A1:A5, C1:C5”)
Using a Named Range
Range(“myRange”)
```
Cells
```
Refernce All Cells
Worksheet.Cells
Rererence Cells with one Parameter
Cells(3) = “C1”
Reference Cells With Two Parameters
Cells(3,3) = “C3”
Cells(3, “E”) = “E3”
```
### Bucles Tips Useful
With … End With
```
With ThisWorkbook.Worksheets(1)
	.Range(“A1”)=Month
End With
```
ofFset
```
For x=0 to 100
	.Range(“A1”).Offset(x,0) = Rnd
Next x
```
## Basic VBA String Operations
Assign Text to String Variable
```
Dim favFood As String
favFood = “Pizza”
‘Strings must be contained within ‘parenthesis
```
Concatenate
```
Dim fullName as String
fullName = “Joe “ & “Brown”
```
* c2
```
Dim fullName as String, firstName as String
firstName = “Joe”
fullName = firstName & “ Brown”
```
* c3
```
Dim fullName as String, firstName as String, lastName as String
firstName = “Joe”
lastName = ”Brown”
fullName = firstName & “ “ & lastName
```
Get Length Of String
```
Dim fullName as String, stringLength as Integer
fullName = “Joe Brown”
stringLength = Len(fullName)
‘stringLength now contains the value 9
```
Trim
```
Dim fullName As String, stringLength as Integer
fullName = “  George Washington  “
stringLength = Len(fullName)
‘stringLength now contains the value 21
fullName= Trim(fullName)
stringLength = Len(fullName)
‘stringLength now contains the value 17
```
Convert Value to String
```
Dim myInt as Integer, myStr as String
myInt = 100
myStr = CStr(myInt)
‘myStr contains “100” and myInt contains 100
```
## VBA Operations to Change String Variables
Slice left
```
Dim fullName as String, firstName as String
fullName = “George Washington”
firstName = Left(fullName, 6)
‘firstName now contains “George”
```
Slice right
```
Dim fullName as String, lastName as String
fullName = “George Washington”
lastName = Right(fullName, 10)
‘firstName now contains “Washington”
```
Middle 
```
Dim fullName as String, lastName as String
fullName = “George Ed Washington”
lastName = Mid(fullName, 8,2)
‘firstName now contains “Ed”
```
Replace Character
```
Dim myStr as String
myStr = “joy”
myStr = Replace(myStr, “j”, “t”)
myStrnow contains “toy”
```
Replace word
```
Dim myStr as String
myStr = “I love pizza”
myStr = Replace(myStr, “pizza”, “fruit”)
‘myStrnow contains “I love fruit”
```
## VBA String Case Operators
Uppercase
```
Dim myStr as String
myStr = “i love pizza”
myStr = Ucase(myStr)
‘myStr now contains “I LOVE PIZZA”
```
LowerCase
```
Dim myStr as String
myStr = “I LOVE PIZZA”
myStr = Lcase(myStr)
‘myStr now contains “i love pizza”
```
Find Character
```
Dim myStr as String, strLocation as Integer
myStr = “I love pizza”
strLocation = InStr(myStr, “z”)
‘strLocation now contains the value 10
```
Find Word
```
Dim myStr as String, strLocation as Integer
myStr = “I love pizza”
strLocation = InStr(myStr, “pizza”)
‘strLocation now contains the value 8
```
Splite Sentence
```
Dim myStr As String, strArray() As String
myStr = "I love pizza"
strArray = Split(myStr)
'strArray now contains a one dimensional ‘array containing “I”, “love” and ‘“pizza” 
```
Splite Comma delimited 
```
Dim myStr As String, strArray() As String
myStr = "john,smith,large,onion"
strArray = Split(myStr, “,”)
'strArray now contains a one dimensional ‘array containing “john”, “smith” and ‘“large” and “onion”
```
### Joining Strings
Join an array of strings into a single string
```
Dim myStr As String, strArray(0 to 2) As String
strArray(0) = "I"
strArray(1) = "love"
strArray(2) = "pizza"
myStr = Join(strArray)
'myStr now contains “I love pizza” 
```
Join an Array of Strings with comma delimiter
```
Dim myStr As String, strArray(0 to 2) As String
strArray(0) = "I"
strArray(1) = "love"
strArray(2) = "pizza"
myStr = Join(strArray, “,”)
'myStr now contains “I,love,pizza” 
```











