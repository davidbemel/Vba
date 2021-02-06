# Book 2010

# Delete Hiden Names
```
Sub DeleteHiddenNames() Dim n As Name
Dim Count As Integer
For Each n In ActiveWorkbook.Names If Not n.Visible Then n.Delete
Count = Count + 1 End If Next n MsgBox Count & “ hidden names were deleted.” End Sub
```
#Arrays formula
```
=LEN(A1:A5)
{=A1:A5*B1:B5}
{=SUM(LEN(A1:A5))}
```
# Count Formulas

=COUNTIF(Region,”North”) =COUNTIF(Sales,300) =COUNTIF(Sales,”>300”) =COUNTIF(Sales,”<>100”) =COUNTIF(Region,”?????”)
=COUNTIF(Region,”*h*”) =COUNTIFS(Month,”Jan”,Sales,”>200”) {=SUM((Month=”Jan”)*(Sales>200))} =COUNTIFS(Month,”Jan”,Region,”North”) {=SUM((Month=”Jan”)*(Region=”North”))}
=COUNTIFS(Month,”Jan”,Region,”North”)+ COUNTIFS(Month,”Jan”,Region,”South”)
{=SUM((Month=”Jan”)*((Region=”North”)+ (Region=”South”)))}
=COUNTIFS(Sales,”>=300”,Sales,”<=400”) {=SUM((Sales>=300)*(Sales<=400))}

#Summing formula examples

=SUMIF(Sales,”>200”) =SUMIF(Month,”Jan”,Sales)
=SUMIF(Month,”Jan”,Sales)+SUMIF(Month,”Feb”,Sales)
{=SUM((Month=”Jan”)*(Region=”North”)*Sales)} 
=SUMIFS(Sales,Month,”Jan”,Region,”North”)
{=SUM((Month=”Jan”)*(Region=”North”)*Sales)} 
=SUMIFS(Sales,Month,”Jan”,Region,”<>North”) 
{=SUM((Month=”Jan”)*(Region<>”North”)*Sales)} =SUMIFS(Sales,Month,”Jan”,Sales,”>=200”)
{=SUM((Month=”Jan”)*(Sales>=200)*(Sales))} =SUMIFS(Sales,Sales,”>=300”,Sales,”<=400”) {=SUM((Sales>=300)*(Sales<=400)*(Sales))}

#Working with Dates and Times

=TRIM(A2)Removes excess spaces. 
 Returns #VALUE! // if there is no second space.
=FIND(“ “,B2,1) //Locates the first space.
=FIND(“ “,B2,C2+1) // Locates the second space. Returns #VALUE! if there is no second space
=IF(ISERROR(D2),C2,D2) 
=LEFT(B2,C2)
=RIGHT(B2,LEN(B2)-E2)
=F2&G2  // Concatenate

#Files

filename       Opens the specified file. The filename is a parameter and does not require a switch.
/r filename   Opens the specified file in read-only mode.
/t filename   Opens the specified file as a template
/n filename   Opens the specified file as a template (same as /t).
/e             Starts Excel without creating a new workbook and without displaying its splash screen.
/p directory   Sets the active path to a directory other than the default directory.
/s              Starts Excel in Safe mode and does not load any add-ins or files in the XLStart or alternate start-up file directories.
/m            Forces Excel to create a new workbook that contains a single Microsoft Excel 4.0 macro sheet (obsolete).

“C:\Program Files\Microsoft Office\Office14\EXCEL.EXE” /p c:\xlfiles

#Registry 
You can use the regedit.exe

 HKEY_CLASSES_ROOT  HKEY_CURRENT_USER  HKEY_LOCAL_MACHINE  HKEY_USERS  HKEY_CURRENT_CONFIG

#Procedures: A procedure is basically a unit of computer code that performs some action. VBA supports two types of procedures: Sub procedures and Function procedures
##Sub
Sub Test()
Sum = 1 + 1 MsgBox “The answer is “ & Sum 
End Sub
## Function
Function 
AddTwo(arg1, arg2) AddTwo = arg1 + arg2
End Function
#Hierachy
Application.Workbooks(“Book1.xlsx”)
Application.Workbooks(“Book1.xlsx”).Worksheets(“Sheet1”)
Application.Workbooks(“Book1.xlsx”).Worksheets(“Sheet1”).Range(“A1”)
#Active objects  (omit a reference)
Worksheets(“Sheet1”).Range(“A1”)
Range(“A1”)
#Objects properties
Worksheets(“Sheet1”).Range(“A1”).Value
#VBA variables
Interest = Worksheets(“Sheet1”).Range(“A1”).Value
#Object methods
Range(“A1”).ClearContents
# Message
Sub SayHello()
Msg = “Is your name “ & Application.UserName & “?” Ans = MsgBox(Msg, vbYesNo) If Ans = vbNo Then
MsgBox “Oh, never mind.” Else MsgBox “I must be clairvoyant!” End If End Sub

Sub ShowValue() Msgbox Worksheets(“Sheet1”).Range(“A1”).Value End Sub

#Value&Formula

Sub ChangeValue() Worksheets(“Sheet1”).Range(“A1”).Value = 123.45 End Sub
If Range(“A1”).HasFormula Then MsgBox Range(“A1”).Formula
Range(“D12”).Formula = “=RAND()*100”

```
Sub ZapRange() Worksheets(“Sheet1”).Range(“A1:C3”).Clear End Sub
```
Sub CopyOne()
Worksheets(“Sheet1”).Range(“A1”).Copy _ Worksheets(“Sheet1”).Range(“B1”)
End Sub
```
#Comments
Worksheets(“Sheet1”).Comments(1)
MsgBox Worksheets(“Sheet1”).Comments(1).Text
MsgBox ActiveSheet.Comments.Count
MsgBox ActiveSheet.Comments(1).Parent.Address
# Application Objects
ActiveCell 
ActiveChart
ActiveSheet
ActiveWindow 
ActiveWorkbook 
Selection
ThisWorkbook
##Examples
ActiveCell.ClearContents
MsgBox ActiveSheet.Name
MsgBox ActiveWorkbook.FullName
Selection.Value = 12
ActiveWindow.RangeSelection.Value = 12
MsgBox ActiveWindow.RangeSelection.Count
#The Range property
object.Range(cell1) object.Range(cell1, cell2)
