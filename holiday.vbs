'-----------------------
'  Finds the next major US holiday and displays the number of days until that event.
'  This file should be saved with a .vbs file extension
'
'Author:  davenport651
'------------------------

currentDate = Date
currentYear = cint(year(currentDate))

'Six corporate holidays: New Years Day(Jan 1), Memorial Day(Last Monday in May), Fourth of July(July 4), Labor Day(First Monday Sept.),
' Thanksgiving Day(fourth Thursday in Nov.), X-Mas Day(Dec. 25)
newYears = "Jan. 1 " & (currentYear + 1)
memDay = LastMonMay & currentYear
july4 = "July 4 " & currentYear
laborDay = FirstMonSept & currentYear
turkeyDay = FourThursNov & currentYear
xMasDay = "Dec. 25 " & currentYear

Set dicHolidays = CreateObject("Scripting.Dictionary")
dicHolidays.Add "Memorial Day",memDay 'key first, next item
dicHolidays.Add "the 4th of July",july4
dicHolidays.Add "Labor Day",laborDay
dicHolidays.Add "Thanksgiving",turkeyDay
dicHolidays.Add "Christmas Day",xMasDay
dicHolidays.Add "New Years", newYears

arrHoliNames = dicHolidays.keys
arrHoliDates = dicHolidays.items

For i=0 to dicHolidays.count -1
intDays = DateDiff("d", currentDate, cdate(arrHoliDates(i)))
if intDays < 0 then
'wscript.echo arrHoliNames(i) & " has passed, continue loop"
else if intDays = 0 then
wscript.echo "Today is " & arrHoliNames(i)
Exit For
else if intDays > 0 then
wscript.echo intDays & " days until " & arrHoliNames(i)
Exit For
End if
End if
End if
Next

'----------Function/Subs listed below-------------------
Function LastMonMay()
dtmDate = cdate("May 31 " & currentYear)
Do
intDayOfWeek = Weekday(dtmDate)
If intDayOfWeek = 2 Then
'Wscript.Echo "The last Monday of May is " & dtmDate & "."
LastMonMay = "May " & day(dtmDate) & " "
Exit Do
Else
dtmDate = dtmDate - 1
End If
Loop
End Function

Function FirstMonSept()
dtmDate = cdate("Sept 1 " & currentYear)
Do
intDayOfWeek = Weekday(dtmDate)
If intDayOfWeek = 2 Then
'Wscript.Echo "The first Monday of Sept. is " & dtmDate & "."
FirstMonSept = "Sept " & day(dtmDate) & " "
Exit Do
End If
dtmDate = dtmDate + 1
Loop
End Function

Function FourThursNov()
dtmDate = cdate("Nov. 1 " & currentYear)
x = 0
Do until x=4
intDayOfWeek = Weekday(dtmDate)
If intDayOfWeek = 5 Then
'Wscript.Echo "A thursday in Nov. is " & dtmDate & "."
x = x+1
End If
dtmDate = dtmDate + 1
Loop
FourThursNov = "Nov. " & day(dtmDate) & " "
End Function
