<%
'*******************************
'* FUNCTION_FormatDate.asp *
'*******************************

'create the required constants for this page
Const C_FormatDate_Full = 0
Const C_FormatDate_DateOnly = 1
Const C_FormatDate_TimeOnly = 2
Const C_FormatDate_MonthAndYear = 3
Const C_FormatDate_YearMonthOnly = 4
Const C_FormatDate_PageHeaderFormat = 5
Const C_FormatDate_Identifier = 6
Const C_FormatDate_GlobalDate = 7
Const C_FormatDate_ShortDate = 8

'****************************************************************************
'FormatDate																	*
'																			*
'this function will format the specified date to the requested format		*
'																			*
'INPUTS:																	*
'   strDate:	String containing the date to be formatted					*
'	lngFormat:	An integer that represents the format of the output date	*
'																			*
'OUTPUT:																	*
'   String that contains the formatted date									*
'****************************************************************************
Function FormatDate(strDate, lngFormat)
	Dim strReturnDate			'string that will be returned after processing
	Dim strDay					'string that will contain the formatted "Day" (01) portion of the date
	Dim strLongDay				'string that will contain the formatted "Day" (Monday) portion of the date
	Dim strMonth				'string that will contain the formatted "Month" (01) portion of the date
	Dim strShortMonth			'string that will contain the formatted "Month" (Jan) portion of the date
	Dim strLongMonth			'string that will contain the formatted "Month" (January) portion of the date
	Dim strYear					'string that will contain the formatted "Year" portion of the date
	Dim strHour					'string that will contain the formatted "Hour" portion of the date
	Dim strMinute				'string that will contain the formatted "Minute" portion of the date
	Dim strSecond				'string that will contain the formatted "Second" portion of the date
	
	If IsNull(strDate) Then
		FormatDate = null
		Exit Function
	End If
	
	'grab the "day" portion of the date
	strDay = Day(strDate)
	If strDay < 10 Then
		strDay = "0" & strDay
	Else
		strDay = CStr(strDay)
	End If
	
	'grab the name of the day
	Select Case Weekday(strDate)
		Case 1
			strLongDay = "Sunday"
		Case 2
			strLongDay = "Monday"
		Case 3
			strLongDay = "Tuesday"
		Case 4
			strLongDay = "Wednesday"
		Case 5
			strLongDay = "Thursday"
		Case 6
			strLongDay = "Friday"
		Case 7
			strLongDay = "Saturday"
	End Select
	
	'grab the "Month" from the date and format it
	strMonth = Month(strDate)
	If strMonth < 10 Then
		strMonth = "0" & strMonth
	Else
		strMonth = CStr(strMonth)
	End If
	Select Case Month(strDate)
		Case 1
			strShortMonth = "Jan"
			strLongMonth = "January"
		Case 2
			strShortMonth = "Feb"
			strLongMonth = "February"
		Case 3
			strShortMonth = "Mar"
			strLongMonth = "March"
		Case 4
			strShortMonth = "Apr"
			strLongMonth = "April"
		Case 5
			strShortMonth = "May"
			strLongMonth = "May"
		Case 6
			strShortMonth = "Jun"
			strLongMonth = "June"
		Case 7
			strShortMonth = "Jul"
			strLongMonth = "July"
		Case 8
			strShortMonth = "Aug"
			strLongMonth = "August"
		Case 9
			strShortMonth = "Sep"
			strLongMonth = "September"
		Case 10
			strShortMonth = "Oct"
			strLongMonth = "October"
		Case 11
			strShortMonth = "Nov"
			strLongMonth = "November"
		Case 12
			strShortMonth = "Dec"
			strLongMonth = "December"
	End Select
	
	'grab the "Year" part of the date
	strYear = Year(strDate)
	If strYear < 10 Then
		strYear = "200" & strYear
	ElseIf strYear < 100 Then
		strYear = "20" & strYear
	ElseIf strYear < 1000 Then
		strYear = "2" & strYear
	Else
		strYear = CStr(strYear)
	End If
	
	'grab the "Hour" part of the date
	strHour = Hour(strDate)
	If strHour < 10 Then
		strHour = "0" & strHour
	Else
		strHour = CStr(strHour)
	End If
	
	'grab the "Minute" part of the date
	strMinute = Minute(strDate)
	If strMinute < 10 Then
		strMinute = "0" & strMinute
	Else
		strMinute = CStr(strMinute)
	End If
		
	'grab the "Second" part of the date
	strSecond = Second(strDate)
	If StrSecond < 10 Then
		strSecond = "0" & strSecond
	Else
		strSecond = CStr(strSecond)
	End If
	
	'now that all the date parts are formatted, put them into the order specified
	Select Case lngFormat
		
		'return the full formated date
		Case C_FormatDate_Full
			strReturnDate = strDay & " " & strLongMonth & " " & strYear & " " & strHour & ":" & strMinute & ":" & strSecond
		
		'return the date only
		Case C_FormatDate_DateOnly
			strReturnDate = strDay & " " & strLongMonth & " " & strYear
		
		'return the time only
		Case C_FormatDate_TimeOnly
			strReturnDate = strHour & ":" & strMinute & ":" & strSecond
		
		'return the month and year
		Case C_FormatDate_MonthAndYear
			strReturnDate = strLongMonth & " " & strYear
		
		'return the year and month
		Case C_FormatDate_YearMonthOnly
			strReturnDate = strYear & strMonth
		
		'returns the date to be displayed in the header of each page
		Case C_FormatDate_PageHeaderFormat
			strReturnDate = strLongDay & ", " & strDay & " " & strLongMonth & " " & strYear
		
		'returns a unique identifier from the data specified
		Case C_FormatDate_Identifier
			strReturnDate = strYear & "-" & strMonth & "-" & strDay & " " & strHour & "." & strMinute
		
		'returns a date & time that can be recognised reguardless of date settings (month as 3 chars)
		Case C_FormatDate_GlobalDate
			strReturnDate = strDay & " " & strShortMonth & " " & strYear & " " & strHour & ":" & strMinute & ":" & strSecond
		
		'returns a date that can be recognised reguardless of date settings (month as 3 chars)
		Case C_FormatDate_ShortDate
			strReturnDate = strDay & " " & strShortMonth & " " & strYear
	End Select
	
	'return the formatted date string
	FormatDate = strReturnDate
End Function
%>