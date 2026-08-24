//the date in srcElement.value was invalid, so force the user to change it
function BadDate(srcElement)
{
	alert('You did not specify a valid date.');
	srcElement.focus();
	srcElement.select();
	return false;
}

//check the date entered in srcElement.value, and replace it with a formatted version ("dd/mm/yyyy" or "dd mmm yyyy")
function CheckDate(srcElement, ReturnAsMedDate)
{
	//create required matching information
	var PatternMatch = /^(\d{1,2})(\/|-| )(1|01|2|02|3|03|4|04|5|05|6|06|7|07|8|08|9|09|10|11|12|JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)(\/|-| )(\d{1,4})$/i;
	var JAN='01';FEB='02';MAR='03';APR='04';MAY='05';JUN='06';JUL='07';AUG='08';SEP='09';OCT='10';NOV='11';DEC='12';
	var M01='Jan';M02='Feb';M03='Mar';M04='Apr';M05='May';M06='Jun';M07='Jul';M08='Aug';M09='Sep';M10='Oct';M11='Nov';M12='Dec';
	
	//create local variables for the processing
	var arrDate;
	var TheDate = srcElement.value.replace(/^\s*([\s\S]*\S+)\s*$|^\s*$/,'$1');
	var TheYear;
	var TheMonth;
	var TheMonthMed; var TheDay;
	var ExtraZero;

	//check that the user entered a date
	if (TheDate.length == 0) return true;
	
	//check the date against the allowed pattern
	arrDate = TheDate.match(PatternMatch);
	if (arrDate == null) return BadDate(srcElement);

	//determine the year
	TheYear = '';
	for (Counter=0;Counter<(4-arrDate[5].length);Counter++) if (Counter==0) TheYear += '2'; else TheYear += '0';
	TheYear += arrDate[5];
	
	//determine the month
	if (arrDate[3].length == 1) TheMonth = '0' + arrDate[3];
	else if (arrDate[3].length == 3) eval('TheMonth = ' + arrDate[3].toUpperCase());
	else TheMonth = arrDate[3].toString();
	eval('TheMonthMed = M' + TheMonth);
	
	//determine the day
	//if (arrDate[1].length == 1) TheDay = '0' + arrDate[1]; else TheDay = arrDate[1];
	TheDay = arrDate[1] * 1;
	
	//check the number of days for each month
	if (TheMonth == APR || TheMonth == JUN || TheMonth == SEP || TheMonth == NOV) {if (TheDay > 30 || TheDay < 1) return BadDate(srcElement);}
	else if (TheMonth==FEB) {if (TheDay > 29 || (TheDay == 29 && !(TheYear % 4 == 0 && (TheYear % 100 != 0 || TheYear % 400 == 0)))) return BadDate(srcElement);}
	else {if (TheDay > 31 || TheDay < 1) return BadDate(srcElement);}
	
	//return the formatted date
	if (TheDay < 10) ExtraZero = '0'; else ExtraZero = '';
	if (ReturnAsMedDate) srcElement.value = ExtraZero + TheDay + ' ' + TheMonthMed + ' ' + TheYear; else srcElement.value = ExtraZero + TheDay + '/' + TheMonth + '/' + TheYear;
	
	//the function succeeded
	return true;
}