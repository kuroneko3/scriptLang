# coding: UTF-8

import sys
import xlwt
import xlutils

from cOutputParam import cOutputParam

left	= 1
top	= 3

def _getDayOfWeek ( dowNum ):
	
	list = [ u"日", u"月", u"火", u"水", u"木", u"金", u"土" ]
	
	return list[dowNum]

def _getLate ( isLate ):
	if ( isLate ):
		return u"遅刻"
	return u""

def _getEarly ( isEarly ):
	if ( isEarly ):
		return u"早退"
	return u""

def _getGoOut ( isGoOut ):
	if ( isGoOut ):
		return u"外出"
	return u""
	

def _setDateParam_ ( sheet, dateParam ):
	
	#	行数を設定
	column = dateParam.day + top
	
	#	日付
	sheet.write ( column, left, dateParam.day )
	
	#	曜日
	sheet.write ( column, left + 1, _getDayOfWeek ( dateParam.dayOfWeek ) )
	
	#	出欠
	sheet.write ( column, left + 2, dateParam.come )
	
	#	遅刻
	sheet.write ( column, left + 3, _getLate ( dateParam.late ) )
	
	#	早退
	sheet.write ( column, left + 4, _getEarly ( dateParam.early ) )
	
	#	外出
	sheet.write ( column, left + 5, _getGoOut ( dateParam.goOut ) )
	
	#	出社時間
	sheet.write ( column, left + 6, dateParam.comeTime )
	
	#	退社時間
	sheet.write ( column, left + 7, dateParam.leaveTime )
	
	#	残業時間
	sheet.write ( column, left + 8, dateParam.overTime )
	
	#	深夜時間
	sheet.write ( column, left + 9, dateParam.nightTime )
	

def _searchDate ( day, param ):

	length = len ( param.m_dateParams )
	
	for i in range ( 0, length ):
		if ( day == param.m_dateParams[i].day ):
			return i
		
	return -1

def _setDateParams ( sheet, param ):
	
	alignment = xlwt.Alignment()
	alignment.horz = 2
	
	bool1	= [ False,	False,	False,	False,	False,	True,	False,	True,	False,	True,	True ]
	bool2	= [ True,	False,	False,	False,	False,	False,	True,	False,	True,	False,	True ]
	text0	= [ u"日\n付", u"曜\n日", u"出\n欠", u"遅\n刻", u"早\n退", u"外\n出", u"出社\n時間", u"退社\n時間", u"残業\n時間", u"深夜\n時間", u"備考"]
	
	#	0～31行目
	for i in range ( 0, 32 ):
		
		paramNum = _searchDate ( i, param )
		if ( paramNum != -1 ):
			_setDateParam ( sheet, param.m_dateParams[paramNum] )
			style	= _getDayStyle ( i, bool1[10], bool2[10] )
			sheet.write ( top + i, left + 10, u"", style )
			continue
		
		for j in range ( 0, 11 ):
			style	= _getDayStyle ( i, bool1[j], bool2[j] )
			style.alignment	= alignment
			if ( i == 0 ):
				sheet.write ( top + i, left + j, text0[j], style )
				
			else:
				if ( j == 0 ):
					sheet.write ( top + i, left + j, i, style )
				else:
					sheet.write ( top + i, left + j, u"", style )
	
	#	32行目
	
	

def _setDataParams_ ( sheet, param ):
    
    for i in range ( 0, 32 ):
        paramNum = _searchDate ( i, param )
        if ( paramNum != -1 ):
            _setDateParam_ ( sheet, param.m_dateParams[paramNum] )
           


def createSheet ( book, param ):
	
	sheet = book.add_sheet ( param.name )
	
	_setDateParams ( sheet, param )
	
	sheet.name = u"テスト"

def addParameterForSheet ( book, param ):
    
    sheet = book.get_sheet( 0 )
    
    _setDataParams_ ( sheet, param )