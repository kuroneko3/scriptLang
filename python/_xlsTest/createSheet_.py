# coding: UTF-8

import sys
import xlwt
import xlutils
from datetime import date

from cOutputParam import cOutputParam

import calendar

#   xls上の表示位置
left	= 0
top	= 4

#   曜日に応じて表示する値を文字を変化させる
def _getDayOfWeek ( dowNum ):
	
	list = [ u"日", u"月", u"火", u"水", u"木", u"金", u"土" ]
	
	return list[dowNum]

#   現在の年月日から曜日を取得する
def _getDayOfWeek_ ( year, month, day ):
    dayOfWeek = date ( year, month, day ).isoweekday () % 7
    return _getDayOfWeek ( dayOfWeek )

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
	

#   string型の時間を実数に変換
#   ex : 6:00 -> 0.25
def _str2time ( timeStr ):
    times = timeStr.split (":") 
    return  ( ( int ( times[0] ) * 60 + int ( times[1] ) ) / ( 60.0 * 24.0 ) )

def _setDateParam_ ( sheet, dateParam ):
	
	#	行数を設定
	column = dateParam.day + top
	
	#	曜日
	#sheet.write ( column, left + 1, _getDayOfWeek ( dateParam.dayOfWeek ) )
	
	#	出欠
	sheet.write ( column, left + 2, dateParam.come )
	
	#	遅刻
	sheet.write ( column, left + 3, _getLate ( dateParam.late ) )
	
	#	早退
	sheet.write ( column, left + 4, _getEarly ( dateParam.early ) )
	
	#	外出
	sheet.write ( column, left + 5, _getGoOut ( dateParam.goOut ) )
	
	#	出社時間
	sheet.write ( column, left + 6, _str2time ( dateParam.comeTime ) )
	
	#	退社時間
	sheet.write ( column, left + 7, _str2time ( dateParam.leaveTime ) )
	
        #   差し引き
        sheet.write ( column, left + 11, _str2time ( dateParam.minusTime ) )

        #   申請残業
        sheet.write ( column, left + 13, _str2time ( dateParam.applyRemainTime ) )


def _setFormula ( sheet, year, month, day ):
    
    #	行数を設定
    column = day + top

    #	曜日
    sheet.write ( column, left + 1, _getDayOfWeek_ ( year, month, day ) )
    
    #	残業時間(単純計算)
    formula = xlwt.Formula ( u'IF(C{0}=\"出\",IF(H{0}<G{0},H{0}+1-$I$4,H{0}-$I$4),\"\")'.format ( column + 1 ) )
    sheet.write ( column, left + 8, formula )
    
    #       残業時間(休憩考慮)
    formula = xlwt.Formula ( u'IF(C{0}=\"出\",IF(H{0}<G{0},I{0}-$K$4+$J$4,IF(H{0}>$K$4,I{0}-($K$4-$J$4),IF(H{0}>$J$4,$J$4-$I$4,I{0}))),0)'.format ( column + 1 ) )
    sheet.write ( column, left + 9, formula )
    
    #   深夜時間(計算)
    formula = xlwt.Formula ( u'IF(C{0}=\"出\",IF(H{0}<G{0},1-$O$4+H{0},IF(H{0}>$O$4,H{0}-$O$4,0)),0)'.format ( column + 1 ) )
    sheet.write ( column, left + 14, formula )
    
    #   残業時間(表示用)
    formula = xlwt.Formula ( u'IF(C{0}=\"出\",J{0}-L{0},\"\")' .format ( column + 1 ) )
    sheet.write ( column, left + 17, formula )
    
    #   深夜時間(表示用)
    formula = xlwt.Formula ( u'IF(C{0}=\"出\",O{0},\"\")' .format ( column + 1 ) )
    sheet.write ( column, left + 18, formula )



def _searchDate ( day, param ):

	length = len ( param.m_dateParams )
	
	for i in range ( 0, length ):
		if ( day == param.m_dateParams[i].day ):
			return i
		
	return -1

def _setDataParams_ ( sheet, param ):

    monthLastDay = calendar.monthrange ( param.m_year, param.m_month )[1]
    
    #   数式を書き出し，勤怠データがある場合は書き出し
    for i in range ( 1, monthLastDay + 1 ):
        paramNum = _searchDate ( i, param )
        if ( paramNum != -1 ):
            _setDateParam_ ( sheet, param.m_dateParams[paramNum] )
        _setFormula (sheet, param.m_year, param.m_month, i )
    
    #  月の日数に応じて必要ない部分を削る 
    for i in range ( monthLastDay + 1, 32 ):
        sheet.write ( i+top, 0, '' )

def addParameterForSheet ( book, param ):
    
    #   書き込むシートを選択する
    sheet = book.get_sheet( 0 )
    
    #   シートにデータを書き込む
    _setDataParams_ ( sheet, param )



