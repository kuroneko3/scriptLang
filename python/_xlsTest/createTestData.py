
from cOutputParam import cOutputParam
from cDateParam import cDateParam

from datetime import date

import random
import calendar

def createTestData ( year, month ) :
    
    params = []
    
    param = cOutputParam ()
    param.m_name	= u"試験  太郎"
    param.m_sheetName   = u"テスト"

    param.m_year = year
    param.m_month = month

    monthLastDay = calendar.monthrange ( year, month )[1]
	
    for i in range ( 1, monthLastDay + 1 ):
        
        dayOfWeek = date ( year, month, i ).isoweekday () % 7
        if ( dayOfWeek == 0 or dayOfWeek == 6 ):
            continue
        
        dateParam = cDateParam ()
        dateParam.day   = i
        hour = random.randrange ( 17,24 )
        if ( random.randrange ( 0, 5 ) == 0 ):
            hour = random.randrange ( 0, 8 )
        minute = random.randrange ( 0, 59 )
        if ( hour == 17 and minute < 30 ):
            minute = 30

        dateParam.leaveTime = str ( hour ) + ':' + str ( minute )
        
        param.m_dateParams.append ( dateParam )

    params.append ( param )
    
    return params
