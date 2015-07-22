
from cOutputParam import cOutputParam
from cDateParam import cDateParam

def createTestData () :
    
    params = []
    
    param = cOutputParam ()
    param.name	= u"試験  太郎"
	
    dateParam = cDateParam ()
    dateParam.day		= 3
    param.m_dateParams.append ( dateParam )
	
    dateParam1 = cDateParam ()
    dateParam1.day		= 15
    param.m_dateParams.append ( dateParam1 )
    
    param.setDayOfWeek ()
    params.append ( param )
    
    return params
