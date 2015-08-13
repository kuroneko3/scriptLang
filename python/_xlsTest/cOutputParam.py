# coding: UTF-8

class cOutputParam:
	
    def __init__ ( self ):
        
        self.m_year         = 2015  #   年
        self.m_month        = 7     #   月
        self.m_name         = u""   #   名前
        self.m_sheetName    = u""   #   シート名
        
        #   日付ごとの勤怠情報を保持する(cDateParamクラス)
        self.m_dateParams = []
        
 
