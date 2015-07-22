# coding: UTF-8

import cDateParam

class cOutputParam:
	
	def __init__ ( self ):
		
		self.m_year		= 2015	#	年
		self.m_month	= 7		#	月
		self.m_name		= u""	#	名前
		
		self.m_dateParams = []
		
		
	def setDayOfWeek ( self ):
		
		length = len ( self.m_dateParams )
		for i in range ( 0, length ):
			self.m_dateParams[i].setDayOfWeek ( self.m_year, self.m_month )
		
	
	def setParameter ( self, year, month, name ):
		
		self.m_year		= year
		self.m_month	= month
		self.m_name		= name
		
		self.setDayOfWeek ()
	
