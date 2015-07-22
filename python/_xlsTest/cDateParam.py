# coding: UTF-8

from datetime import date

class cDateParam:
	
	def __init__ ( self ):
		self.day		= 0				#	日付
		self.dayOfWeek	= 0				#	曜日, 0 = 日, 1 = 月・・・
		self.come		= u"出"			#	出欠
		self.late		= False			#	遅刻
		self.early		= False			#	早退
		self.goOut		= False			#	外出
		self.comeTime	= u"9:00"		#	出社時間
		self.leaveTime	= u"17:30"		#	退社時間
		self.overTime	= u"0:00"		#	残業時間
		self.nightTime	= u"0:00"		#	深夜時間
	
	def setDayOfWeek ( self, year, month ):
		self.dayOfWeek	= date ( year, month, self.day ).isoweekday () % 7
	
	def setParam ( self, come, late, early, goOut ):
		self.come		= come
		self.late		= late
		self.early		= early
		self.goOut		= goOut
		
	def setTime ( self, comeTime, leaveTime, overTime, nightTime ):
		self.comeTime	= comeTime
		self.leaveTime	= leaveTime
		self.overTime	= overTime
		self.nightTime	= nightTime
	
