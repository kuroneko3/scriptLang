# coding: UTF-8

from datetime import date

class cDateParam:
	
    def __init__ ( self ):
        self.day                = 0         #	日付
        self.come               = u"出"     #	出欠
        self.late               = False     #	遅刻
        self.early              = False     #	早退
        self.goOut              = False     #	外出
        self.comeTime           = u"9:00"   #	出社時間
        self.leaveTime          = u"17:30"  #	退社時間
        self.minusTime          = u"0:00"   #   差し引き
        self.applyRemainTime    = u"0:00"   #   申請残業


