# coding: UTF-8

import sys
from xlrd import open_workbook
from xlutils.copy import copy
import xlwt
import xlutils

from cOutputParam import cOutputParam
from createSheet_ import addParameterForSheet
from cDateParam import cDateParam

from createTestData import createTestData

if __name__ == "__main__":
	
	argvs	= sys.argv
	argc	= len ( argvs )
	
	#   パラメータとして年月が渡されていない場合
	if ( argc != 3 ):
		print 'need year and month'
		quit ()
	
	#   テンプレートを読み込み
	temp = open_workbook( 'template.xls', formatting_info=True )
	
	#	仮パラメータ生成
	params = createTestData ()
	
	
	#   テンプレートをコピー
	wbook = copy ( temp )
	
	#   値を設定
	addParameterForSheet ( wbook, params[0] )
	
	#	xlsファイルを出力する
	outputFileName = str ( argvs[1] ) + '_' + str ( argvs[2] ) + '.xls'
	wbook.save ( outputFileName )
