# coding: UTF-8

import sys
from xlrd import open_workbook
from xlutils.copy import copy
import xlwt
import xlutils

from cOutputParam import cOutputParam
from cDateParam import cDateParam
from createSheet_ import addParameterForSheet

from createTestData import createTestData

def outputAttendanceBook ( outputFileName, attendanceData ):

    #   テンプレートファイル名
    templateFileName = 'template.xls'

    #   テンプレートを読み込み
    tempBook = open_workbook ( templateFileName, formatting_info = True )

    #   テンプレートを出力用にコピー
    outBook = copy ( tempBook )

    #   シート名一覧を生成
    sheetNames = tempBook.sheet_names ()

    #   値を設定
    addParameterForSheet ( outBook, attendanceData, sheetNames )

    #   xlsファイルを出力する
    outBook.save ( outputFileName )


if __name__ == "__main__":

    argvs   = sys.argv
    argc    = len ( argvs )

    #   パラメータとして年月が渡されていない場合
    if ( argc != 3 ):
        print '年，月が必要'
        quit ()

    #   仮パラメータを生成
    testData = createTestData ( int ( argvs[1] ), int ( argvs[2] ) )

    #   出力ファイル名を生成
    outputFileName = str ( argvs[1] ) + '_' + str ( argvs[2] ) + '.xls'

    outputAttendanceBook ( outputFileName, testData[0] )

