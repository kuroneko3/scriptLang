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

    #   �e���v���[�g�t�@�C����
    templateFileName = 'template.xls'

    #   �e���v���[�g��ǂݍ���
    tempBook = open_workbook ( templateFileName, formatting_info = True )

    #   �e���v���[�g���o�͗p�ɃR�s�[
    outBook = copy ( tempBook )

    #   �V�[�g���ꗗ�𐶐�
    sheetNames = tempBook.sheet_names ()

    #   �l��ݒ�
    addParameterForSheet ( outBook, attendanceData, sheetNames )

    #   xls�t�@�C�����o�͂���
    outBook.save ( outputFileName )


if __name__ == "__main__":

    argvs   = sys.argv
    argc    = len ( argvs )

    #   �p�����[�^�Ƃ��ĔN�����n����Ă��Ȃ��ꍇ
    if ( argc != 3 ):
        print '�N�C�����K�v'
        quit ()

    #   ���p�����[�^�𐶐�
    testData = createTestData ( int ( argvs[1] ), int ( argvs[2] ) )

    #   �o�̓t�@�C�����𐶐�
    outputFileName = str ( argvs[1] ) + '_' + str ( argvs[2] ) + '.xls'

    outputAttendanceBook ( outputFileName, testData[0] )

