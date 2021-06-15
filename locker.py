import ctypes
from logging import log
import win32com.client
import sys

from os.path import join, isfile
from loguru import logger
from time import sleep

# logger.add('DriveLocker.log')

FILE_KEY = 'DriveLocker.key'
ACC_CONTENT = 'c8524dbce0b244b02915bbe38a87586e'
TIMEOUT_LOOP = 3

def block():
    logger.info('Block desctop!')
    ctypes.windll.user32.LockWorkStation()


def get_obj_items():
    try:
        strComputer = sys.argv[1]
    except IndexError:
        strComputer = "."
    objWMIService = win32com.client.Dispatch( "WbemScripting.SWbemLocator" )
    objSWbemServices = objWMIService.ConnectServer( strComputer, "root/CIMV2" )
    return objSWbemServices.ExecQuery( "SELECT * FROM Win32_LogicalDisk" )


def scan_usbDrive_from_key():
    logger.debug('Start scan devices usb')
    for dr in get_obj_items():
        try:
            if dr.DriveType == 2:
                fileKeyPath = join(dr.Caption, '\\', FILE_KEY)
                logger.debug('File key path: ' + str(fileKeyPath))
                if isfile(fileKeyPath):
                    logger.debug('File key exists!!')
                    return fileKeyPath
                else:
                    logger.debug('Non search key file in ' + str(dr.Caption))
            else:
                continue
        except:
            return False




def locker():
    k = scan_usbDrive_from_key()
    if k:
        with open(k ,'r') as fk:
            if fk.read() != ACC_CONTENT:
                block()
    else:
        block()


def loop():
    logger.info('Start Drive locker!')
    while True:
        sleep(TIMEOUT_LOOP)
        locker()


if __name__ == '__main__':
    loop()


# for d in get_obj_items():
#     try:
#         print(d.Caption)
#         print(d.Description)
#         print(d.DriveType) # 2 its USB
#         print()
#         print()
#     except:
#         print('Error type')

