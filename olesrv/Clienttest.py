import win32com.client
import time

def main_prog():


    client = win32com.client.Dispatch("OLE1c77Bank.TerminalCom")  # "OLE1c77Bank.TerminalCom")
    #print(client.initchannel())
    ver = client.get_version()
    print(ver)
    client.pathkey = 'D:\py\prj\SrvOleDFS\olesrv\pb_2836413030.jks'
    client.password = '2808Andru1977'
    client.initdfs()
    # # client.TestEthConn('192.168.3.3')
    client.fnrro="4000225321"
    client.statusrro()
    print(client.res_status_smena)
    print(client.res_status)
    print(client.discription_error)
    #
    # # i = 0
    # # while i < 2:
    # #   i+=1
    # # time.sleep(10)
    # #   print('open COM.....')
    # # time.sleep(3)
    # # print(client.testConn())
    # # print('1000.60')
    # # print(client.inputCom('10000.60'))
    # # print('------------')


    # print(client.error)
    # print(client.responseCode)

if __name__ == '__main__':
    main_prog()