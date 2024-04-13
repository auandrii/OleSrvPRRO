# -*- coding: utf-8 -*-
import hashlib
import sys
import serial
import json
import logging
import pathlib
import time
import socket
import grpc
import rro_pb2 as vbk
import rro_pb2_grpc as vbkrpc
import settings
import EUSignCP as EU
import pythoncom
#from xml.etree import ElementTree


PATH = pathlib.Path(__file__).parent.absolute()
VERSION = '1.2'
x = settings.CODE_STR

logging.basicConfig(filename=f'{PATH}/{settings.NAME_FILE_LOGT}', format='%(asctime)s - %(message)s',
                    level=logging.INFO)

debugging = settings.DEBUG

if debugging:
    from win32com.server.dispatcher import DefaultDebugDispatcher
    useDispatcher = DefaultDebugDispatcher
else:
    useDispatcher = None


def logging_report(txtreport, type_report='') -> None:
    if type_report == 'excp':
        logging.exception(txtreport)
    else:
        if debugging == 1:
            logging.info(txtreport)

def table_wares(wares_string):

    return None

def parser_string(str_pars):

    str_pars = str_pars.split('&')
    new_string = ''
    for _ in str_pars:
        if x.get(_) != None:
            new_string += x.get(_)
        else:
            new_string += _
    return new_string


def create_check_prro(checktype, fiscal_number, edrpo, id_prro, date_check=0, last_hash_check=0,
                      check_xml='') -> bytes:
    check = f'<?xml version="1.0" encoding="windows-1251"?><RQ V="1">' \
            f'<DAT FN="{fiscal_number}" TN="{edrpo}" ZN="" DT="{id_prro}" V="1">'
    if checktype == '2':  # Z-report
        logging_report(check_xml,)
        check_xml = check_xml.split(';')
        check += f'<Z NO="1">'
        smi = 0
        smo = 0
        for i in check_xml:
            string_xml = i.split(':')
            name = parser_string(string_xml[0])

            if name != '':
                check += f'<M NM="{name}" SMI="{string_xml[1]}" SMO="{string_xml[2]}" T="{string_xml[3]}"/>'
                smi += int(string_xml[4])
                smo += int(string_xml[5])
        check += f'<NC NI="{smi}" NO="{smo}"/></Z>'
    elif checktype == '3':  # 0-check
        check += f'<C T="108"></C>'
    elif checktype == '0' or checktype == '1':  # check_order or returne
        check_xml = check_xml.split('&\?')
        head = check_xml[0].split('&\:')
        summa = head[0]
        num_chek = head[1]
        nal_beznal = head[2]
        nal_beznal_txt = parser_string(head[5])
        fiscal_number = head[3]
        date_check = head[4]
        check += f'<C T="{checktype}">'
        check_xml = check_xml[1].split('&\;')  # +"&\:"
        count = 0
        for i in check_xml:

            string_xml = i.split('&\:')
            name = parser_string(string_xml[0])
            if name != '':
                count += 1
                check += f'<P N="{count}" C="{string_xml[4]}"'
                check += f' NM="{name}" SM="{string_xml[1]}" Q="{string_xml[2]}" PRC="{string_xml[3]}" TX="0"></P>'
                logging_report(f'NAME:{name} - summa:{string_xml[1]} quantity:{string_xml[2]} - price:{string_xml[3]}')
        check += f'<M N="{count+1}" T="{nal_beznal}" NM="{nal_beznal_txt}" SM="{summa}" M="{summa}" RM="0"/>'
        check += f'<E N="{count+2}" NO="{num_chek}" SM="{summa}" FN="{fiscal_number}" TS="{date_check}"/>'
        check += f'</C>'
    check += f'<TS>{date_check}</TS></DAT><MAC ID="">{last_hash_check.strip()}</MAC></RQ>'
    logging_report(f'checkdfs  XML in {check}')
    return str.encode(check, encoding='windows-1251')

class TerminalCom:
    # _reg_clsctx_ = pythoncom.CLSCTX_LOCAL_SERVER
    # _reg_clsid_ = "{746890E4-EB54-478C-A60F-301B1D661A8F}"
    #_reg_clsid_ = "{746890E4-EB54-478C-A60F-301B1D661A8F}"  # test
    _reg_clsid_ = pythoncom.CreateGuid()
    _reg_progid_ = "OLE1c77Bank.TerminalComT"
    _reg_desc_ = "OLE1c77Bank.TerminalCom.V3"               # test
    _public_methods_ = ['get_version',
                        'to_utf8',
                        'hash256str',
                        'hash256str_',
                        'checkdfs',
                        'initdfs',
                        'inforrodfs',
                        'initCom',
                        'openCom',
                        'gui_progressbar',
                        'statusrro',
                        'inputCom',
                        'testConn',
                        # 'outputCom',
                        # 'test_eth_conn',
                        'inputEth']
    _public_attrs_ = ['res_status_smena',
                      'hash256string',
                      'res_status',
                      'res_id',
                      'discription_error',
                      'password',
                      'pathkey',
                      'error',
                      'responseCode',
                      'errorDescription',
                      # 'pan',
                      # 'date',
                      # 'time',
                      # 'rrn',
                      'receipt',
                      'fnrro']
    # _readonly_attrs_ = []
    
    def __init__(self):
        self.error = True
        self.errorDescription = ''
        self.responseCode = 0
        self.receipt = ''
        self.ip_adress = ""
        self.port = 2000
        self.fnrro = ''
        self.pIface = None
        self.pathkey = ''
        self.password = ''
        self.res_id = 0
        self.res_status = ''
        self.discription_error = ''
        self.hash256string = ''
        self.res_status_smena = 0
        self.ser_com_port = None
        logging_report(f'Create OLE server')

    def get_version(self):
        logging_report(f'Get version - {VERSION}')
        logging_report(self.text)
        return VERSION

    def gui_progressbar(self):
       pass

    def to_utf8(self, text):
        logging_report(f'TEXT in {text}')
        return text

    def hash256str(self, name_file):
        b_chek_xml = ''
        logging_report(f'hash256str open file {PATH}/{name_file}.xml')
        ftxt = open(f'{PATH}/{name_file}.xml', mode='r', encoding='windows-1251')
        for line in ftxt:
            b_chek_xml += line.strip('\r\n')
        encoded = b_chek_xml.encode(encoding='windows-1251')
        result = hashlib.sha256(encoded)
        self.hash256string = result.hexdigest()
    
    def hash256str_(self, name_file):
        # logging_report(f'Hashing 256 file check XML {name_file}')
        result = hashlib.sha256(name_file)
        self.hash256string = result.hexdigest()

    def _initchannel(self):
        logging_report(f'Init channel DFS server - {settings.SERVER}:{settings.PORT}')
        channel = grpc.secure_channel(f'{settings.SERVER}:{settings.PORT}', grpc.ssl_channel_credentials())
        return vbkrpc.ChkIncomeServiceStub(channel)
        
    def initdfs(self):
        self.error = True
        self.discription_error = ''
        logging_report(f'Status ini PB {self.pIface}')
        logging_report(f'KEYPASS {self.pathkey}')
        if len(self.password) == 0:
            self.discription_error = f'KEY_PASS Empty password key sign'
            logging_report(self.discription_error, 'excp')
            return self.error

        try:
            EU.EULoad()
            logging_report(f'EULoad -> завантаженa OK ')
        except Exception as e:
            self.discription_error = f'Load Dll EULoad failed {e}'
            logging.exception(self.discription_error, 'excp')
            return self.error
        self.pIface = EU.EUGetInterface()
        try:
            self.pIface.Initialize()
        except Exception as e:
            EU.EUUnload()
            self.discription_error = f'Initialize failed {e}'
            logging.exception(self.discription_error, 'excp')
            return self.error

        self.pathkey = str.strip(self.pathkey).encode()
        self.password = str.strip(self.password).encode()
        try:
            logging_report(f'try pIface reset privatkey')
            self.pIface.ResetPrivateKey()
            self.pIface.ReadPrivateKeyFile(self.pathkey, self.password, None)
        except Exception as e:
            self.pIface.Finalize()
            EU.EUUnload()
            self.discription_error = f'Key sign reading failed {e}'
            logging.exception(self.discription_error, 'excp')
            return self.error
        logging_report(f'read-pryvat-key and init PRRO')
        # self.error = False
        return False  # False it is all good

    def sign_check(self, b_chek_xml):
        lSign = []

        try:
            self.pIface.SignDataInternal(True, b_chek_xml, len(b_chek_xml), None, lSign)
        except Exception as e:
            self.pIface.Finalize()
            EU.EUUnload()  #
            self.discription_error = f'ID:-1  Status:SignXML failed   Error_message:SignXML failed'
            logging_report(f'SignXML failed {e}', 'excp')
            return self.error
        return lSign[0]

    def unsign_check(self, sign_data):
        lSign=[]
        try:
            self.pIface.GetDataFromSignedData(None, sign_data, len(sign_data), lSign)
            logging_report(f'Data From Signed Data  {lSign[0]}')
            return lSign[0].decode("utf-8", "replace")
        except Exception as e:
            logging_report(f'Error Get Data From Signed Data {e}', 'excp')
        # self.res_id = lSign[0].decode("utf-8", "replace")
        return ''

    def checkdfs(self, data_time, check_xml, localnumber, checktype, last_hash_check, edrpo, id_prro, id_cancel=''):
        self.res_id = 0
        self.res_status = ''
        self.discription_error = ''
        self.error = True
        strtime = data_time  # .encode('utf-8')
        logging_report(f'str_time {strtime} id_cancel={id_cancel}')
        check_xml = create_check_prro(checktype, self.fnrro, edrpo, id_prro, date_check=strtime,
                                      last_hash_check=last_hash_check, check_xml=check_xml)
        self.hash256str_(check_xml)
        check_sign = self.sign_check(check_xml)
        stub = self._initchannel()
        res = []
        _count = 5
        if checktype == '0':
            checktype = '1'
        while _count > 0:
            try:
                res = stub.sendChkV2(vbk.Check(rro_fn=str(self.fnrro),
                                               date_time=int(strtime),
                                               check_sign=check_sign,
                                               local_number=int(localnumber),
                                               check_type=int(checktype),
                                               id_cancel=id_cancel))
                self.error = res.status
                break
            except Exception as e:
                logging_report(f'DFS infoRro failed count - {_count} {e}', 'excp')
                self.error = True
            time.sleep(5)
            _count -= 1
        logging_report(self.error)
        if _count == 0:
            self.discription_error = f'Error fiscalization check in PRRO, repeat later'
            logging_report(self.discription_error, 'excp')
            return self.error

        logging_report(f'DFS id -> {res.id }')    
        logging_report(f'DFS status -> {res.status }')    
        logging_report(f'DFS error_message -> {res.error_message}')
        lSign = ''
        if len(res.id_sign) > 0:
            lSign = self.unsign_check(res.id_sign)
        self.res_id = res.id
        if len(res.data_sign) > 0:
            lSign = self.unsign_check(res.data_sign)
        self.res_data_sign = lSign
        self.res_status = res.status
        self.discription_error = res.error_message.strip()
      
        return self.error   # f'ID:{res.id}  Status:{res.status}'

    def inforrodfs(self):
        self.res_status = ''
        self.discription_error = ''
        self.res_status_smena = 0
        b_fnrro = str.encode(self.fnrro)
        self.error = True
        logging_report(f'GET Status PRRO FNPRRO -> {b_fnrro}')
        check_sign = self.sign_check(b_fnrro)
        stub = self._initchannel()
        try:
            res = stub.infoRro(vbk.CheckRequest(rro_fn_sign=check_sign))
        
        except Exception as e:
            self.discription_error = f'Error get info PRRO  {e}'
            logging_report(self.discription_error, 'excp')
            return self.error
        try:
            logging_report(f'DFS Статус відповіді {res.status}')
            logging_report(f'DFS Статус ПРРО {res.status_rro}')
            logging_report(f'DFS Статус зміни {res.open_shift}')
            logging_report(f'DFS Стан ПРРО {res.online}')
            logging_report(f'DFS Останній підписант {res.last_signer}')
            logging_report(f'DFS Назва {res.name}')
            logging_report(f'DFS Назва ТО {res.name_to}')
            logging_report(f'DFS Адреса ТО {res.addr}')
            logging_report(f'DFS Платник єдиного податку {res.single_tax}')
       
            logging_report(f'DFS Дозволено роботу в офлайн режимі {res.offline_allowed }')
            logging_report(f'DFS Додаткова кількість офлайн номерів {res.add_num}')
            logging_report(f'DFS Податковий номер платника ПДВ {res.pn}')
            for resoper in res.operators:
                logging_report(f'DFS Касири serial  {resoper.serial }')
                logging_report(f'DFS Касири status {resoper.status}')
                logging_report(f'DFS Касири senior {resoper.senior}')
                logging_report(f'DFS Касири isname {resoper.isname}')
            logging_report(f'DFS Податковий номер платника {res.tins}')
            logging_report(f'DFS Локальний номер каси {res.lnum}')
            logging_report(f'DFS Назва платника {res.name_pay}')
        except Exception as e:
            self.discription_error = f'DFS Exeption {e}'
            logging.exception(self.discription_error, 'excp')
            return self.error

        self.res_status = res.status
        self.discription_error = ''
        if res.open_shift:
            self.res_status_smena = 1
        else:
            self.res_status_smena = 0
        self.discription_error = f'Статус відповіді:{res.status}    Статус ПРРО:{res.status_rro} ' \
                                 f' Статус зміни:{ res.open_shift}  Стан ПРРО:{res.online}'
        return False

    def statusrro(self):
        b_fnrro = str.encode(self.fnrro)
        self.error = True
        logging_report(f'GET Status PRRO FNPRRO -> {b_fnrro}')
        check_sign = self.sign_check(b_fnrro)
        stub = self._initchannel()
        try:
            res = stub.statusRro(vbk.CheckRequest(rro_fn_sign=check_sign))

        except Exception as e:
            self.discription_error = f'Error get info PRRO  {e}'
            logging_report(self.discription_error, 'excp')
            return self.error
        try:
            logging_report(f'Status PRRO :{res.status} open_shift:{res.open_shift}  online:{res.online}')
        except:
            logging_report(f'ERRORe Statusrro','excp')
        if res.open_shift:
            self.res_status_smena = 1
        else:
            self.res_status_smena = 0
        self.res_status = res.status
        self.discription_error = f'Статус відповіді:{res.status}   ' \
                             f' Статус зміни:{res.open_shift}  Стан ПРРО:{res.online}'
        return False

    def serial_ports(self):
        if sys.platform.startswith('win'):
            ports = ['COM%s' % (i + 1) for i in range(2, 25)]
        else:
            raise EnvironmentError('Unsupported platform')
        result = []
        for port in ports:
             try:
                logging_report(f'{port}')
                s = serial.Serial(port, 115200)
                logging_report(f's  = {s}')
                s.close()
                result.append(port)
                logging_report(f'{result}')
             except (OSError, serial.SerialException):
                logging_report(f'ERROR FIND COM {serial.SerialException} ', 'excp')
        return result

    def initCom(self, comPort, comSpeed=115200, timeout=60, bytesize=8, parity='N', stopbits=1):
        self.ser_com_port = serial.Serial(baudrate = comSpeed, port = comPort, bytesize = bytesize, parity = parity,
                                          stopbits = stopbits, timeout = None)
        self.error = False
        logging_report(f'INIT COM port {self.ser_com_port} ')
        return self.error

    def test_eth_conn(self, ip_adress, port=2000):
        self.ip_adress = ip_adress
        self.port = port
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.connect((self.ip_adress, self.port))
            s.sendall(b'{"method":"PingDevice","step":0}')
            data = s.recv(1024)
            terminator = data.index(b'\x00')
            data = data[:terminator]
            try:
                ans = json.loads(data)
            except ValueError as e:
                logging_report(f'ERROR outputEth load json {e}', 'excp')
            
            self.error = str(ans['error'])
        logging_report(f'TEST connect {ip_adress} port {port} ')
        return False  # ==0

    def inputEth(self, InText):
        command =  str.encode('{ "method": "Purchase", "step": 0, "params": { "amount": "'+str(InText)+'", "discount": "", "merchantId": "0" } }')#+(b'\x00')
        self.error = 'true'
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            logging_report(f'Send {self.ip_adress} port {self.port} input {InText} data {command}')
            s.connect((self.ip_adress, self.port))
            s.sendall(command)
            data=s.recv(2048)
            try:
                terminator = data.index(b'\x00')
                data=data[:terminator].decode('utf8')
            except ValueError as e:
                logging_report(f'ERROR outputEth find terminator {e}')
            try:
                ans = json.loads(data)
            except ValueError as e:
                logging.exception(f'ERROR outputEth load json {e}')
        self.error = str(ans['error'])
        #logging_report(f'Geting data {ans['errorDescription'].decode('utf8')}')
        self.errorDescription =str(ans['errorDescription'])
        params = ans['params']  
        self.responseCode = params.get('responseCode', '0')
        self.date = str(params.get('date', ' '))
        self.time = str(params.get('time', ' '))
        self.pan = str(params.get('pan', ' ') )
        self.rrn = str(params.get('rrn', ' ') )
        self.receipt = str(params.get('receipt', ' '))
        logging_report(f'STOP outputEth {self.error}:{self.errorDescription}:{self.responseCode}:{self.date}:{self.time}')
        if self.receipt != ' ':
            logging_report(f'Receipt {self.receipt}')
        
        return str(self.error)
        

    def openCom(self):
        logging_report(f'OPEN COM port {self.ser_com_port} ')
        self.error = True
        self.discription_error = ''
        if not self.ser_com_port.is_open:
            try:
                self.ser_com_port.open()
                self.error = False
            except Exception as e:
                self.discription_error = f'ERROR open com {self.ser_com_port.port} port {e}'
                logging_report(self.discription_error, 'excp')
        return self.error
   
    def inputCom(self, InText):
        command = str.encode('{ "method": "Purchase", "step": 0, "params": { "amount": "'+str(InText)+'", "discount": "", "merchantId": "0" } }')#+(b'\x00')
        self.error = True
        if not self.ser_com_port.is_open:
            logging_report(f'COM {self.ser_com_port.port} is close. try opened')
            if self.openCom():  # True if error
                return self.error
        logging_report(f'START SEND COM {command}')
 
        self.ser_com_port.write(command)
        logging_report(f'STOP SEND')
        return self.outputCom()
    
    def outputCom(self):
        ans = ''
        params =[]
        sout = b''
        logging_report(f'START get data')

        self.ser_com_port.timeout = 60
        while True:
                c = self.ser_com_port.read(self.ser_com_port.inWaiting())
                #logging_report(f'!!!data fffff-{c}')
                sout += c
                if c.find(b'\x00') > -1:
                    break

        sout = sout.decode('utf8')
        terminator = sout.index('\x00')
        sout = sout[:terminator]

        logging_report(f'Reading  data OK')

        try:
            ans = json.loads(sout)
        except ValueError as e:
            self.discription_error = f'ERROR outputcom load json {e}'
            logging_report(self.discription_error, 'excp')
        if str(ans['error']) == 'true':
            self.error = True
        else:
            self.error = False

        self.errorDescription =str(ans['errorDescription'])
        params = ans['params']  
        self.responseCode = params.get('responseCode', '0')
        date = str(params.get('date', ' '))
        time = str(params.get('time', ' '))
        # pan = str(params.get('pan', ' '))
        # rrn = str(params.get('rrn', ' '))
        self.receipt = str(params.get('receipt', ' '))
        logging_report(f'STOP output com {self.error}:{self.errorDescription}:{self.responseCode}:{date}:{time}')
        #if self.receipt != ' ':
        logging_report(f'Receipt {self.receipt}')
        return self.error

    def testConn(self): #!!!!!
        logging_report(f'try test COM ports')
        logging_report(f'try open {self.ser_com_port.port}')
        self.error = True
        try: 
            if self.ser_com_port.is_open:
                logging_report(f'try send in port {self.ser_com_port.port}')
                self.ser_com_port.write(b'{"method":"PingDevice","step":0}')
                logging_report(f'try get data from port {self.ser_com_port.port}')
                return self.outputCom()
        except Exception as e:
            self.discription_error = f'ERROR testcom {e}'
            logging_report(self.discription_error, 'excp')
            return self.error
        

if __name__ == '__main__':
    param = sys.argv[1:]
    if '--register' in param or '--unregister' in param:
        import win32com.server.register
        win32com.server.register.UseCommandLine(TerminalCom, debug=debugging)
    #else:
        # start the server.
        #from win32com.server import localserver
        #localserver.serve(['{0998C9DA-DED7-4B04-B937-B37671831CCC}'])
