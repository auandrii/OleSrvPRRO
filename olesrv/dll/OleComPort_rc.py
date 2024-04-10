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
# from xml.etree import ElementTree
import EUSignCP as EUSignCP

PATH = pathlib.Path(__file__).parent.absolute()
VERSION = '1.1'

logging.basicConfig(filename=f'{PATH}/{settings.NAME_FILE_LOG}', format='%(asctime)s - %(message)s',
                    level=logging.INFO)

debugging = 1
if debugging:
    from win32com.server.dispatcher import DefaultDebugDispatcher

    useDispatcher = DefaultDebugDispatcher
else:
    useDispatcher = None


class TerminalCom:
    # _reg_clsctx_ = pythoncom.CLSCTX_LOCAL_SERVER
    # _reg_clsid_ = "{746890E4-EB54-478C-A60F-301B1D661A8F}"
    _reg_clsid_ = "{746890E4-EB54-478C-A60F-301B1D661A8F}"  # test
    _reg_progid_ = "OLE1c77Bank.TerminalCom"
    _reg_desc_ = "OLE1c77Bank.TerminalCom.V2"  # test
    _public_methods_ = ['get_version',
                        'toUTF8',
                        'hash256str',
                        'checkdfs',
                        'initdfs',
                        'inforrodfs',
                        'initCom',
                        'openCom',
                        'testConn',
                        'inputCom',
                        'outputCom',
                        'test_eth_conn',
                        'inputEth']
    _public_attrs_ = ['res_status_smena',
                      'hash256string',
                      'ukr',
                      'res_status',
                      'res_id',
                      'res_errore',
                      'password',
                      'pathkey',
                      'error',
                      'responseCode',
                      'errorDescription',
                      'pan',
                      'date',
                      'time',
                      'rrn',
                      'receipt',
                      'fnrro']

    # _readonly_attrs_ = []

    def __init__(self):
        self.error = 'true'
        self.errorDescription = ''
        self.noCalls = 0
        self.responseCode = 0
        self.date = ''
        self.time = ''
        self.pan = ''
        self.rrn = ''
        self.receipt = ''
        self.ip_adress = ""
        self.port = 2000
        self.fnrro = ''
        self.pIface = ''
        self.pathkey = ''
        self.password = ''
        self.fnrro = ''
        self.res_id = 0
        self.res_status = ''
        self.res_errore = ''
        self.ukr = ''
        self.hash256string = ''
        self.res_status_smena = 0
        self.text = ''
        self.ser_com_port = serial.Serial()

    def get_version(self):
        logging.info(f'Get version - {VERSION}')
        return VERSION

    def toUTF8(self, text):
        logging.info(f'TEXT in {text}')

    def hash256str(self, name_file):
        b_chek_xml = ''
        ftxt = open(f'{PATH}/{name_file}.xml', mode='r', encoding='windows-1251')
        for line in ftxt:
            b_chek_xml += line.strip('\r\n')

        encoded = b_chek_xml.encode(encoding='windows-1251')
        result = hashlib.sha256(encoded)
        self.hash256string = result.hexdigest()

    def initchannel(self):
        channel = grpc.secure_channel(f'{settings.SERVER}:{settings.PORT}', grpc.ssl_channel_credentials())
        return vbkrpc.ChkIncomeServiceStub(channel)

    def initdfs(self):
        logging.info(f'Status ini PB {self.pIface}')
        logging.info(f'KEYPASS {self.pathkey}')
        if len(self.password) == 0:
            logging.info(f'KEYPASS пустой пароль ')
        try:
            EUSignCP.EULoad()
            logging.info(f'EULoad -> завантаженa OK ')
        except:
            logging.exception(f'EULoad load failed ')
            return False
        self.pIface = EUSignCP.EUGetInterface()
        try:
            self.pIface.Initialize()
        except Exception as e:
            logging.exception(f'Initialize failed {e}')
            EUSignCP.EUUnload()
            return False

        self.pathkey = str.encode(str.strip(self.pathkey))
        logging.info(f'set-path')
        self.password = str.encode(str.strip(self.password))
        logging.info(f'set-pass')
        try:
            logging.info(f'try pIface reset privatkey')
            self.pIface.ResetPrivateKey()
            logging.info(f'Reset Pkey - OK')

            if not self.pIface.IsPrivateKeyReaded():
                logging.info(f'if not pfice')
                logging.info(f'set-pass{self.pathkey}-{len(self.password)}')
                self.pIface.ReadPrivateKeyFile(self.pathkey, self.password, None)
                logging.info(f'if not pface - OK')
        except Exception as e:
            logging.exception(f'Key reading failed {e}')
            self.pIface.Finalize()
            EUSignCP.EUUnload()
            return False
        logging.info(f'read-pryvat-key')
        if self.pIface.IsPrivateKeyReaded():
            logging.info(f'Key success read')
        else:
            logging.info(f'Key reading failed ')
            self.pIface.Finalize()
            EUSignCP.EUUnload()
            return False
        return True

    def checkdfs(self, data_time, chek_xml, localnumber, checktype):
        self.res_id = 0
        self.res_status = ''
        self.res_errore = ''
        b_chek_xml = ''
        ftxt = open(f'{PATH}/{chek_xml}.xml', mode='r', encoding='windows-1251')
        for line in ftxt:
            b_chek_xml += line.strip('\r\n')

        logging.info(f'checkdfs  XML in {b_chek_xml}')
        b_chek_xml = str.encode(b_chek_xml, encoding='windows-1251')
        lSign = []

        try:
            self.pIface.SignDataInternal(True, b_chek_xml, len(b_chek_xml), None, lSign)
        except Exception as e:
            logging.exception(f'SignXML failed {e}')
            self.pIface.Finalize()
            EUSignCP.EUUnload()
            return f'ID:-1  Status:SignXML failed   Error_message:SignXML failed'

        stub = self.initchannel()
        res = []
        strtime = data_time.encode('utf-8')
        logging.info(f'strtime {strtime}')
        _count = 5
        while _count > 0:
            f_error = False
            try:
                res = stub.sendChkV2(vbk.Check(rro_fn=str(self.fnrro), date_time=int(strtime),
                                               check_sign=lSign[0], local_number=int(localnumber),
                                               check_type=int(checktype)))
                break
            except Exception as e:
                logging.exception(f'DFS infoRro failed count - {_count} {e}')
                f_error = True
            time.sleep(5)
            _count -= 1
        if f_error:
            return f'ID:-1  Status:DFS infoRro failed   Error_message:DFS infoRro failed'
        logging.info(f'DFS id -> {res.id}')
        logging.info(f'DFS status -> {res.status}')
        logging.info(f'DFS error_message -> {res.error_message}')
        if len(res.id_sign) > 0:
            try:
                self.pIface.GetDataFromSignedData(None, res.id_sign, len(res.id_sign), lSign)
                logging.info(f'res id_sign {lSign[0]}')
            except Exception as e:
                logging.exception(f'No TRUE t.id_sign {e}')
        # self.res_id = lSign[0].decode("utf-8", "replace")
        self.res_id = res.id
        if len(res.data_sign) > 0:
            try:
                self.pIface.GetDataFromSignedData(None, res.data_sign, len(res.data_sign), lSign)
                logging.info(f't.data_sign {lSign[0]}')
            except Exception as e:
                logging.exception(f'No TRUE t.data_sign {e}')
        self.res_data_sign = lSign[0].decode("utf-8", "replace")
        self.res_status = res.status
        self.res_errore = res.error_message.strip()

        return f'ID:{res.id}  Status:{res.status}'

    def inforrodfs(self):
        self.res_id = ''
        self.res_status = ''
        self.res_errore = ''
        self.res_status_smena = 0
        b_fnrro = str.encode(self.fnrro)
        logging.info(f'Status RRO FNRRO -> {b_fnrro}')

        lSign = []
        if self.pIface.IsPrivateKeyReaded():
            logging.info(f'Key success read')
        else:
            logging.info(f'Key reading failed')
            self.pIface.Finalize()
            EUSignCP.EUUnload()
            return False
        try:
            self.pIface.SignDataInternal(True, b_fnrro, len(b_fnrro), None, lSign)
        except Exception as e:
            logging.exception(f'SignData failed {e}')
            self.pIface.Finalize()
            EUSignCP.EUUnload()
            return False
        res = []
        stub = self.initchannel()
        try:
            res = stub.infoRro(vbk.CheckRequest(rro_fn_sign=lSign[0]))

        except Exception as e:
            logging.exception(f'DFS infoRro failed {e}')

            return False
        try:
            logging.info(f'DFS Статус відповіді {res.status}')
            logging.info(f'DFS Статус ПРРО {res.status_rro}')
            logging.info(f'DFS Статус зміни {res.open_shift}')
            logging.info(f'DFS Стан ПРРО {res.online}')
            logging.info(f'DFS Останній підписант {res.last_signer}')
            logging.info(f'DFS Назва {res.name}')
            logging.info(f'DFS Назва ТО {res.name_to}')
            logging.info(f'DFS Адреса ТО {res.addr}')
            logging.info(f'DFS Платник єдиного податку {res.single_tax}')

            logging.info(f'DFS Дозволено роботу в офлайн режимі {res.offline_allowed}')
            logging.info(f'DFS Додаткова кількість офлайн номерів {res.add_num}')
            logging.info(f'DFS Податковий номер платника ПДВ {res.pn}')
            logging.info(f'DFS Касири serial  {res.operators[0].serial}')
            logging.info(f'DFS Касири status {res.operators[0].status}')
            logging.info(f'DFS Касири senior {res.operators[0].senior}')
            logging.info(f'DFS Касири isname {res.operators[0].isname}')
            logging.info(f'DFS Податковий номер платника {res.tins}')
            logging.info(f'DFS Локальний номер каси {res.lnum}')
            logging.info(f'DFS Назва платника {res.name_pay}')
        except Exception as e:
            logging.exception(f'DFS Exeption {e}')
        self.res_id = ''
        self.res_status = res.status
        self.res_errore = ''
        if res.open_shift:
            self.res_status_smena = 1
        else:
            self.res_status_smena = 0
        return f'Статус відповіді:{res.status}    Статус ПРРО:{res.status_rro}  Статус зміни:{res.open_shift} ' \
               f'Стан ПРРО:{res.online}'

    def serial_ports(self):
        if sys.platform.startswith('win'):
            ports = ['COM%s' % (i + 1) for i in range(2, 25)]
        else:
            raise EnvironmentError('Unsupported platform')
        result = []
        for port in ports:
            try:
                logging.info(f'{port}')
                s = serial.Serial(port, 115200)
                logging.info(f's  = {s}')
                s.close()
                result.append(port)
                logging.info(f'{result}')
            except (OSError, serial.SerialException):
                logging.exception(f'ERROR FIND COM {serial.SerialException} ')
        return result

    def initCom(self, comPort, comSpeed=115200, timeout=60, bytesize=8, parity='N', stopbits=1):

        self.ser_com_port.baudrate = comSpeed
        self.ser_com_port.port = comPort
        self.ser_com_port.bytesize = bytesize
        self.ser_com_port.parity = parity
        self.ser_com_port.stopbits = stopbits
        self.ser_com_port.timeout = None
        self.error = 'False'
        logging.info(f'INIT COM port {self.ser_com_port} ')
        return True

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
                logging.exception(f'ERROR outputEth load json {e}')

            self.error = str(ans['error'])
        logging.info(f'TEST connect {ip_adress} port {port} ')
        return True

    def inputEth(self, InText):
        command = str.encode('{ "method": "Purchase", "step": 0, "params": { "amount": "' + str(
            InText) + '", "discount": "", "merchantId": "0" } }')  # +(b'\x00')
        self.error = 'true'
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            logging.info(f'Send {self.ip_adress} port {self.port} input {InText} data {command}')
            s.connect((self.ip_adress, self.port))
            s.sendall(command)
            data = s.recv(2048)
            try:
                terminator = data.index(b'\x00')
                data = data[:terminator].decode('utf8')
            except ValueError as e:
                logging.info(f'ERROR outputEth find terminator {e}')
            try:
                ans = json.loads(data)
            except ValueError as e:
                logging.exception(f'ERROR outputEth load json {e}')
        self.error = str(ans['error'])
        # logging.info(f'Geting data {ans['errorDescription'].decode('utf8')}')
        self.errorDescription = str(ans['errorDescription'])
        params = ans['params']
        self.responseCode = params.get('responseCode', '0')
        self.date = str(params.get('date', ' '))
        self.time = str(params.get('time', ' '))
        self.pan = str(params.get('pan', ' '))
        self.rrn = str(params.get('rrn', ' '))
        self.receipt = str(params.get('receipt', ' '))
        logging.info(f'STOP outputEth {self.error}:{self.errorDescription}:{self.responseCode}:{self.date}:{self.time}')
        if self.receipt != ' ':
            logging.info(f'Receipt {self.receipt}')

        return str(self.error)

    def openCom(self):
        logging.info(f'OPEN COM port {self.ser_com_port} ')
        if not self.ser_com_port.is_open:
            try:
                self.ser_com_port.open()
                self.error = (str(self.ser_com_port.is_open)).capitalize()
            except Exception as e:
                logging.exception(f'ERROR open com port {e}')
                self.error = 'True'
        return str(self.ser_com_port.is_open)

    def inputCom(self, InText):
        command = str.encode('{ "method": "Purchase", "step": 0, "params": { "amount": "' + str(
            InText) + '", "discount": "", "merchantId": "0" } }')  # +(b'\x00')
        self.error = 'true'
        if not self.ser_com_port.is_open:
            logging.info(f'COM {self.ser_com_port.port} is close. try opened')
            if self.openCom() == 'false':
                return 0
        logging.info(f'START SEND COM {command}')

        self.ser_com_port.write(command)
        logging.info(f'STOP SEND')
        self.outputCom()
        return (str(self.error)).capitalize()

    def outputCom(self):
        ans = ''
        params = []
        sout = b''
        logging.info(f'START get data')
        if 1 == 0:
            try:
                sout = self.ser_com_port.read_until(b'\x00')
                logging.info(f'SOUT get data = {sout}')

                logging.info(f'DECODE get data')
                sout = sout.decode('utf8')
                logging.info(f'DECODE END get data')

            except serial.SerialException as e:
                logging.info(f'ERROR outputcom {e} ')
            logging.info(f'data {sout}')
        else:
            self.ser_com_port.timeout = 60
            while True:
                c = self.ser_com_port.read(self.ser_com_port.inWaiting())
                # logging.info(f'!!!data fffff-{c}')
                sout += c
                if c.find(b'\x00') > -1:
                    break
                # logging.info(f'!!!data {sout}-{c}')
            # logging.info(f'&&&&data {sout}-{c}')
            # logging.info(f'!!!data {sout}')
            # if sout.find(b'\x00'):
            #     logging.info(f'Finding x00 - {sout}')
            #     break
            sout = sout.decode('utf8')
        terminator = sout.index('\x00')
        sout = sout[:terminator]

        logging.info(f'Reading  data OK')

        try:
            ans = json.loads(sout)
        except ValueError as e:
            logging.exception(f'ERROR outputcom load json {e}')
        self.error = str(ans['error'])
        self.errorDescription = str(ans['errorDescription'])
        params = ans['params']
        self.responseCode = params.get('responseCode', '0')
        self.date = str(params.get('date', ' '))
        self.time = str(params.get('time', ' '))
        self.pan = str(params.get('pan', ' '))
        self.rrn = str(params.get('rrn', ' '))
        self.receipt = str(params.get('receipt', ' '))
        logging.info(f'STOP outputcom {self.error}:{self.errorDescription}:{self.responseCode}:{self.date}:{self.time}')
        if self.receipt != ' ':
            logging.info(f'Receipt {self.receipt}')

    def testConn(self):  # !!!!!
        params = []
        sout = b''
        logging.info(f'try ports')
        port = self.ser_com_port
        logging.info(f'try open {port}')
        try:
            if self.ser_com_port.is_open:
                logging.info(f' {port.is_open}')
                logging.info(f'try send {port}')
                self.ser_com_port.write(b'{"method":"PingDevice","step":0}')
                logging.info(f'try get {port}')
                self.outputCom()
                logging.info(f' self.error {self.error}')
                if self.error == 'False':
                    logging.info(f'Brack {self.error}')
        except Exception as e:
            logging.exception(f'ERROR testcom {e}')
        return str(self.error)


if __name__ == '__main__':
    param = sys.argv[1:]
    if '--register' in param or '--unregister' in param:
        import win32com.server.register

        win32com.server.register.UseCommandLine(TerminalCom, debug=debugging)
    # else:
    # start the server.
    # from win32com.server import localserver
    # localserver.serve(['{0998C9DA-DED7-4B04-B937-B37671831CCC}'])
