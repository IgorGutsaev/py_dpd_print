from zeep import Client
from zeep.plugins import HistoryPlugin
import pyodbc
from lxml import etree
from dpd_settings import settings
import win32api
import os
import srv
import logging
from logging.handlers import RotatingFileHandler
import datetime



# declare global variables
history = HistoryPlugin()

# set 1
client1 = Client(wsdl=settings.wsdl1, plugins=[history])
factory1 = client1.type_factory('ns0')
auth1 = factory1.auth(clientNumber=settings.clientNumber,
                      clientKey=settings.clientKey)
# set 2
client2 = Client(wsdl=settings.wsdl2, plugins=[history])
factory2 = client2.type_factory('ns0')
auth2 = factory2.auth(clientNumber=settings.clientNumber,
                      clientKey=settings.clientKey)


def init_logger():
    ''' init logger with unicode support '''
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    handler = RotatingFileHandler(
        filename='dpd_print.log', maxBytes=5000000, backupCount=1, encoding='utf-8')
    formatter = logging.Formatter(
        '%(asctime)s %(levelname)s\t%(funcName)s() <%(lineno)s> %(message)s')
    handler.setFormatter(formatter)
    root_logger.addHandler(handler)


def process_start():
    logging.debug('Start process!')
    # get new order for print
    order_code = get_new_order()

    if order_code:
        # get order details
        details = get_order_details(order_code)
        logging.info('Order Packed: {}'.format(order_code))
        logging.debug('details:\n' + str(details))

        # create dict for order
        if len(details) > 0:
            order = init_order(order_code, details)
            logging.debug('order:\n' + str(order))
            if order:
                try:
                    # register order in dpd
                    order_status = dict()
                    order_status = createOrder(**order)
                    order_status['pack_date'] = details['ExecuteDate']
                    order_status['ship_date'] = order['datePickup']
                    update_order_status(order_code, order_status)
                    logging.debug('order_status:\n' + str(order_status))
                    if 'status' in order_status and order_status['status'] == 'OK':
                        # get dpd label and save it
                        resp = createLabelFile(order_status['orderNum'], order['cargoNumPack'])

                        if 'file' in resp:
                            print_label(resp, order_code, order_status['orderNum'])
                        else:
                            raise 'No PDF data found in response. Order {0}'.format(order_code)
                except Exception as ex:
                    logging.error('{0}'.format(ex))


def createLabelFile(dpd_order_num, label_qty):
    ''' get label for print in PDF format, use set 2 '''
    # createLabelFile(getLabelFile: ns0:dpdGetLabelFile) -> return: ns0:dpdOrderLabelsFile
    # ns0:dpdGetLabelFile(auth: ns0:auth, fileFormat: ns0:fileFormat, pageSize: ns0:pageSize, order: ns0:orderParam[])
    # ns0:orderParam(orderNum: xsd:string, parcelsNumber: xsd:int)
    # ns0:auth(clientNumber: xsd:long, clientKey: xsd:string)
    # ns0:createLabelFileResponse(return: ns0:dpdOrderLabelsFile)
    try:
        order = factory2.orderParam(
            orderNum=dpd_order_num, parcelsNumber=label_qty)
        dpdGetLabelFile = factory2.dpdGetLabelFile(
            auth=auth2, fileFormat='PDF', pageSize='A6', order=[order])   
        response = client2.service.createLabelFile(
            getLabelFile=dpdGetLabelFile)
        return response
    except Exception as ex:
        logging.error('{0}'.format(ex))


def print_label(response, order_code, dpd_order_num):
    try:
        pdf_data = response['file']
        path_file = os.path.join(os.getcwd(), 'Labels')
        os.makedirs(path_file, exist_ok=True)
        file_name = os.path.join(
            path_file, '{0}-{1}.pdf'.format(order_code, dpd_order_num))
        logging.info('Saving dpd label {0}'.format(file_name))
        with open(file_name, 'wb') as f:
            f.write(pdf_data)
        if 'status' in response['order'][0]:
            win32api.ShellExecute(0, 'open', settings.print_app, '{0} {1}'.format(
                settings.print_command, file_name), '', 1)
            logging.info('File printed: {0}'.format(file_name))
            sql = "UPDATE [dbo].[dpd_orders] SET [printed] = 1 WHERE order_code = '{0}'".format(
                order_code)
            execute_sql(sql)
            delete_old_files(path_file)
    except Exception as ex:
        logging.error('{0}'.format(ex))


def delete_old_files(path_file):
    max_date = datetime.datetime.now() - datetime.timedelta(days=0, seconds=0, microseconds=0,
                                                            milliseconds=0, minutes=10, hours=0, weeks=0)
    files = os.listdir(path_file)
    for _t1 in files:
        path_to_file = os.path.join(path_file, _t1)
        if datetime.datetime.fromtimestamp(os.path.getctime(path_to_file)) < max_date:
            try:
                os.remove(path_to_file)
            except Exception as ex:
                logging.error('{0}'.format(ex))


def update_order_status(order_code, order_status):
    if 'status' in order_status:
        sql = '''EXEC [dpd].[update_order_status]
                    @order_code = N'{0}',
                    @status = N'{1}',
                    @error_msg = N'{2}',
                    @dpd_id = N'{3}',
                    @pack_date = N'{4}',
                    @ship_date = N'{5}';'''.format(order_code,
                                                   order_status['status'],
                                                   order_status['errorMessage'],
                                                   order_status['orderNum'],
                                                   order_status['pack_date'].isoformat(
                                                       timespec='seconds'),
                                                   order_status['ship_date'])
    else:
        sql = '''EXEC [dpd].[update_order_status]
            @order_code = N'{0}',
            @status = N'{1}',
            @error_msg = N'{2}',
            @pack_date = N'{3}',
            @ship_date = N'{4}';'''.format(order_code, 'OrderError', 'Error create order', order_status['pack_date'].isoformat(timespec='seconds'), order_status['ship_date'])
    execute_sql(sql)


def init_order(order_code, details):
    ''' create order datatype from details'''
    ''' 
    structure of received data:
    {'Type': 23, 'InputDate': datetime.datetime(2018, 5, 7, 3, 1, 26), 'ExecuteDate': datetime.datetime(2018, 5, 8, 7, 12, 9, 260000), 'Status': 'Упакован', 'StatusCode': '60', 'BoxQty': 1,
    'TotalWeight': Decimal('0.820000'), 'ReceiverName': 'АРСЕНЬЕВА А.Б.', 'TerminalNo': 'MQ23', 'Phone': '89639262357', 'AddressLV': 'Д. 28А КОМСОМОЛЬСКАЯ УЛ.', 'AreaLV': 'МОСКОВСКАЯ ОБЛАСТЬ',
    'CityLV': 'НОГИНСК', 'PostIndexLV': '142400', 'Region': None, 'Region_type': None, 'District': None, 'District_type': None, 'City': None, 'City_type': None, 'Place': None, 'Place_type': None,
    'Street': None, 'Street_type': None, 'House': None, 'House_type': None, 'Building': None, 'Building_type': None, 'Structure': None, 'Structure_type': None,
    'Flat': None, 'Flat_type': None, 'Zip': None, 'Pretty': None, 'Precision': None, 'Recall': None, 'Warnings': None, 'FIAS_City': None}

    ns0:address( house: xsd:string, houseKorpus: xsd:string, str: xsd:string, vlad: xsd:string, extraInfo: xsd:string, office: xsd:string, flat: xsd:string, workTimeFrom: xsd:string, workTimeTo: xsd:string, dinnerTimeFrom: xsd:string, dinnerTimeTo: xsd:string, contactFio: xsd:string, contactPhone: xsd:string, contactEmail: xsd:string, instructions: xsd:string, needPass: xsd:boolean)
     
    '''
    try:
        order = dict()

        # check if terminal number exist
        if details['TerminalNo']:
            # terminal
            address = factory1.address(terminalCode=details['TerminalNo'],
                                       name=details['ReceiverName'],
                                       contactFio=details['ReceiverName'],
                                       contactPhone=details['Phone'])
            order['serviceVariant'] = 'ДТ'
        else:
            # home address
            address = factory1.address(flat=details['Flat'] if details['Flat'] else '',
                                       str=details['Structure_type'] + ' ' +
                                       details['Structure'] if details['Structure'] else '',
                                       houseKorpus=details['Building'] if details['Building'] else '',
                                       house=details['House'] if details['House'] else '',
                                       streetAbbr=details['Street_type'],
                                       street=details['Street'],
                                       city=details['City_type'] + ' ' +
                                       details['City'] if details['City'] else details['Place_type'] +
                                       ' ' + details['Place'],
                                       region=details['Region'] +
                                       ' ' + details['Region_type'],
                                       index=details['Zip'],
                                       countryName='Россия',
                                       name=details['ReceiverName'],
                                       contactFio=details['ReceiverName'],
                                       contactPhone=details['Phone'])
            order['serviceVariant'] = 'ДД'

        order['datePickup'] = get_pick_date()
        order['orderNumberInternal'] = order_code
        order['cargoNumPack'] = details['BoxQty']
        order['cargoWeight'] = details['TotalWeight']
        order['receiverAddress'] = address

        return order
    except Exception as ex:
        logging.error('{0}'.format(ex))
        return None


def get_pick_date():
    sql = "SELECT VarValue FROM dbo.Variables WHERE VarName = 'DPDPickDate';"
    table = execute_sql_fetch(sql)
    if len(table) > 0:
        return table[0]['VarValue']
    else:
        return None


def get_order_details(order_code):
    sql = "EXEC	[dpd].[GetOrderDetails] @OrderCode = N'{0}'".format(order_code)
    table = execute_sql_fetch(sql)
    if len(table) > 0:
        return table[0]
    else:
        return None


def get_new_order():
    sql = '''SELECT TOP (1) order_code
            FROM dpd_orders
            WHERE status = 'NEW'
            ORDER BY pack_date; '''
    table = execute_sql_fetch(sql)
    if len(table) > 0:
        return table[0]['order_code']
    else:
        return ''


def dpdGetOrderStatus(order_no):  # not used
    ''' get status of submited order, use set 1 '''
    # ns0:getOrderStatus(orderStatus: ns0:dpdGetOrderStatus)
    # ns0:dpdGetOrderStatus(auth: ns0:auth, order: ns0:internalOrderNumber[])
    # ns0:internalOrderNumber(orderNumberInternal: xsd:string, datePickup: xsd:date)

    order = factory1.internalOrderNumber(orderNumberInternal=order_no)
    dpdGetOrderStatus = factory1.dpdGetOrderStatus(auth=auth1, order=[order])
    try:
        response = client1.service.getOrderStatus(
            orderStatus=dpdGetOrderStatus)
        logging.info('Request\n' +
                     etree.tounicode(history.last_sent['envelope'], pretty_print=True))
        logging.info('Response\n' +
                     etree.tounicode(history.last_received['envelope'], pretty_print=True))
    except Exception as ex:
        print('{0}'.format(ex))


def createOrder(datePickup,
                orderNumberInternal,
                cargoNumPack,
                cargoWeight,
                receiverAddress,
                serviceVariant):
    ''' register new order in dpd server, use set 1 '''
    '''
    ns0:createOrder(orders: ns0:dpdOrdersData)
    ns0:dpdOrdersData(auth: ns0:auth, header: ns0:header, order: ns0:order[])
    ns0:header(datePickup: xsd:date, payer: xsd:long, senderAddress: ns0:address, pickupTimePeriod: xsd:string, regularNum: xsd:string)
    ns0:order(orderNumberInternal: xsd:string, serviceCode: xsd:string, serviceVariant: xsd:string, cargoNumPack: xsd:int, cargoWeight: xsd:double, cargoVolume: xsd:double, cargoRegistered: xsd:boolean, cargoValue: xsd:double, cargoCategory: xsd:string, deliveryTimePeriod: xsd:string, paymentType: xsd:string, extraParam: ns0:parameter[], dataInt: ns0:dataInternational, receiverAddress: ns0:address, returnAddress: ns0:address, extraService: ns0:extraService[], parcel: ns0:parcel[], unitLoad: ns0:unitLoad[])
    ns0:address(code: xsd:string, name: xsd:string, terminalCode: xsd:string, addressString: xsd:string, countryName: xsd:string, index: xsd:string, region: xsd:string, city: xsd:string, street: xsd:string, streetAbbr: xsd:string, house: xsd:string, houseKorpus: xsd:string, str: xsd:string, vlad: xsd:string, extraInfo: xsd:string, office: xsd:string, flat: xsd:string, workTimeFrom: xsd:string, workTimeTo: xsd:string, dinnerTimeFrom: xsd:string, dinnerTimeTo: xsd:string, contactFio: xsd:string, contactPhone: xsd:string, contactEmail: xsd:string, instructions: xsd:string, needPass: xsd:boolean)
    ns0:dpdOrderStatus(orderNumberInternal: xsd:string, orderNum: xsd:string, status: xsd:string, errorMessage: xsd:string)
    '''
    try:
        # header
        senderAddress = factory1.address(code=settings.senderAddress_code)
        header = factory1.header(datePickup=datePickup,
                                 pickupTimePeriod=settings.pickupTimePeriod, senderAddress=senderAddress)

        # order
        # orderNumberInternal = orderNumberInternal  # 'Test01'
        # cargoNumPack = cargoNumPack  # количество мест
        # cargoWeight = cargoWeight  # Вес отправки, кг
        # receiverAddress = receiverAddress
        # serviceVariant = serviceVariant

        serviceCode = settings.serviceCode  # settings.serviceCode
        cargoRegistered = settings.cargoRegistered  # Ценный груз
        cargoValue = settings.cargoValue  # Сумма объявленной ценности
        cargoCategory = settings.cargoCategory  # Содержимое отправки

        order = factory1.order(orderNumberInternal=orderNumberInternal,
                               serviceCode=serviceCode,
                               serviceVariant=serviceVariant,
                               cargoNumPack=cargoNumPack,
                               cargoWeight=cargoWeight,
                               cargoRegistered=cargoRegistered,
                               cargoValue=cargoValue,
                               cargoCategory=cargoCategory,
                               receiverAddress=receiverAddress)
        orders = factory1.dpdOrdersData(
            auth=auth1, header=header, order=[order])

        response = client1.service.createOrder(orders=orders)
        logging.info('Request\n' +
                     etree.tounicode(history.last_sent['envelope'], pretty_print=True))
        logging.info('Response\n' +
                     etree.tounicode(history.last_received['envelope'], pretty_print=True))
        return response[0]
    except Exception as ex:
        logging.error('{0}'.format(ex))
        return None


def execute_sql(sql):
    ''' execute sql command on db '''
    logging.debug(sql)
    try:
        with pyodbc.connect(settings.conn_str) as conn:
            cursor = conn.cursor()
            cursor.execute(sql)
            cursor.commit()
            return 1
    except Exception as ex:
        logging.error('{0}'.format(ex))
        return None


def execute_sql_fetch(sql):
    ''' execute sql command on db '''
    logging.debug(sql)
    try:
        with pyodbc.connect(settings.conn_str) as conn:
            cursor = conn.cursor()
            cursor.execute(sql)
            table = [dict(zip([column[0] for column in cursor.description], row))
                     for row in cursor.fetchall()]
            # result = cursor.fetchall()
            cursor.commit()
            return table
    except Exception as ex:
        logging.error('{0}'.format(ex))
        return None


if __name__ == '__main__':
    init_logger()
    # process_start()

    srv = srv.srv('DPD_PRINT')
    while True:
        srv.next_run()
        process_start()
        srv.upd_last_run()

        # print(srv.last_run)
