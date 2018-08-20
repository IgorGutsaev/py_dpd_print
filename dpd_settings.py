class settings:
    # declare global variables
    conn_str = 'DRIVER={SQL Server};SERVER=RU-LOB-SQL01;DATABASE=FiluetWH;UID=exuser;PWD=good4you'
    clientNumber = 1001038335
    clientKey = '1F9A998B47EFA546C081F99C99A1E0F57E9F103F'

    # set 1
    # wsdl1 = 'http://wstest.dpd.ru/services/order2?wsdl' # test env
    wsdl1 = 'http://ws.dpd.ru/services/order2?wsdl' # prod env

    # set 2
    # wsdl2 = 'http://wstest.dpd.ru/services/label-print?wsdl' # test env
    wsdl2 = 'http://ws.dpd.ru/services/label-print?wsdl' # prod env


    # createOrder
    senderAddress_code='FILUET_SVO'
    pickupTimePeriod = '9-18'
    serviceCode = 'ECN'
    ''' BZP	DPD 18:00
        ECN	DPD ECONOMY
        CUR	DPD CLASSIC domestic
        NDY	DPD EXPRESS
        CSM	DPD Online Express
        PCL	DPD Online Classic
        DPI	DPD CLASSIC international IMPORT
        DPE	DPD CLASSIC international EXPORT
        MAX	DPD MAX domestic
        MXO	DPD Online Max '''

    serviceVariant = 'ДТ'
    ''' ДД – от двери отправителя до двери получателя;
        ДТ – от двери отправителя до терминала DPD;
        ТД – от терминала DPD до двери получателя;
        ТТ – от терминала DPD до терминала DPD. '''
    
    cargoRegistered = 0  # Ценный груз
    cargoValue = 0  # Сумма объявленной ценности
    cargoCategory = 'Косметика'  # Содержимое отправки

    # print label
    print_app = 'C:\\Program Files\\Tracker Software\\PDF Editor\\PDFXEdit.exe'
    print_command = '/print:showui=no;printer="PZ01"'
    

    




