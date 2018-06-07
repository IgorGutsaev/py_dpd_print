import sys
from datetime import datetime, timedelta
import pyodbc


class filuet_log:
    ''' class to save logs in db '''

    def __init__(self, app_name='', history_days=90):
        self.__app_name = app_name
        # sets the number of days to keep the logs history
        self.__history_days = history_days
        self.__conn_str = 'DRIVER={SQL Server};SERVER=RU-LOB-SQL01;DATABASE=FiluetWH;UID=exuser;PWD=good4you'

    def INFO(self, log_text='', order_no=''):
        f_name = self.__func_name()
        self.__add_log(log_text=log_text, order_no=order_no,
                       f_name=f_name, log_level='INFO')

    def DEBUG(self, log_text='', order_no=''):
        f_name = self.__func_name()
        self.__add_log(log_text=log_text, order_no=order_no,
                       f_name=f_name, log_level='DEBUG')

    def ERROR(self, log_text='', order_no=''):
        f_name = self.__func_name()
        self.__add_log(log_text=log_text, order_no=order_no,
                       f_name=f_name, log_level='ERROR')

    def clean_logs_db(self):
        min_date = datetime.now() + timedelta(days=-self.__history_days)
        sql = "DELETE FROM [Log] WHERE ([Log].log_datetime < '{0:%Y-%m-%d}') AND ([Log].log_app='{1}')".format(
            min_date, self.__app_name)
        self.__execute_sql(sql)

    def __func_name(self):
        ''' get name of function called logging '''
        frame = sys._getframe(2)
        res = ''
        while (frame):
            f_name = frame.f_code.co_name
            if f_name == '<module>':
                break
            res = '.' + f_name + '()' + res
            frame = frame.f_back
        return res[1:]

    def __execute_sql(self, sql):
        ''' execute sql command on db '''
        # print(sql)
        try:
            with pyodbc.connect(self.__conn_str) as conn:
                cursor = conn.cursor()
                cursor.execute(sql)
                cursor.commit()
        except pyodbc.DatabaseError as ex:
            print(ex)
            pass

    def __add_log(self, log_text='', order_no='', f_name='', log_level='INFO'):
        if(len(log_text)>10000):
            log_text='Message too long. Truncated.\n'+log_text[:10000]
        log_text=log_text.replace("'","''")
        sql = '''INSERT INTO [dbo].[Log]
            ([log_datetime]
            ,[log_text]
            ,[log_order_no]
            ,[log_func_name]
            ,[log_app]
            ,[log_level])
        VALUES
            (GETDATE(),N'{0}',N'{1}',N'{2}',N'{3}',N'{4}')'''.format(log_text, order_no, f_name, self.__app_name, log_level)
        self.__execute_sql(sql)
        self.clean_logs_db()


def start():
    def start1():
        log.DEBUG('test log 2', '123456789')
        

    log.INFO(log_text='test log', order_no='test order')
    start1()


if __name__ == '__main__':
    log = filuet_log(app_name='test app')
    start()
