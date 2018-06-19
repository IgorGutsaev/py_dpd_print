import pyodbc
from datetime import datetime, timedelta, time
from filuet_log import filuet_log
# import time

class srv:
    ''' class to get running schedule and assist'''

    def __init__(self, srv_name=''):
        self.srv_name = srv_name
        self.__conn_str = 'DRIVER={SQL Server};SERVER=RU-LOB-SQL01;DATABASE=FiluetWH;UID=exuser;PWD=good4you'
        self.log = filuet_log(app_name='SRV')
        sql='''SELECT s.id , s.srv_name , s.file_name , s.week_days , s.start_time , s.end_time , s.period_sec , s.last_run
                FROM dbo.Services AS s
                WHERE s.srv_name = N'{0}';'''.format(self.srv_name)
        tbl = self.__execute_sql(sql)
        if len(tbl)==1:
            self.id=tbl[0][0]
            self.file_name=tbl[0][2]
            self.week_days=tbl[0][3]
            self.start_time=datetime.strptime(tbl[0][4],'%H:%M:%S').time()
            self.end_time=datetime.strptime(tbl[0][5],'%H:%M:%S').time()
            self.period_sec=tbl[0][6]
            self.last_run=tbl[0][7]
            if self.last_run is None:
                self.last_run = datetime(2000,1,1,0,0,0)
    
    def upd_last_run(self):
        sql='''SET NOCOUNT ON UPDATE dbo.Services SET dbo.Services.last_run = GETDATE()
                WHERE dbo.Services.id = {0}; SELECT @@ROWCOUNT;'''.format(self.id)
        self.__execute_sql(sql)
        self.last_run=datetime.now()

    def next_run(self):
        now = datetime.now()
        if str(now.isoweekday()) in self.week_days:
            if self.start_time <= now.time() <= self.end_time:
                while (self.last_run+timedelta(seconds=self.period_sec))>datetime.now() :
                    # print(datetime.now())
                    pass
            elif now.time() > self.end_time:
                raise SystemExit

    def __execute_sql(self, sql):
        ''' execute sql command on db '''
        self.log.DEBUG(sql)
        try:
            with pyodbc.connect(self.__conn_str) as conn:
                cursor = conn.cursor()
                cursor.execute(sql)
                table = cursor.fetchall()
                cursor.commit()
                return table
        except pyodbc.DatabaseError as ex:
            self.log.ERROR(log_text='{0}'.format(ex))
            return None

if __name__ == '__main__':
    srv=srv('DPD_TERMINALS_LIMITS')
    n=1
    while n<5:
        srv.next_run()
        # time.sleep(5)
        srv.upd_last_run()
        print(srv.last_run)
        print(n)
        n+=1


