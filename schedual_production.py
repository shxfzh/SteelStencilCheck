import xlrd
import pymysql as sql

class ImportStep(object):
    def __init__(self,host='10.18.91.156', user='mds', password='mds@user2020MYSQL', database='ican_resource_monitor', port=3306):
        self.conn = sql.connect(host=host, user=user, password=password, database=database, port=port)
        try:
            self.conn
        except Exception as error:
            print('failed to connect to the database:{}'.format(database))
        self.cursor = self.conn.cursor()

    def importFile(self,path):
        try:
            data = xlrd.open_workbook(path)
            table = data.sheet_by_index(0)
            nrows = table.nrows
            for i in range(1, nrows):
                ret = self.converData(table.row_values(i))
                print(ret)
                sqlstr = 'insert into `t_schedual_production` (production_line,lot_num,board_code,level,\
                    schedual_qty_a,schedual_qty_b,steel_stencil_name) values {}' .format(ret)
                try:
                    self.cursor.execute(sqlstr)
                except:
                    self.conn.rollback()
            self.cursor.close()
            self.conn.commit()
            self.conn.close()
        except:
            print('error')

    def converData(self,rowData):
        column_used = {'线体': 0, "批次": 1, "单板代码": 3, "排序": 5, "计划完成a面": 16, "计划完成b面": 17, "钢网名称": 12}
        ret = [rowData[j] for j in column_used.values()]
        ret[0] = str(ret[0])
        ret[1] = str(ret[1])
        ret[2] = str(ret[2])
        ret[3] = int(ret[3])
        ret[4] = int(ret[4])
        ret[5] = int(ret[5])
        ret[6] = str(ret[6])
        return tuple(ret)

if __name__ == "__main__":
    path = r'c://users/Administrator/Desktop/1.xls'
    istp = ImportStep('127.0.0.1', 'root', 'root', 'test', 3306)
    istp.importFile(path)