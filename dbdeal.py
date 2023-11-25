import xlrd


class excel_ex:
    def __init__(self):
        self.device=[]
        self.repair=[]
        self.defect=[]
        self.measure=[]

    def get_data(self,path):
        data = xlrd.open_workbook(path)
        dv_table = data.sheet_by_name("设备表")
        r_table = data.sheet_by_name("检修记录表")
        df_table = data.sheet_by_name("缺陷表")
        m_table = data.sheet_by_name("反措表")
        self.device = data(dv_table)
        self.repair = data(r_table)
        self.defect = data(df_table)
        self.measure = data(m_table)

    def data(self,table):
        device = []
        rowNum = table.nrows
        colNum = table.ncols
        for i in range(rowNum):
            dev = []
            for j in range(colNum):
                dev.append(table.cell_value(i,j))
            device.append(dev)
        return device
    
    def Db_helper(self):
        
