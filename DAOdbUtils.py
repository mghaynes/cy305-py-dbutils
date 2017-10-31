import os, tkinter, pypyodbc, tkinter.messagebox
import csv, datetime
import win32com.client


pypyodbc.lowercase = False

# python3 python3 "z:\S&F\Courses\It305\ay181\admin\database_grader\DBHW3_grader.pyw"
cdtDict = {}
sections = []
sec = ""
debug = 0 #Set from 0 or 2 to get varying levels of output; 0=no output, 2=very verbose

class Table:
    def __init__(self, dbPath, tableName, type='TABLE'):
        self._dbPath = dbPath
        self._conn = None
        self._cur = None
        self._is_connected = False
        self._tableName = tableName
        self._tableType = type
        if self._ConnectToDB():
            self._rows = self.RowCount()
            self._columns = self.ColCount()
            self._columnNames = self.ColumnNames()
            self._columnTypes = self.ColumnTypes()
            if type == 'TABLE':
                self._primaryKeys = self.PrimaryKeys()
                self._foreignKeys, self._foreignKeysTables = self.ForeignKeys()
            self._CloseConnection()


    def PrintTable(self):
        print('TABLE NAME:', self._tableName)
        print('TABLE TYPE:', self._tableType)
        print('TABLE ROWS:', self._rows)
        print('TABLE COLUMNS:', self._columns)
        print('TABLE FIELD NAMES:', self._columnNames)
        print('TABLE FIELD TYPES:', self._columnTypes)
        if self._tableType == 'TABLE':
            print('PRIMARY KEY(S):', self._primaryKeys)
            print('FOREIGN KEY(S):', self._foreignKeys)
            print('FOREIGN KEY(S) TABLES:',self._foreignKeysTables)


    def _ConnectToDB(self):
        try:  # Try to connect to the database
            self._conn = pypyodbc.connect(
                r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" + "Dbq={0};".format(self._dbPath))
            self._cur = self._conn.cursor()
            self._is_connected = True
            return 1
        except Exception as e:
            print('Cannot open DB. Check if path valid.')
            return 0

    def _CloseConnection(self):
        try:
            self._cur.close()
            del self._cur
            self._cur = None
            self._conn.close()
            del self._conn
            self._conn = None
            self._is_connected = False
            return 1
        except:
            print('No connection to close')
            return 0


    # returns the name of the table
    def TableName(self):
        return self._tableName


    # returns the number of rows in a table
    def RowCount(self):
        if not self._is_connected:
            self._ConnectToDB()
        numRows = self._cur.execute('SELECT COUNT(*) AS count FROM [' + self._tableName + ']').fetchone()[0]
        if not self._is_connected:
            self._CloseConnection()
        return numRows


    # returns the number of columns in a table
    def ColCount(self):
        if not self._is_connected:
            self._ConnectToDB()
        self._cur.execute('SELECT * FROM [' + self._tableName + ']').fetchone()
        numCols = len(self._cur.description)
        if not self._is_connected:
            self._CloseConnection()
        return numCols


    # returns the names of the columns in a table
    def ColumnNames(self):
        if not self._is_connected:
            self._ConnectToDB()
        fieldNames = []
        for row in self._cur.columns(table=self._tableName):
            fieldNames.append("[" + row[3] + "]")
        if not self._is_connected:
            self._CloseConnection()
        return fieldNames


    # returns the types of the columns in a table
    def ColumnTypes(self):
        if not self._is_connected:
            self._ConnectToDB()
        fieldTypes = []
        for row in self._cur.columns(table=self._tableName):
            fieldTypes.append(row[5])
        if not self._is_connected:
            self._CloseConnection()
        return fieldTypes


    # Get a row of valid data from a table
    def GetValidRow(self):
        if not self._is_connected:
            self._ConnectToDB()
            self._ConnectToDB()
        row = self._cur.execute('SELECT * FROM [' + self._tableName + ']').fetchone()
        if not self._is_connected:
            self._CloseConnection()
        return row


    def PrimaryKeys(self, debug=0):
        PKs=[]
        if not self._is_connected:
            self._ConnectToDB()
        # Get unique indexes for the table
        rows = self._cur.statistics(table=self._tableName, unique=True)
        for row in rows:
            # Ignore the index of the whole table
            if row[8] == None:
                continue
            else:
                # append column name
                PKs.append(row[8])
            if debug > 1:
                print(row)
        if not self._is_connected:
            self._CloseConnection()
        return PKs


    def ForeignKeys(self, debug=0):
        FKs=[]
        FKTables = []
        if not self._is_connected:
            self._ConnectToDB()
        rows = self._cur.statistics(table=self._tableName)
        for row in rows:
            if row[8] == None:
                continue
            elif row[8] != row[5] and row[5] != 'PrimaryKey':
                FKs.append(row[8])
                FKTables.append(row[5].replace(self._tableName,''))
                if debug > 1:
                    print(row)
        for row in self._cur.statistics(table=self._tableName):
            print(row)
        if not self._is_connected:
            self._CloseConnection()
        return FKs, FKTables




    def ExecuteQuery(self):
        if not self._is_connected:
            self._ConnectToDB()
        sql = '{CALL ' + self._tableName + '}'
        # rows = self._cur.execute(sql)._last_executed
        rows = self._cur.execute(sql)
        for row in rows:
            print(row)
        if not self._is_connected:
            self._CloseConnection()

    def GetRecords(self):
        if not self._is_connected:
            self._ConnectToDB()
        sql = 'SELECT * FROM [' + self._tableName + ']'
        rows = self._cur.execute(sql).fetchall()
        if not self._is_connected:
            self._CloseConnection()
        return rows

    def PrintRecords(self):
        rows = self.GetRecords()
        for row in rows:
            print(row)
        return 1

'''END TABLE CLASS '''

# Note: Table1 should be the 'correct' table/query
def GradeTables(table1, table2):
    scoreVector = [0,0,0]
    if table1._tableName == table2._tableName:
        scoreVector[0] = 1
    if table1._rows == table2._rows:
        scoreVector[1] += 1
        table1_records = table1.GetRecords()
        table2_records = table2.GetRecords()
        scoreVector[2] = 1
        for rowNum in range(table1._rows):
            table_intersection = set(table1_records[rowNum]).intersection(set(table2_records[rowNum]))
            if len(table_intersection) != len(table1_records[rowNum]):
                scoreVector[2] = 0
                break
    if table1._columns == table2._columns:
        scoreVector[1] += 1
    return scoreVector



def FindBestTable(solnTable, tableNameList, dbPath):
    bestScore = [0,0,0]
    bestTableName = ''
    badTableNames = []
    for tableName in tableNameList:
        try:
            nextTable = Table(dbPath, tableName)
        except Exception as e:
            print('Error TABLE:',tableName,e)
            badTableNames.append(tableName)
            continue
        scoreVector = GradeTables(solnTable, nextTable)
        if sum(scoreVector) > sum(bestScore):
            bestTableName = tableName
            bestScore = scoreVector
        if sum(bestScore) == 4:
            break
    return bestTableName, bestScore, badTableNames



def GetTableNames(cur):
    tableList = []
    for row in cur.tables():
        if row[3] == 'TABLE':
            if not row[2].startswith('~'):
                tableList.append(row[2])
    return tableList

def GetQueryNames(cur):
    queryList = []
    for row in cur.tables():
        if row[3] == 'VIEW':
            if not row[2].startswith('~'):
                queryList.append(row[2])
    return queryList


'''-----------------------------------------------------------------------------------------------'''
'''-----------------------------------------------------------------------------------------------
    BELOW ALL OLD CODE NEEDS TO BE INTEGRATED INTO TABLE CLASS
   -----------------------------------------------------------------------------------------------'''
'''-----------------------------------------------------------------------------------------------'''



def main():
    # dbPath = r"\\usmasvddeecs\eecs\S&F\Courses\IT305\libraries\program_tracker_hw5_soln.accdb"
    # SolnDBPath = r"./program_tracker_hw5_soln.accdb"
    SolnDBPath = r"./DBProject181_soln.accdb"
    # dbPath = r"\\usmasvddeecs\eecs\S&F\Courses\IT305\libraries\program_tracker_hw2(soln).accdb"
    # studentDBPath = r"\\usmasvddeecs\eecs\Cadet\Courses\CY305\HAYNES\F3\FOWLER.CHRISTOPHER\database\hw5\program_tracker_hw5.accdb"

    SolnDBEngine = win32com.client.Dispatch("DAO.DBEngine.120")

    SolnWS = SolnDBEngine.Workspaces(0)
    try:
        SolnDB = SolnWS.OpenDatabase(SolnDBPath)
    except:
        print('Error opening database')
        return 0

    # Get a table
    try:
        TableDef = SolnDB.TableDefs('SoldierCompletesTraining')
    except:
        print('Error opening table')
        return 0


    #Note that the ColumnWidths have some weird conversion of 1 in = 1440
    LookupFields = ['DisplayControl','RowSourceType','RowSource','BoundColumn','ColumnCount', 'ColumnWidths',
                    'LimitToList']
    # Get field names
    for Field in TableDef.Fields:
        print('Name:',Field.Name,'Type:',Field.Type,'Size',Field.Size)
        for property in Field.Properties:
            if property.Name in LookupFields:
                print(property.Name,': ', property.Value)


    # Get primary keys
    print('PRIMARY KEYS:',)
    for idx in TableDef.Indexes:
        if idx.Primary:
            for field in idx.Fields:
                print(field.Name)
    print()


    # Get foreign keys
    print ('FOREIGN KEYS:')
    for rel in SolnDB.Relations:
        print('Name:',rel.Name)
        print('Foreign Table:',rel.Table,'\tTable:',rel.ForeignTable,'\tAttributes:',rel.Attributes)
        for field in rel.Fields:
            print('Field in Foreign Table:',field.Name,'\tField in Table:',field.ForeignName)
    print()

    # Get SQL
    for query in SolnDB.QueryDefs:
        if '~' not in query.Name:
            print('QUERY:',query.Name)
            print(query.SQL)

    #Loop through all the records in a table
    table = SolnDB.OpenRecordset('Platoon')
    #First get number of records
    print(table.RecordCount)
    while not table.EOF:
        record = table.GetRows()
        print(record)




if __name__ == "__main__":
    main()
