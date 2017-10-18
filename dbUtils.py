import os, tkinter, pypyodbc, tkinter.messagebox
import csv, datetime

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
        numRows = self._cur.execute('SELECT COUNT(*) AS count FROM [' + self._tableName + ']').fetchone()[0]
        return numRows


    # returns the number of columns in a table
    def ColCount(self):
        self._cur.execute('SELECT * FROM [' + self._tableName + ']').fetchone()
        numCols = len(self._cur.description)
        return numCols


    # returns the names of the columns in a table
    def ColumnNames(self):
        fieldNames = []
        for row in self._cur.columns(table=self._tableName):
            fieldNames.append("[" + row[3] + "]")
        return fieldNames


    # returns the types of the columns in a table
    def ColumnTypes(self):
        fieldTypes = []
        for row in self._cur.columns(table=self._tableName):
            fieldTypes.append(row[5])
        return fieldTypes


    # Get a row of valid data from a table
    def GetValidRow(self):
        connectionStartOpen = True
        if self._cur == None:
            connectionStartOpen == False
            self._ConnectToDB()
        row = self._cur.execute('SELECT * FROM [' + self._tableName + ']').fetchone()
        if not connectionStartOpen:
            self._CloseConnection()
        return row

    # Generates data that should not be in the table based on column type
    # def GetBadVal(self, cnt):
    #     if self._columnTypes[cnt] == 'DATETIME':
    #         return (datetime.datetime.now(),)
    #     elif self._columnTypes[cnt] == 'VARCHAR':
    #         return ('TEST',)
    #     else:
    #         return (-99999,)

    # Deletes a specific record (NOTE: Will not work on relationnship table)
    # def DeleteRecord(key, debug=0):
    #     sql = "SELECT * FROM [" + self._tableName + "]"
    #     try:
    #         rows = self._cur.execute(sql).fetchall()
    #     except Exception as e:
    #         if debug==2:
    #             print('Delete select error',e)
    #     idTuple = (rows[-1][0],)
    #     if debug==2:
    #         print('Last record:', rows[-1])
    #     sql = "DELETE FROM [" + self._tableName + "]" + " WHERE " + key + " = ?"
    #     if debug==2:
    #         print(sql)
    #     try:
    #         self._cur.execute(sql, idTuple)
    #     except Exception as e:
    #         print('Delete error:',e)


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
        if not self._is_connected:
            self._CloseConnection()
        return FKs, FKTables

    # def PrimaryKeys(self, debug=0):
    #     PKs = []
    #     #Iterate through each field and check if it's a primary key
    #     for cnt,field in enumerate(self._columnNames):
    #         try:
    #             # Get a row of existing data from the table. Try to change field to None and insert into table.
    #             validRow = self.GetValidRow()
    #             testRow = list(validRow)
    #             testRow[cnt] = None
    #             qNums = '?,'*len(testRow)
    #             qNums = qNums[:-1]
    #             sql = "INSERT INTO [" + self._tableName + '](' + ','.join(self._columnNames) + ')' + " VALUES ("+qNums+")"
    #             if debug == 2:
    #                 print('\t',sql, '; PARAMS:', testRow)
    #             self._cur.execute(sql,testRow)
    #             # If we were able to insert into table, there is no primary key! Delete last entry
    #             self.DeleteRecord(field, debug=debug)
    #         # Field may be a primary key depending on exception thrown when attempting to insert
    #         except Exception as e:
    #             eStr = str(e)
    #             if debug == 2:
    #                 print('PK error:',e)
    #             # Assume it's a primary key if field is autocount, can't insert null, or requires value in field
    #             if 'You tried to assign the Null value to a variable that is not a Variant data type' in eStr or \
    #                 'Index or primary key cannot contain a Null value.' in eStr or \
    #                 'You must enter a value in the' in eStr:
    #                 if field not in PKs:
    #                     PKs.append(field)
    #     if debug:
    #         print('\tPRIMARY KEYS:',PKs)
    #     return PKs


    # ForeignKeys finds the foreign keys for a table. REQUIRES tables primary keys already identified.
    # def ForeignKeyTables(self, debug=0):
    #     FKs = []
    #     # ASSUMPTION: If relationship table (i.e. composite key), then no foreign keys
    #     if len(self._primaryKeys) > 1:
    #         return FKs
    #     for cnt,field in enumerate(self._columnNames):
    #         # Keep valid primary key
    #         if field in self._primaryKeys: continue
    #         testFields = self._primaryKeys+[field]
    #         testFields = ','.join(testFields)
    #         testData = []
    #         # Change pkey values so it will throw error based on foreign key
    #         for key in self._primaryKeys:
    #             idx = self._columnNames.index(key)
    #             badVal = self.GetBadVal(idx)[0]
    #             testData.append(badVal)
    #         # Use some bad value for the foreign key field
    #         badVal = self.GetBadVal(cnt)[0]
    #         testData.append(badVal)
    #         Qs = '?,'*len(testData)
    #         Qs = Qs[:-1]+")"
    #         try:
    #             sql = "INSERT INTO [" + self._tableName + '] (' + testFields + ") VALUES (" + Qs
    #             if debug == 2:
    #                 print(sql, tuple(testData))
    #             self._cur.execute(sql,tuple(testData))
    #             # If successful then not a foreign key, so delete last entry
    #             self.DeleteRecord(debug=debug)
    #         except Exception as e:
    #             eStr = str(e)
    #             if debug==2:
    #                 print('FK error:',e)
    #             # if related record required, then it's a foreign key
    #             if ('You cannot add or change a record because a related record is required in table') in eStr:
    #                 tableName = eStr.split("'")[-2].split('.')[-1]
    #                 # solnTableName = MatchTables(dbDict, tableName)
    #                 # key = [dbDict[solnTableName]['PrimaryKeys'],solnTableName]
    #                 if tableName not in FKs:
    #                     FKs.append(tableName)
    #     if debug:
    #         print('\tFOREIGN KEYS:',FKs,'\n')
    #     return FKs

    def Procedures(self):
        if self._cur == None:
            self._ConnectToDB()
        sql = '{CALL ' + self._tableName + '}'
        rows = self._cur.execute(sql)._last_executed
        # rows = self._cur.procedures()
        # for row in rows:
        print(rows)

        if self._cur != None:
            self._CloseConnection()

'''END TABLE CLASS '''

def GradeTables(table1, table2):
    scoreVector = [0,0,0]
    if table1._tableName == table2._tableName:
        scoreVector[0] = 1
    if table1._rows == table2._rows:
        scoreVector[1] += 1
        table1._ConnectToDB()
        table2._ConnectToDB()
        table1_records = table1._cur.execute('SELECT * FROM [' + table1._tableName + ']')
        table2_records = table2._cur.execute('SELECT * FROM [' + table2._tableName + ']')
        scoreVector[2] = 1
        for rowNum in range(table1._rows):
            rec_1 = table1_records.fetchone()
            rec_2 = table2_records.fetchone()
            table_intersection = set(rec_1).intersection(set(rec_2))
            if len(table_intersection) != len(rec_1):
                scoreVector[2] = 0
                break
        table1._CloseConnection()
        table2._CloseConnection()
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
    dbPath = r"\\usmasvddeecs\eecs\S&F\Courses\IT305\libraries\program_tracker_hw5_soln.accdb"
    # dbPath = r"\\usmasvddeecs\eecs\S&F\Courses\IT305\libraries\program_tracker_hw2(soln).accdb"
    # studentDBPath = r"\\usmasvddeecs\eecs\Cadet\Courses\CY305\HAYNES\F3\FOWLER.CHRISTOPHER\database\hw5\program_tracker_hw5.accdb"

    # Connect to DB
    try:  # Try to connect to the database
        conn = pypyodbc.connect(
            r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" + "Dbq={0};".format(dbPath))
    except:
        print('Cannot open DB')
        return 0

    # if we have a good connection to the database
    cur = conn.cursor()
    tableList = GetTableNames(cur)
    queryList = GetQueryNames(cur)
    for table in queryList:
        solnTable = Table(dbPath, table, type='QUERY')
        solnTable.PrintTable()
        solnTable.Procedures()
        print('\n')
    # print('A row:',solnTable.GetValidRow())
    #     print(solnTable.Statistics())
    # studentTable = Table(studentDBPath, 'EmployeeBioBrief')
    # print(gradeTables(solnTable, studentTable))

if __name__ == "__main__":
    main()
