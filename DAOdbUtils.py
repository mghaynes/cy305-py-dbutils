import win32com.client
import collections

cdtDict = {}
sections = []
sec = ""
debug = 0 #Set from 0 or 2 to get varying levels of output; 0=no output, 2=very verbose

Lookup = collections.namedtuple('Lookup', ['DisplayControl', 'RowSourceType', 'RowSource', 'BoundColumn',
                                            'ColumnCount', 'ColumnWidths', 'LimitToList'])
ColumnMeta = collections.namedtuple('ColumnMeta',['Name','Type','Size'])


class Relationship:
    def __init__(self, Name, Table, Field, RelatedTable, RelatedField, Attributes):
        self.Name = Name
        self.Table = Table
        self.Field = Field
        self.RelatedTable = RelatedTable
        self.RelatedField = RelatedField
        self._attributes = Attributes
        if Attributes==0:
            self.JoinType = 'INNER'
            self.ReferentialIntegrity = True
        elif Attributes==2:
            self.JoinType = 'INNER'
            self.ReferentialIntegrity = False
        elif Attributes == 16777216:
            self.JoinType = 'OUTER RELATED'
            self.ReferentialIntegrity = True
        elif Attributes == 16777218:
            self.JoinType = 'OUTER RELATED'
            self.ReferentialIntegrity = False
        elif Attributes == 33554432:
            self.JoinType = 'OUTER TABLE'
            self.ReferentialIntegrity = True
        elif Attributes == 33554434:
            self.JoinType = 'OUTER TABLE'
            self.ReferentialIntegrity = False
        else:
            self.JoinType = 'UNKNOWN'
            self.ReferentialIntegrity = None


    def __str__(self):
        return 'Table: '+self.Table+'\tField: '+self.Field+'\tRelated Table: '+\
               self.RelatedTable+'\tRelated Field: '+self.RelatedField+'\tRelationship Type: '+self.JoinType+\
               '\tReferential Integrity: '+str(self.ReferentialIntegrity)

    ''' END RELATIONSHIP CLASS'''

class DataBase:
    def __init__(self, dbPath, debug=0):
        self._dbEngine = win32com.client.Dispatch("DAO.DBEngine.120")
        self._ws = self._dbEngine.Workspaces(0)
        self._dbPath = dbPath
        self._db = self._ws.OpenDatabase(self._dbPath)
        self._debug = debug
        self.TableNames = self.TableList(debug=self._debug)
        self.QueryNames = self.TableList(isTable=False, debug=self._debug)
        self.Tables = self.LoadTables(self.TableNames)
        self.Queries = self.LoadTables(self.QueryNames, isTable=False)
        self.Relationships = self.GetRelationships(debug=self._debug)

    # For query list, isTable must be False
    def TableList(self, isTable=True, debug=0):
        table_list = []
        if isTable:
            tables = self._db.TableDefs
        else:
            tables = self._db.QueryDefs
        if debug and isTable:
            print('TABLES:')
        elif debug and not isTable:
            print('QUERIES')
        for table in tables:
            if not table.Name.startswith('MSys') and not table.Name.startswith('~'):
                table_list.append(table.Name)
                if debug:
                    print(table.Name)
        return table_list


    def LoadTables(self, table_list, isTable=True):
        tables = {}
        for table in table_list:
            if isTable:
                tables[table] = Table(self._db.TableDefs(table), dbPath=self._dbPath)
            else:
                tables[table] = Table(self._db.QueryDefs(table), isTable=isTable, dbPath=self._dbPath)
        return tables


    '''' Attributes translations
        0 = Enforce referential integrity (RI), Inner join
        2 = Referential integrity (RI) not enforced, Inner join
        16777216 = RI, outer join on related table
        16777218 = No RI, outer join on related table
        33554434 = No RI, outer join on table
        33554432 = RI, outer join on table'''
    def GetRelationships(self, debug=0):
        relationships=[]
        if debug:
            print('RELATIONSHIPS')
        for rltn in self._db.Relations:
            for field in rltn.Fields:
                new_relationship = Relationship(Name=rltn.Name,RelatedTable=rltn.Table, Table=rltn.ForeignTable,
                                                RelatedField=field.Name,Field=field.ForeignName,
                                                Attributes=rltn.Attributes)
                if debug:
                    print(new_relationship)
                relationships.append(new_relationship)
        return relationships

    ''' END DATABASE CLASS '''

class Table:
    def __init__(self, table_meta=None, isTable=True, dbPath=None, debug=1):
        if table_meta==None:
            return
        self._dbEngine = win32com.client.Dispatch("DAO.DBEngine.120")
        self._ws = self._dbEngine.Workspaces(0)
        self._dbPath = dbPath
        self._TableMetaData = table_meta
        self.Name = table_meta.Name
        self.debug = debug
        if isTable:
            self.TableType = 'TABLE'
            self.RecordCount = table_meta.RecordCount
            self.PrimaryKeys = self.GetPrimaryKeys()
        else:
            self.TableType = 'QUERY'
            self.SQL = self.GetSQL(table_meta)
            if dbPath != None:
                self.RecordCount = self.QueryRecordCount()
        self.ColumnMetaData = self.GetColumnMetaData(table_meta)
        self.ColumnCount = len(self.ColumnMetaData)


    def __str__(self):
        column_tuples = [(field.Name, field.Type, field.Size) for field in self.ColumnMetaData]
        if self.TableType == 'TABLE':
            return 'Table Name: {:25}Type: {:10}Row Count: {:<10}Column Count: {}\nColumns: {}\nPrimary Keys: ' \
                   '{}'.format(self.Name, self.TableType,self.RecordCount, self.ColumnCount,
                                              column_tuples, ', '.join(self.PrimaryKeys))
        elif self.TableType == 'QUERY':
            return 'Query Name: {:25}Type: {:10}Row Count: {:<10}Column Count: {}\nColumns: {}\nSQL: ' \
                   '{}'.format(self.Name, self.TableType,self.RecordCount, self.ColumnCount, column_tuples, self.SQL)
        else:
            return ''
                # self._rows = self.RowCount(self.debug)

    def QueryRecordCount(self):
        self._db = self._ws.OpenDatabase(self._dbPath)
        num_rows = self._db.OpenRecordset(self.Name).RecordCount
        self._db.Close()
        return num_rows


    # returns the names of the columns in a table
    def GetColumnMetaData(self, table_meta, debug=0):
        columns = []
        if debug:
            print('TABLE:', table_meta.Name)
        for Field in table_meta.Fields:
            if Field.Type == 1:
                type = 'Yes/No'
            elif Field.Type == 4:
                if Field.Attributes in [17,18]:
                    type = 'Autonumber'
                else:
                    type = 'LongInteger'
            elif Field.Type == 7:
                type = 'Double'
            elif Field.Type == 8:
                type = 'Date/Time'
            elif Field.Type == 10:
                type = 'ShortText'
            else:
                type = 'UNKNOWN'
            column_meta = ColumnMeta(Field.Name, type, Field.Size)
            columns.append(column_meta)
            if debug:
                print('Field Name:', column_meta.Name,'Type:',column_meta.Type,'Size',column_meta.Size)
        return columns


    def GetLookups(self, field_meta, debug=0):
        # Note that the ColumnWidths have some weird conversion of 1" = 1440
        LookupFields = ['RowSourceType','RowSource','BoundColumn','ColumnCount', 'ColumnWidths',
                        'LimitToList']
        lookup = None
        for property in field_meta.Properties:
            if property.Name == 'DisplayControl':
                if property.Value == 111:
                    display_control = 'Combo box'
                if property.Value == 110:
                    display_control = 'List box'
                if property.Value == 109:
                    display_control = 'Text box'
            if property.Name == 'RowSourceType':
                row_source_type = property.Value
            if property.Name == 'RowSource':
                row_source = property.Value
            if property.Name == 'BoundColumn':
                bound_column = property.Value
            if property.Name == 'ColumnCount':
                column_count = property.Value
            if property.Name == 'ColumnWidths':
                column_widths = property.Value
            if property.Name == 'LimitToList':
                limit_to_list = property.Value
            if debug > 1 and property.Name in LookupFields:
                print(property.Name,': ', property.Value)
            if debug > 1 and property.Name == 'DisplayControl':
                print(property.Name, ': ', display_control)
        lookup = Lookup(display_control, row_source_type, row_source, bound_column, column_count, column_widths,
                        limit_to_list)
        return lookup


    def GetPrimaryKeys(self, debug=0):
        PKs=[]
        for idx in self._TableMetaData.Indexes:
            if idx.Primary:
                for field in idx.Fields:
                    PKs.append(field.Name)
        if debug:
            print(self.Name.upper(),'primary keys:', ','.join(PKs))
        return PKs


    def GetSQL(self, query, debug=0):
        if '~' not in query.Name:
            if debug:
                print('QUERY SQL for',query.Name)
                print(query.SQL)
            return query.SQL
        else:
            return 0


    def GetRecords(self, debug=0):
        self._db = self._ws.OpenDatabase(self._dbPath)
        table = self._db.OpenRecordset(self.Name)
        records = []
        while not table.EOF:
            record = table.GetRows()
            records.append(record)
            if debug>1:
                print(record)
        self._db.Close()
        return records


    def GetFieldObject(self, name):
        return self._TableMetaData.Fields(name)


'''END TABLE CLASS '''


def ListProperties(object):
    for property in object.Properties:
        try:
            print(property.Name, ':', property.Value)
        except:
            print(property.Name)


# Note: Table1 should be the 'correct' table/query. Table 2 is compared against Table 1.
# def GradeTables(table1, table2, verbose=False):
#     scoreVector = [0,0,0]
#     if table1._tableName == table2._tableName:
#         scoreVector[0] = 1
#     if table1._rows == table2._rows:
#         scoreVector[1] += 1
#         table1_records = table1.GetRecords()
#         table2_records = table2.GetRecords()
#         scoreVector[2] = 1
#         for rowNum in range(table1._rows):
#             table_intersection = set(table1_records[rowNum]).intersection(set(table2_records[rowNum]))
#             if len(table_intersection) != len(table1_records[rowNum]):
#                 scoreVector[2] = 0
#                 break
#     if table1._columns == table2._columns:
#         scoreVector[1] += 1
#     return scoreVector



'''-----------------------------------------------------------------------------------------------'''
'''-----------------------------------------------------------------------------------------------'''
def main():
    SolnDBPath = r"./DBProject181_soln.accdb"
    SolnDB = DataBase(SolnDBPath)
    # Print meta data on all the tables in the database
    for table in SolnDB.TableNames:
        print(SolnDB.Tables[table], '\n')
    # Print meta data on all the queries in the database
    for query in SolnDB.QueryNames:
        print(SolnDB.Queries[query], '\n')
    # print all the relationships in the table
    for relationship in SolnDB.Relationships:
        print(relationship,'\n')
    # print all the records in a table (Note: If debug < 2, it doesn't print anything. Just returns the records)
    print('Platoon Table Records')
    SolnDB.Tables['Platoon'].GetRecords(debug=2)
    print()
    # print the lookups for a field (Note: If debug < 2, it doesn't print anything. Just returns the Lookup tuple)
    print('Lookups for soldierTrained field in SoldierCompletesTraining')
    table = SolnDB.Tables['SoldierCompletesTraining']
    field = table.GetFieldObject('soldierTrained')
    table.GetLookups(field, debug=2)
    print()
    # print the properties for some metadata (e.g. Table, Query, or Field)
    print('Table Properties')
    ListProperties(table._TableMetaData)


if __name__ == "__main__":
    main()
