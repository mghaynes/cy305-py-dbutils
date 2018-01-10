import win32com.client
import collections
import json

cdtDict = {}
sections = []
sec = ""
debug = 0  # Set from 0 or 2 to get varying levels of output; 0=no output, 2=very verbose

Lookup = collections.namedtuple('Lookup', ['DisplayControl', 'RowSourceType', 'RowSource', 'BoundColumn',
                                           'ColumnCount', 'ColumnWidths', 'LimitToList'])
ColumnMeta = collections.namedtuple('ColumnMeta', ['Name', 'Type', 'Size'])

Relationship = collections.namedtuple('Relationship', ['Table', 'Field', 'RelatedTable', 'RelatedField',
                                                       'EnforceIntegrity', 'JoinType', 'Attributes'])

# class Relationship(collections.namedtuple('Relationship', ['Table', 'Field', 'RelatedTable', 'RelatedField',
#                                                            'EnforceIntegrity', 'JoinType','Attributes'])):


class DataBase:
    def __init__(self, dbPath, debug=0):
        self._dbEngine = win32com.client.Dispatch("DAO.DBEngine.120")
        self._ws = self._dbEngine.Workspaces(0)
        self._dbPath = dbPath
        self._db = self._ws.OpenDatabase(self._dbPath)
        self._debug = debug
        self.TableNames = self.TableList(debug=self._debug)
        self.QueryNames = self.TableList(isTable=False, debug=self._debug)
        self.Relationships = self.GetRelationships(debug=self._debug)
        self.Tables = self.LoadTables(self.TableNames)
        self.Queries = self.LoadTables(self.QueryNames, isTable=False)

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
                if table in self.Relationships:
                    tables[table].ForeignKeys = self.Relationships[table]
            else:
                tables[table] = Table(self._db.QueryDefs(table), isTable=isTable, dbPath=self._dbPath)
        return tables


    '''' Attributes translations (I THINK!)
        0 = Enforce referential integrity (RI), Inner join
        2 = Referential integrity (RI) not enforced, Inner join
        16777216 = RI, outer join on related table
        16777218 = No RI, outer join on related table
        33554434 = No RI, outer join on table
        33554432 = RI, outer join on table'''
    def GetRelationships(self, debug=1):
        relationships = dict()
        for rltn in self._db.Relations:
            if rltn.ForeignTable not in relationships:
                relationships[rltn.ForeignTable] = dict()
            if rltn.Table not in relationships[rltn.ForeignTable]:
                relationships[rltn.ForeignTable][rltn.Table] = dict()
            for field in rltn.Fields:
                if rltn.Attributes == 0:
                    JoinType = 'INNER'
                    ReferentialIntegrity = True
                elif rltn.Attributes == 2:
                    JoinType = 'INNER'
                    ReferentialIntegrity = False
                elif rltn.Attributes == 16777216:
                    JoinType = 'OUTER RELATED'
                    ReferentialIntegrity = True
                elif rltn.Attributes == 16777218:
                    JoinType = 'OUTER RELATED'
                    ReferentialIntegrity = False
                elif rltn.Attributes == 33554432:
                    JoinType = 'OUTER TABLE'
                    ReferentialIntegrity = True
                elif rltn.Attributes == 33554434:
                    JoinType = 'OUTER TABLE'
                    ReferentialIntegrity = False
                else:
                    JoinType = 'UNKNOWN'
                    ReferentialIntegrity = None
                new_rltn = Relationship(Table=rltn.ForeignTable, Field=field.ForeignName, RelatedTable=rltn.Table,
                                        RelatedField=field.Name, EnforceIntegrity=ReferentialIntegrity,
                                        JoinType=JoinType, Attributes=rltn.Attributes)
                relationships[rltn.ForeignTable][rltn.Table][field.ForeignName] = new_rltn
                # if debug:
                #     print(relationships)
        if debug:
            for table_name in relationships.keys():
                for foreign_name in relationships[table_name].keys():
                    for field_name in relationships[table_name][foreign_name].keys():
                        print(relationships[table_name][foreign_name][field_name])
        return relationships

    ''' END DATABASE CLASS '''

class Table:
    def __init__(self, table_meta=None, isTable=True, dbPath=None, debug=0):
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
            self.ForeignKeys = ''
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
            if self.ForeignKeys:
                fk_list = [str(r2) for k, r in self.ForeignKeys.items() for k2, r2 in r.items()]
            else:
                fk_list = ['']
            return 'Table Name: {:25}Type: {:10}Row Count: {:<10}Column Count: {}\nColumns: {}\nPrimary Keys: ' \
                   '{}\nForeign Keys: {}'.format(self.Name, self.TableType,self.RecordCount, self.ColumnCount,
                                                 column_tuples, ', '.join(self.PrimaryKeys),
                                                 '\n'.join(fk_list))
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

    # def GetForeignKeys(self, debug=0):
    #     FKs = dict()
    #     self._db = self._ws.OpenDatabase(self._dbPath)
    #     if self.Name in self._db.Relationships:
    #         pass
            # for rltn_name in self._db.Relationships[self.Name].keys():
                # FKs[rltn_name] = self._db.Relationships[self.Name][rltn_name]
        # self._db.Close()
        # for foreign_name in FKs.keys():
        #     for field_name in FKs[foreign_name].keys():
        #         print(FKs[foreign_name][field_name])

        # return FKs


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
            temp_rec = []
            record = table.GetRows()
            for item in record:
                temp_rec.append(list(item)[0])
            records.append(temp_rec)
            if debug > 1:
                print(temp_rec)
        self._db.Close()
        return records


    def GetFieldObject(self, name):
        return self._TableMetaData.Fields(name)

    def GetFields(self):
        fields = []
        for column in self.ColumnMetaData:
            fields.append(column.Name)
        return fields

    def GetTypes(self):
        types = []
        for column in self.ColumnMetaData:
            types.append(column.Type)
        return types

    def GetSizes(self):
        sizes = []
        for column in self.ColumnMetaData:
            sizes.append(column.Size)
        return sizes

'''END TABLE CLASS '''


def ListProperties(object):
    for property in object.Properties:
        try:
            print(property.Name, ':', property.Value)
        except:
            print(property.Name)

class TableScore(collections.namedtuple('TableScore',['NameScore','RowCountScore','ColCountScore','FieldNameScore',
                                                      'FieldTypeScore','FieldSizeScore','RowsScore', 'SamePriKeysScore',
                                                      'DiffPriKeysScore', 'Correct_Num_Rltns', 'Fld', 'Rltd_Tbl',
                                                      'Rltd_Fld', 'Join', 'Integrity'])):
    def __str__(self):
        name_str = 'Table name score : {}\n'.format(self.NameScore)
        row_cnt_str = 'Row count score: {}\n'.format(self.RowCountScore)
        col_cnt_str = 'Column count score: {}\n'.format(self.ColCountScore)
        field_name_str = 'Num Scoring field names: {}\n'.format(self.FieldNameScore)
        field_type_str = 'Num Scoring field types: {}\n'.format(self.FieldTypeScore)
        field_size_str = 'Num Scoring field sizes: {}\n'.format(self.FieldSizeScore)
        sm_pri_keys_str = 'Matching primary keys score: {}\n'.format(self.SamePriKeysScore)
        diff_pri_keys_str = 'Different primary keys score: {}\n'.format(self.DiffPriKeysScore)
        num_rltns_str = 'Correct number of relationships (1 or 0): {}\n'.format(self.Correct_Num_Rltns)
        fld_str = 'Matching relationship fields: {}\n'.format(self.Fld)
        rltd_tbl_str = 'Matching relationship related tables: {}\n'.format(self.Rltd_Tbl)
        rltd_fld_str = 'Matching relationship related fields: {}\n'.format(self.Rltd_Fld)
        join_str =  'Matching join types: {}\n'.format(self.Join)
        integrity_str = 'Matching referential integrity values: {}\n'.format(self.Integrity)
        if self.RowsScore == 1:
            row_score_str = 'Records Score: 4 (Exact)\t'
        elif self.RowsScore == 3/4:
            row_score_str = 'Records Score: 3 (Exact, columns out of order)'
        elif self.RowsScore == 2/4:
            row_score_str = 'Records Score: 2 (Exact, rows out of order)'
        elif self.RowsScore == 1/4:
            row_score_str = 'Records Score: 1 (Exact, rows and columns out of order)'
        else:
            row_score_str = 'Records Score: 0'
        return name_str+row_cnt_str+col_cnt_str+field_name_str+field_type_str+field_size_str+sm_pri_keys_str+ \
               diff_pri_keys_str+num_rltns_str+fld_str+rltd_tbl_str+rltd_fld_str+join_str+integrity_str+row_score_str


# base_table_score allocates 20% fields, 40% PKs, 40% relationships (doesn't check table values)
base_table_score = TableScore(NameScore=.05, RowCountScore=0, ColCountScore=0, FieldNameScore=.05, FieldTypeScore=.1,
                              FieldSizeScore=0, RowsScore=0, SamePriKeysScore=.4, DiffPriKeysScore=0,
                              Correct_Num_Rltns=.025, Fld=.075, Rltd_Tbl=.1, Rltd_Fld=.1, Join=.025, Integrity=.075)
# no relations table score allocations 20 (doesn't check table values or relationships)
# no_rltns_table_score = TableScore(NameScore=.1, RowCountScore=.1, ColCountScore=.1, FieldNameScore=.1,
#                                   FieldTypeScore=.1, FieldSizeScore=.1, RowsScore=.1, SamePriKeysScore=.1,
#                                   DiffPriKeysScore=.1, Correct_Num_Rltns=.1, Fld=.1, Rltd_Tbl=.1, Rltd_Fld=.1, Join=.1,
#                                   Integrity=0)
# base2_table_score = TableScore(NameScore=.1, RowCountScore=.1, ColCountScore=.1, FieldNameScore=.1, FieldTypeScore=.1,
#                               FieldSizeScore=.1, RowsScore=.1, SamePriKeysScore=.1, DiffPriKeysScore=.1,
#                               Correct_Num_Rltns=.1, Fld=.1, Rltd_Tbl=.1, Rltd_Fld=.1, Join=.1, Integrity=0)
# base3_table_score = TableScore(NameScore=.1, RowCountScore=.1, ColCountScore=.1, FieldNameScore=.1, FieldTypeScore=.1,
#                               FieldSizeScore=.1, RowsScore=.1, SamePriKeysScore=.1, DiffPriKeysScore=.1,
#                               Correct_Num_Rltns=.1, Fld=.1, Rltd_Tbl=.1, Rltd_Fld=.1, Join=.1, Integrity=0)


def GradeRelationships(rltn_dict1, rltn_dict2):
    correct_num_rltns = fld = rltd_fld = rltd_tbl = join = integrity = 0
    # if no relationships then return all 1s
    if rltn_dict1 == '':
        return 1, 1, 1, 1, 1, 1
    num_rltns = len(rltn_dict1.keys())
    if num_rltns == len(rltn_dict2.keys()):
        correct_num_rltns = 1
    # print(rltn_dict1.keys())
    # print(correct_num_rltns)
    for rltd_tbl1_key in rltn_dict1:
        # obvious flaw here is as long as key in once will keep getting credit
        #  even if should be in multiple times but not
        if rltd_tbl1_key in rltn_dict2:
            rltd_tbl += 1
            # print(rltd_tbl1_key)
            # same potential flaw as above
            for field1 in rltn_dict1[rltd_tbl1_key]:
                if field1 in rltn_dict2[rltd_tbl1_key]:
                    fld += 1
                    print(field1)
                    rltn1 = rltn_dict1[rltd_tbl1_key][field1]
                    rltn2 = rltn_dict2[rltd_tbl1_key][field1]
                    if rltn1.RelatedField == rltn2.RelatedField:
                        rltd_fld += 1
                    if rltn1.JoinType == rltn2.JoinType:
                        join += 1
                    if rltn1.EnforceIntegrity == rltn2.EnforceIntegrity:
                        integrity += 1
    rltd_tbl /= num_rltns
    fld /= num_rltns
    rltd_fld /= num_rltns
    join /= num_rltns
    integrity /= num_rltns
    print('related field:{}\njoin:{}\nintegrity:{}'.format(rltd_fld, join, integrity))
    return correct_num_rltns, fld, rltd_tbl, rltd_fld, join, integrity

def AssessTableEntries(table1, table2):
    table1_recs = table1.GetRecords()
    table2_recs = table2.GetRecords()
    exact_rec_score = 4
    # check exact table match (i.e. row,col values all match)
    for cnt, row in enumerate(table1_recs):
        if row != table2_recs[cnt]:
            exact_rec_score = 3
            break
    # check row values match (i.e. col order doesn't matter)
    if exact_rec_score == 3:
        for cnt, row in enumerate(table1_recs):
            if set(row).intersection(table2_recs[cnt]) != set(row):
                exact_rec_score = 2
                break
    # check out of order exact records match (i.e. rows out of order, but col order still matters)
    if exact_rec_score == 2:
        for row in table1_recs:
            if row not in table2_recs:
                exact_rec_score = 1
                break
    # check if recs in table but out of order (col order doesn't matter)
    if exact_rec_score == 1:
        for row in table1_recs:
            any_score = False
            for row2 in table2_recs:
                if set(row).intersection(row2) == set(row):
                    any_score = True
                    break
            if not any_score:
                exact_rec_score = 0
                break
    return exact_rec_score

# Note: Table1 should be the 'correct' table/query. Table 2 is compared against Table 1.
# The scores are returned as percentages. For example, if you had 2 of 3 primary keys correct the
# score returned is 0.67 (this makes it easier to multiply by whatever rubric you want to use)
def AssessTables(table1, table2):
    name_score = row_count_score = col_count_score = field_name_score = field_type_score = field_size_score = \
        exact_rec_score = 0
    if table1.Name == table2.Name:
        name_score = 1
    if table1.RecordCount == table2.RecordCount:
        row_count_score = 1
    if table1.ColumnCount == table2.ColumnCount:
        col_count_score = 1
    # primary keys intersection returns primary keys in common between table1 and table2
    pk_same = len(set(table1.PrimaryKeys).intersection(table2.PrimaryKeys)) / len(table1.PrimaryKeys)
    # this finds any keys that the student (table2) has that are not in the solution (table1)
    pk_diff = len(set(table2.PrimaryKeys).difference(table1.PrimaryKeys)) / (len(table1.GetFields()) -
                                                                             len(table1.PrimaryKeys))
    correct_num_rltns, fld, rltd_tbl, rltd_fld, join, integrity = GradeRelationships(table1.ForeignKeys,
                                                                                     table2.ForeignKeys)
    table1_fields = table1.GetFields()
    table2_fields = table2.GetFields()
    table1_types = table1.GetTypes()
    table2_types = table2.GetTypes()
    table1_sizes = table1.GetSizes()
    table2_sizes = table2.GetSizes()
    for cnt, field in enumerate(table1_fields):
        if field in table2_fields:
            table2_idx = table2_fields.index(field)
            field_name_score += 1
            if table1_types[cnt] == table2_types[table2_idx]:
                field_type_score += 1
            if table1_sizes[cnt] == table2_sizes[table2_idx]:
                field_size_score += 1
    field_name_score /= len(table1_fields)
    field_type_score /= len(table1_types)
    field_size_score /= len(table1_sizes)
    if row_count_score:
        exact_rec_score = AssessTableEntries(table1, table2)
    exact_rec_score /= 4
    table_score = TableScore(name_score, row_count_score, col_count_score, field_name_score, field_type_score,
                             field_size_score, exact_rec_score, pk_same, pk_diff, correct_num_rltns, fld, rltd_tbl,
                             rltd_fld, join, integrity)
    return table_score


def ScoreTable(assessed_table, score_vector=base_table_score):
    table_score = 0
    for cnt in range(len(assessed_table)):
        table_score += assessed_table[cnt]*score_vector[cnt]
    # print('Table Score: {}%'.format(table_score*100))


'''-----------------------------------------------------------------------------------------------'''
'''-----------------------------------------------------------------------------------------------'''


def main():
    SolnDBPath = r"./DBProject181_soln.accdb"
    StudentDBPath = r"./DBProject181.accdb"
    SolnDB = DataBase(SolnDBPath)
    StudentDB = DataBase(StudentDBPath)
    # Print meta data on all the tables in the database
    for table in SolnDB.TableNames:
        print(SolnDB.Tables[table], '\n')
    # Print meta data on all the queries in the database
    for query in SolnDB.QueryNames:
        print(SolnDB.Queries[query], '\n')
    # print all the relationships in the table
    # for relationship in SolnDB.Relationships:
    #    print(json.dumps(relationship))
    print(json.dumps(SolnDB.Relationships))
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
    # table_assessment = AssessTables(SolnDB.Tables['SoldierCompletesTraining'],
    #                                 StudentDB.Tables['SoldierCompletesTraining'])
    table_assessment = AssessTables(SolnDB.Tables['Platoon'], StudentDB.Tables['Platoon'])
    print()
    print('Comparing "SoldierCompletesTraining" tables...')
    print(table_assessment)
    ScoreTable(table_assessment)


if __name__ == "__main__":
    main()
