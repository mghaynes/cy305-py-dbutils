

# you'll need to import these libraries
# pip install pypiwin32
import win32com.client
import Levenshtein
import numpy as np
# these are built in to python
import collections
import json
import re
import itertools
import copy


'''
TO DO: 
1) Check if the field is supposed to be hidden for queries
2) Add functions to check lookup properties

'''

debug = 0  # Set from 0 or 2 to get varying levels of output; 0=no output, 2=very verbose (NOT IMPLEMENTED YET)
too_many_penalty = .05  # penalty for selecting too many items

Lookup = collections.namedtuple('Lookup', ['DisplayControl', 'RowSourceType', 'RowSource', 'BoundColumn',
                                           'ColumnCount', 'ColumnWidths', 'LimitToList'])
ColumnMeta = collections.namedtuple('ColumnMeta', ['Name', 'Type', 'Size'])

Relationship = collections.namedtuple('Relationship', ['Table', 'Field', 'RelatedTable', 'RelatedField',
                                                       'EnforceIntegrity', 'JoinType', 'Attributes'])


'''-----------------------------------------------------------------------------------------------------------------'''
'''                                               CLASS: DATABASE                                                   '''
'''    DataBase class loads key properties of database to include relationships, table, and query properties        '''


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


'''---------------------------------------------- END DATABASE CLASS ------------------------------------------------'''


'''-----------------------------------------------------------------------------------------------------------------'''
'''                                               CLASS: TABLE                                                      '''
''' DataBase class permits various operations on tables/queries to include getting records, SQL, lookups, keys,     '''
''' and more.                                                                                                       '''

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

    def hasColumn(self, name):
        column_meta = self.ColumnMetaData
        found = False
        for col in column_meta:
            if name in col.Name:
                return True
        return False

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
                print('Field Name:', column_meta.Name, 'Type:', column_meta.Type, 'Size', column_meta.Size)
        return columns


    def GetLookups(self, fieldName, debug=0):
        # Note that the ColumnWidths uses twips a unit of measure where 1 in = 1440 twips, 1 cm = 567 twips
        LookupFields = ['RowSourceType', 'RowSource', 'BoundColumn', 'ColumnCount', 'ColumnWidths',
                        'LimitToList']
        field_meta = self.GetFieldObject(fieldName)
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


'''---------------------------------------------- END TABLE CLASS ------------------------------------------------'''


# CLASS TableScore
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


def AssignTableWeights(NameScore=0, RowCountScore=0, ColCountScore=0, FieldNameScore=0, FieldTypeScore=0,
                       FieldSizeScore=0, RowsScore=0, SamePriKeysScore=0, DiffPriKeysScore=0, Correct_Num_Rltns=0,
                       Fld=0, Rltd_Tbl=0, Rltd_Fld=0, Join=0, Integrity=0):
    return TableScore(NameScore, RowCountScore, ColCountScore, FieldNameScore, FieldTypeScore,
                              FieldSizeScore, RowsScore, SamePriKeysScore, DiffPriKeysScore,
                              Correct_Num_Rltns, Fld, Rltd_Tbl, Rltd_Fld, Join, Integrity)


# base_table_score allocates 20% fields, 40% PKs, 40% relationships (doesn't check table values)
base_table_weight = AssignTableWeights(NameScore=.05, FieldNameScore=.05, FieldTypeScore=.1, SamePriKeysScore=.4,
                                      Correct_Num_Rltns=.025, Fld=.075, Rltd_Tbl=.1, Rltd_Fld=.1, Join=.025,
                                      Integrity=.075)


# CLASS
class QueryScore(collections.namedtuple('QueryScore',['SELECTscore', 'FROMscore', 'CRITERIAscore', 'GROUPBYscore',
                                                      'TOTALSscore', 'SORTscore', 'WHEREpenalty', 'HAVINGpenalty',
                                                      'GROUPBYpenalty', 'SORTpenalty', 'MatchScore'])):
    def __str__(self):
        select_str = 'SELECT score : {}\n'.format(self.SELECTscore)
        from_str = 'FROM score: {}\n'.format(self.FROMscore)
        criteria_str = 'CRITERIA score: {}\n'.format(self.CRITERIAscore)
        groupby_str = 'GROUPBY score: {}\n'.format(self.GROUPBYscore)
        totals_str = 'TOTALS score: {}\n'.format(self.TOTALSscore)
        sort_str = 'SORT score: {}\n'.format(self.SORTscore)
        where_penalty_str = 'WHERE penalty: {}\n'.format(self.WHEREpenalty)
        having_penalty_str = 'HAVING penalty: {}\n'.format(self.HAVINGpenalty)
        groupby_penalty_str = 'GROUPBY penalty: {}\n'.format(self.GROUPBYpenalty)
        sort_penalty_str = 'SORT penalty: {}\n'.format(self.SORTpenalty)
        row_score_str = 'Exact records match: {}'.format(self.MatchScore)
        return select_str+from_str+criteria_str+groupby_str+totals_str+sort_str+where_penalty_str+ \
               having_penalty_str+groupby_penalty_str+sort_penalty_str+row_score_str

def AssignQueryWeights(SELECTscore=0, FROMscore=0, CRITERIAscore=0, GROUPBYscore=0, TOTALSscore=0,
                       SORTscore=0, WHEREpenalty=0, HAVINGpenalty=0, GROUPBYpenalty=0, SORTpenalty=0, MatchScore=0):
    return QueryScore(SELECTscore, FROMscore, CRITERIAscore, GROUPBYscore, TOTALSscore,
                       SORTscore, WHEREpenalty, HAVINGpenalty, GROUPBYpenalty, SORTpenalty, MatchScore)

# default query weighting
base_query_weight = AssignQueryWeights(SELECTscore=0.2, FROMscore=0.2, CRITERIAscore=0.25, GROUPBYscore=0.125,
                                       TOTALSscore=0.125, SORTscore=0.1, WHEREpenalty=.1, HAVINGpenalty=.1,
                                       GROUPBYpenalty=.1, SORTpenalty=.1, MatchScore=0)


# returns Lenvenshtein distance between a target string and a list of strings. (CURRENTLY NOT USED)
# VARIABLE: Target  TYPE: String
# VARIABLE: Options TYPE: List (with elements being strings)
def BestMatch(target, options):
    best_distance = float("inf")
    best_option = ''
    for option in options:
        distance = Levenshtein.distance(target, option)
        if distance == best_distance:
            print("In BestMatch have two options with same Levenshtein distance. Check it out")
        if distance < best_distance:
            best_distance = distance
            best_option = option
    return best_distance, best_option


def ListProperties(object):
    for property in object.Properties:
        try:
            print(property.Name, ':', property.Value)
        except:
            print(property.Name)


def GradeRelationships(rltn_dict1, rltn_dict2):
    correct_num_rltns = fld = rltd_fld = rltd_tbl = join = integrity = 0
    # if no relationships then return all 1s
    if rltn_dict1 == '':
        return 1, 1, 1, 1, 1, 1
        # return 0, 0, 0, 0, 0, 0
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


def ExactRecordsMatch(table1_recs, table2_recs):
    exact_rec_score = 4
    # check exact table match (i.e. row,col values all match)
    for cnt, row in enumerate(table1_recs):
        if exact_rec_score == 4:
            if row != table2_recs[cnt]:
                exact_rec_score = 3
        # check row values match (i.e. col order doesn't matter) (allows for extra columns)
        if exact_rec_score == 3:
            if set(row).intersection(table2_recs[cnt]) != set(row):
                return 2
    return exact_rec_score

def AssessTableEntries(table1, table2):
    table1_recs = table1.GetRecords()
    table2_recs = table2.GetRecords()
    exact_rec_score = ExactRecordsMatch(table1_recs, table2_recs)
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
    global too_many_penalty
    name_score = row_count_score = col_count_score = field_name_score = field_type_score = field_size_score = \
        exact_rec_score = excess_fields = 0
    score_report = []
    score_report += ['{} TABLE\n'.format(table1.Name)]
    if table1.Name == table2.Name:
        name_score = 1
        score_report += ['\t-Table names match\n']
    else:
        score_report += ['\t-Table names DO NOT match\n\t\tSoln: {}\n\t\tStdnt: {}\n'.format(table1.Name, table2.Name)]
    if table1.RecordCount == table2.RecordCount:
        row_count_score = 1
    if table1.ColumnCount == table2.ColumnCount:
        col_count_score = 1
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
    if len(table2_fields) > len(table1_fields):
        excess_fields = len(table2_fields) - len(table1_fields)
    field_name_score *= (1-(excess_fields*too_many_penalty)) / len(table1_fields)
    if field_name_score == 1:
        score_report += ['\t-Fields match\n']
    else:
        score_report += ['\t-Fields DO NOT match\n\t\tSoln: {}\n\t\tStdnt: {}\n'.format(table1_fields, table2_fields)]
    field_type_score *= (1-(excess_fields*too_many_penalty)) / len(table1_types)
    if field_type_score == 1:
        score_report += ['\t-Field types match\n']
    else:
        score_report += ['\t-Field types DO NOT match\n\t\tSoln: {}\n\t\tStdnt: {}\n'.format(table1_sizes, table2_sizes)]
    field_size_score *= (1-(excess_fields*too_many_penalty)) / len(table1_sizes)
    if field_size_score == 1:
        score_report += ['\t-Field sizes match\n']
    else:
        score_report += ['\t-Field sizes DO NOT match\n\t\tSoln: {}\n\t\tStdnt: {}\n'.format(table1_fields, table2_fields)]
    # primary keys intersection returns primary keys in common between table1 and table2
    pk_same = len(set(table1.PrimaryKeys).intersection(table2.PrimaryKeys)) / len(table1.PrimaryKeys)
    # this finds number keys that the student (table2) has that are not in the solution (table1)
    pk_diff = len(set(table2.PrimaryKeys).difference(table1.PrimaryKeys))
    pk_same *= (1-(pk_diff*too_many_penalty))
    if pk_same == 1:
        score_report += ['\t-Primary keys match\n']
    else:
        score_report += ['\t-Prmary keys DO NOT match\n\t\tSoln: {}\n\t\tStdnt: {}\n'.format(table1.PrimaryKeys,
                                                                                            table2.PrimaryKeys)]
    correct_num_rltns, fld, rltd_tbl, rltd_fld, join, integrity = GradeRelationships(table1.ForeignKeys,
                                                                                     table2.ForeignKeys)
    if sum([fld, rltd_tbl, rltd_fld, join, integrity]) == 5:
        score_report += ['\t-Relationships match\n']
    else:
        score_report += ['\t-Relationships DO NOT match\n\t\tSoln: {}\n\t\tStdnt: {}\n'.format(table1.ForeignKeys,
                                                                                            table2.ForeignKeys)]
    if row_count_score:
        exact_rec_score = AssessTableEntries(table1, table2)
    if exact_rec_score:
        score_report += ['\t-Records match\n']
    else:
        score_report += ['\t-Records DO NOT match']
    exact_rec_score /= 4
    table_score = TableScore(name_score, row_count_score, col_count_score, field_name_score, field_type_score,
                             field_size_score, exact_rec_score, pk_same, pk_diff, correct_num_rltns, fld, rltd_tbl,
                             rltd_fld, join, integrity)
    # print(''.join(score_report))
    return table_score, score_report


def ScoreTable(assessed_table, score_vector=base_table_weight):
    table_score = 0
    for cnt in range(len(assessed_table)):
        table_score += assessed_table[cnt]*score_vector[cnt]
    return table_score


def ScoreQuery(assessed_query, score_vector=base_query_weight):
    query_score = 0
    for cnt in range(len(assessed_query)):
        if not isinstance(assessed_query[cnt], bool):
            query_score += assessed_query[cnt]*score_vector[cnt]
    return query_score


def GetNumberMatches(reference_list, list2, debug=True):
    count = 0
    matches = []
    copy_list = copy.deepcopy(reference_list)
    # if debug:
    #     print('LIST1:', reference_list)
    #     print('LIST2:', list2)
    for item in list2:
        if item in copy_list:
            count += 1
            matches.append(item)
            copy_list.remove(item)
    # if debug:
    #     print('Num Matches: {}\nMatches: {}'.format(count, matches))
    return count, matches


def CleanStatement(statement):
    clean = statement.strip().replace('(', '').replace(')', '').replace('Max', '').replace('Count', '')\
                             .replace('Min', '').replace('Avg', '').replace('Sum', '').replace('StDev', '')\
                             .replace('Var', '').replace('First', '').replace('Last', '')
    return clean

def GetFieldsFromCompoundField(compound_field):
    fields = []
    for field in compound_field.split('.'):
        if '(' in field:
            fields.append(field.split('(')[1])
        elif ')' in field:
            fields.append(field.split(')')[0])
        else:
            fields.append(field)
    return fields


def GetPenaltyMultiple(soln_list, student_list):
    global too_many_penalty
    penalty_multiple = 0
    num_in_soln = np.size(soln_list)
    num_in_student = np.size(student_list)
    if num_in_student > num_in_soln:
        penalty_multiple = too_many_penalty * (num_in_student - num_in_soln)
    if penalty_multiple > .9:
        penalty_multiple = .9
    return penalty_multiple, num_in_soln, num_in_student


# ========================== FOLLOWING FUNCTIONS USED TO ASSESS THE QUERY SQL STATEMENT ============================== #
# Generically, each access query has following rows: field, table, total, sort, criteria. Additionally, have to
# check if tables have correct relationships.
# NOTE: NEED TO ADD WAY TO CHECK IS SHOW BOX CHECKED -- THERE IS A HIDDEN TRUE/FALSE STATEMENT
'''-----------------------------------------------------------------------------------------------------------------'''
'''                         FOLLOWING FUNCTIONS USED TO ANALYZE 'SELECT' STATEMENT                                  '''


def AssessQuerySelect(soln_select, student_select, debug=True):
    if debug:
        print('\n\tASSESSING SELECT STATEMENT')
        print('\t\tSOLN: ', soln_select)
        print('\t\tSTUDENT: ', student_select)
    if soln_select == student_select:
        return 1, ['\tSELECT statements match\n']
    if student_select is None:
        return 0, ['\tSELECT statements DO NOT match\n\t\tNo student SELECT statement\n']
    # Stripping SELECT statement (this is specific to way Access stores as 'SELECT x, y,x\r'
    soln_fields = soln_select.strip('\r').split('SELECT ')[1].split(', ')
    student_fields = student_select.strip('\r').split('SELECT ')[1].split(', ')
    # Split elements on '.' (Access puts table on left of '.' and field name on right)
    soln_select_elements = []
    student_select_elements = []
    for compound_field in soln_fields:
        soln_select_elements += GetFieldsFromCompoundField(compound_field)
    for compound_field in student_fields:
        student_select_elements += GetFieldsFromCompoundField(compound_field)
    # Check to see how many field,table matches between two queries
    select_cnt, matches = GetNumberMatches(soln_select_elements, student_select_elements, debug)
    penalty_factor, num_elements, student_elements = GetPenaltyMultiple(soln_select_elements, student_select_elements)
    compare_ratio = (select_cnt / num_elements) * (1 - penalty_factor)  # penalty for choosing too much stuff
    select_report = []
    if compare_ratio >= 1:
        select_report += ['\tSELECT statements match\n']
    else:
        select_report += ['\tSELECT statements DO NOT match\n\t\tSOLN Select: {}\n\t\tSTDNT Select: {}\n\t\tMatching '
                          'elements: {}\n\t\tScore: {} Matches / {} Possible * {} Penalty Factor '
                          '= {}\n'.format(soln_select_elements, student_select_elements, matches, select_cnt,
                                          len(soln_select_elements), (1-penalty_factor), compare_ratio)]
    return compare_ratio, select_report


'''                                     END FROM STATEMENT ANALYSIS                                                 '''
'''-----------------------------------------------------------------------------------------------------------------'''


'''-----------------------------------------------------------------------------------------------------------------'''
'''                         FOLLOWING 4 FUNCTIONS USED TO ANALYZE 'FROM' STATEMENT                                  '''
# Purpose of this function is to return the tables, fields, and join types used in query
# statement is SQL FROM line with 'FROM' already stripped
def GetKeyFromElements(statement, debug=True):
    x = 0  # used for debugging purposes only
    all_joins = []  # list of all tables, fields, and join types in query. Initialized to empty list.
    cur_joins = [1]  # list of elements found in current parsing. Initialized to not empty for while loop.
    while len(cur_joins) > 0:
        # Use regular expression to find elements in statement of format:
        #  '<TableName1> <INNER|RIGHT|LEFT> JOIN <TableName2> ON <TableName1.FieldName1> = <TableName2.FieldName2>'
        cur_joins = re.findall(r'\(?\w+ \w+ JOIN \[?\w+\]? ON \w+\.\w+ = \w+\.\w+\)?', statement)
        # Below loop accounts for nesting elements.
        for join in cur_joins:
            # Replace nested elements. Have to do this to go 'up' hierarchy
            statement = statement.replace(join, 'BLAH'+str(x))
            if debug:  # print statement after replacing found elements
                print('De-nesting Iter {}: {}'.format(x+1, statement))
            # Use of x in the 'BLAH' replacement potentially helps with debugging. Otherwise not needed
            x += 1
            # Strip out key elements (i.e. table names, fields, and join types)
            key_elements = re.findall(r'\w+ JOIN', join)  # Find elements of format '<INNER|RIGHT|LEFT> JOIN'
            for element in re.findall(r'\w+\.\w+', join):  # Find elements of format '<TableName>.<FieldName>'
                key_elements += element.split('.')  # Split table name and field name and add to list
            # Add key elements from JOIN sub-statement to the master list of joins
            all_joins.append(key_elements)
    if debug:  # print found relationships
        for cnt, join in enumerate(all_joins):
            print('Relationship {}: {}'.format(cnt, join))
    return all_joins


# Check table relationships. If no relationship, add table name to list. If relationship, strip key elements
def BreakdownQueryFromStmt(from_statement, debug=True):
    # Stripping 'FROM' from statement to allow additional manipulation.
    statement1 = from_statement.strip('\r').split('FROM ')[1]
    stmt_relationships = []
    for sub_statmenet in statement1.split(', '):  # if no relationship, tables separated by commas
        if 'JOIN' not in sub_statmenet:  # if no relationship, no JOIN in statement
            stmt_relationships.append([sub_statmenet])
        else:  # if relationship exists, get key elements (tables, fields, relationship type)
            relationships = GetKeyFromElements(statement1, debug)
            for rltn in relationships:
                stmt_relationships.append(rltn)
    if debug:
        print(stmt_relationships)
    return stmt_relationships


# Compare all possible permutations and return the best possible value
def CompareStuff(soln_compare, student_compare, num_choose, debug=True):
    if debug:
        print('Comparing Stuff')
    best_comp = []
    best_comp_val = possible_elements = student_elements = 0
    # for item in soln_compare:
    #     possible_elements += len(item)
    # for item in student_compare:
    #     student_elements += len(item)
    for permute in itertools.permutations(soln_compare, num_choose):
        iter_score = 0
        for cnt, item in enumerate(student_compare):
            if cnt+1 <= len(permute):
                score, matches = GetNumberMatches(permute[cnt], item, debug)
                iter_score += score
        if iter_score > best_comp_val:
            best_comp_val = iter_score
            best_comp = permute
    penalty_factor, possible_elements, student_elements = GetPenaltyMultiple(soln_compare, student_compare)
    compare_ratio = (best_comp_val / possible_elements) * (1-penalty_factor)  # penalty for choosing too much stuff
    compare_report = '\n\t\tScore: {} Matches / {} Possible * {} Penalty = {}'.format(best_comp_val, possible_elements,
                                                                               1-penalty_factor, compare_ratio)
    if debug:
        print('Best comparison: {}'.format(best_comp))
        print('Raw comparison score: {}\t# possible elements: {}\t'
              '# student elements: {}'.format(best_comp_val, possible_elements, student_elements))
        print('Final score: {}'.format(compare_ratio))
    return compare_ratio, best_comp, compare_report


# The SQL FROM statement shows which tables were used in the query and the relationship between those tables
def AssessQueryFrom(soln_from_statement, student_from_statement, debug=True):
    if debug:
        print('\n\tASSESSING FROM STATEMENTS')
        print('\t\tSolution FROM Statement:', soln_from_statement)
        print('\t\tSolution FROM Statement:', student_from_statement)
    if soln_from_statement == student_from_statement:
        return 1, ['\tFROM statements match\n']
    if student_from_statement is None:
        return 0, ['\tFROM statements DO NOT match\n\t\tNo STDNT FROM statmenet\n']
    soln_relationships = BreakdownQueryFromStmt(soln_from_statement, debug)
    student_relationships = BreakdownQueryFromStmt(student_from_statement, debug)
    comp_results = CompareStuff(soln_relationships, student_relationships, len(soln_relationships), debug)
    if comp_results[0] >= 1:
        from_report = ['\tFROM statements match\n']
    else:
        from_report = ['\tFROM statements DO NOT match\n\t\tSOLN rltnships: {}\n\t\tSTDNT rltnships: '
                       '{}'.format(comp_results[1], student_relationships) + comp_results[2] + '\n']
    return comp_results[0], from_report


'''                                     END FROM STATEMENT ANALYSIS                                                 '''
'''-----------------------------------------------------------------------------------------------------------------'''


'''-----------------------------------------------------------------------------------------------------------------'''
'''                    FOLLOWING 3 FUNCTIONS USED TO ANALYZE 'AND' AND 'OR' CRITERIA                                '''
# This function recursively calls itself. Isolates each individual element in a conditional logic statement.
def GetConditionalElements(statement):
    #remove all paranthesis and totals key words from statement
    # temp_statement = ''.join(statement.split('(')).strip()
    # statement = ''.join(temp_statement.split(')')).strip()
    statement = CleanStatement(statement)
    # print(statement)
    elements = []
    # list of conditional statments we check for
    symbols = [' And ', ' Or ', '>=', '<=', '=', '>', '<', 'Between', 'Is Null']
    for symbol in symbols:
        if symbol in statement:
            temp_elements = statement.split(symbol)  # split statement on symbol
            if len(temp_elements) > 1:
                for cnt in range(len(temp_elements) - 1):
                    elements.append(symbol)  # append as many symbols as appear in statement
                for element in temp_elements:
                    elements += GetConditionalElements(element)  # recursively call function on each substatement
            break  # if found a symbol exit loop to prevent duplicates
    if not elements and statement:  # if elements list is empty and statement is not empty, add operand to list
        elements.append(statement)
    return elements


def BreakdownCriteriaStatement(full_statement):
    num_elements = num_stmts = 0
    complete_elements_list = []
    for OR_line in full_statement.split(' OR '):
        AND_line = OR_line.split(' AND ')
        line_elements_list = []
        for AND_stmt in AND_line:
            base_elements_list = GetConditionalElements(AND_stmt)
            num_elements += len(base_elements_list)
            num_stmts += 1
            line_elements_list += [base_elements_list]
        complete_elements_list.append(line_elements_list)
    return num_elements, num_stmts, complete_elements_list


def AssessQueryCriteria(soln_where, soln_having, student_where, student_having, debug=True):
    if debug:
        print('\n\tASSESSING WHERE/HAVING')
        print('SOLN WHERE:', soln_where)
        print('SOLN HAVING:', soln_having)
        print('STUDENT WHERE:', student_where)
        print('STUDENT HAVING:', student_having)
    if soln_where == student_where and soln_having == student_having:
        return 1, ['\tAND/OR statements match\n']
    if student_where is None and student_having is None:
        return 0, ['\tAND/OR statements DO NOT match\n\t\tNo STDNT AND/OR statmenet\n']
    # Stripping WHERE and HAVING statements (specific to way Access stores SQL statements)
    # Consider various situations
    if student_where is None and student_having is None:
        return 0
    if soln_where is not None and soln_having is None:
        soln_stripped_stmt = soln_where.strip().split('WHERE ')[1]
        if student_where is not None and student_having is None:  # compare where's
            student_stripped_stmt = student_where.strip().split('WHERE ')[1]
        if student_where is None and student_having is not None:  # compare where to have
            student_stripped_stmt = student_having.strip().split('HAVING ')[1]
        if student_where is not None and student_having is not None:  # just compare where's
            student_stripped_stmt = student_where.strip().split('WHERE ')[1] + ' OR ' + \
                                    student_having.strip().split('HAVING ')[1]
    if soln_where is None and soln_having is not None:
        soln_stripped_stmt = soln_having.strip().split('HAVING ')[1]
        if student_where is not None and student_having is None:  # compare having to where
            student_stripped_stmt = student_where.strip().split('WHERE ')[1]
        if student_where is None and student_having is not None:  # compare havings
            student_stripped_stmt = student_having.strip().split('HAVING ')[1]
        if student_where is not None and student_having is not None:  # tricky case
            student_stripped_stmt = student_where.strip().split('WHERE ')[1] + ' OR ' + \
                                    student_having.strip().split('HAVING ')[1]
    if soln_where is not None and soln_having is not None:
        # combine statments with an OR
        soln_stripped_stmt = soln_where.strip().split('WHERE ')[1] + ' OR ' + soln_having.strip().split('HAVING ')[1]
        if student_where is not None and student_having is None:  # compare having to where
            student_stripped_stmt = student_where.strip().split('WHERE ')[1]
        if student_where is None and student_having is not None:  # compare havings
            student_stripped_stmt = student_having.strip().split('HAVING ')[1]
        if student_where is not None and student_having is not None:  # combine both with an ' OR '
            student_stripped_stmt = student_where.strip().split('WHERE ')[1] + ' OR ' + \
                                    student_having.strip().split('HAVING ')[1]
    # 'OR' indicates criteria on separate lines so first split on 'OR'
    # 'AND' indicates criteria in separate fields so second split on 'AND'
    # 'And' or 'Or' in indicates criteria on the same field, so look at those last

    num_soln_elements, num_soln_stmts, soln_elements_list = BreakdownCriteriaStatement(soln_stripped_stmt)
    num_stdnt_elements, num_stdnt_stmts, stdnt_elements_list = BreakdownCriteriaStatement(student_stripped_stmt)
    final_list = []
    extra_stmt = 0
    for permute in list(itertools.permutations(stdnt_elements_list, len(stdnt_elements_list))):  #Permute student OR
        # print('PERMUTE: ', permute)
        permute_list = []
        for cnt, row in enumerate(permute):
            stmt_permute = []
            for cnt2, stmts in enumerate(list(itertools.permutations(row, len(row)))):  # Permute student AND
                # print('\tROW:', list(stmts))
                stmt_permute.append(list(stmts))
            permute_list.append(stmt_permute)
        # print('Permute list:', permute_list)  # create list of permuted AND/OR combinations
        final_list.append(list(itertools.product(*permute_list)))  # all combination taking one from each list
    # print('PERMUTED LIST')
    if num_stdnt_elements > num_soln_elements:
        extra_stmt = num_stdnt_elements - num_soln_elements
    best_score = 0
    best_match = ''
    for item in final_list:
        for item2 in item:
            temp_score = 0
            temp_list = []
            # print(item2)
            for cnt, item6 in enumerate(soln_elements_list):
                # print('SOLN Element:', item6, '\nSTDNT Element:', item2[cnt])
                if cnt + 1 <= len(item2):
                    for cnt2, item7 in enumerate(item6):
                        if cnt2+1 <= len(item2[cnt]):
                            num_matches, matches = GetNumberMatches(item7, item2[cnt][cnt2])
                            # print('MATCHES', matches, num_matches)
                            temp_score += num_matches
                            if num_matches > 0 and item2[cnt] not in temp_list:
                                temp_list.append(item2[cnt])
            if temp_score > best_score:
                best_score = temp_score
                best_match = temp_list
            # for item3 in item2:
            #     print(item3)
                # for item4 in item3:
                #     print(item4)
                    # print(item6)
    # print('BEST MATCH: {}\nBEST SCORE: {}'.format(best_match, best_score))

    final_criteria_score = (best_score / num_soln_elements) * (1 - (too_many_penalty * (extra_stmt)))
    if final_criteria_score >= 1:
        criteria_report = ['\tAND/OR statements match\n']
    else:
        criteria_report = ['\tAND/OR statements DO NOT match\n\t\tSOLN criteria: {}\n\t\tSTDNT criteria: {}\n\t\tBest match'
                       ': {}\n\t\tScore: {} Matching / {} Possible * {} Penalty = '
                       '{}\n'.format(soln_elements_list, stdnt_elements_list, best_match, best_score,
                                   num_soln_elements, (1 - (too_many_penalty * (extra_stmt))), final_criteria_score)]
    return final_criteria_score, criteria_report


'''                                     END CRITERIA ANALYSIS                                                       '''
'''-----------------------------------------------------------------------------------------------------------------'''

# Checks for correct relationships in query
def AssessQueryTotalsFunctions(soln_totals, student_totals, debug=True):
    if debug:
        print('\n\tASSESSING TOTALS STATEMENT')
        print('SOLN: ', soln_totals)
        print('STUDENT: ', student_totals)
    if student_totals is None:
        return 0
    # Stripping SELECT statement
    soln_fields = soln_totals.strip('\r').split('SELECT ')[1].split(', ')
    student_fields = student_totals.strip('\r').split('SELECT ')[1].split(', ')
    # See which statments have totals functions, then add them to list
    soln_totals_elements = []
    student_totals_elements = []
    for compound_field in soln_fields:
        if '(' in compound_field:
            temp_elements = []
            temp_elements.append(compound_field.split('(')[0])
            temp_elements += GetFieldsFromCompoundField(compound_field)
            soln_totals_elements.append(temp_elements)
    for compound_field in student_fields:
        if '(' in compound_field:
            temp_elements = []
            temp_elements.append(compound_field.split('(')[0])
            temp_elements += GetFieldsFromCompoundField(compound_field)
            student_totals_elements.append(temp_elements)
    # Check to see how many field,table matches between two queries
    # select_cnt, matches = GetNumberMatches(soln_select_elements, student_select_elements, debug)
    num_totals = len(soln_totals_elements)
    compare_ratio, best_match, report = CompareStuff(soln_totals_elements, student_totals_elements, num_totals, False)
    if debug:
        print('Solution Totals: {}'.format(soln_totals_elements))
        print('Student Totals: {}'.format(student_totals_elements))
        print('Best Match: {}'.format(best_match))
        print('# Correct: {}\t# Select: {}'.format(np.size(soln_totals_elements), np.size(best_match)))
    # penalty_factor, num_elements, student_elements = GetPenaltyMultiple(soln_select_elements, student_select_elements)
    # compare_ratio = (select_cnt / num_elements) * (1 - penalty_factor)  # penalty for choosing too much stuff
    # return compare_ratio
    return soln_totals_elements, student_totals_elements, best_match

# NOTE: This function is almsot verbatim same as AssessQuerySelect function; consider combining for efficiency?
def AssessQueryGroupby(soln_groupby, student_groupby, debug=True):
    if debug:
        print('\n\tASSESSING GROUP BY STATEMENT')
        print('SOLN: ', soln_groupby)
        print('STUDENT: ', student_groupby)
    if student_groupby is None:
        return 0
    # Stripping GROUP BY statement'
    soln_fields = soln_groupby.strip('\r').split('GROUP BY ')[1].split(', ')
    student_fields = student_groupby.strip('\r').split('GROUP BY ')[1].split(', ')
    # Split elements on '.' (Access puts table on left of '.' and field name on right)
    soln_groupby_elements = []
    student_groupby_elements = []
    soln_display_groupby = []
    stdnt_display_groupby = []
    for compound_field in soln_fields:
        soln_groupby_elements += GetFieldsFromCompoundField(compound_field)
        soln_display_groupby += [['GROUP BY'] + GetFieldsFromCompoundField(compound_field)]
    for compound_field in student_fields:
        student_groupby_elements += GetFieldsFromCompoundField(compound_field)
        stdnt_display_groupby += [['GROUP BY'] + GetFieldsFromCompoundField(compound_field)]
    # Check to see how many field,table matches between two queries
    groupby_cnt, matches = GetNumberMatches(soln_groupby_elements, student_groupby_elements, debug)
    compare_ratio, best_match, report = CompareStuff(soln_display_groupby, stdnt_display_groupby,
                                                     len(soln_display_groupby), False)
    if debug:
        print('Solution group by: {}'.format(soln_display_groupby))
        print('Student group by: {}'.format(stdnt_display_groupby))
        print('Best Match: {}'.format(best_match))
        print('# Correct: {}\t# Groupby: {}'.format(np.size(best_match), np.size(soln_display_groupby)))
    # penalty_factor, num_elements, student_elements = GetPenaltyMultiple(soln_groupby_elements, student_groupby_elements)
    # compare_ratio = (groupby_cnt / num_elements) * (1 - penalty_factor)  # penalty for choosing too much stuff
    # return compare_ratio
    return soln_display_groupby, stdnt_display_groupby, best_match

def AssessTotalsRow(soln_groupby, student_groupby, soln_select, student_select, debug=True):
    if soln_select == student_select and soln_groupby == student_groupby:
        return 1, ['\tTOTALS functions match\n']
    totals_score = 0
    totals_report = ''
    soln_groupby_elements = soln_totals_elements = stdnt_groupby_elements = stdnt_totals_elements = \
        best_groupby = best_totals = []
    groupby_penalty = totals_functions_exist = False
    if soln_groupby is not None:  # If there is a GROUP BY in solution
        soln_groupby_elements, stdnt_groupby_elements, best_groupby = AssessQueryGroupby(soln_groupby, student_groupby,
                                                                                         debug)
        totals_functions_exist = True
    if '(' in soln_select or ')' in soln_select:  # If there is a totals function in solution
        soln_totals_elements, stdnt_totals_elements, best_totals = AssessQueryTotalsFunctions(soln_select,
                                                                                                student_select, debug)
        totals_functions_exist = True
    if not totals_functions_exist:
        return 0, ''
    num_matches = np.size(best_groupby) + np.size(best_totals)
    num_possible = np.size(soln_groupby_elements) + np.size(soln_totals_elements)
    if len(soln_groupby_elements) > 0 or len(soln_totals_elements) > 0:
        totals_score = num_matches / num_possible
    # extra_stmts = np.size(soln_groupby_elements) + np.size(soln_totals_elements) - np.size(stdnt_groupby_elements) - \
    #               np.size(stdnt_totals_elements)  # not using so don't penalize twice (already penalized in select)
    if totals_score == 1:
        totals_report = ['\tTOTALS functions match\n']

    else:
        totals_report = ['\tTOTALS functions DO NOT match\n']
    totals_report[0] += '\t\tSOLN totals: {}\n\t\tSTDNT totals: {}\n\t\tBest match: {}\n\t\tScore: {} Matching / {} ' \
                        'Possible = {}\n\t\tNote -- No penalty for extra elements, since penalized in SELECT analysis' \
                        '.\n'.format(soln_groupby_elements+soln_totals_elements,
                                     stdnt_groupby_elements+stdnt_totals_elements, best_groupby+best_totals, num_matches,
                                     num_possible, totals_score)
    return totals_score, totals_report


# Need to add something for ascending vs descending
def AssessQuerySort(soln_sort, student_sort, debug=True):
    if debug:
        print('\n\tASSESSING SORT')
        print('Soln Sort:', soln_sort)
        print('Student Sort:', student_sort)
    if soln_sort == student_sort:
        return 1, ['\tORDER BY statements match\n']
    if student_sort is None:
        return 0, ['\tORDER BY statements DO NOT match\n\t\tNo STDNT ORDER BY statmenet\n']
    sort_score = order_score = direction_score = 0
    soln_sort = CleanStatement(soln_sort)
    student_sort = CleanStatement(student_sort)
    # Stripping ORDER BY statement (specific to way Access stores SQL statements)
    soln_stripped_sort = soln_sort.strip(';').split('ORDER BY ')[1].split(', ')
    student_stripped_sort = student_sort.strip(';').split('ORDER BY ')[1].split(', ')
    first_time_through_loop = True
    all_soln_elements = []
    all_stdnt_elements = []
    for cnt, soln_field in enumerate(soln_stripped_sort):
        soln_elements = soln_field.split(' DESC')
        if len(soln_elements) > 1:
            soln_elements[1] = 'DESC'
        all_soln_elements.append(soln_elements)
        for cnt2, student_field in enumerate(student_stripped_sort):
            student_elements = student_field.split(' DESC')
            if len(student_elements) > 1:
                student_elements[1] = 'DESC'
            if first_time_through_loop:
                all_stdnt_elements.append(student_elements)
            # print('Stdnts', student_elements)
            if soln_elements[0] in student_elements[0]:
                sort_score += 1
                if cnt == cnt2:
                    order_score += 1
                if len(soln_elements) == len(student_elements):
                    direction_score += 1
        first_time_through_loop = False
    if debug:
        print('SOLN elements:', all_soln_elements)
        print('STDNT elements:', all_stdnt_elements)
    num_elements = len(soln_stripped_sort)
    if debug:
        print('Fields Score: {}\nOrder score: {}\nDirection score: {}'.format(sort_score, order_score, direction_score))
    final_score = (sort_score + order_score + direction_score) / num_elements / 3
    if final_score >= 1:
        return 1, ['\tORDER BY statements match\n']
    else:
        report = ['\tORDER BY statements DO NOT match\n\t\tSOLN ordering: {}\n\t\tSTDNT ordering" {}\n\t\tSort fields: '
                  '{} / {}\tSort direction: {} / {}\tField ordering: {} / {}\n\t\tSort score: '
                  '{}'.format(all_soln_elements, all_stdnt_elements, sort_score, num_elements, direction_score,
                              num_elements, order_score, num_elements, final_score)]
        return final_score, report


def FindSubStatement(statement_list, substring):
    if statement_list is None:
        return None
    for substatement in statement_list:
        if substring in substatement:
            return substatement

###  Used to check if two SQL queries are the same.
###  query1: should be the SQL attribute from the Table class.
###  string: the exact SQL string from the 'answer' with a SELECT
###     FROM and WHERE on seperate lines. 
def AssessStringQuery(query1, string):
    SQL1_parts = query1.strip().strip().split('\r\n')
    SQL2_parts = string.strip().split('\n')
    return any(map(lambda x,y:y == x,SQL1_parts,SQL2_parts))


def AssessQuery(query1, query2, debug=False):
    if debug:
        print('ASSESSING QUERY')
    row_count_score = exact_rec_score = select_score = from_score = criteria_score \
                    = groupby_score = totals_score = sort_score = 0
    where_penalty = having_penalty = groupby_penalty = sort_penalty = False
    query_report = ['{} QUERY\n'.format(query1.Name)]
    if query1.SQL == query2.SQL:
        query_report += ['\tExact SQL match']
        if debug:
            print(''.join(query_report))
        return QueryScore(1, 1, 1, 1, 1, 1, where_penalty, having_penalty, groupby_penalty, sort_penalty, 4), \
               query_report

    if query1.RecordCount == query2.RecordCount:
        row_count_score = 1
    if row_count_score:
        exact_rec_score = AssessTableEntries(query1, query2)
    # If some variation of exact record match then return
    if exact_rec_score == 4:
        query_report += ['\tExact record match']
        if debug:
            print(''.join(query_report))
        return QueryScore(1, 1, 1, 1, 1, 1, where_penalty, having_penalty, groupby_penalty, sort_penalty, 4), \
               query_report
    SQL1_parts = query1.SQL.strip().split('\n')
    SQL2_parts = query2.SQL.strip().split('\n')
    # first element of any query SQL is the select statement, so see if they are selecting correct fields
    soln_criteria_statements = []
    student_criteria_statements = []

    # Assess the 'SELECT' statement
    soln_select = FindSubStatement(SQL1_parts, 'SELECT')
    student_select = FindSubStatement(SQL2_parts, 'SELECT')
    if soln_select is not None:  # If there is a SELECT in solution
        select_score, select_report = AssessQuerySelect(soln_select, student_select, debug)
        query_report += select_report

    # Assess the 'FROM' statement
    soln_from = FindSubStatement(SQL1_parts, 'FROM')
    student_from = FindSubStatement(SQL2_parts, 'FROM')
    if soln_from is not None:  # If there is a FROM in solution
        from_score, from_report = AssessQueryFrom(soln_from, student_from, debug)
        query_report += from_report

    # Assess 'WHERE' and 'HAVING' criteria
    soln_where = FindSubStatement(SQL1_parts, 'WHERE')
    soln_having = FindSubStatement(SQL1_parts, 'HAVING')
    student_where = FindSubStatement(SQL2_parts, 'WHERE')
    student_having = FindSubStatement(SQL2_parts, 'HAVING')
    if soln_where is not None or soln_having is not None:  # If there is WHERE or HAVING in solution, assess
        criteria_score, criteria_report = AssessQueryCriteria(soln_where, soln_having, student_where, student_having,
                                                              debug)
        query_report += criteria_report
    if soln_where is None and student_where is not None:
        where_penalty = True  # Penalty for using WHERE when not supposed to
    if soln_having is None and student_having is not None:
        having_penalty = True  # Penalty for using HAVING when not supposed to

    # Assess 'GROUPBY' and Totals functions
    soln_groupby = FindSubStatement(SQL1_parts, 'GROUP BY')
    student_groupby = FindSubStatement(SQL2_parts, 'GROUP BY')
    totals_score, totals_report = AssessTotalsRow(soln_groupby, student_groupby, soln_select, student_select, debug)
    if (soln_groupby is None and student_groupby is not None) or ('(' not in soln_select and '(' in student_select):
        groupby_penalty = True  # Penalty for using totals functions when not supposed to
    if len(totals_report) > 0:
        query_report += totals_report

    # Assess 'SORT'
    soln_sort = FindSubStatement(SQL1_parts, 'ORDER')
    student_sort = FindSubStatement(SQL2_parts, 'ORDER')
    if soln_sort is not None:  # If there is ORDER in solution, assess
            sort_score, sort_report = AssessQuerySort(soln_sort, student_sort, debug)
            query_report += sort_report
    if soln_sort is None and student_sort is not None:
        pass  # Penalty for sorting when not supposed to

    if debug:
        print('\nSELECT score: {}\nFROM score: {}\nWHERE/HAVING score: {}\nGROUP BY score: {}\nTOTALS score: {}'
          '\nSORT score: {}'.format(select_score, from_score, criteria_score, groupby_score, totals_score, sort_score))
        print('\n{}'.format(query1.SQL))
        print(query2.SQL)
    query_results = QueryScore(select_score, from_score, criteria_score, groupby_score, totals_score, sort_score, \
           where_penalty, having_penalty, groupby_penalty, sort_penalty, exact_rec_score)
    if debug:
        print(''.join(query_report))
    return query_results, query_report
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
    table.GetLookups('soldierTrained', debug=2)
    print()
    # print the properties for some metadata (e.g. Table, Query, or Field)
    # print('Table Properties')
    # ListProperties(table._TableMetaData)
    # print('\nField Properties')
    # ListProperties(field)
    # print(field.Properties['ColumnHidden'].Value)
    # print('\nQuery Properties')
    # ListProperties(SolnDB.Queries['APFTStars']._TableMetaData)

    table_assessment, report = AssessTables(SolnDB.Tables['SoldierCompletesTraining'],
                                    StudentDB.Tables['SoldierCompletesTraining'])
    # table_assessment = AssessTables(SolnDB.Tables['Platoon'], StudentDB.Tables['Platoon'])
    print()
    print('Comparing "SoldierCompletesTraining" tables...')
    # print(table_assessment)
    print(''.join(report))
    print('\nFinal Table Score: ', ScoreTable(table_assessment))
    query_assessment, q_report = AssessQuery(SolnDB.Queries['APFTStars'], StudentDB.Queries['APFTStars'])
    # query_assessment, q_report = AssessQuery(SolnDB.Queries['Junior25BList'], StudentDB.Queries['Junior25BList'])
    # query_assessment, q_report = AssessQuery(SolnDB.Queries['Max2017APFTScores'], StudentDB.Queries['Max2017APFTScores'],debug=False)
    # query_assessment, q_report = AssessQuery(SolnDB.Queries['MostRecentlyPromoted'], StudentDB.Queries['MostRecentlyPromoted'])
    # query_assessment, q_report = AssessQuery(SolnDB.Queries['Q42017Awards'], StudentDB.Queries['Q42017Awards'])
    # query_assessment, q_report = AssessQuery(SolnDB.Queries['SoldierNames'], StudentDB.Queries['SoldierNames'])
    # query_assessment, q_report = AssessQuery(SolnDB.Queries['SoldiersTrainedOnTARPandCRM'], StudentDB.Queries['SoldiersTrainedOnTARPandCRM'])
    # query_assessment, q_report = AssessQuery(SolnDB.Queries['UntrainedLeaders'], StudentDB.Queries['UntrainedLeaders'])
    print()
    print(''.join(q_report))
    print('\nFinal Query Score: ', ScoreQuery(query_assessment))

if __name__ == "__main__":
    main()
