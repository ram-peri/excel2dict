#!/usr/bin/python

#Author: Ram Peri
#Email: <peri> <dot> <ram> <at> <gmail> <dot> <com>

import re, os, pdb, xlrd, json, datetime, logging

s_logger = logging.getLogger('root')

s_sheet_names = [u'Oxford Direct']

s_tables = {
'key1' : {
        'sheet_name' : 'Shee1',
        'table_start' :  'A1',
        'table_direction' : 	'LeftToRight',
        'table_max' :  50,
        'table_columns' : ['col1', 'col2',
                           '', '',
                           'col5', 'col6',
                           'formula:row_counter=row_counter+2']},
}

class ExcelToErpTable:
    def __init__(self, table_name, entry):
        self.m_table_name            = table_name
        self.m_sheet_name            = entry['sheet_name']
        self.m_table_start           = entry['table_start']
        self.m_table_direction       = entry['table_direction']
        self.m_table_max             = entry['table_max']
        self.m_table_columns         = entry['table_columns']
        self.m_is_top_level          = entry.get('is_top_level', False)
        self.m_merge_columns         = entry.get('merge_columns', False)
        self.m_relative_table_start   = entry.get('relative_table_start', {})

        self.m_first_nonempty_column = next(i for i, j in enumerate(self.m_table_columns) if j)
        
        s_logger.debug('First non empty column for table[%s] is [%d]', self.m_table_name, self.m_first_nonempty_column)
        
        if self.m_first_nonempty_column > len(self.m_table_columns):
            s_logger.error('There are no non empty fields in [%s]', self.m_table_columns)

    def __print__(self):
        res = ""
        res += self.m_table_name + " : \n" 
        res += "Is top?     : " + str(self.m_is_top_level) + "\n" 
        res += "Sheet       : " + self.m_sheet_name + "\n" 
        res += "Table Start : " + self.m_table_start + "\n" 
        res += "Table Dir   : " + self.m_table_direction + "\n" 
        res += "Table Max   : " + str(self.m_table_max) + "\n" 
        res += "Table Cols  : " + str(self.m_table_columns) + "\n" 
        return res
            
    def __repr__(self):
        return self.__print__()


class ExcelToErpTables:
    s_excel_erp_tables = None

    def __init__(self, tables, logger):
        self.m_tables = {}
        for table_name, table_entry in tables.iteritems():
            self.m_tables[table_name] = ExcelToErpTable(table_name, table_entry)

        self.m_table_recursion_counter = 0
        self.m_cell_re                    = re.compile("([A-Z]*)([0-9]*)")

    def __print__(self):
        res = "Tables:"
        for (table_name, table) in self.m_tables.iteritems():
            res += "\t\t" + str(table)
        return res

    def pretty(self, d, indent=0):
        res = ""
        if isinstance(d, dict):
            res += '\t' * indent + "{" 
            for key, value in d.iteritems():
                res += '\t' * indent + (("%30s: %s\n") % (str(key), self.pretty(value, indent+1)) )
            res += '\t' * (indent+4) + '}\n'
        elif isinstance(d, list) or isinstance(d, tuple):
            for value in d:
                res+= self.pretty(value, indent+1)
        else:
            res += str(d)
        return res

    def __repr__(self):
        return self.__print__()
            
    def parse_value(self, cell, field_name):
        val = cell.value
        if cell.ctype == xlrd.XL_CELL_DATE:
            try:
                val = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(val)) #+ 1462 * 0))
                val = val.strftime('%Y-%m-%d')
            except Exception, e:
                s_logger.warning('Error while processing date field[%s]', field_name)
                val = None
        elif '_date' in field_name.lower() or 'date_' in field_name.lower():
            s_logger.warn( 'Non date type field:[%s] value[%s]', field_name, val)
            val = val.strip()
            d1 = None
            try:
                d1  = datetime.datetime.strptime(val, "%m/%d/%Y")
            except:
                try:
                    d1 = datetime.datetime.strptime(val, "%m/%d/%y")
                except:                        
                    try:
                        d1 = datetime.datetime.strptime(val, "%Y-%m-%d")
                    except:
                        pass
            if not d1:
                return None
            val = d1.strftime("%Y-%m-%d")
            s_logger.debug('New date value [%s]', val)
        else:
            if val == u'Yes' : val = 1
            elif val == u'No' : val = 0
        return val
    
    def get_cell_coordinates(self, name):
        if not name:
            return (-1, -1)
        m = self.m_cell_re.match(name)
        (col_name, row) = m.group(1), int(m.group(2)) -1
        col = reduce(lambda s,a:s*26+ord(a)-ord('A')+1, col_name, 0)
        return (row, col-1)

    def get_val(self, sheet, cell_name, default=None):
        (row, col) = self.get_cell_coordinates(cell_name)
        #self.m_logger.debug( 'Cell Name[%s], row[%d], col[%d] cellname[%s]', cell_name, row, col, xlrd.cellname(row, col))
        val = sheet.cell(row, col)
        if val: return val
        return default

    def process_table(self, sheets, excel_table, new_start=None):
        self.m_table_recursion_counter+=1
        table_start = excel_table.m_table_start
        if new_start:
            s_logger.info('New start [%s], will override tablestart[%s]', new_start, excel_table.m_table_start)
            table_start = new_start
        
        s_logger.debug('PROC: Table[%s], sheet[%s], start[%s], dir[%s], max[%s] recursion[%d]',
                       excel_table.m_table_name, excel_table.m_sheet_name, table_start,
                       excel_table.m_table_direction, excel_table.m_table_max, self.m_table_recursion_counter)

        (row, col) = self.get_cell_coordinates(table_start)

        if not excel_table.m_sheet_name in sheets:
            s_logger.warn('Could not find sheet name[%s], moving on ...', excel_table.m_sheet_name)
            self.m_table_recursion_counter -= 1
            return []
        
        sheet = sheets[excel_table.m_sheet_name]
        entries = []
        row_counter = 0
        found_entry = True
                
        while row_counter < excel_table.m_table_max and found_entry:
            #pdb.set_trace()
            s_logger.debug('Row %d/%d in table %s', row_counter, excel_table.m_table_max, excel_table.m_table_name)
            new_entry = {}
            col_counter = -1

            for field_name in excel_table.m_table_columns:
                if isinstance(field_name, basestring) and field_name.startswith('fixed_column:'):
                    field_name = field_name[13:]
                    fixed_column = True
                    #Do not advance the cursor for fixed_column
                else:
                    col_counter = col_counter + 1
                    fixed_column = False
                
                #Skip empty fields
                if field_name == '':
                    continue
                
                s_logger.debug('Processing field :%s: of %s', str(field_name), str(excel_table.m_table_columns))


                if excel_table.m_table_direction == 'TopToBottom':
                    rowno = row+row_counter
                    if fixed_column:
                        (orig_row, orig_col) = self.get_cell_coordinates(excel_table.m_table_start)
                        colno = orig_col
                    else:
                        colno = col+col_counter
                else:
                    colno = col + row_counter
                    if fixed_column:
                        (orig_row, orig_col) = self.get_cell_coordinates(excel_table.m_table_start)
                        rowno = orig_row
                    else:
                        rowno = row + col_counter
                
                cell_name = xlrd.cellname(rowno, colno)
                
                if isinstance(field_name, list):
                    new_table_entry = self.m_tables.get(field_name[0], None)
                    if not new_table_entry:
                        s_logger.error('Unable to find table[%s] in global tables, need to skip it ...', field_name[0])
                        continue

                    relative_table_start = None
                    if field_name[0] in excel_table.m_relative_table_start:
                        s_logger.info('Found relative start for %s in %s', field_name[0], excel_table.m_relative_table_start)
                        relative_table_start = cell_name
                    
                    val = self.process_table(sheets, new_table_entry, relative_table_start)

                    #For a table, we break when we get back an empty table
                    #used to be, now we continue processing.
                    if col_counter == excel_table.m_first_nonempty_column and not val and relative_table_start:
                        s_logger.debug('Found empty table at (row:%d, col:%d), [%s], ending parent table',
                                       rowno, colno, xlrd.cellname(rowno, colno)) 
                        found_entry = False
                        break
                else:
                    #Process formula fields
                    if field_name.startswith('formula:'):
                        s_logger.debug('Processing formula[%s] Before row[%d], col[%d] [%s]',
                                            field_name[8:], row_counter, col_counter, xlrd.cellname(rowno, colno))
                        exec(field_name[8:])
                        s_logger.debug('After row[%d], col[%d] [%s]',
                                            row_counter, col_counter, xlrd.cellname(rowno, colno))
                        continue
                    
                    #Check if we found an empty row, i.e. we have reached the end of data
                    if rowno >= sheet.nrows or colno >= sheet.ncols:
                        s_logger.debug('We have reached the end of this sheet, nothing more to process')
                        found_entry = False
                        break
                    
                    if col_counter == excel_table.m_first_nonempty_column and sheet.cell_type(rowno, colno) in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                        s_logger.debug('Found empty entry at (row:%d, col:%d) [%s], end of table',
                                       rowno, colno, xlrd.cellname(rowno, colno)) 
                        found_entry = False
                        break
                    reason = ""
                    cell  = sheet.cell(rowno, colno)
                    
                    try:
                        val   = self.parse_value(cell, field_name)
                    except Exception, e:
                        error   =   'Error while processing sheet[%s], cell[%s], for field[%s]' % (excel_table.m_sheet_name, cell, field_name)
                        s_logger.error(error)
                        self.m_errors.append(error)
                        if len(self.m_errors) < 100:
                            continue
                        else:
                            raise
                if val != None and val != '':
                    if isinstance(field_name, list):  #Yo, this is a group, treat it accordingly
                        if excel_table.m_merge_columns:
                            s_logger.debug('val is[%s]', val)
                            if val and len(val[0]) >= 2:
                                new_entry.update(val[0][2])  #We assume there is only one value for children when using merge columns
                        else:
                            new_entry[field_name[0]] = val
                    else:
                        new_entry[field_name] = val
                else:
                    self.m_errors += "Error on sheet[%s] Cell[%s]\n" % (excel_table.m_sheet_name, cell_name)
                s_logger.debug( "Sheet Name[%s], Cell name[%s], Field name[%s], Value[%s]", excel_table.m_sheet_name, cell_name, field_name, val)
            
            s_logger.info('Done processing fields for table[%s] row[%d] value[%s]', excel_table.m_table_name, row_counter, new_entry)
            
            if excel_table.m_is_top_level:
                entries = new_entry
            elif len(new_entry) > 0 and found_entry:
                entries.append((0,0,new_entry))
            else:
                pass #This is the end of child table
            row_counter +=  1
        
        if excel_table.m_table_max > 1 and row_counter >= excel_table.m_table_max:
            s_logger.warn('Table [%s], has exceeded the maximum number of entries, will bailing out', excel_table.m_table_name)
        
        self.m_table_recursion_counter -= 1

        return entries

    def process_file(self, filename, errors, email_text):
        self.m_errors = ""
        self.m_processing_successful = False

        wb = xlrd.open_workbook(filename, xlrd.Book.datemode)

        sheets = {}

        #Only look for sheets we know about.
        for sheet_name in s_sheet_names:
            try:
                sheets[sheet_name] = wb.sheet_by_name(sheet_name)
            except xlrd.biffh.XLRDError, e:
                pass
        
        #If this is a provider quote, ignore all other sheets.
        process_type1 = True
        customer_info = { }
        for x in ['known file1', 'known file2', 'known file3']:
            t_sheet_name = self.m_tables[x].m_sheet_name
            if t_sheet_name in sheets:
                sheets =  { t_sheet_name : sheets[t_sheet_name] }
                excel_table = self.m_tables[x]
                process_type1 = False
                break
        if process_type1:
            excel_table = self.m_tables['type1']
            customer_info = { 'is_company' : 1, 'customer' : 1, 'supplier' : 0 }
            if s_general_information_sheet_name not in sheets:
                error = 'Could not find sheet %s in TYPE1, cannot process' % s_general_information_sheet_name
                errors.append(error)
                s_logger.error(error)
                return {}
        
        s_logger.info( 'sheets is: %s, top level table is[%s]', sheets, excel_table.m_table_name)        
        customer_info.update(self.process_table(sheets, excel_table))
        
        s_logger.info( 'FINAL:\n%s\n', str(customer_info))
        self.m_processing_successful = True
        return customer_info


ExcelToErpTables.s_excel_erp_tables = ExcelToErpTables(s_tables, None)

s_results = {
    "output.xls" : {},
}


def test_sheets():
    res =True
    tests_base = "/tmp"
    for filename in s_results.keys():
        customer_info = ExcelToErpTables.s_excel_erp_tables.process_file(os.path.join(tests_base, filename))
        customer_info_str = str(customer_info)
        expected_result = s_results[filename]
        expected_result_str = str(expected_result)
        if customer_info != expected_result:
            res = False
            s_logger.error('Result---%s--- does not match expected_result---%s----', customer_info_str[:min(len(customer_info_str), 100)],
                           expected_result_str[:min(len(expected_result_str), 100)])
    return res

#Test routines
#TODO: 
#1. Store the hash table and run comparison
#2. Add tests for all sub routines
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO,
                        format='%(asctime)s %(levelname)s %(message)s')
    filename = "~/sample.xls"
    ExcelToErpTables.s_excel_erp_tables.process_file(filename, None, None)
