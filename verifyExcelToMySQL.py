from os.path import join, dirname, abspath
import pymysql
import xlrd, datetime
import timeit

# Set system encoding to UTF-8 to avoid 
# potential error on some characters
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

# Connect to the MySQL database
db = pymysql.connect("localhost","root","m2860321","excelData")
cursor = db.cursor()

# Read in Excel file and Setup variables
fname = join(dirname(dirname(abspath(__file__))), 'test_data', 'sale.xls')
xl_workbook = xlrd.open_workbook(fname)
# get worksheets
sheet_names = xl_workbook.sheet_names()
# use the first worksheet
xl_sheet = xl_workbook.sheet_by_name(sheet_names[0])

# Start verifying the database 
errorCount = 0
# set up timer
start_time = timeit.default_timer()
for i in range(1, xl_sheet.nrows):
    Row_ID = str(int(xl_sheet.row(i)[0].value))
    sql = "select * from saleData WHERE Row_ID = " + Row_ID
    cursor.execute(sql)
    for row in cursor:
        for j in range(1, len(row)):
            # for date from Excel, convert it to desired format
            if j == 17: continue
            if j == 2 or j == 20:
                a1 = xl_sheet.row(i)[j].value
                a1_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(a1, xl_workbook.datemode))
                fromExcel = a1_as_datetime.date().isoformat()
            else:
                fromMySQL = row[j]
                # make sure same format and types are compared
                if '.' in fromMySQL:
                    fromMySQL = int(float(fromMySQL))
                    fromExcel = int(xl_sheet.row(i)[j].value)
                elif fromMySQL.isdigit():
                    fromExcel = str(int(xl_sheet.row(i)[j].value))
                else:
                    fromExcel = str(xl_sheet.row(i)[j].value)
                if fromExcel != fromMySQL:
                    errorCount += 1
                    print 'failed ' + str(errorCount) + " case"
    print "checked " + str(i) + " records"
elapsed = timeit.default_timer() - start_time
print "finished in " + str(elapsed) + "s"
print str(errorCount) + " failed"