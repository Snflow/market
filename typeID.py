import simplejson
import xlrd
from pprint import pprint

file = xlrd.open_workbook('invTypes.xls')
sh = file.sheet_by_index(0)

outfile = open('data','w')
outfile.write('[')

#sh.nrows: the number of rows in the sheet
for i in range(2,sh.nrows+1):
    line = i-1
    if (sh.cell(line,11).ctype == 2 and sh.cell(line,9).value != 0):
        unit_typeID = sh.cell(i-1,0).value
        unit_typeID = int(unit_typeID)
        unit_name = sh.cell(i-1,2).value
        unit = {'ID': unit_typeID, 'name': unit_name}
        unit_obj = simplejson.dumps(unit)
        unit_json = simplejson.loads(unit_obj)
        print unit_json["ID"]
        outfile.write(unit_obj+',')

end = {'ID': 'end'}
end_obj = simplejson.dumps(end)
outfile.write(end_obj+']')
outfile.close()

print "Total rows in the sheet:", sh.nrows
