#! /usr/bin/env python

# Author: Urucas <info@urucas.com>
# Version: 1.0.0
class xls2json:

    yeison = None

    def parse(self, xlsPath, hasKeys, jsonPath):

        import xlrd
        wb = xlrd.open_workbook(xlsPath)
        sh = wb.sheet_by_index(0)

        if hasKeys:
            data = self.parseRowsWithFieldKeys(sh)
        else:
            data = self.parseRows(sh)

        import json
        self.yeison = json.dumps(data)
        if jsonPath == None:
            return self.yeison
           
        with open(jsonPath, 'w') as f:
            f.write(self.yeison)

    def parseRows(self, sheet):
        rows = []
        for i in range(0, sheet.nrows):
            row_values = sheet.row_values(i)
            row = []
            import unicodedata
            for val in row_values:
                val = unicodedata.normalize('NFKD', unicode(val)).encode('ascii','ignore') 
                row.append(val)
            rows.append(row)
        return rows    

    
    def parseRowsWithFieldKeys(self, sheet):

        rows = []
        keys = sheet.row_values(0)
        import unicodedata
        for i in range(1, len(keys)):
            val = unicodedata.normalize('NFKD', unicode(keys[i])).encode('ascii','ignore')  
            keys[i] = val

        for i in range(1, sheet.nrows):
            row_values = sheet.row_values(i)
            _row = {}
            for j in range(0, len(keys)):
                val = unicodedata.normalize('NFKD', unicode(row_values[j])).encode('ascii','ignore')  
                _row[keys[j]] = val
            rows.append(_row)
        return rows    


    def getJSON(self):
        return self.yeison

def main():
    
    import argparse
    parser = argparse.ArgumentParser(description="xls2jsonpie is a simple python command-libe tool to convert your xls files to json.")
    parser.add_argument('--xls', help='xls file path', type=str, default=None)
    parser.add_argument('--hasFieldKeys', help="Set a boolean param to check if the first row of the sheet has the field keys", type=bool, default=False)
    parser.add_argument('--json', help='json file path to dump the data', type=str, default=None)

    args = parser.parse_args()
    import os
    if args.xls == None or not os.path.exists(args.xls):
       die("xls path not found")
        
    x2j = xls2json()
    x2j.parse(args.xls, args.hasFieldKeys, args.json)

def die(msg):
    print msg
    import sys
    sys.exit(1)

if __name__ == '__main__':
    main()

