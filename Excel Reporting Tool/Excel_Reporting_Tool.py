import openpyxl
import xml
import itertools
from datetime import date, datetime
from dateutil.parser import parse
import re

class Report_Generator():

    def Load(self):

        #Loads the xml file. Controls the layout of the html
        xml_tree = xml.etree.ElementTree.parse("xml_test.xml")
        xml_root = xml_tree.getroot()
        xml_iter = xml_tree.iter()

        #Loads the excel file into memory
        self.excel_workbook = openpyxl.load_workbook("Money Model.xlsm",data_only=True)

        self.parse_instructions(xml_iter)

    def parse_instructions(self,xml_iter):

        html_output = open("report.html","w")

        report = """<html>
        <head></head>"""

        for node in xml_iter:
            if node.tag=="txt":
                report += "\n" + "<body><p>" + node.text + "</p></body>"
        
            elif node.tag=="SQL":
                report +="\n" + "<body>" + self.parse_SQL(node.text) + "</body>"
                print("FINISHED LOGIC")

        report += "</html>"
        html_output.write(report)
        print("WRITTEN REPORT")
        html_output.close()

    def parse_SQL(self,SQL):

        SQL_Operators = ['SELECT','FROM','WHERE','AND','SUM','AVG','COUNT','MAX','MIN']
        SQL_Maths_Operators = ['SUM','AVG','COUNT','MAX','MIN']
        SQL_Command_Pos = []     
        SQL_Command_Dict = {}

        SQL_No_Line = SQL.replace("\n"," ").replace('\r',' ') + " "

        for string in SQL_Operators:
            for m in re.finditer(string, SQL_No_Line):
                SQL_Command_Pos.append([m.start(),m.end(),string])
        SQL_Command_Pos.sort()

        for i in range(0,len(SQL_Command_Pos)): SQL_Command_Dict[SQL_Command_Pos[i][2]] = []

        for i in range(0,len(SQL_Command_Pos)):
            pm=''
            count=0         
            for m in re.finditer(' ', SQL_No_Line):
                if i < (len(SQL_Command_Pos)-1):
                    if m.start() >= SQL_Command_Pos[i][1] and m.end()<=SQL_Command_Pos[i+1][0]:
                        chunk_string = " ".join(self.get_string(SQL_No_Line,count,m,pm,SQL_Command_Pos[i][1]).split())
                        if chunk_string != SQL_Command_Pos[i][2]:
                            SQL_Command_Dict[SQL_Command_Pos[i][2]].append(self.get_string(SQL_No_Line,count,m,pm,SQL_Command_Pos[i][1]))
                elif i == (len(SQL_Command_Pos)-1):
                    if m.start()>= SQL_Command_Pos[i][1]:
                        chunk_string = " ".join( self.get_string(SQL_No_Line,count,m,pm,SQL_Command_Pos[i][1]).split())
                        if chunk_string != SQL_Command_Pos[i][2]:
                            SQL_Command_Dict[SQL_Command_Pos[i][2]].append(self.get_string(SQL_No_Line,count,m,pm,SQL_Command_Pos[i][1]))
                    count = 0
                pm=m
                count +=1

        SQL_Select = []
        SQL_Math = []
        SQL_Where = []
        SQL_New_Math = []

        current_key = 'SELECT'
        for i in range(1,len(SQL_Command_Dict[current_key])):
            if "," in SQL_Command_Dict[current_key][i] or i == len(SQL_Command_Dict[current_key])-1:
                SQL_Select.append([SQL_Command_Dict[current_key][i-1].replace("_"," "), SQL_Command_Dict[current_key][i].strip(",")])

         #Maths Selectors      
        for i in range(0,len(SQL_Maths_Operators)):
            if SQL_Maths_Operators[i] in SQL_Command_Dict.keys():
                for j in  range(0,len(SQL_Command_Dict[SQL_Maths_Operators[i]])):
                    SQL_Math.append([SQL_Command_Dict[SQL_Maths_Operators[i]][j].strip('()'),SQL_Maths_Operators[i]])
                #del SQL_Command_Dict[SQL_Maths_Operators[i]]
        to_be_removed = []

        for i in range(0,len(SQL_Math)):
            print("LOOPING")
            if i != len(SQL_Math)-1:
                print("NOT AT THE END")
                if SQL_Math[i][0] in SQL_Math[i+1]:
                    print('FOUND A DUPE')
                    SQL_Math[i].append(SQL_Math[i+1][1])
                    to_be_removed.append(i+1)

        for i in range(0,len(to_be_removed)):
            del SQL_Math[to_be_removed[i]-i]

        current_key = 'AND'
        for i in range(0,len(SQL_Command_Dict[current_key])):
            if SQL_Command_Dict[current_key][i] == '=':
                if i+1 == len(SQL_Command_Dict[current_key])-1:
                    SQL_Where.append([SQL_Command_Dict[current_key][i-1],SQL_Command_Dict[current_key][i],SQL_Command_Dict[current_key][i+1].strip("'")])
                else:
                    contains = True
                    if "'" in SQL_Command_Dict[current_key][i+1] and SQL_Command_Dict[current_key][i+1].index("'") ==0 :
                        SQL_Where.append([SQL_Command_Dict[current_key][i-1],SQL_Command_Dict[current_key][i],' '.join([SQL_Command_Dict[current_key][i+1],SQL_Command_Dict[current_key][i+2]]).strip("'")])
            elif  SQL_Command_Dict[current_key][i] == "BETWEEN":
                SQL_Where.append([ SQL_Command_Dict[current_key][i-1],SQL_Command_Dict[current_key][i], SQL_Command_Dict[current_key][i+1], SQL_Command_Dict[current_key][i+2]])

        SQL_Command_Dict['SELECT'] = SQL_Select
        SQL_Command_Dict['MATHS'] = SQL_Math
        SQL_Command_Dict['WHERE'] = SQL_Where

        del SQL_Math
        del SQL_Select
        del SQL_Where
        del SQL_Command_Dict['AND']

        active_worksheet = self.excel_workbook[SQL_Command_Dict['FROM'][0]]

        current_key = 'SELECT'
        self.get_column(SQL_Command_Dict,active_worksheet,current_key,'WHERE')
        current_key = 'MATHS'
        self.get_column(SQL_Command_Dict,active_worksheet,current_key,'MATHS')
        


        print(SQL_Command_Dict)
        
        return self.SQL_Logic(active_worksheet,SQL_Command_Dict)

    def SQL_Logic(self,active_worksheet,SQL_Command_Dict):
        
        Output_Sum = 0
        Output_List = []
        Min_Max_Count = []
        Output_String = ''

        #Logic Bit of the code
        #First loop goes through all of the rows in the spreadsheet

        for i in range(1,active_worksheet.max_row):
            count = 0
            #This goes through the commands given by the SQL
            for sublist in SQL_Command_Dict['WHERE']:
                sheet_data = active_worksheet.cell(row=i,column=sublist[-1]).value
                #Checks to see if the cell value is a datetime or not. Following operations can only be done on datetime objects
                if sublist[1] == 'BETWEEN':
                    if type(sheet_data) is datetime and self.validate_date(sublist[2]):
                            date1 = sheet_data.date()
                            #Converts the two dates provided by SQL_Where into datetime objects so they can be compared to the cell datetime
                            str_to_date1 =  self.to_date(sublist[2])
                            str_to_date2 = self.to_date(sublist[3])
                            if str_to_date1<date1 and date1<str_to_date2:
                                #If matches the conditions adds 1 to the count so that the final check knows if all checks have passed for a row.
                                count += 1
                elif sublist[1]== '=':
                    if sheet_data == sublist[2]:
                        count +=1

            #If all the checks have passed, the counter will equal the total number of clauses given. Will then add data from the column specified by SQL_Sun_Index
            if count == len(SQL_Command_Dict['WHERE']):
               
                flat_list = []

                for j in range(0,len(SQL_Command_Dict['WHERE'])):
                    sheet_value = active_worksheet.cell(row=i,column=SQL_Command_Dict['WHERE'][j][-1]).value
                    if type(sheet_value) is datetime:
                        flat_list.append(datetime.strftime(sheet_value,'%d-%m-%Y'))
                    else:
                        flat_list.append(sheet_value)
                for sublist in SQL_Command_Dict['MATHS']:
                        flat_list.append('£'+ str(active_worksheet.cell(row=i,column=sublist[-1]).value))
                        Min_Max_Count.append(active_worksheet.cell(row=i,column=sublist[-1]).value)
                Output_List.append(flat_list)
        
        print(Output_List)
        checked_list = []
        for sublist in Output_List:
            for i in range(0,len(sublist)):
                print(i)
                if sublist[i] not in checked_list or "£" in sublist[i]:
                    checked_list.append(sublist[i])
                else:
                    sublist[i] = ""
        del checked_list

        Output_String += '<table> \n<tr>'
        for substring in SQL_Command_Dict['WHERE']:
            Output_String +=  '\n<th>' + substring[0] + '</th>'

        Output_String += '</tr>'

        for sublist in Output_List:
            Output_String += '<tr><td>' + '</td><td>'.join(sublist) + '</td></tr>'
        
        Output_String += '</table>'

        for sublist in SQL_Command_Dict['MATHS']:
            if 'MAX' in sublist: Output_String += '\n <p>Max Value : ' + str(max(Min_Max_Count))+ '</p>'
            if 'MIN' in sublist: Output_String += '\n <p>Min Value : ' + str(min(Min_Max_Count))+ '</p>'
            if 'AVG' in sublist: Output_String += '\n <p>Average Value : ' + str(sum(Min_Max_Count)/len(Min_Max_Count))+ '</p>'
            if 'SUM' in sublist: Output_String += '\n <p>Total Cost : ' + str(sum(Min_Max_Count))+ '</p>'
            if 'COUNT' in sublist: Output_String += '\n <p>Total Number of Elements : ' + str(len(Min_Max_Count))+ '</p>'
            

        return  Output_String

    def to_date(self,str_to_date):
        return datetime.strptime(str_to_date, '%d/%m/%Y').date()

    def get_string(self,string,count,m,pm,SQL_Command_Pos):
        if count == 0:
            return (string[SQL_Command_Pos:m.start()])
        else:
            return (string[pm.end():m.start()])

    def validate_date(self,test):
        try:
           datetime.strptime(test,'%d/%m/%Y')
        except ValueError:
            return False
        else:
            return True

    def get_column(self,dict,active_worksheet,current_key,other_key):
        for i,j in itertools.product(range(1,13),range(0,len(dict[current_key]))):
            if active_worksheet.cell(row=1,column=i).value == dict[current_key][j][0]:
                for x in range(len(dict[other_key])):
                    if dict[current_key][j][0] or dict[current_key][j][1] in dict[other_key][i]:
                        dict[other_key][j].append(i)

if __name__ == '__main__':
    Report_Generator().Load()
