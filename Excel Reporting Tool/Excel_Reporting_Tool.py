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


        report += "</html>"
        html_output.write(report)
        html_output.close()

    def parse_SQL(self,SQL):

        SQL_Operators = ['SELECT','FROM','WHERE','AND','SUM','AVG','COUNT','MAX','MIN']
        SQL_Maths_Operators = ['SUM','AVG','COUNT','MAX','MIN']
        SQL_Command_Pos = []
        
        SQL_Command_Dict = {}

        SQL_Sheet = ""
        SQL_Sum = ""
        SQL_Sum_index =1
        Output_String = ""

        SQL_No_Line = SQL.replace("\n"," ").replace('\r',' ') + " "

        for string in SQL_Operators:
            for m in re.finditer(string, SQL_No_Line):
                SQL_Command_Pos.append([m.start(),m.end(),string])

        SQL_Command_Pos.sort()

        for i in range(0,len(SQL_Command_Pos)): SQL_Command_Dict[SQL_Command_Pos[i][2]] = []

        print(SQL_Command_Pos)

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
                    
        print(SQL_Command_Dict)

        SQL_Select = []
        SQL_Math = []
        SQL_Where = []

        for i in range(1,len(SQL_Command_Dict['SELECT'])):
            if "," in SQL_Command_Dict['SELECT'][i] or i == len(SQL_Command_Dict['SELECT'])-1:
                SQL_Select.append([SQL_Command_Dict['SELECT'][i-1].replace("_"," "), SQL_Command_Dict['SELECT'][i].strip(",")])
               
        for i in range(0,len(SQL_Maths_Operators)):
            for j in  range(0,len(SQL_Command_Dict[SQL_Maths_Operators[i]])):
                SQL_Math.append(SQL_Maths_Operators[i],SQL_Command_Dict[SQL_Maths_Operators[i][j].strip('()')])
                              
                
        SQL_Command_Dict['SELECT'] = SQL_Select
        SQL_Command_Dict['Maths'] = SQL_Math

        del SQL_Math
        del SQL_Select

        active_worksheet = self.excel_workbook['Transactions']

        #Gets the column indice for the excel spredsheet using the names from SQL Iterates through the top 13 rows of the spreadsheet and compares them to the list of Names provided
        #If it matches the coulumn indices is saved along with the abriviated name into a new list. Column_Indicies
        for columns,SQL_Columns_Count in itertools.product(range(1,13),range(0,len(SQL_Columns))):
            if active_worksheet.cell(row=1,column=columns).value == SQL_Columns[SQL_Columns_Count][0]:
                column_indices.append([columns,SQL_Columns[SQL_Columns_Count][1]])

        #Does a similiar thing to the above loop however is just looking for the column the action is going to be performed on.
        for columns in range(1,13):
            if active_worksheet.cell(row=1,column=columns).value == SQL_Sum[0]:
                SQL_Sum_index=columns
        
        print(column_indices)

        Output_String = self.SQL_Logic(active_worksheet,SQL_Where,column_indices,SQL_Sum_index,SQL_Columns)

        return Output_String

    def SQL_Logic(self,active_worksheet,SQL_Where,column_indices,SQL_Sum_Index,SQL_Columns):
        
        Output_Sum = 0
        Output_String = ""
        Output_List = []
        #Logic Bit of the code
        #First loop goes through all of the rows in the spreadsheet

        for i in range(12792,active_worksheet.max_row):
            count = 0
            #This goes through the commands given by the SQL
            for j in range(0, len(SQL_Where)):
                #This goes through the column indicies list. Contains the number and Shorthand name of the column
                for n in range (0, len(column_indices)):
                    #This checks to see if the command is going to be on the correct column
                    if SQL_Where[j][0] == column_indices[n][1]:
                        sheet_data = active_worksheet.cell(row=i,column=column_indices[n][0]).value
                        #Checks to see if the cell value is a datetime or not. Following operations can only be done on datetime objects
                        if type(sheet_data) is datetime:
                            #Checks to see if the current SQL command is a date and is in the correct date format, by checking if it contains a -
                            if self.split_date(SQL_Where[j][1]):
                                date_list = SQL_Where[j][1].split("-")
                                date1 = sheet_data.date()
                                #Converts the two dates provided by SQL_Where into datetime objects so they can be compared to the cell datetime
                                str_to_date1 =  self.to_date(date_list[0])
                                str_to_date2 = self.to_date(date_list[1])
                                if str_to_date1<date1 and date1<str_to_date2:
                                    #If matches the conditions adds 1 to the count so that the final check knows if all checks have passed for a row.
                                    count += 1
                        else:
                            #If cell value is not a datetime checks to see if the cell value matches the string provided by SQL_Where. If true adds 1 to the counter.
                            if sheet_data == SQL_Where[j][1]:
                                count +=1

            #If all the checks have passed, the counter will equal the total number of clauses given. Will then add data from the column specified by SQL_Sun_Index
            if count == len(SQL_Where):
               
                flat_list = []

                for j in range(0,len(column_indices)):
                    sheet_value = active_worksheet.cell(row=i,column=column_indices[j][0]).value
                    if type(sheet_value) is datetime:
                        flat_list.append(datetime.strftime(sheet_value,'%d-%m-%Y'))
                    else:
                        flat_list.append(sheet_value)

                Output_Sum += active_worksheet.cell(row=i,column=SQL_Sum_Index).value
                flat_list.append('£'+ str(active_worksheet.cell(row=i,column=SQL_Sum_Index).value))

                Output_List.append(flat_list)


        Output_String = ''.join(self.list_to_HTML_Table(Output_List)) + '\n <p> Total Cost : ' + str(Output_Sum) + '</p>'
        return Output_String

    def list_to_HTML_Table(self,list):
        yield '<table>'
        for sublist in list:
            yield '<tr><td>'
            yield '</td><td>'.join(sublist)
            yield '</td></tr>'
        yield '</table>'

    def to_date(self,str_to_date):
        return datetime.strptime(str_to_date, '%d/%m/%Y').date()

    def get_string(self,string,count,m,pm,SQL_Command_Pos):
        if count == 0:
            return (string[SQL_Command_Pos:m.start()])
        else:
            return (string[pm.end():m.start()])

    def split_date(self,test):
        try:
            key, value = test.split("-")
        except ValueError:
            return False
        else:
            return True



if __name__ == '__main__':
    Report_Generator().Load()
