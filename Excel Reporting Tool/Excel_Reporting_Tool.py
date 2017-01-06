import openpyxl
import xml
import itertools
from datetime import date, datetime
from dateutil.parser import parse

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
            print(node.tag)
            if node.tag=="txt":
                report += "\n" + "<body><p>" + node.text + "</p></body>"
        
            elif node.tag=="SQL":
                report +="\n" + "<body><p>" + str(self.parse_SQL(node.text)) + "</p></body>"


        report += "</html>"
        html_output.write(report)
        html_output.close()

    def parse_SQL(self,SQL):

        SQL_Command_List = []
        SQL_Columns = []
        SQL_Where = []
        column_indices = []

        SQL_Sheet = ""
        SQL_Sum = ""
        SQL_Sum_index =1
        Output_Sum = 0

        #Splits the text that comes in into understandable instructions
        #First splits the text from the xml file into lines.
        #Then splits the text up determined by "," Creates a new list which contains, eg [['Select][Account=AT,Transaction=TR]]
        #Then finds a command word. eg SELECT, Will the go one down in the list and then perform actions as required on the data.
        #For SELECT it will split the string at the =. Then add the two halves to a list for further use.

        SQL_List = str.splitlines(SQL)

        for z in SQL_List:
            space_list= z.split(None,1)
            for i in space_list:
                comma_list = i.split(",")
                SQL_Command_List.append(comma_list)

        for i in range (0,len(SQL_Command_List)):
            if SQL_Command_List[i][0] == 'SELECT':
                for string in SQL_Command_List[i+1]:
                    split_string = string.split("=",1)
                    SQL_Columns.append([split_string[0],split_string[1]])
            elif SQL_Command_List[i][0] == "FROM":
                SQL_Sheet = SQL_Command_List[i+1]
            elif SQL_Command_List[i][0] == "WHERE":
                for string in SQL_Command_List[i+1]:
                    split_string = string.split("=",1)
                    SQL_Where.append([split_string[0],split_string[1]])
            elif SQL_Command_List[i][0] == "SUM":
                SQL_Sum = SQL_Command_List[i+1]

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
                Output_Sum += active_worksheet.cell(row=i,column=SQL_Sum_index).value
                                    
        return Output_Sum

    def to_date(self,str_to_date):
        return datetime.strptime(str_to_date, '%d/%m/%Y').date()

    def split_date(self,test):
        try:
            key, value = test.split("-")
        except ValueError:
            return False
        else:
            return True



if __name__ == '__main__':
    Report_Generator().Load()
