# -*- coding: utf-8 -*-
import openpyxl
import xml
import itertools
from datetime import date, datetime
from dateutil.parser import parse
import re
import logging
import logging.handlers
from operator import itemgetter
import codecs
import tkinter as tk
from tkinter import filedialog as fd
import webbrowser
import collections

class Report_Generator():

    def Main(self):

        html_class_arr = []

        report = """<meta charset="UTF-8"> <html>
        <head>
        <!-- Plotly.js -->
        <script src="http://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
        <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
        
         """

        xml_file_path,excel_file_path = tk_windows().tk_window()
        xml_iter,excel_workbook = self.Load(xml_file_path,excel_file_path)

        report += "<style type='text/css'>"

        for node in xml_iter:

            if node.tag =="Style":
                for subnode in  node:
                    flat_list = ['','']
                    temp_string = ""

                    if subnode.tag == "Class":
                            flat_list[0],flat_list[1], temp_string = self.Style_Class(subnode)

                            html_class_arr.append(flat_list)
                            report += temp_string

                report += "\n</style> \n </head> \n <body>"

            elif node.tag=="txt":

                for sublist in html_class_arr:
                    if 'Text' == sublist[1]:
                        report+= '<p class="' + sublist[0] + '">'

                else:
                    report += '<p>'

                report += node.text + "</p>"
        
            elif node.tag=="SQL":

                SQL_Operators = ['SELECT','FROM','WHERE','AND','SUM','AVG','COUNT','MAX','MIN','GROUP BY',]
                SQL_Maths_Operators = ['SUM','AVG','COUNT','MAX','MIN']
            
                SQL_Command_Dict = {}
                Output_List = []
                Group_By = []

                self.parse_SQL(node.text,SQL_Command_Dict,SQL_Operators,SQL_Maths_Operators)
                active_worksheet = excel_workbook[SQL_Command_Dict['FROM'][0]]
                self.find_columns(active_worksheet,SQL_Command_Dict)
                self.SQL_Logic(active_worksheet,SQL_Command_Dict,Group_By,Output_List)
                output = self.output_string_manipulation(SQL_Command_Dict,Output_List,Group_By,html_class_arr)

                report +="\n" + '<div> ' + output  + "</div> </body>"

        report += "</html>"
        html_output = codecs.open("report.html","w","utf-8")
        html_output.write(report)
        html_output.close()

        webbrowser.open_new_tab('report.html')

    def Style_Class(self,node):

        try:

            str_Class_Name = node.find("Class_Name").text 
            if str_Class_Name is None: str_Class_Name = '';
            str_Applys_to = node.find("Apply_To").text
            if str_Applys_to is None: str_Applys_to = '';
            str_Background_Colour = node.find("Background_Colour").text
            if str_Background_Colour is None: str_Background_Colour = '';
            str_Text_Colour = node.find("Text_Colour").text
            if str_Text_Colour is None: str_Text_Colour = '';
            str_Text_Size = node.find("Text_Size").text
            if str_Text_Size is None: str_Text_Size = '';
            str_Text_Alignment = node.find("Text_Alignment").text
            if str_Text_Alignment is None: str_Text_Alignment = '';
            str_Text_Font = node.find("Text_Font").text
            if str_Text_Font is None: str_Text_Font = '';

            #str_Table_Width = node.find("Table_Width").text
            str_Table_Border = node.find("Table_Border").text
            if str_Table_Border is None: str_Table_Border = '';
            str_Table_Border_Collapse = node.find("Table_Border_Collapse").text
            if str_Table_Border_Collapse is None: str_Table_Border_Collapse = '';
            str_Table_Border_Spacing = node.find("Table_Border_Spacing").text
            if str_Table_Border_Spacing is None: str_Table_Border_Spacing = '';
            str_Table_Padding = node.find("Table_Padding").text
            if str_Table_Padding is None: str_Table_Padding = '';

        except Exception as e:
            pass

        str_html_Class = ""
        str_html_Class += "." + str_Class_Name + " {"
        str_html_Class += "\n background: "+str_Background_Colour + ";"
        str_html_Class += "\n color: " + str_Text_Colour + ";"
        str_html_Class += "\n font-size: " + str_Text_Size + ";"
        str_html_Class += "\n text-align: " + str_Text_Alignment + ";"
        str_html_Class += "\n font-family: " + str_Text_Font + ";"

        str_html_Class += "\n border: " + str_Table_Border + ";"
        str_html_Class += "\n border-collapse: " + str_Table_Border_Collapse + ";"
        str_html_Class += "\n border-spacing: " + str_Table_Border_Spacing + ";"
        str_html_Class += "\n padding:" + str_Table_Padding + ";"
        str_html_Class += "\n }"


        return str_Class_Name, str_Applys_to, str_html_Class

    def Load(self,xml_file_path,excel_file_path):
        #Loads the xml file. Controls the layout of the html

        try:
            xml_tree = xml.etree.ElementTree.parse(xml_file_path)
            xml_root = xml_tree.getroot()
            xml_iter = xml_tree.iter()

            #Loads the excel file into memory
            print("here")
            excel_workbook = openpyxl.load_workbook(excel_file_path,data_only=True)
            print("now here")
        except Exception as e:
            return;

        return xml_iter,excel_workbook

    def parse_SQL(self,SQL,SQL_Command_Dict,SQL_Operators,SQL_Maths_Operators):

        self.get_parsed_SQL(SQL,SQL_Command_Dict,SQL_Operators)

        SQL_Group_By =[]
        SQL_Select = []
        SQL_Math = {}
        SQL_Where = {}
        SQL_New_Math = []

        current_key = 'GROUP BY'	
        try:
            self.get_comma_list(SQL_Group_By,SQL_Command_Dict[current_key])
        except KeyError as e:
            pass

        current_key = 'SELECT'
        self.get_comma_list(SQL_Select,SQL_Command_Dict[current_key])

        current_key = 'AND'
        SQL_Where =  self.parse_And(SQL_Command_Dict[current_key])		
        #Maths Selectors      

        try:
            for key in SQL_Maths_Operators:
                SQL_Math.update(self.parse_Maths(SQL_Command_Dict[key],key))
                del SQL_Command_Dict[key]
        except KeyError as e:
            pass

        SQL_Command_Dict['GROUP BY'] = SQL_Group_By
        SQL_Command_Dict['SELECT'] = SQL_Select
        SQL_Command_Dict['WHERE'] = SQL_Where
        SQL_Command_Dict['MATHS'] = SQL_Math

        del SQL_Group_By
        del SQL_Select
        del SQL_Where
        del SQL_Math
        del SQL_Command_Dict['AND']

        if SQL_Command_Dict['MATHS']  and  not SQL_Command_Dict['GROUP BY']:
            raise NameError('Group by Required When Maths operators present') 

    def find_columns(self,active_worksheet,SQL_Command_Dict):

        current_key = 'SELECT'
        self.get_column(active_worksheet,SQL_Command_Dict[current_key],current_key,SQL_Command_Dict)
        #self.get_column(active_worksheet,SQL_Command_Dict[current_key],'WHERE',SQL_Command_Dict)
        self.get_Where_Columns(SQL_Command_Dict['WHERE'],active_worksheet)
        try:		
            self.get_column(active_worksheet,SQL_Command_Dict[current_key],'GROUP BY',SQL_Command_Dict)
        except KeyError:
            pass
        current_key = 'MATHS'
        self.get_Where_Columns(SQL_Command_Dict[current_key],active_worksheet)

    def get_Where_Columns(self,Main_Dict,active_worksheet):

        key_list = Main_Dict.keys()
        for i,key_string in itertools.product( range(1,active_worksheet.max_column),key_list):
            if active_worksheet.cell(row=1,column=i).value == key_string:
                for dict_list in Main_Dict[key_string]: 
                    dict_list.append(i)

    def SQL_Logic(self,active_worksheet,SQL_Command_Dict,Group_By,Output_List):

        #Logic Bit of the code
        #First loop goes through all of the rows in the spreadsheet

        for i in range(1,active_worksheet.max_row):
            count = 0
            #This goes through the commands given by the SQL
            for key,dict_list in SQL_Command_Dict['WHERE'].items():		              
                for sublist in dict_list:
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
                        if sheet_data == sublist[0]:
                            count +=1

            #If all the checks have passed, the counter will equal the total number of clauses given. Will then add data from the column specified by SQL_Sun_Index
            if count == len(SQL_Command_Dict['WHERE']):
                self.return_table_data(SQL_Command_Dict,active_worksheet,i,Group_By,Output_List)

        del SQL_Command_Dict['WHERE'],active_worksheet

    def output_string_manipulation(self,SQL_Command_Dict,Output_List,Group_By, html_group_arr):

        #Creates the variables needed for the function. 
        Output_String = ''
        self.graph_output = ""
        str_html_class_name = ""

        header_list_group = []
        header_list_select = []

        Group_By_Output_List = []
        Group_By_Dict = collections.OrderedDict()


        #This creates the group by dictionary from the group by array. Tests to see if a group by is needed.
        if not Group_By[0]:
            pass
        else:			
            print('Not Sorted Group By : ', Group_By)
            Group_By.sort(key=itemgetter(0))
            print('Sorted Group By : ', Group_By)
            self.return_group_by(Group_By,Group_By_Dict,header_list_group,SQL_Command_Dict)
            print(Group_By_Dict)
            Group_By_Dict
         

        #Output string. Writen in html. Adds a button to toggle the table below it.               
        Output_String +=  "<div> \n <details> \n <summary> Click to reveal non grouped data </summary>"

        #Tests for any css classes that are inculded. If it finds one that matches the correct type it will apply the class to the table otherwise output just the table.
        for sublist in html_group_arr:
            if 'Table' == sublist[1]:
                Output_String += '\n <table class="' + sublist[0] + '" > \n <tr> \n '
                str_html_class_name = sublist[0]
                break

        else:
            Output_String += '\n <table> \n <tr> \n'

        #Creates the table structure for the select table
        Output_String += '\n'.join(self.generate_html_table(Output_List,SQL_Command_Dict['SELECT'],header_list_select,1,str_html_class_name))
        str_html_class_name = ""


        Output_String +=  "\n </details> \n </div>"
        
        del Output_List, header_list_select,SQL_Command_Dict['SELECT']

        Output_String += '<div> \n <details> \n <summary> Click to reveal grouped data </summary>'
            
        #Embeded in a try incase group by doesnt exist. Creates a header list for the table from the SQL command dictionary. Takes from both the group by and maths keys.
        try:
            temp_string = ''
            header_list = []
            graph_list = []
            

            for sublist in SQL_Command_Dict['GROUP BY']:
                header_list.append(sublist[1])

            for k in SQL_Command_Dict['MATHS']:
                header_list.append(k)

            #Tests for any css classes that are inculded. If it finds one that matches the correct type it will apply the class to the table otherwise output just the table.
            for sublist in html_group_arr:
                if 'Table' == sublist[1]:
                    Output_String += '\n <table class="' + sublist[0] + '" >' 
                    str_html_class_name = sublist[0]
                    break
            else:
                Output_String += '\n <table>'


            #Creates the html table header.
            Output_String += '\n'.join(self.generate_group_by_table(header_list,str_html_class_name))
            #Recursive function that generates the table from the grouped dictionary
            temp_string, graph_list = self.get_recursive_dict(Group_By_Dict,temp_string,graph_list,0,str_html_class_name)
            Output_String += temp_string
            Output_String += '\n </table>'
        except KeyError: pass		
        Output_String += '\n </details> \n </div>'

        del Group_By_Output_List,header_list_group

        Output_String += '<div> \n <details> \n <summary> Click to reveal the graphs </summary>'
        Output_String += self.graph_output
        Output_String += '\n </details> \n </div>'

        return  Output_String

    def generate_group_by_table(self,header_list,str_html_class_name):
        yield '\n <tr class="' + str_html_class_name +'" > \n'
        for header in header_list:
            yield '     <th class="'+ str_html_class_name +'">' + header + '</th>'
        yield ' </tr> \n'
            
    def get_recursive_dict(self,recdict,out_string,graph_out,count=0,str_html_class_name=""):

        for subkey in recdict:
            if count == 0:
                out_string += '<tr class="'+ str_html_class_name +'"> \n'
            if isinstance(recdict[subkey],dict):
                if count >0:
                    out_string+= '</tr> \n <tr class="'+ str_html_class_name +'"> \n <td class="'+ str_html_class_name +'"></td>'
                out_string += '<td>' + subkey +'</td>'
                #graph_list.append(subkey)
                out_string, graph_out = self.get_recursive_dict(recdict[subkey],out_string,graph_out,count+1,str_html_class_name)

            elif isinstance(recdict[subkey],list):
                for s in recdict[subkey]:
                    out_string += '<td class="'+ str_html_class_name +'">' + str(s) + '</td>'
                    graph_out.append(s)
                out_string += '</tr>'

            if count == 0:
                out_string += '</tr> \n'
        return out_string, graph_out

    def return_table_data(self,SQL_Command_Dict,active_worksheet,i,Group_By,Output_List):

        flat_list = []
        Group_By_Flat_List = []

        for select_sublist in SQL_Command_Dict['SELECT']:
            sheet_value = active_worksheet.cell(row=i,column=select_sublist[-1]).value
            if type(sheet_value) is datetime:
                pos = str([s for s in select_sublist[:-1] if '%' in s])
                flat_list.append(datetime.strftime(sheet_value,pos))
            else:
                if sheet_value is None: sheet_value = ''			
                flat_list.append(sheet_value)
        try:

            for group_by_sublist in SQL_Command_Dict['GROUP BY']:
                group_by_value = active_worksheet.cell(row=i,column=group_by_sublist[-1]).value
                if type(active_worksheet.cell(row=i,column=group_by_sublist[-1]).value) is datetime:
                    pos = str([s for s in select_sublist[:-1] if '%' in s])
                    Group_By_Flat_List.append(datetime.strftime(active_worksheet.cell(row=i,column=group_by_sublist[-1]).value,pos))
                else:
                    if group_by_value is None: group_by_value = ''				
                    Group_By_Flat_List.append(group_by_value)

        except KeyError:
            pass

        for key, dict_list in SQL_Command_Dict['MATHS'].items():
            for maths_sublist in dict_list:
                flat_list.append('£'+ str(active_worksheet.cell(row=i,column=maths_sublist[-1]).value))
                Group_By_Flat_List.append(active_worksheet.cell(row=i,column=maths_sublist[-1]).value)


        Group_By.append(Group_By_Flat_List)
        Output_List.append(flat_list)

    def get_parsed_SQL(self,SQL,SQL_Command_Dict,SQL_Operators):

        SQL_Command_Pos = [] 
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
                    if m.start()> SQL_Command_Pos[i][1]:                     
                        chunk_string = " ".join( self.get_string(SQL_No_Line,count,m,pm,SQL_Command_Pos[i][1]).split())
                        if chunk_string != SQL_Command_Pos[i][2]:
                            SQL_Command_Dict[SQL_Command_Pos[i][2]].append(self.get_string(SQL_No_Line,count,m,pm,SQL_Command_Pos[i][1]))
                    count = 0
                pm=m
                count +=1

    def is_in(self,i_str,i_list):
        if i_str not in i_list:
            i_list.append(i_str)        

    def return_maths(self,SQL_Command_Dict,Min_Max_Count,m_list,header_list):
        for sublist in SQL_Command_Dict:
            if 'MAX' in sublist:
                m_list.append( str(max(Min_Max_Count)))
                self.is_in('MAX',header_list)
            if 'MIN' in sublist: 
                m_list.append(str(min(Min_Max_Count)))
                self.is_in('MIN',header_list)
            if 'AVG' in sublist: 
                m_list.append((sum(Min_Max_Count)/len(Min_Max_Count)))
                self.is_in('AVG',header_list)
            if 'SUM' in sublist: 
                m_list.append(str(sum(Min_Max_Count)))
                self.is_in('SUM',header_list)
            if 'COUNT' in sublist: 
                m_list.append(str(len(Min_Max_Count)))
                self.is_in('COUNT',header_list)

    def get_comma_list(self,list_dict,dict_arr):

        for i in range(0,len(dict_arr)):
            flat_list =[]
            if dict_arr[i] == 'AS':
                
                temp_string = ''
                flat_list.append(dict_arr[i-1])   
                for substring in dict_arr[i+1:]:
                    temp_string += substring + ' '
                    if ',' in substring: break
                flat_list.append(temp_string)

            elif i <=1 and len(dict_arr) <=2: flat_list.append(dict_arr[i])
            for x,sublist in enumerate(flat_list):
                if 'TO_CHAR' in sublist:
                    tempstring = ''
                    pos = 1
                    for m in re.finditer('TO_CHAR', sublist): pos = m.end()
                    for substring in sublist[pos:]:
                        if '%' in substring: break
                        if substring not in {'(',')',','}:
                            tempstring += substring
                    flat_list[x] = tempstring
                    if '%' in sublist:
                        pos = sublist.index('%')
                        flat_list.append(sublist[pos:-2])

            if flat_list:
                for x,substring in enumerate( flat_list):  
                    flat_list[x] = substring.strip().strip(',').replace('_',' ').replace('"','').replace("'",'')
                list_dict.append(flat_list)

    def parse_Maths(self,Main_dict,key):

        flat_dict = {}
        for i,sublist in enumerate(Main_dict):
            flat_list = []
            if 'AS' in sublist:
                flat_list.append(Main_dict[i+1])
                flat_list.append(key)                
                new_key= Main_dict[i-1].strip('(').strip(')')

                if new_key in flat_dict:
                    flat_dict[new_key].append(flat_list)
                else:
                    flat_dict[new_key] = []
                    flat_dict[new_key].append(flat_list)

        return flat_dict

    def parse_And(self,dict_arr):

        temp_list = []
        pre_export_list = []
        flat_dict = {}
        for sublist in dict_arr:
            temp_list.append(sublist.replace('(','').replace(')',''))

        dict_arr = temp_list
        for i,sublist in enumerate(dict_arr):
            flat_list = []

            if sublist in {'=','<','>'}:                
                temp_string = ''
                if dict_arr[i+1].startswith("'"):   
                    temp_string += dict_arr[i+1]
                    for a_sublist in dict_arr[i+1:]:
                        if a_sublist.endswith("'"):
                            if a_sublist not in temp_string: temp_string +=' '+ a_sublist
                            break
                        elif not a_sublist.startswith("'"): temp_string +=' '+ a_sublist
                    flat_list.append(temp_string.replace("'",''))
                    flat_list.append(sublist)

                if dict_arr[i-1].replace('_',' ') in flat_dict:
                    flat_dict[dict_arr[i-1].replace('_',' ')].append(flat_list)
                else:
                    flat_dict[dict_arr[i-1].replace('_',' ')] = []
                    flat_dict[dict_arr[i-1].replace('_',' ')].append(flat_list)
        return flat_dict

    def to_date(self,str_to_date): return datetime.strptime(str_to_date, '%d/%m/%Y').date()

    def get_string(self,string,count,m,pm,SQL_Command_Pos):
        if count == 0:
            return (string[SQL_Command_Pos:m.start()])
        else:
            return (string[pm.end():m.start()])

    def validate_date(self,test):
        try:
           datetime.strptime(str(test),'%d/%m/%Y')
        except ValueError:
            return False
        else:
            return True

    def get_column(self,active_worksheet,current_key,other_key,SQL_Dict):
        
        for i in range(1,active_worksheet.max_column):
            for dict_list in current_key:
                if active_worksheet.cell(row=1,column=i).value == dict_list[0]:
                    for sublist in SQL_Dict[other_key]:
                        if dict_list[0] in sublist or dict_list[1] in sublist:
                            if other_key == 'GROUP BY':
                                sublist.insert(0,dict_list[0])
                            sublist.append(i)                            
            for sublist in SQL_Dict[other_key]:					
                if type(sublist[-1]) is not int or self.validate_date(sublist[-1]):                  			
                    if sublist[0] == active_worksheet.cell(row=1,column=i).value:
                        sublist.append(i)		
                        			
    def generate_html_table(self,Output_List,dict_arr,header_list,current_table,str_html_class_name):

        Output_String =""
        for substring in dict_arr:
            yield '   <th class="'+ str_html_class_name +'">' + str(substring[1]) + '</th>'

        for h_string in header_list:
            yield  '   <th class="'+ str_html_class_name +'">' + str(h_string) + '</th>'
        yield '  </tr>'

        for sublist in Output_List: 
            yield '\n  <tr class="'+ str_html_class_name +'"> \n'
            for l_string in sublist:
                yield'    <td class="'+ str_html_class_name +'" >'+ str(l_string) + '</td>'
            yield '\n  </tr> \n'

        yield '</table>'

    def remove_same(self,list):
        checked_list = []
        for sublist in list:
            for i in range(0,len(sublist)):
                if sublist[i] not in checked_list or "£" in sublist[i]:
                    checked_list.append(sublist[i])
                else:
                    sublist[i] = ""
        del checked_list

    def return_group_by(self,Group_By,Group_Dict,header_list_group,SQL_Command_Dict):
        Grouped_Items = []
        Temp_Grouped_Items = []

        group_length =  len(SQL_Command_Dict['GROUP BY'])

        for sublist in Group_By:
            self.recursive_dict(Group_Dict,0,group_length,sublist)

        self.return_dict(Group_Dict,SQL_Command_Dict,header_list_group)                     

    def return_dict(self,recdict,SQL_Command_Dict,header_list_group,count=0,title_list=''):

        for k,v in recdict.items():
            if isinstance(v, dict):
                title_list += k + ': '

                if k in title_list[title_list.find(':'):]:
                    title_list = title_list.replace(title_list[: title_list.find(':')+2],'')
                    
                self.return_dict(v,SQL_Command_Dict,header_list_group,count+1,title_list)
            else:
                int_list = v
                title_list += ': ' + k

                if k in title_list[title_list.rfind(':'):]:
                    new_string = title_list[:title_list.find(':')] + title_list[title_list.rfind(':'):]

                title_list = new_string
                maths_dict = {}
                maths_list = []
                for key,sub_dict in SQL_Command_Dict['MATHS'].items():
                    self.return_maths(sub_dict,int_list,maths_list,header_list_group)   
                    self.return_graph(int_list,'box',title_list)
                    maths_dict[key] = maths_list
                recdict[k] = maths_dict

    def recursive_dict(self,recdict,count,groupint,l):
        if count == (groupint-1):
            if l[count] not in recdict.keys():
                recdict[l[count]]=[]
            recdict[l[count]].append(l[count+1])
        else:
            if l[count] not in recdict.keys():
                recdict[l[count]] = collections.OrderedDict()
            self.recursive_dict(recdict[l[count]],count+1,groupint,l)

    def create_nested_dict(self,Group_Dict,key):
        flat_dict = {}
        for dict_list in Group_Dict[key]:
            if dict_list in flat_dict:
                flat_dict[dict_list].update(dict_list)
            else:
                flat_dict[dict_list] = []
                flat_dict[dict_list].append(dict_list)
        Group_Dict[key] = flat_dict

    def return_graph(self,int_list,type,title):
        self.graph_output += '<div> <div id="graph' + title + '"></div>'
        self.graph_output += '<script> var trace1 = { x:' + str(int_list) + ', \n type:"' + type + '''",\n name: "Set 1"};
        var data = [trace1]; \n var layout = {title: " ''' +str(title) +'''"};
        Plotly.newPlot("graph''' + title + '", data, layout) </script> \n </div>'

class tk_windows ():

    def tk_window(self):

        self.xml_file_path = ""
        self.excel_file_path = ""
        self.graph_selection_out = ""

        root = tk.Tk()

        self.graph_selection = tk.StringVar()

        frame_button = tk.Frame(root)
        frame_radio_button = tk.Frame(root)

        frame_button_xml = tk.Frame(frame_button)
        frame_button_excel = tk.Frame(frame_button)

        label_xml = tk.Label(frame_button_xml)
        label_excel = tk.Label(frame_button_excel)

        button_xml = tk.Button(frame_button_xml,text='Pick Xml File',command=lambda:  self.xml_button(label_xml))
        button_excel = tk.Button(frame_button_excel,text='Pick Excel File',command=lambda: self.excel_button(label_excel))
        button_close = tk.Button(root,text="Continue",command = root.destroy)

        
        radio_graph_bar = tk.Radiobutton(frame_radio_button,text="Bar Chart",variable=self.graph_selection,value="bar",command=self.rb_select)
        radio_graph_line = tk.Radiobutton(frame_radio_button,text="Line Chart",variable=self.graph_selection,value="line",command=self.rb_select)
        radio_graph_box = tk.Radiobutton(frame_radio_button,text="Box Plot",variable=self.graph_selection,value="box",command=self.rb_select)

        button_xml.pack()
        button_excel.pack()
        label_xml.pack()
        label_excel.pack()

        frame_button_xml.pack()
        frame_button_excel.pack()

        radio_graph_bar.pack()
        radio_graph_line.pack()
        radio_graph_box.pack()

        frame_button.grid(row=0, column=0)
        frame_radio_button.grid(row=0, column=1)

        #frame_button.pack()
        #frame_radio_button.pack()
        button_close.grid(columnspan=2)
        root.mainloop()

        return self.xml_file_path,self.excel_file_path,self.graph_selection_out

    def xml_button(self,label_xml):

        self.xml_file_path = fd.askopenfilename(filetypes=(("XML Files","*.xml"),("All Files","*.")))
        label_xml.config(text=self.xml_file_path)
    
    def excel_button(self,label_excel):

        self.excel_file_path = fd.askopenfilename(filetypes=(("Excel Files","*.xlsm"),("All Files","*.")))
        label_excel.config(text=self.excel_file_path)

    def rb_select(self):
        self.graph_selection_out = self.graph_selection.get()

if __name__ == '__main__':
    Report_Generator().Main()

