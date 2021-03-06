import xlrd  
import pyodbc
import json
from datetime import datetime
import os
import sys
import shutil
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import shutil
import re
notapp="nan"
try:
    configfile=('C:\\Users\\lygio.joy\\python\\config\\pythonconfig.txt')
    json_file= open(configfile,"r",encoding="utf-8")
    data=json.load(json_file)
    json_file.close()
    server=data['server']
    db=data['DBName']
    #Database connection and table information
    connection = pyodbc.connect(r'Driver={SQL Server};Server=%s;Database=%s;Trusted_Connection=yes;' %(server,db))  
    cursor=connection.cursor()
    DBname=data['DBName']
    tableschema=data['tableschema']
    tablename=data['tablename']
    location = data['FileLocation']
    Archive=data['ArchiveFile']
    Error=data['Error']
    dropscript='use'+" "+ str(DBname) +'\n if exists (select 1 from information_schema.tables where table_schema=' +"'"+str(tableschema) +"'"+ 'and table_name=' +"'"+str(tablename) +"')\n"+'BEGIN \n DROP TABLE '+ str(DBname) +"."+str(tableschema)+"."+str(tablename) +'\n END \n'
    cursor.execute(dropscript)
    connection.commit()
    #Location of excel 
    FileList=os.listdir(location)
    for xlfile in FileList:
        File=os.path.join(location,xlfile)
        if File.endswith('.xlsx') or File.endswith('.xls'):
            df = pd.read_excel(File)
            column_name=df.columns
            num_columns=df.shape[1]
            num_rows=df.shape[0]
            data_type=data['datatype']
            columns=data['columnname']
            cLen=len(columns)
            tLen=len(data_type)
            # sql syntax for Drop and create table script
            CreateScript='use'+" "+ str(DBname) +'\n if not exists (select 1 from information_schema.tables where table_schema=' +"'"+str(tableschema) +"'"+ 'and table_name=' +"'"+str(tablename) +"')\n"+'BEGIN \n CREATE TABLE ' + str(DBname) +"."+str(tableschema)+"."+str(tablename)+'\n (id1 int identity(1,1),filename nvarchar(max) NULL,\n'
    #sql syntax on insert script
            InsertScript='INSERT INTO '+str(DBname) +"."+str(tableschema)+"."+str(tablename) + ')'
    #creating sql create table script with excel column header as column names and datatype on nvarchar(255)
            header=0
            for header in range(num_columns):
                columnname=column_name[header]
                columnname=columnname.replace(" ","")
                columnname=columnname.replace(" ","")
                columnname=columnname.replace("-","")
                columnname=columnname.replace(",","")
                columnname=columnname.replace("(","")
                columnname=columnname.replace(")","")
            #print (columnname)
                c=0
                d=0
                while c<cLen and d<tLen :
                    if (columnname.lower()==columns[c].lower()):
                        CreateScript +=columns[c] + " " + data_type[d] +" NULL,\n"
                    c=c+1
                    d=d+1
            CreateScript=CreateScript.rstrip(',\n') + ') END'
            #need to add commit command
            cursor.execute(CreateScript)
            connection.commit()
            column=0
    #creating sql insert script with column names
            #for column in range(num_columns):
            #    columnname=column_name[column]
            #    columnname=columnname.replace(" ","")
            #    columnname=columnname.replace("-","")
            #    columnname=columnname.replace(",","")
            #    columnname=columnname.replace("(","")
            #    columnname=columnname.replace(")","")
            #    c=0
            #    d=0
                #while c<cLen and d<tLen :
                #   if (columnname.lower()==columns[c].lower()):
                #        InsertScript +="["+columns[c] +"], \n" 
                #    c=c+1
                #   d=d+1
    #add word 'values' to insert script 
            #InsertScript=InsertScript.rstrip(', \n') +"\n)"
    #reset column variable to 0
            #column=0
    #get each value from excel column and add to the insert script
            for rows in range(num_rows):
                InsertScript += "SELECT '"+File+"',"
                for column in range(num_columns):
                    columnname=column_name[column]
                    columnname=columnname.replace(" ","")
                    columnname=columnname.replace("-","")
                    columnname=columnname.replace(",","")
                    columnname=columnname.replace("(","")
                    columnname=columnname.replace(")","")
                    value=df[column_name [column]][rows]
                    d=0
                    c=0
                    while c<cLen and d<tLen :
                        
                        if columnname.lower()==columns[c].lower() and data_type[d]=='datetime':
                            try:
                                value=value.split()
                                value=value[0] + " "+value[1]+" " +value[2]
                                value=pd.to_datetime(value)
                                value=datetime.date(value)
                                value=str(value)
                                value=value.replace("0001-01-01","NULL")
                                if(value=='NULL'):
                                    InsertScript +="NULL,"
                                else:
                                    InsertScript +="'"+str(value)+"',"
                            except Exception as e:
                                print("Error:",e)
                                src=location+"\\"+xlfile
                                Failed=Error+"\\"+xlfile
                                shutil.move(src,Failed)
                                sys.exit()
                        elif columnname.lower()==columns[c].lower() and data_type[d]=='int':
                            try:
                                if (value==notapp):
                                    InsertScript+="'"+'NULL'+"',"
                                else:
                                    value=int(value)
                                    if(type(value)==int):
                                        InsertScript += str(value)+","
                                    else:
                                        print("record in this column " + columns[c] + " is not a "+ data_type[d] +" bad value is "+ str(value) +" in row "+ str(rows))
                                        sys.exit()
                            except Exception as e:
                                print("Error:",e)
                                src=location+"\\"+xlfile
                                Failed=Error+"\\"+xlfile
                                shutil.move(src,Failed)
                                sys.exit()
                        elif columnname.lower()==columns[c].lower() and data_type[d]=='float':
                            try:
                                value=float(value)
                                if(type(value)==float):
                                    InsertScript +="'"+str(value)+"',"
                                else:
                                    print("record in this column " + columns[c] + " is not a "+ data_type[d] +" bad value is "+ value +" in row "+ rows)
                                    sys.exit()
                            except Exception as e:
                                print("Error:",e)
                                src=location+"\\"+xlfile
                                Failed=Error+"\\"+xlfile
                                shutil.move(src,Failed)
                                sys.exit()
                        elif columnname.lower()==columns[c].lower():
                            try:
                                value=str(value)
                                value=value.replace("'","''")
                                if (value==notapp):
                                    InsertScript += "''" +","
                                else:
                                    InsertScript +="'"+str(value)+"',"
                            except Exception as e:
                                print("Error:",e)
                                src=location+"\\"+xlfile
                                Failed=Error+"\\"+xlfile
                                shutil.move(src,Failed)
                                sys.exit()
                         
                        c=c+1
                        d=d+1

                InsertScript=InsertScript.rstrip(',')+" \n"
            InsertScript=InsertScript.replace("SELECT","UNION ALL SELECT")
            InsertScript=InsertScript.replace(")UNION ALL"," ")
            try:
                cursor.execute(InsertScript)
                connection.commit()
                src=location+"\\"+xlfile
                backup=Archive+"\\"+xlfile
                shutil.move(src,backup)
            except Exception as e:
                print("insert failed for file "+ File)
                print("Error:",e)
                src=location+"\\"+xlfile
                Failed=Error+"\\"+xlfile
                shutil.move(src,Failed) 
except Exception as e:
    print("Error:",e)