from tkinter import *
from tkinter import messagebox
import xml.etree.ElementTree as et
from tkinter import filedialog
import pandas as pd
import re

class Sortng_App:
    def __init__(self,root) -> None:
        self.root=root
        self.root.title("Oerlikon Filtering App")
        self.root.geometry("700x500+0+0")
        self.root.configure(background='white')
        #self.logo_incon = PhotoImage(file=r"C:\Users\ZakariaH\Documents\File Sorting App\images\logo.png") 
        title=Label(self.root,text="Filtering App",font=("Arial",14,"bold"),bg="yellow",fg="red",anchor=W,compound=LEFT) #image=self.logo_incon,
        title.place(x=0,y=0,relwidth=0.5,relheight=0.1)

        self.config_name=StringVar(self.root,value="Config file location")
        self.filter_name=StringVar(self.root,value="Filter file location")
        self.save_name=StringVar(self.root,value="Save file location")
        #=======Section 1: ====================
        lbl_config_folder = Label(self.root,text="Select config file",font=("Arial",12,"normal"),bg='white',fg="black",anchor=W).place(x=50,y=100)
        txt_config_name = Entry(self.root,font=("Arial",10,"normal"),fg="black",state='readonly',bg='lightyellow',textvariable=self.config_name)
        txt_config_name.place(x=200,y=100,height=20,width=300)
        btn_browse_config_folder = Button(self.root,text="Browse",command=self.browse_config_folder,font=("Arial",12,"bold"),fg="black",bg="lightyellow",padx=20,activebackground='white',activeforeground='red',cursor='hand2')
        btn_browse_config_folder.place(x=550,y=100,height=20,width=100)
        
        lbl_filter_folder = Label(self.root,text="Select filter file",font=("Arial",12,"normal"),bg='white',fg="black",anchor=W).place(x=50,y=140)
        txt_filter_name = Entry(self.root,font=("Arial",10,"normal"),fg="black",state='readonly',bg='lightyellow',textvariable=self.filter_name)
        txt_filter_name.place(x=200,y=140,height=20,width=300)
        btn_browse_filter_folder = Button(self.root,text="Browse",command=self.browse_filter_folder,font=("Arial",12,"bold"),fg="black",bg="lightyellow",padx=20,activebackground='white',activeforeground='red',cursor='hand2')
        btn_browse_filter_folder.place(x=550,y=140,height=20,width=100)
        
        lbl_save_folder = Label(self.root,text="Select save location",font=("Arial",12,"normal"),bg='white',fg="black",anchor=W).place(x=50,y=180)
        txt_save_name = Entry(self.root,font=("Arial",10,"normal"),fg="black",state='readonly',bg='lightyellow',textvariable=self.save_name)
        txt_save_name.place(x=200,y=180,height=20,width=300)
        btn_browse_save_folder=Button(self.root,text="Browse",command=self.browse_save_folder,font=("Arial",12,"bold"),fg="black",bg="lightyellow",padx=20,activebackground='white',activeforeground='red',cursor='hand2')
        btn_browse_save_folder.place(x=550,y=180,height=20,width=100)
         #=======Section 2: ====================
        hr = Label(self.root, text="", bg='lightgrey').place(x=50, y=250, height=2, width=600)
        #lbl__support_ext=Label(self.root,text="Various Supported Extenstions",font=("Arial",12,"bold"),bg='white',fg="black",anchor=W).place(x=50,y=190)
        #=======Section 3: ====================
        Frame1 = Frame(self.root,bg='white',bd=4,relief=RIDGE)
        Frame1.place(x=50,y=300,width=350,height=150)
        self.lbl_ERP_type = Label(Frame1,text="ERP Type",font=("Arial",12,"bold"),bg='white',fg="black",anchor=W).place(x=10,y=10)
        self.var_ERP_type=StringVar(self.root,1)
        self.radio_ERP_type = Radiobutton(Frame1,text="SAP",value=1,variable=self.var_ERP_type,font=("Arial",12,"bold"),bg='white',fg="black",anchor=W).place(x=10,y=40)
        self.radio_ERP_type = Radiobutton(Frame1,text="Navision",value=2,variable=self.var_ERP_type,font=("Arial",12,"bold"),bg='white',fg="black",anchor=W).place(x=10,y=70)

        #=======Section 4: ====================
        btn_start = Button(self.root,text="Start",command=self.start,font=("Arial",14,"bold"),fg="black",bg="lightyellow",padx=20,activebackground='white',activeforeground='red',cursor='hand2')
        btn_start.place(x=550,y=300,height=60,width=100)
        btn_exit = Button(self.root,text="Exit",command=root.destroy,font=("Arial",14,"bold"),fg="black",bg="lightyellow",padx=20,activebackground='white',activeforeground='red',cursor='hand2').place(x=550,y=400,height=40,width=100)

        #=======Section 5: ====================
       

        #=======Section 6: ====================
    def browse_config_folder(self):
            
        op = filedialog.askopenfilename(title="Select Config file",filetypes=(("config files","*.config"),("all files","*.*")))
        if op != None:
            #print(op)
            self.config_name.set(op)

    def browse_filter_folder(self):
        op = filedialog.askopenfilename(title="Select Filter file",filetypes=(("Excel files","*.xlsx"),("all files","*.*")))
        if op != None:
            #print(op)
            self.filter_name.set(op)
                
    def browse_save_folder(self):
        op = filedialog.asksaveasfilename(title="Select Save file",filetypes=(("config files","*.config"),("all files","*.*")))
        if op != None:
            #print(op)
            self.save_name.set(op)

            
    #=======Section 7: ====================        


    def nav(self):
        #path_filter,path_config,path_save
        db_map = {'BKR':'65e18c5b-f0ca-46d5-823f-d2a60c26295b', # 04 Balzers Korea
        'BJP':'af596506-052b-4563-9277-b6d4f589b354', # 11 Nihon Balzers Coating
        'BTH':'5525704c-d6b4-4c5f-8245-927b7acf876c', # 12 BTH-live
        'BVN':'e30d0278-1a0f-47cd-9517-300226d48eec', # 13 Balzers Vietnam
        'BIN':'0a404fab-d717-48c3-b471-23563ecb7813', # 14 Balzers India
        'BMY':'3d404ea8-b725-48e1-8c34-7a2974bf594f', # 15 Balzers Malaysia
        'BPH':'69cf8c77-40d8-4b60-9339-a2caef902296', # 16 Balzers Philippines
        'BCN':'0a55f0ab-8f58-4a59-9bfd-2219cd71e09e', # 21 Balzers Coating (SuZhou) Co
        'BMX':'f948eb31-78d3-4e4b-80c6-a02e35a51f7b', # 17 Balzers Mexico
	    'BUS':'8f6aee57-c8d1-43f0-b247-1c3c470342f6', # 18 Balzers USA
	    'BAR':'81794210-2fd8-4589-bb80-3f4068b32724', # 19 Balzers Argentina
	    'BBR':'8b9bcce2-0fdd-4169-9b1c-6f67c5383884', # 20 Balzers Brazil
        'BDE':'e160715e-4937-48d3-b9da-50d0d592fd9a', # 03 Balzers Germany
	    'BFL':'5753e96e-74ab-4318-9ffc-4d48dc443b08', # 08 Balzers BFL
	    'BFL':'cd82e2a8-c91b-47b0-83fb-41c259b4dda8', # 09 Balzers_BCH
	    'BFR':'dc0e43f4-66ba-4e48-86bc-975d14ceb1f5' # 10 Balzers Coating S_A_S
          
        }

        rev_db_map = {v: k for k, v in db_map.items()}
        #print(rev_db_map)
        #path_filter = "C:\\Users\\ZakariaH\\Desktop\\Oerlikon\\client_sent_doc\\Filtering Template\\Navsion Filtering Rules Form.xlsx"
        #path_filter = path_filter
        df_nav= pd.read_excel(str(self.filter_name.get().replace('/','\\\\')),sheet_name=1)
        #df_nav.head()
        #df_nav.Database.unique()
        #df_nav[(df_nav['Database']=='BDE') & (df_nav['Field Type']=='GL Account')]['Value'].to_list()
        #df_nav[(df_nav['Database']=='BDE') & (df_nav['Field Type']=='Vendor')]['Value'].to_list()


        account_dict = {}
        vendor_dict = {}

        for i in df_nav.Database.unique():
            account = df_nav[(df_nav['Database']== i) & (df_nav['Field Type']=='GL Account')]['Value'].to_list()
            Vendor = df_nav[(df_nav['Database']== i) & (df_nav['Field Type']=='Vendor')]['Value'].to_list()
            account_dict[i] = account
            vendor_dict[i] = Vendor

        #print(vendor_dict)

        def vendor_Account(print_value='default',switch=0):

            """1: Account Master
            2: Vendor Master
            3: Vendor Transaction
            4: Vendor and GL Account"""

            on=switch

            pattern_account_master = '\(\s\'\w{2}.*\)'
            replace_account_master = str(account_list).replace('[','(').replace(']',')')
            ##########################################
            pattern_vendor_master = '\(\s\'\w{2}.*\)'
            replace_vendor_master = str(account_list).replace('[','(').replace(']',')')
            #########################################
            pattern_account_transaction ='(end NOT IN\s)(\(\s\'\w+\'\,.*\'\w+\'\s\))'
            replace_account_transaction = str(f"end NOT IN {str(account_list).replace('[','(').replace(']',')')}")
            ##########################
            pattern_vendor_transaction ='(AND \[No_\] NOT IN\s)(\(\s\'\w+\'\,.*\'\w+\'\s\)\s\))'
            replace_vendor_transaction = str(f"AND [No_] NOT IN {str(vendor_list).replace('[','(').replace(']',')')}")
            ##########################

            sql = file.find('sql')
            sql=sql.text
            sql=sql.strip()
            sql=" ".join(sql.split())                    

            if on == 1:
                sql = re.sub(pattern_account_master,replace_account_master,sql)
                print(f"\n\n{print_value}\n\n")
                print(sql)
            elif on == 2:
                sql = re.sub(pattern_vendor_master,replace_vendor_master,sql)
                print(f"\n\n{print_value}\n\n")
                print(sql)
            elif on == 3:
                sql = re.sub(pattern_vendor_transaction,replace_vendor_transaction,sql)
                print(f"\n\n{print_value}\n\n")
                print(sql)
            elif on == 4:
                sql = re.sub(pattern_vendor_transaction,replace_vendor_transaction,sql)
                sql = re.sub( pattern_account_transaction,replace_account_transaction,sql)
                print(f"\n\n{print_value}\n\n")
                print(sql)
            else:      
                print(f"Please select a number")
        tree = et.parse(self.config_name.get())
        #tree = et.parse(path_config)
        # Read root node
        root = tree.getroot()
        # find child node
        systems = root.find('systems')
        ###############################
        database = None    
        # loop through systems
        for system in systems:
            #print(system.attrib['id'])
            system_id = system.attrib['id']  
            database = rev_db_map[system_id]
            print(database)
            ######################
            account_list = account_dict[database]
            print(account_list[:5])
            vendor_list = vendor_dict[database]
            print(vendor_list[:5])
            ######################
            requireSystem= system
            #print(requireSystem)
            #####################
            fileGroups = requireSystem.find('fileGroups')
            print(fileGroups)
            ###################
            for fileGroup in fileGroups:
                print(f"\n\n{fileGroup.attrib['name']}")

            ##################
                files =fileGroup
            ##################
                for file in files:

                    if file.attrib !={}:
                        #print(file.attrib)

                        if file.attrib['name'] == 'Account master':
                            #print(file.attrib['name'])
                            vendor_Account('Account master',switch=1) 
                        #------------------------------------------------------------#    
                        elif file.attrib['name'] == 'Vendor master':
                            vendor_Account('Vendor master',switch=2)
                        #------------------------------------------------------------# 
                        elif file.attrib['name'] == 'Purchase invoice':
                            vendor_Account('Purchase invoice',switch=3)            
                        #------------------------------------------------------------#      
                        elif file.attrib['name'] == 'Purchase invoice line':
                            vendor_Account('Purchase invoice line',switch=4)
                        #------------------------------------------------------------#       
                        elif file.attrib['name'] == 'Purchase credit memo':
                            vendor_Account('Purchase credit memo',switch=3)
                            #----------------------------------------------------------#     
                        elif file.attrib['name'] == 'Purchase credit memo line':
                            vendor_Account('Purchase credit memo line',switch=4)

                            #------------------------------------------------------------#     
                        elif file.attrib['name'] == 'Vendor ledger':
                            vendor_Account('Vendor ledger',switch=3)
                            #------------------------------------------------------------#     
                        elif file.attrib['name'] == 'GL Entry':
                            vendor_Account('GL Entry',switch=4)
            tree.write(str(self.save_name.get()).replace('/','\\\\'))

########################### sap Start ############################################

    
    def sap(self):
        import pandas as pd 
        import numpy as np
        path= str(self.filter_name.get().replace('/','\\\\'))
        conf_file_loc=str(self.config_name.get().replace('/','\\\\'))
        #path = r'C:\Users\ZakariaH\Desktop\Oerlikon\client_sent_doc\Filtering Template\SAP Filtering Rules Form.xlsx'
        #conf_file_loc=r'C:\Users\ZakariaH\Desktop\Oerlikon\Automation\output_sap_3_0.xml'
        #pd.set_option('display.max_colwidth', None)
        def read_data(path,system_name,sheet_name=1,dtype='object'):
            """
            read data from given location.
            path = location of the file
            system_name = name of the system ex: 'MEP100'
            sheet_name= optional default to 1
            dtype= optional default to 'object'
            """
            df = pd.read_excel(path,sheet_name=sheet_name,dtype=dtype)
            df.columns = df.columns.str.strip()
            df = df[df['SAP Name']==system_name]
            return df

        def data_transform(data,column,filter):
            """
            column : 'Field Type'
            filter : 'GL Account'
            """
            df = data
            df = df[df[column]==filter]
            df = df.fillna(method='ffill',axis=1)
            return df

        def data_transform_row_num(data,groupby_col='Company Code',row_num='row_num',start=0):
            """
            groupby_col: Company Code
            """
            df = data
            df[row_num] = df.groupby([groupby_col]).cumcount()+ start
            return df

        #######################################################
        def data_transform_max_value_col(data):
            """
            adding max value column
            """
            df1 = data
            df2 = df1.groupby(['Company Code'])[['row_num','id']].max()
            df2 = df2.astype('object')
            df = df1.merge(df2,on=['id'],how='left')
            df = df.rename(columns={'row_num_x':'row_num','row_num_y':'max_value'})
            df = df.copy()
            return df
        #######################################################

        def not_null_to_1(row):
            if row['max_value'] == row['max_value']:
                return 1
            else:
                return 0


        def account_filter_bseg(row):
            """
            inatializing filter.

            """

            if row['row_num'] == 0 and row['id'] == 1:
                return  str("((BUKRS = ") + "'"+str(f"{row['Company Code']}")+"' AND HKONT NOT BETWEEN '" + row['Value From'] + "' AND '" + row['Value To']+"'"
            #################################################

            elif row['row_num'] == 0 and row['row_num'] != 1:
                return  str("(BUKRS = ") + "'"+str(f"{row['Company Code']}")+"' AND HKONT NOT BETWEEN '" + row['Value From'] + "' AND '" + row['Value To']+"'"
            else:
                return str("AND HKONT NOT BETWEEN '") + row['Value From'] + "' AND '" + row['Value To']+ "'"


        def add_backward_bracket(row):
            if row['max_value'] ==1 and row['Min_Max_id']!=2:
                return  row['filter']+") OR"
            #############################################
            elif row['max_value'] ==1 and row['Min_Max_id']==2:
                return  row['filter']+")) OR"

            else:
                return row['filter']

        def remove_first_and_last_character_from_each_line(data):
            data['filter'] = data['filter'][1:-1]
            return data



        ####################################### Account filter bseg ################################################
        def sap_mep_100_acc_bseg(path):
            mep_100 = read_data(path,'MEP100')
            mep_100 = data_transform(mep_100,'Field Type','GL Account')
            mep_100 = data_transform_row_num(mep_100)
            #mep_100.head()
            mep_100 = data_transform_row_num(mep_100,row_num='id',groupby_col='SAP Name',start=1)
            #mep_100.head()

            mep_100 = data_transform_max_value_col(mep_100)
            #mep_100.head()

            mep_100['max_value']=mep_100.apply(not_null_to_1,axis=1)
            #mep_100.head(5)
            def min_max_id(row):
                mi=np.min(mep_100.id,axis=0)
                ma=np.max(mep_100.id,axis=0)
                if row['id'] == mi:
                    return 1
                elif  row['id'] == ma:
                    return 2
                else:
                    return 0
            mep_100['Min_Max_id']=mep_100.apply(min_max_id,axis=1)
            #mep_100.head()

            mep_100['filter'] = mep_100.apply(account_filter_bseg,axis=1)
            #mep_100.head(5)

            mep_100['filter'] = mep_100.apply(add_backward_bracket,axis=1)
            #mep_100.head(20)

            mep_100_acc_filter_1=mep_100['filter']
            #mep_100_acc_filter_1.head()

            #######################################
            mep_100_acc_filter_2_unique_com_code =str(list(mep_100['Company Code'].unique())).replace("\n","").replace("\'","").replace(" ","")
            mep_100_acc_filter_2_unique_com_code=mep_100_acc_filter_2_unique_com_code.split(',')
            mep_100_acc_filter_2_unique_com_code
            mep_100_acc_filter_2_str = str(f"(BUKRS IN ({mep_100_acc_filter_2_unique_com_code} ) AND KOART = 'K') OR").replace("[","").replace("]","")
            mep_100_acc_filter_2_li =mep_100_acc_filter_2_str.split(',')

            i=0
            mep_100_acc_filter_2_new_list=[]
            while i<len(mep_100_acc_filter_2_li):
                mep_100_acc_filter_2_new_list.append(str(mep_100_acc_filter_2_li[i:i+3]).replace("[",'').replace("]",'')+',')
                i+=3
            mep_100_acc_filter_2_new_list[-1] = replace_last_occurrence(str( mep_100_acc_filter_2_new_list[-1]), ',', '')
            mep_100_acc_filter_2=pd.DataFrame(mep_100_acc_filter_2_new_list,columns=['filter'])

            mep_100_acc_filter_2=mep_100_acc_filter_2['filter'].str.replace('"','')
            mep_100_acc_filter_2=mep_100_acc_filter_2.str.replace("  ",'')
            

            #######################################

            mep_100_acc_filter_3_unique_com_code =str(list(mep_100['Company Code'].unique())).replace("\n","").replace("\'","").replace(" ","")
            mep_100_acc_filter_3_unique_com_code=mep_100_acc_filter_3_unique_com_code.split(',')
            mep_100_acc_filter_3_unique_com_code
            mep_100_acc_filter_3_str = str(f"(BUKRS NOT IN ({mep_100_acc_filter_3_unique_com_code} ))").replace("[","").replace("]","")
            mep_100_acc_filter_3_li =mep_100_acc_filter_3_str.split(',')

            i=0
            mep_100_acc_filter_3_new_list=[]
            while i<len(mep_100_acc_filter_3_li):
                mep_100_acc_filter_3_new_list.append(str(mep_100_acc_filter_3_li[i:i+3]).replace("[",'').replace("]",'')+',')
                i+=3
            mep_100_acc_filter_3_new_list[-1] = replace_last_occurrence(str( mep_100_acc_filter_3_new_list[-1]), ',', '')

            mep_100_acc_filter_3=pd.DataFrame(mep_100_acc_filter_3_new_list,columns=['filter'])


            mep_100_acc_filter_3=mep_100_acc_filter_3['filter'].str.replace('"','')
            mep_100_acc_filter_3=mep_100_acc_filter_3.str.replace("  ",'')
            mep_100_acc_filter_3


            #################################
            mep_100_acc_filters =pd.concat([mep_100_acc_filter_1,mep_100_acc_filter_2,mep_100_acc_filter_3])
            return mep_100_acc_filters



        ################################
        ####### Vendor Group ############

        def sap_mep_100_vendor(path):
            mep_100 = read_data(path,'MEP100')
            mep_100 = data_transform(mep_100,'Field Type','Vendor Group')
            mep_100_vendor_group_list = mep_100['Value From'].to_list()

            df_sap_vendor_group_str_man = str(f"(KTOKK NOT IN ({mep_100_vendor_group_list} ))")
            li =df_sap_vendor_group_str_man.split(',')
            li

            i=0
            new_list=[]
            while i<len(li):
                new_list.append(str(li[i:i+3]).replace('"','').replace('[','').replace(']',''))
                i+=3

            #new_list

            ##########################
            df_sap_vendor_group_filter_1 = list(str(f"MANDT = '@system:client' AND").split('\n'))
            df_sap_vendor_group_filter_1=pd.DataFrame(df_sap_vendor_group_filter_1,columns=['filter'])
            #df_sap_vendor_group_filter_1
            #########################

            df_sap_vendor_group_filter_2=pd.DataFrame(new_list,columns=['filter'])
            df_sap_vendor_group_filter_2
            df_sap_vendor_group_filter_2=df_sap_vendor_group_filter_2.apply(remove_first_and_last_character_from_each_line,axis=1)
            #df_sap_vendor_group_filter_2
            df_sap_vendor_group_filters=pd.concat([df_sap_vendor_group_filter_1,df_sap_vendor_group_filter_2])
            #df_sap_vendor_group_filters
            return df_sap_vendor_group_filters['filter']

        def replace_last_occurrence(s, old, new):
            return (s[:s.rfind(old)] + new + s[s.rfind(old) + len(old):])

        #################################
        def sap_eldim_vendor(path):
            eldim_vendor = read_data(path,'Eldim')
            # selece dataframe by sap name and field type
            eldim_vendor = eldim_vendor[eldim_vendor['Field Type']=='Vendor']
            #eldim_vendor
            eldim_vendor_li=eldim_vendor['Value From'].to_list()
            eldim_vendor_li =list(set(eldim_vendor_li))
            ###################################################################
            eldim_vendor_str = str(f"LIFNR NOT IN ({eldim_vendor_li})")
            eldim_vendor_str
            eldim_vendor_list =eldim_vendor_str.split(',')
            #print(len(set(eldim_vendor_list)))


            ####################################################################
            eldim_vendor_nested_list=[]

            i = 0
            while i < len(eldim_vendor_list):
                if i == 0:
                    eldim_vendor_nested_list.append(str(eldim_vendor_list[i:i+3]).replace('"','')+',')
                    i+=3  

                else:
                    eldim_vendor_nested_list.append(str(eldim_vendor_list[i:i+3]).replace('"','')+',')
                    i+=3 
            eldim_vendor_nested_list[-1]=replace_last_occurrence(str(eldim_vendor_nested_list[-1]), ',', '')  
            #eldim_vendor_nested_list
            ######################################################################
            df_eldim_vendor_filter_1 = list(str(f"MANDT = '@system:client' AND").split('\n'))
            df_eldim_vendor_filter_1 = pd.DataFrame(df_eldim_vendor_filter_1,columns=['filter'])
            df_eldim_vendor_filter_1=df_eldim_vendor_filter_1['filter']
            #########################################################################

            df_eldim_vendor_filter_2 = pd.DataFrame(eldim_vendor_nested_list,columns=['filter'])
            #df_eldim_vendor_filter_2 = df_eldim_vendor_filter_2.apply(remove_first_and_last_character_from_each_line,axis=1)
            df_eldim_vendor_filter_2 = df_eldim_vendor_filter_2['filter'].str.replace('[','',regex=False).str.replace(']','',regex=False)
            #df_eldim_vendor_filter_2=df_eldim_vendor_filter_2.str.replace(']','',regex=False)
            df_eldim_vendor_filters = pd.concat([df_eldim_vendor_filter_1,df_eldim_vendor_filter_2],axis=0)
            #df_eldim_vendor_filters
            return df_eldim_vendor_filters

        ################################# 

        def eldim_acc_filters(path,s_file='ska1',a_name='SAKNR'):
            """
            path : location of the file
            s_file : name of source file {'SKA1' or 'SKAT'}
            a_name : naming of account {'SAKNR' or 'HKONT'}
            """

            switch = s_file
            eldim_account = read_data(path,'Eldim')
            #eldim_account = data_transform(eldim_account,'Field Type','GL Account')
            #eldim_account
            eldim_account_li=eldim_account['Value From'].to_list()
            eldim_account_li =list(set(eldim_account_li))
            #print(eldim_account_li)
            ###################################################################
            eldim_account_str = str(f"{a_name} NOT IN ({eldim_account_li} )")
            eldim_account_str
            eldim_account_list =eldim_account_str.split(',')
            #print(len(set(eldim_account_list)))


            ####################################################################
            eldim_account_nested_list=[]
            i = 0
            while i < len(eldim_account_list):
            #while i < len(eldim_account_list):
                if i ==0:
                    eldim_account_nested_list.append(str(eldim_account_list[i:i+2]).replace('"','')+',')
                    i+=2
                elif i != len(eldim_account_list)-3:
                    eldim_account_nested_list.append(str(eldim_account_list[i:i+3]).replace('"','')+',')
                    i+=3       
                else:

                    eldim_account_nested_list.append(str(eldim_account_list[i:i+3]).replace('"','').replace(']',''))

                    i+=3
            #eldim_account_nested_list
            ######################################################################
            df_eldim_account_filter_1 = list(str(f"MANDT = '@system:client' AND").split('\n'))
            df_eldim_account_filter_1 = pd.DataFrame(df_eldim_account_filter_1,columns=['filter'])
            df_eldim_account_filter_1=df_eldim_account_filter_1['filter']
            #########################################################################

            df_eldim_account_filter_2 = list(str(f"SPRAS IN ( 'E', 'D', 'F', 'S', 'U' ) AND").split('\n'))
            df_eldim_account_filter_2 = pd.DataFrame(df_eldim_account_filter_2,columns=['filter'])
            df_eldim_account_filter_2=df_eldim_account_filter_2['filter']

            #########################################################################

            df_eldim_account_filter_3 = pd.DataFrame(eldim_account_nested_list,columns=['filter'])
            #df_eldim_account_filter_3 = df_eldim_account_filter_3.apply(remove_first_and_last_character_from_each_line,axis=1)
            df_eldim_account_filter_3 = df_eldim_account_filter_3['filter'].str.replace('[','',regex=False).str.replace(']','',regex=False)

            if switch == 'ska1':
                df_eldim_account_filters = pd.concat([df_eldim_account_filter_1,df_eldim_account_filter_3],axis=0)
                return df_eldim_account_filters

            elif switch == 'skat':
                df_eldim_account_filters = pd.concat([df_eldim_account_filter_1,df_eldim_account_filter_2,df_eldim_account_filter_3],axis=0)
                return df_eldim_account_filters   
            elif switch == 'bseg':
                return df_eldim_account_filter_3

        ###############################################################################
        ############# Vendor Group ####################################################

        def sap_mep_014_vendor(path):

        # selece dataframe by sap name and field type
            mep_014_vendor = read_data(path,'MEP14')
            mep_014_vendor = mep_014_vendor[mep_014_vendor['Field Type']=='Vendor Group']
            mep_014_vendor_li=mep_014_vendor['Value From'].to_list()
            mep_014_vendor_li =list(set(mep_014_vendor_li))
            ###################################################################
            mep_014_vendor_str = str(f"KTOKK NOT IN ({mep_014_vendor_li} ))")
            mep_014_vendor_str
            mep_014_vendor_list =mep_014_vendor_str.split(',')
            print(len(set(mep_014_vendor_list)))


            ####################################################################
            mep_014_vendor_nested_list=[]

            i = 0
            while i < len(mep_014_vendor_list):
                if i ==0:
                    mep_014_vendor_nested_list.append(str(mep_014_vendor_list[i:i+2]).replace('"','').replace(']',''))
                    i+=2                                           
                else:
                    mep_014_vendor_nested_list.append(str(mep_014_vendor_list[i:i+3]).replace('"','').replace(']',''))

                    i+=3
            #mep_014_vendor_nested_list[-1] = str(mep_014_vendor_nested_list[-1]).replace(',','')
            #mep_014_vendor_nested_list
            ######################################################################
            df_mep_014_vendor_filter_1 = list(str(f"MANDT = '@system:client' AND").split('\n'))
            df_mep_014_vendor_filter_1 = pd.DataFrame(df_mep_014_vendor_filter_1,columns=['filter'])
            df_mep_014_vendor_filter_1=df_mep_014_vendor_filter_1['filter']
            #########################################################################

            df_mep_014_vendor_filter_2 = pd.DataFrame(mep_014_vendor_nested_list,columns=['filter'])
            df_mep_014_vendor_filter_2 = df_mep_014_vendor_filter_2.apply(remove_first_and_last_character_from_each_line,axis=1)
            df_mep_014_vendor_filter_2 = df_mep_014_vendor_filter_2['filter'].str.replace('[','',regex=False)
            #df_mep_014_vendor_filter_2
            df_mep_014_vendor_filters = pd.concat([df_mep_014_vendor_filter_1,df_mep_014_vendor_filter_2],axis=0)
            return df_mep_014_vendor_filters
        ############################################################################################################
        ############################### SAP Account mep14 ##########################################################
        def acc_filter_mep14(path):
            mep_014_account = read_data(path,'MEP14')

            #df_sap_mep_014

            # selece dataframe by sap name and field type
            df_sap_mep100_gl_acc = mep_014_account[mep_014_account['Field Type']=='GL Account']  

            # select df_sap_bde_gl_account where Value is null
            df_sap_mep100_gl_acc_null = df_sap_mep100_gl_acc[df_sap_mep100_gl_acc['Value To'].isnull()]
            #df_sap_mep100_gl_acc_null.head()
            #df_sap_mep100_gl_acc_null.fillna(method='ffill',axis=1)

            # fillna with horizontal axis
            df_sap_mep100_gl_acc = df_sap_mep100_gl_acc.fillna(method='ffill',axis=1)
            #df_sap_mep100_gl_acc.head()
            # Generate row number of the dataframe by group
            df_sap_mep100_gl_acc['Row Number'] = df_sap_mep100_gl_acc.groupby(['Company Code']).cumcount()
            df_sap_mep100_gl_acc['id'] = df_sap_mep100_gl_acc.groupby(['SAP Name']).cumcount()+1
            #df_sap_mep100_gl_acc

            #df_sap_mep100_gl_acc.columns
            df_1=df_sap_mep100_gl_acc
            df_2 = df_sap_mep100_gl_acc.groupby(['Company Code'])[['Row Number','id']].max()
            df_2=df_2.astype('object')
            df=df_1.merge(df_2,on=['id'],how='left')

            df.rename(columns={'Row Number_x':'Row Number','Row Number_y':'Max Value'},inplace=True)
            #df

            def not_null_to_1(row):
                if row['Max Value'] == row['Max Value']:
                    return 1
                else:
                    return 0

            df['Max Value']=df.apply(not_null_to_1,axis=1)
            #df.head(30)

            #######################
            def min_max_id(row):
                mi=np.min(df.id,axis=0)
                ma=np.max(df.id,axis=0)
                if row['id'] == mi:
                    return 1
                elif  row['id'] == ma:
                    return 2
                else:
                    return 0

            ########################

            df['Min_Max_id']=df.apply(min_max_id,axis=1)
            #df

            ##########################################################
            def forward_brac(row):

                if row['Row Number'] == 0 and row['id'] == 1:
                    return  str("((BUKRS = ") + "'"+str(f"{row['Company Code']}")+"' AND HKONT NOT BETWEEN '" + row['Value From'] + "' AND '" + row['Value To']+"'"
                #################################################

                elif row['Row Number'] == 0 and row['Row Number'] != 1:
                    return  str("(BUKRS = ") + "'"+str(f"{row['Company Code']}")+"' AND HKONT NOT BETWEEN '" + row['Value From'] + "' AND '" + row['Value To']+"'"
                else:
                    return str("AND HKONT NOT BETWEEN '") + row['Value From'] + "' AND '" + row['Value To']+ "'"
            #########################################################


            df['filter'] = df.apply(forward_brac,axis=1)
            #df.head(20)

            def backward_brac(row):
                if row['Max Value'] ==1 and row['Min_Max_id']!=2:
                    return  row['filter']+") OR"
                #############################################
                elif row['Max Value'] ==1 and row['Min_Max_id']==2:
                    return  row['filter']+")) OR"

                else:
                    return row['filter']

            df['filter'] = df.apply(backward_brac,axis=1)
            #df.head(20)
            #df.tail(10)

            df_filter_1= df[['id','filter']].copy()
            df_filter_1=df_filter_1['filter']

            ################################################################################################
            ## Filter data frame 2
            df_filter_2_unique_com_code =str(list(df['Company Code'].unique())).replace("\n","").replace("\'","").replace(" ","")
            df_filter_2_unique_com_code=df_filter_2_unique_com_code.split(',')
            df_filter_2_unique_com_code
            df_man_str = str(f"(BUKRS IN ({df_filter_2_unique_com_code} ) AND KOART = 'K') ").replace("[","").replace("]","")
            li =df_man_str.split(',')

            i=0
            new_list=[]
            while i<len(li):
                new_list.append(str(li[i:i+3]).replace("[",'').replace("]",''))
                i+=3
            df_filter_2=pd.DataFrame(new_list,columns=['filter'])
            #df_filter_2

            def remove_first_and_last_character_from_each_line(df):
                df['filter'] = df['filter'][1:-1]
                return df

            df_filter_2=df_filter_2.apply(remove_first_and_last_character_from_each_line,axis=1)
            df_filter_2=df_filter_2['filter'].str.replace('"','')
            df_filter_2=df_filter_2.str.replace("  ",'')
            #df_filter_2= df_filter_2[['filter']]
            #df_filter_2
            filters=pd.concat([df_filter_1,df_filter_2],axis=0)
            return filters
        ##############################################################################################################################
        ##############################################################################################################################

        def company_filter_mep14(path):
        ############# Company Code ################################################
        # selece dataframe by sap name and field type
            mep_014_company = read_data(path,'MEP14')
            mep_014_company = mep_014_company['Company Code'].dropna().unique()

            mep_014_company
            mep_014_company_li =list(set(mep_014_company))
            ###################################################################
            mep_014_company_str = str(f"BUKRS IN ({mep_014_company_li} ))")
            mep_014_company_str
            mep_014_company_list =mep_014_company_str.split(',')
            print(len(set(mep_014_company_list)))

            ####################################################################
            mep_014_company_nested_list=[]

            i = 0
            while i < len(mep_014_company_list):
                if i ==0:
                    mep_014_company_nested_list.append(str(mep_014_company_list[i:i+3]).replace('"','').replace(']',''))
                    i+=3                                           
                else:
                    mep_014_company_nested_list.append(str(mep_014_company_list[i:i+3]).replace('"','').replace(']',''))

                    i+=3
            #mep_014_company_nested_list
            ######################################################################
            df_mep_014_company_filter_1 = list(str(f"MANDT = '@system:client' AND").split('\n'))
            df_mep_014_company_filter_1 = pd.DataFrame(df_mep_014_company_filter_1,columns=['filter'])
            df_mep_014_company_filter_1=df_mep_014_company_filter_1['filter']
            #########################################################################

            df_mep_014_company_filter_2 = pd.DataFrame(mep_014_company_nested_list,columns=['filter'])
            df_mep_014_company_filter_2 = df_mep_014_company_filter_2.apply(remove_first_and_last_character_from_each_line,axis=1)
            df_mep_014_company_filter_2 = df_mep_014_company_filter_2['filter'].str.replace('[','',regex=False)
            #df_mep_014_company_filter_2
            df_mep_014_company_filters = pd.concat([df_mep_014_company_filter_1,df_mep_014_company_filter_2],axis=0)
            return df_mep_014_company_filters

        ############################################################################################################
        ########################## CNP #############################################################################
        ############################################################################################################
        def filter_cnp_company(path):
            ############# Company Code ################################################
            # selece dataframe by sap name and field type
            cnp_company = read_data(path,'CNP')
            cnp_company = cnp_company['Company Code'].dropna().unique()

            cnp_company
            cnp_company_li =list(set(cnp_company))
            ###################################################################
            cnp_company_str = str(f"BUKRS IN ({cnp_company_li} )")
            cnp_company_str
            cnp_company_list =cnp_company_str.split(',')
            print(len(set(cnp_company_list)))


            ####################################################################
            cnp_company_nested_list=[]

            i = 0
            while i < len(cnp_company_list):
                if i ==0:
                    cnp_company_nested_list.append(str(cnp_company_list[i:i+3]).replace('"','').replace(']',''))
                    i+=3                                           
                else:
                    cnp_company_nested_list.append(str(cnp_company_list[i:i+3]).replace('"','').replace(']',''))

                    i+=3
            cnp_company_nested_list[-1]=replace_last_occurrence(str(cnp_company_nested_list[-1]), ',', '')
            
            #cnp_company_nested_list
            ######################################################################
            df_cnp_company_filter_1 = list(str(f"MANDT = '@system:client' AND").split('\n'))
            df_cnp_company_filter_1 = pd.DataFrame(df_cnp_company_filter_1,columns=['filter'])
            df_cnp_company_filter_1=df_cnp_company_filter_1['filter']
            #########################################################################

            df_cnp_company_filter_2 = pd.DataFrame(cnp_company_nested_list,columns=['filter'])
            #df_cnp_company_filter_2 = df_cnp_company_filter_2.apply(remove_first_and_last_character_from_each_line,axis=1)
            df_cnp_company_filter_2 = df_cnp_company_filter_2['filter'].str.replace('[','',regex=False)
            #df_cnp_company_filter_2
            df_cnp_company_filters = pd.concat([df_cnp_company_filter_1,df_cnp_company_filter_2],axis=0)
            return df_cnp_company_filters

        ############################################################################################################

        def filter_cnp_account_bseg(path,sys_name='CNP',field_type='GL Account',sheet_name=1,dtype='object'):
            
            # select dataframe by sap name
            df = pd.read_excel(path,sheet_name,dtype=dtype)
            df.columns = df.columns.str.strip()
            
            df = df[df['SAP Name']==sys_name]
            # selece dataframe by sap name and field type 
            df = df[df['Field Type']==field_type] 

            # select df_sap_bde_gl_account where Value is null
            df_sap_cnp_account_null = df[df['Value To'].isnull()]
            df_sap_cnp_account = df.fillna(method='ffill',axis=1)

            # Generate row number of the dataframe by group
            df_sap_cnp_account['Row Number'] = df_sap_cnp_account.groupby(['Company Code']).cumcount()
            df_sap_cnp_account['id'] = df_sap_cnp_account.groupby(['SAP Name']).cumcount()+1

            df_1=df_sap_cnp_account
            df_2 = df_sap_cnp_account.groupby(['Company Code'])[['Row Number','id']].max()
            df_2=df_2.astype('object')
            df=df_1.merge(df_2,on=['id'],how='left')

            df.rename(columns={'Row Number_x':'Row Number','Row Number_y':'Max Value'},inplace=True)

            def not_null_to_1(row):
                if row['Max Value'] == row['Max Value']:
                    return 1
                else:
                    return 0

            df['Max Value']=df.apply(not_null_to_1,axis=1)

            #######################
            def min_max_id(row):
                mi=np.min(df.id,axis=0)
                ma=np.max(df.id,axis=0)
                if row['id'] == mi:
                    return 1
                elif  row['id'] == ma:
                    return 2
                else:
                    return 0
            ########################

            df['Min_Max_id']=df.apply(min_max_id,axis=1)

            ##########################################################
            def forward_brac(row):

                if row['Row Number'] == 0 and row['id'] == 1:
                    return  str("((BUKRS = ") + "'"+str(f"{row['Company Code']}")+"' AND HKONT NOT BETWEEN '" + row['Value From'] + "' AND '" + row['Value To']+"'"
                #################################################

                elif row['Row Number'] == 0 and row['Row Number'] != 1:
                    return  str("(BUKRS = ") + "'"+str(f"{row['Company Code']}")+"' AND HKONT NOT BETWEEN '" + row['Value From'] + "' AND '" + row['Value To']+"'"
                else:
                    return str("AND HKONT NOT BETWEEN '") + row['Value From'] + "' AND '" + row['Value To']+ "'"
            #########################################################
            df['filter'] = df.apply(forward_brac,axis=1)

            def backward_brac(row):
                if row['Max Value'] ==1 and row['Min_Max_id']!=2:
                    return  row['filter']+") OR"
                #############################################
                elif row['Max Value'] ==1 and row['Min_Max_id']==2:
                    return  row['filter']+")) OR"

                else:
                    return row['filter']

            df['filter'] = df.apply(backward_brac,axis=1)
            #df.head(20)
            #df.tail(10)
            df_filter_1= df[['id','filter']].copy()
            df_filter_1=df_filter_1['filter']
            ################################################################################################
            ## Filter data frame 2
            df_filter_2_unique_com_code =str(list(df['Company Code'].unique())).replace("\n","").replace("\'","").replace(" ","")
            df_filter_2_unique_com_code=df_filter_2_unique_com_code.split(',')
            df_filter_2_unique_com_code
            df_man_str = str(f"(BUKRS IN ({df_filter_2_unique_com_code} ) AND KOART = 'K') ").replace("[","").replace("]","")
            li =df_man_str.split(',')

            i=0
            nested_list=[]
            while i<len(li):
                nested_list.append(str(li[i:i+2]).replace("[",'').replace("]",'')+',')
                i+=2
                
            nested_list[-1] = replace_last_occurrence(str( nested_list[-1]), ',', '')
                
            df_filter_2=pd.DataFrame(nested_list,columns=['filter'])
            
            

            def remove_first_and_last_character_from_each_line(df):
                df['filter'] = df['filter'][1:-1]
                return df

            #df_filter_2=df_filter_2.apply(remove_first_and_last_character_from_each_line,axis=1)
            df_filter_2=df_filter_2['filter'].str.replace('"','')
            df_filter_2=df_filter_2.str.replace("  ",'')
            #df_filter_2= df_filter_2[['filter']]
            #df_filter_2
            filters=pd.concat([df_filter_1,df_filter_2],axis=0)
            filters
            #filters.to_csv("filter.csv",sep='|', index=False)
            #filters.tail(20
            return filters
        #############################################################################################################################


        def filter_cnp_vendor(path):
            cnp_vendor = read_data(path,'CNP')
            # selece dataframe by sap name and field type
            cnp_vendor = cnp_vendor[cnp_vendor['Field Type']=='Vendor']
            #cnp_vendor
            cnp_vendor_li=cnp_vendor['Value From'].to_list()
            cnp_vendor_li =list(set(cnp_vendor_li))
            ###################################################################
            cnp_vendor_str = str(f"LIFNR NOT IN ({cnp_vendor_li})")
            cnp_vendor_str
            cnp_vendor_list =cnp_vendor_str.split(',')
            print(len(set(cnp_vendor_list)))

            ####################################################################
            cnp_vendor_nested_list=[]

            i = 0
            while i < len(cnp_vendor_list):
                if i ==0:
                    cnp_vendor_nested_list.append(str(cnp_vendor_list[i:i+2]).replace('"','').replace(']',''))
                    i+=2                                           
                else:
                    cnp_vendor_nested_list.append(str(cnp_vendor_list[i:i+3]).replace('"','').replace(']',''))

                    i+=3
            #cnp_vendor_nested_list
            ######################################################################
            df_cnp_vendor_filter_1 = list(str(f"MANDT = '@system:client' AND").split('\n'))
            df_cnp_vendor_filter_1 = pd.DataFrame(df_cnp_vendor_filter_1,columns=['filter'])
            df_cnp_vendor_filter_1=df_cnp_vendor_filter_1['filter']
            #########################################################################

            df_cnp_vendor_filter_2 = pd.DataFrame(cnp_vendor_nested_list,columns=['filter'])
            #df_cnp_vendor_filter_2 = df_cnp_vendor_filter_2.apply(remove_first_and_last_character_from_each_line,axis=1)
            df_cnp_vendor_filter_2 = df_cnp_vendor_filter_2['filter'].str.replace('[','',regex=False)
            #df_cnp_vendor_filter_2
            df_cnp_vendor_filters = pd.concat([df_cnp_vendor_filter_1,df_cnp_vendor_filter_2],axis=0)
            df_cnp_vendor_filters
            
            return df_cnp_vendor_filters
        #############################################################################################################################

        def filter_cnp_account(path):
            cnp_account = read_data(path,'CNP')
            # selece dataframe by sap name and field type
            cnp_account = cnp_account[cnp_account['Field Type']=='GL Account']
            #cnp_account
            cnp_account_li=cnp_account['Value From'].to_list()
            cnp_account_li =list(set(cnp_account_li))
            ###################################################################
            cnp_account_str = str(f"SAKNR NOT IN ({cnp_account_li} )")
            cnp_account_str
            cnp_account_list =cnp_account_str.split(',')
            print(len(set(cnp_account_list)))

            ####################################################################
            cnp_account_nested_list=[]

            i = 0
            while i < len(cnp_account_list):
                if i ==0:
                    cnp_account_nested_list.append(str(cnp_account_list[i:i+2]).replace('"','').replace(']',''))
                    i+=2                                           
                else:
                    cnp_account_nested_list.append(str(cnp_account_list[i:i+4]).replace('"','').replace(']',''))

                    i+=4
            #cnp_account_nested_list
            ######################################################################
            df_cnp_account_filter_1 = list(str(f"MANDT = '@system:client' AND KTOPL ='SACH'").split('\n'))
            df_cnp_account_filter_1 = pd.DataFrame(df_cnp_account_filter_1,columns=['filter'])
            df_cnp_account_filter_1=df_cnp_account_filter_1['filter']
            #########################################################################

            df_cnp_account_filter_2 = pd.DataFrame(cnp_account_nested_list,columns=['filter'])
            #df_cnp_account_filter_2 = df_cnp_account_filter_2.apply(remove_first_and_last_character_from_each_line,axis=1)
            df_cnp_account_filter_2 = df_cnp_account_filter_2['filter'].str.replace('[','',regex=False)
            #df_cnp_account_filter_2
            df_cnp_account_filters = pd.concat([df_cnp_account_filter_1,df_cnp_account_filter_2],axis=0)
            df_cnp_account_filters
            return df_cnp_account_filters
            #########################################################################
        def filter_acc_ven_list(path,sys_name,field_type='GL Account',file_group='Master',leng=False,sheet_name=1,dtype='object'):
            import pandas as pd
            """
            path : location
            sys_name : sap system name [MEP100,'CNP']
            field_type : field types ['Vendor','Vendor Group', GL Account]
            file_group : ['Master','Purchasing','Payment']
            leng : [True, False]
            """
            df = pd.read_excel(path,sheet_name,dtype=dtype)
            df.columns = df.columns.str.strip()
            
            df = df[df['SAP Name']==sys_name]
            df = df[df['Field Type']==field_type]
            
            if field_type == 'GL Account' and file_group=='Master':
                ac_type = 'SAKNR'
            elif field_type == 'GL Account' and file_group == 'Purchasing':
                ac_type = 'HKONT'
            elif field_type =='Vendor'and (file_group == 'Master' or  file_group == 'Purchasing'):
                ac_type ='LIFNR'
            elif field_type =='Vendor Group' and (file_group =='Master'or file_group == 'Purchasing' or file_group == 'Payment') :
                ac_type ='KTOKK'
                
            ####################################################################
                
            filter_li = df['Value From'].to_list()
            filter_li = list(set(filter_li))
            filter_str = str(f"{ac_type} NOT IN ({filter_li})")
            filter_li = filter_str.split(',')

            ####################################################################
            filter_nested_list = []

            i = 0
            while i < len(filter_li):
                
                if i == 0:
                    filter_nested_list.append(str(filter_li[i:i+2]).replace('"','').replace(']','')+',')
                    i += 2                                           
                else:
                    filter_nested_list.append(str(filter_li[i:i+4]).replace('"','').replace(']','')+',')

                    i += 4
            
            filter_nested_list[-1] = replace_last_occurrence(str(filter_nested_list[-1]), ',', '')
            
            ######################################################################
            
            
            if field_type =='Vendor' or field_type =='Vendor Group':
                filter_line_1_MANDT = list(str(f"MANDT = '@system:client' AND").split('\n'))
                    
            elif field_type =='GL Account':
                filter_line_1_MANDT = list(str(f"MANDT = '@system:client' AND KTOPL ='SACH'").split('\n'))
                
            ######################################################################
            if field_type =='GL Account' and file_group =='Master' and leng==True and sys_name =='CNP' :
                filter_line_2_SPRAS = list(str(f"SPRAS IN ( 'E', 'D', 'F', 'S', 'U', '1' ) AND").split('\n'))
                
                
            else :
                filter_line_2_SPRAS = list(str(f"SPRAS IN ( 'E', 'D', 'F', 'S', 'U' ) AND").split('\n'))
        
            ########################################################################   
            filter_line_1_MANDT = pd.DataFrame(filter_line_1_MANDT,columns = ['filter'])
            filter_line_1_MANDT = filter_line_1_MANDT['filter']
            
            #########################################################################
            
            filter_line_2_SPRAS = pd.DataFrame( filter_line_2_SPRAS,columns=['filter'])
            filter_line_2_SPRAS = filter_line_2_SPRAS['filter']
            ########################################################################

            filter_3_li = pd.DataFrame(filter_nested_list,columns=['filter'])
            filter_3_li = filter_3_li['filter'].str.replace('[','',regex=False)
            
            ########################################################################
            if leng == True:
                filters = pd.concat([filter_line_1_MANDT, filter_line_2_SPRAS, filter_3_li],axis = 0)
                return filters
            
            elif leng == False:
                filters = pd.concat([filter_line_1_MANDT, filter_3_li],axis = 0)
                return filters    
                


        ############################################ Filter End ####################################################


        ############################################ XML part ######################################################

        import xml.etree.ElementTree as et
        #################################################################################
        #####                            XML Class                                  #####
        #################################################################################

        class filter_modify:
            def __init__(self,tree, filter_loc,list_of_filters):
                """
                tree = tree # tree is the tree object
                filter_loc = filter_loc # filter_loc is the location of the filter in the xml file
                list_of_filters = list_of_filters # list_of_filters is the list of filters to be added
                """
                self.tree = tree
                self.filter_loc = filter_loc
                self.list_of_filters = list_of_filters
                self.root = self.tree.getroot()

            def sap_filter_modify(self):


                li = self.list_of_filters
                ##################################################
                for child in self.root.findall(self.filter_loc):
                    for subchild in child.findall('.//filter'):
                        child.remove(subchild)
                ##################################################
                for i in li:
                    c = et.Element('filter')
                    c.text = str(i)
                    self.root.find(self.filter_loc).append(c)
                return self.tree


        ########################################################################################
        ## Find the file group name and system id in the xml file

        ########################################################################################


        def find_file_group_name_and_system_id(tree,file_group_name,sourceTable):
            dic = {}  
            root = tree.getroot()
            for child in root.findall('systems/system'):
                s_id =child.attrib['id']
                s_name = child.attrib['name']
                for subchild in child.findall(f'fileGroups/fileGroup[@name="{file_group_name}"]/file[@sourceTable="{sourceTable}"]'):
                    s_id =child.attrib['id']
                    s_name = child.attrib['name']
                    f_id = subchild.attrib['id']
                    f_name= subchild.attrib['sourceTable'] 
                    dic[s_name]={'system_id':s_id,'file_name':f_name,'file_id':f_id,'file_group':file_group_name}

            return dic    


        ########################################################################################

        def filter_cnp_com (path,sys_name='CNP',sheet_name=1,dtype='object',file_group='Purchasing'):
            
            df = pd.read_excel(path,sheet_name,dtype=dtype)
            df.columns = df.columns.str.strip()
            
            df = df[df['SAP Name']==sys_name]
            #df = read_data(path,'CNP')
            df_filter_3_unique_com_code =str(list(df['Company Code'].unique())).replace('[','').replace(']','').replace(' ','').strip()
            df_filter_3_unique_com_code=df_filter_3_unique_com_code.split(',')
            #########
            cnp_company_filter_str_1 = f"MANDT = '@system:client' AND" 
            cnp_company_filter_str_2 = f"""BSCHL IN ( '21', '22', '23', '31', '32', '33' ) AND
            BUDAT BETWEEN '@firstDay' AND '@lastDay' AND""".split('\n')
            cnp_company_filter_str_1 = pd.Series(cnp_company_filter_str_1).str.strip()
            cnp_company_filter_str_2 = pd.Series(cnp_company_filter_str_2).str.strip()
            ##########
            df_man_str = str(f"BUKRS IN ({df_filter_3_unique_com_code})")
            li =df_man_str.split(',')
            li
            i=0
            nested_list=[]
            while i<len(li):
                if i != len(li):
                    nested_list.append(str(li[i:i+4]).replace("[",'').replace("]",'')+','.strip())
                    i+=4
                else:
                    nested_list.append(str(li[i:i+3]).replace("[",'').replace("]",'')+',')
                    i+=3
            nested_list[-1] = replace_last_occurrence(str( nested_list[-1]), ',', '')
            com_filter_2 = pd.Series(nested_list)
            com_filter_2=com_filter_2.str.replace('"','')
            if file_group == 'Purchasing':
                com_filters = pd.concat([cnp_company_filter_str_1,cnp_company_filter_str_2,com_filter_2])
                return com_filters
            else:
                com_filters = pd.concat([cnp_company_filter_str_1,com_filter_2])
                return com_filters
        ############################################ XML part End ##################################################

        ############################################################################################################
        ########################################## main work #######################################################
        ############################################################################################################


        sap = {'"dataSourceId"':['"05"','"06"','"07"','"22"','"23"','"24"'] ,
                '"file_group"': ['"Master data"','"Purchasing data"','"Payment data"'],  
                '"sourceTable"': ['"LFA1"','"BSEG"','"LFB1"','"SKA1"','"SKAT"','"T001"','"CEPC"','"CSKS"','"T024E"','"BSIK"','"BSAK"']}

        dataSourceId_mep100 = sap['"dataSourceId"'][0]
        dataSourceId_eldim = sap['"dataSourceId"'][1]
        dataSourceId_mep014 = sap['"dataSourceId"'][2]
        dataSourceId_cnp = sap['"dataSourceId"'][4]

        file_group_name_master = sap['"file_group"'][0]
        file_group_name_purchasing = sap['"file_group"'][1]
        file_group_name_payment = sap['"file_group"'][2]


        sourceTable_lfa1 = sap['"sourceTable"'][0]
        sourceTable_bseg = sap['"sourceTable"'][1]
        sourceTable_lfb1 = sap['"sourceTable"'][2]
        sourceTable_ska1 = sap['"sourceTable"'][3]
        sourceTable_skat = sap['"sourceTable"'][4]
        sourceTable_t001 = sap['"sourceTable"'][5]
        sourceTable_cepc = sap['"sourceTable"'][6]
        sourceTable_csks = sap['"sourceTable"'][7]
        sourceTable_t024e = sap['"sourceTable"'][8]
        sourceTable_bsik = sap['"sourceTable"'][9]
        sourceTable_bsak = sap['"sourceTable"'][10]

        ############################################################################################################
        ########################################## MEP 100  ########################################################
        ############################################################################################################

        filter_loc_mep_100_master_lfa1=f'systems/system[@dataSourceId={dataSourceId_mep100}]/fileGroups/fileGroup[@name={file_group_name_master}]/file[@sourceTable={sourceTable_lfa1}]/filters'
        filter_loc_mep_100_purchasing_lfa1=f'systems/system[@dataSourceId={dataSourceId_mep100}]/fileGroups/fileGroup[@name={file_group_name_purchasing}]/file[@sourceTable={sourceTable_lfa1}]/filters'
        filter_loc_mep_100_payment_lfa1=f'systems/system[@dataSourceId={dataSourceId_mep100}]/fileGroups/fileGroup[@name={file_group_name_payment}]/file[@sourceTable={sourceTable_lfa1}]/filters'
        filter_loc_mep_100_purchasing_bseg=f'systems/system[@dataSourceId={dataSourceId_mep100}]/fileGroups/fileGroup[@name={file_group_name_purchasing}]/file[@sourceTable={sourceTable_bseg}]/filters'

        mep_fi_lfa1 = list(sap_mep_100_vendor(path))
        mep_fi_bseg = list(sap_mep_100_acc_bseg(path))

        tree = et.parse(conf_file_loc)
        tree_mep_100_lfa1_master = filter_modify(tree,filter_loc_mep_100_master_lfa1,mep_fi_lfa1).sap_filter_modify()
        #tree_mep_100_lfa1_master.write(r'C:\Users\ZakariaH\Desktop\Oerlikon\Automation\output_sap_4_0.xml')
        tree_mep_100_lfa1_purchasing = filter_modify(tree_mep_100_lfa1_master,filter_loc_mep_100_purchasing_lfa1,mep_fi_lfa1).sap_filter_modify()
        #tree_mep_100_lfa1_purchasing.write(r'C:\Users\ZakariaH\Desktop\Oerlikon\Automation\output_sap_4_0.xml')
        tree_mep_100_lfa1_payment = filter_modify(tree_mep_100_lfa1_purchasing,filter_loc_mep_100_payment_lfa1,mep_fi_lfa1).sap_filter_modify()
        #tree_mep_100_lfa1_payment.write(r'C:\Users\ZakariaH\Desktop\Oerlikon\Automation\output_sap_4_0.xml')
        tree_mep_100_bseg = filter_modify(tree_mep_100_lfa1_payment,filter_loc_mep_100_purchasing_bseg,mep_fi_bseg).sap_filter_modify()
        #tree_mep_100_bseg.write(r'C:\Users\ZakariaH\Desktop\Oerlikon\Automation\output_sap_4_0.xml')

        ##########################################################################################################
        ####################################### SAP Eldim ########################################################
        ##########################################################################################################

        filter_loc_eldim_master_lfa1=f'systems/system[@dataSourceId={dataSourceId_eldim}]/fileGroups/fileGroup[@name={file_group_name_master}]/file[@sourceTable={sourceTable_lfa1}]/filters'
        #print(filter_loc_eldim_master_lfa1)
        filter_loc_eldim_master_lfb1=f'systems/system[@dataSourceId={dataSourceId_eldim}]/fileGroups/fileGroup[@name={file_group_name_master}]/file[@sourceTable={sourceTable_lfb1}]/filters'
        #print(filter_loc_eldim_master_lfb1)
        filter_loc_eldim_master_ska1=f'systems/system[@dataSourceId={dataSourceId_eldim}]/fileGroups/fileGroup[@name={file_group_name_master}]/file[@sourceTable={sourceTable_ska1}]/filters'
        filter_loc_eldim_master_skat=f'systems/system[@dataSourceId={dataSourceId_eldim}]/fileGroups/fileGroup[@name={file_group_name_master}]/file[@sourceTable={sourceTable_skat}]/filters'
        #print(filter_loc_eldim_master_ska1t)
        filter_loc_eldim_purchasing_lfa1=f'systems/system[@dataSourceId={dataSourceId_eldim}]/fileGroups/fileGroup[@name={file_group_name_purchasing}]/file[@sourceTable={sourceTable_lfa1}]/filters'
        filter_loc_eldim_purchasing_bseg=f'systems/system[@dataSourceId={dataSourceId_eldim}]/fileGroups/fileGroup[@name={file_group_name_purchasing}]/file[@sourceTable={sourceTable_bseg}]/filters'

        filter_eldim_master_lfa1 = list(sap_eldim_vendor(path))
        #print(filter_eldim_master_lfa1)
        filter_eldim_master_lfb1 = list(sap_eldim_vendor(path))
        #print(filter_eldim_master_lfb1)
        filter_eldim_master_ska1 = list(eldim_acc_filters(path,s_file='ska1',a_name='SAKNR'))
        #print(filter_eldim_master_ska1)
        filter_eldim_master_skat = list(eldim_acc_filters(path,s_file='skat',a_name='SAKNR'))
        #print(filter_eldim_master_ska1t)
        filter_eldim_purchasing_lfa1 = list(sap_eldim_vendor(path))
        #print(filter_eldim_purchasing_lfa1)
        filter_eldim_purchasing_bseg = list(eldim_acc_filters(path,s_file='bseg',a_name='HKONT'))
        #print(filter_eldim_purchasing_bseg)

        tree_eldim_master_lfa1 = filter_modify(tree_mep_100_bseg,filter_loc_eldim_master_lfa1,filter_eldim_master_lfa1).sap_filter_modify()
        #print(tree_eldim_master_lfa1)
        tree_eldim_master_lfb1 = filter_modify(tree_eldim_master_lfa1,filter_loc_eldim_master_lfb1,filter_eldim_master_lfb1).sap_filter_modify()
        tree_eldim_master_ska1 = filter_modify(tree_eldim_master_lfb1,filter_loc_eldim_master_ska1,filter_eldim_master_ska1).sap_filter_modify()
        tree_eldim_master_ska1t = filter_modify(tree_eldim_master_ska1,filter_loc_eldim_master_skat,filter_eldim_master_skat).sap_filter_modify()
        tree_eldim_purchasing_lfa1 = filter_modify(tree_eldim_master_ska1t,filter_loc_eldim_purchasing_lfa1,filter_eldim_purchasing_lfa1).sap_filter_modify()
        tree_eldim_purchasing_bseg = filter_modify(tree_eldim_purchasing_lfa1,filter_loc_eldim_purchasing_bseg,filter_eldim_purchasing_bseg).sap_filter_modify()

        ##########################################################################################################
        ####################################### SAP MEP 014 ######################################################
        ##########################################################################################################

        filter_loc_014_master_lfa1=f'systems/system[@dataSourceId={dataSourceId_mep014}]/fileGroups/fileGroup[@name={file_group_name_master}]/file[@sourceTable={sourceTable_lfa1}]/filters'
        filter_loc_014_master_t001=f'systems/system[@dataSourceId={dataSourceId_mep014}]/fileGroups/fileGroup[@name={file_group_name_master}]/file[@sourceTable={sourceTable_t001}]/filters'
        filter_loc_014_purchasing_lfa1=f'systems/system[@dataSourceId={dataSourceId_mep014}]/fileGroups/fileGroup[@name={file_group_name_purchasing}]/file[@sourceTable={sourceTable_lfa1}]/filters'
        filter_loc_014_purchasing_bseg=f'systems/system[@dataSourceId={dataSourceId_mep014}]/fileGroups/fileGroup[@name={file_group_name_purchasing}]/file[@sourceTable={sourceTable_bseg}]/filters'
        filter_loc_014_purchasing_t001=f'systems/system[@dataSourceId={dataSourceId_mep014}]/fileGroups/fileGroup[@name={file_group_name_purchasing}]/file[@sourceTable={sourceTable_t001}]/filters'
        filter_loc_014_payment_lfa1=f'systems/system[@dataSourceId={dataSourceId_mep014}]/fileGroups/fileGroup[@name={file_group_name_payment}]/file[@sourceTable={sourceTable_lfa1}]/filters'

        filter_014_lfa1 = list(sap_mep_014_vendor(path))
        #print(filter_014_lfa1)
        filter_014_t001 = list(company_filter_mep14(path))
        #print(filter_014_t001)
        filter_014_purchasing_bseg = list(acc_filter_mep14(path))
        #print(filter_014_purchasing_bseg)

        tree_mep_014_master_lfa1 = filter_modify(tree_eldim_purchasing_bseg,filter_loc_014_master_lfa1,filter_014_lfa1).sap_filter_modify()
        tree_mep_014_master_t001 = filter_modify(tree_mep_014_master_lfa1,filter_loc_014_master_t001,filter_014_t001).sap_filter_modify()
        tree_mep_014_purchasing_lfa1 = filter_modify(tree_mep_014_master_t001,filter_loc_014_purchasing_lfa1,filter_014_lfa1).sap_filter_modify()
        tree_mep_014_purchasing_bseg = filter_modify(tree_mep_014_purchasing_lfa1,filter_loc_014_purchasing_bseg,filter_014_purchasing_bseg).sap_filter_modify()
        tree_mep_014_purchasing_t001 = filter_modify(tree_mep_014_purchasing_bseg,filter_loc_014_purchasing_t001,filter_014_t001).sap_filter_modify()
        tree_mep_014_payment_lfa1 = filter_modify(tree_mep_014_purchasing_t001,filter_loc_014_payment_lfa1,filter_014_lfa1).sap_filter_modify()

        ##########################################################################################################
        ####################################### SAP CNP ##########################################################
        ##########################################################################################################

        filter_loc_cnp_master_cepc=f'systems/system[@dataSourceId={dataSourceId_cnp}]/fileGroups/fileGroup[@name={file_group_name_master}]/file[@sourceTable={sourceTable_cepc}]/filters'
        filter_loc_cnp_master_csks=f'systems/system[@dataSourceId={dataSourceId_cnp}]/fileGroups/fileGroup[@name={file_group_name_master}]/file[@sourceTable={sourceTable_csks}]/filters'
        filter_loc_cnp_master_lfa1=f'systems/system[@dataSourceId={dataSourceId_cnp}]/fileGroups/fileGroup[@name={file_group_name_master}]/file[@sourceTable={sourceTable_lfa1}]/filters'
        filter_loc_cnp_master_ska1=f'systems/system[@dataSourceId={dataSourceId_cnp}]/fileGroups/fileGroup[@name={file_group_name_master}]/file[@sourceTable={sourceTable_ska1}]/filters'
        filter_loc_cnp_master_t001=f'systems/system[@dataSourceId={dataSourceId_cnp}]/fileGroups/fileGroup[@name={file_group_name_master}]/file[@sourceTable={sourceTable_t001}]/filters'
        filter_loc_cnp_master_t024e=f'systems/system[@dataSourceId={dataSourceId_cnp}]/fileGroups/fileGroup[@name={file_group_name_master}]/file[@sourceTable={sourceTable_t024e}]/filters'

        filter_loc_cnp_purchasing_lfa1=f'systems/system[@dataSourceId={dataSourceId_cnp}]/fileGroups/fileGroup[@name={file_group_name_purchasing}]/file[@sourceTable={sourceTable_lfa1}]/filters'
        filter_loc_cnp_purchasing_bsik=f'systems/system[@dataSourceId={dataSourceId_cnp}]/fileGroups/fileGroup[@name={file_group_name_purchasing}]/file[@sourceTable={sourceTable_bsik}]/filters'
        filter_loc_cnp_purchasing_bsak=f'systems/system[@dataSourceId={dataSourceId_cnp}]/fileGroups/fileGroup[@name={file_group_name_purchasing}]/file[@sourceTable={sourceTable_bsak}]/filters'
        filter_loc_cnp_purchasing_bseg=f'systems/system[@dataSourceId={dataSourceId_cnp}]/fileGroups/fileGroup[@name={file_group_name_purchasing}]/file[@sourceTable={sourceTable_bseg}]/filters'
        filter_loc_cnp_purchasing_t001=f'systems/system[@dataSourceId={dataSourceId_cnp}]/fileGroups/fileGroup[@name={file_group_name_purchasing}]/file[@sourceTable={sourceTable_t001}]/filters'

        filter_loc_cnp_payment_lfa1=f'systems/system[@dataSourceId={dataSourceId_cnp}]/fileGroups/fileGroup[@name={file_group_name_payment}]/file[@sourceTable={sourceTable_lfa1}]/filters'

        filter_cnp_vendor_master_lfa1 = list(filter_acc_ven_list(path,'CNP',field_type='Vendor',file_group='Master'))
        filter_cnp_vendor_purchasing_lfa1 = list(filter_acc_ven_list(path,'CNP',field_type='Vendor',file_group='Master'))
        #print(filter_cnp_vendor_master)
        filter_cnp_company_t001 = list(filter_cnp_com(path,file_group='Master'))
        #print(filter_cnp_company_t001)
        filer_cnp_account_master = list(filter_acc_ven_list(path,'CNP',field_type='GL Account',file_group='Master'))
        #print(filer_cnp_account_master)
        filter_cnp_acc_bseg = list(filter_cnp_account_bseg(path))
        #print(filter_cnp_acc_bseg)
        filter_cnp_bsik = list(filter_cnp_com(path,file_group='Purchasing'))
        #print(filter_cnp_bsik)
        filter_cnp_bsak = list(filter_cnp_com(path,file_group='Purchasing'))
        #print(filter_cnp_bsak)

        tree_CNP_master_cepc = filter_modify(tree_mep_014_payment_lfa1,filter_loc_cnp_master_cepc,filter_cnp_company_t001).sap_filter_modify()
        tree_CNP_master_csks = filter_modify(tree_CNP_master_cepc,filter_loc_cnp_master_csks,filter_cnp_company_t001).sap_filter_modify()
        tree_CNP_master_lfa1 = filter_modify(tree_CNP_master_csks,filter_loc_cnp_master_lfa1,filter_cnp_vendor_master_lfa1).sap_filter_modify()
        tree_CNP_master_ska1 = filter_modify(tree_CNP_master_lfa1,filter_loc_cnp_master_ska1,filer_cnp_account_master).sap_filter_modify()
        tree_CNP_master_t001 = filter_modify(tree_CNP_master_ska1,filter_loc_cnp_master_t001,filter_cnp_company_t001).sap_filter_modify()
        tree_CNP_master_t024e = filter_modify(tree_CNP_master_t001,filter_loc_cnp_master_t024e,filter_cnp_company_t001).sap_filter_modify()

        tree_CNP_purchasing_lfa1 = filter_modify(tree_CNP_master_t024e,filter_loc_cnp_purchasing_lfa1,filter_cnp_vendor_purchasing_lfa1).sap_filter_modify()
        tree_CNP_purchasing_bsik = filter_modify(tree_CNP_purchasing_lfa1,filter_loc_cnp_purchasing_bsik,filter_cnp_bsik).sap_filter_modify()
        tree_CNP_purchasing_bsak = filter_modify(tree_CNP_purchasing_bsik,filter_loc_cnp_purchasing_bsak,filter_cnp_bsak).sap_filter_modify()
        tree_CNP_purchasing_bseg = filter_modify(tree_CNP_purchasing_bsak,filter_loc_cnp_purchasing_bseg,filter_cnp_acc_bseg).sap_filter_modify()
        tree_CNP_purchasing_t001 = filter_modify(tree_CNP_purchasing_bseg,filter_loc_cnp_purchasing_t001,filter_cnp_company_t001).sap_filter_modify()

        tree_CNP_payment_lfa1 = filter_modify(tree_CNP_purchasing_t001,filter_loc_cnp_payment_lfa1,filter_cnp_vendor_master_lfa1).sap_filter_modify()

        tree_CNP_payment_lfa1.write(str(self.save_name.get()).replace('/','\\\\'))



        

############################ Sap End ##############################################    
    def start(self):

        try:
    
            if self.var_ERP_type.get()=='1':
                self.sap()
                messagebox.showinfo('SAP Filter','SAP Filter Completed')
            elif self.var_ERP_type.get()=='2':
                self.nav()
                messagebox.showinfo('Navision Filter','Navision Filter Completed')
        except Exception as e:

            messagebox.showerror('Error',message=e)
#===============================================================
root = Tk()
obj=Sortng_App(root)
root.mainloop()
