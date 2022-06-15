from tkinter import *
from tkinter import ttk, filedialog
import xml.etree.ElementTree as et
import pandas as pd
import re


class Sortng_App:
    def __init__(self,root) -> None:
        self.root=root
        self.root.title("Oerlikon Filtering App")
        self.root.geometry("700x500+0+0")
        self.root.configure(background='white')
        self.logo_incon = PhotoImage(file=r"C:\Users\ZakariaH\Documents\File Sorting App\images\logo.png")
        title=Label(self.root,text="Filtering App",font=("Arial",14,"bold"),bg="yellow",fg="red",anchor=W,image=self.logo_incon,compound=LEFT)
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
        btn_start = Button(self.root,text="Start",command=self.test,font=("Arial",14,"bold"),fg="black",bg="lightyellow",padx=20,activebackground='white',activeforeground='red',cursor='hand2')
        btn_start.place(x=550,y=300,height=60,width=100)
        btn_exit = Button(self.root,text="Exit",command=root.destroy,font=("Arial",14,"bold"),fg="black",bg="lightyellow",padx=20,activebackground='white',activeforeground='red',cursor='hand2').place(x=550,y=400,height=40,width=100)

        #=======Section 5: ====================
    def browse_config_folder(self):
            
        op = filedialog.askopenfilename(title="Select Config file",filetypes=(("config files","*.config"),("all files","*.*")))
        if op != None:
            print(op)
            self.config_name.set(op)

    def browse_filter_folder(self):
        op = filedialog.askopenfilename(title="Select Filter file",filetypes=(("Excel files","*.xlsx"),("all files","*.*")))
        if op != None:
            print(op)
            self.filter_name.set(op)
                
    def browse_save_folder(self):
        op = filedialog.asksaveasfilename(title="Select Save file",filetypes=(("config files","*.config"),("all files","*.*")))
        if op != None:
            print(op)
            self.save_name.set(op)

    def test(self):
        #path_filter,path_config,path_save
        db_map = {'BKR':'65e18c5b-f0ca-46d5-823f-d2a60c26295b', # 04 Balzers Korea
        'BJP':'af596506-052b-4563-9277-b6d4f589b354', # 11 Nihon Balzers Coating
        'BTH':'5525704c-d6b4-4c5f-8245-927b7acf876c', # 12 BTH-live
        'BVN':'e30d0278-1a0f-47cd-9517-300226d48eec', # 13 Balzers Vietnam
        'BIN':'0a404fab-d717-48c3-b471-23563ecb7813', # 14 Balzers India
        'BMY':'3d404ea8-b725-48e1-8c34-7a2974bf594f', # 15 Balzers Malaysia
        'BPH':'69cf8c77-40d8-4b60-9339-a2caef902296', # 16 Balzers Philippines
        'BCN':'0a55f0ab-8f58-4a59-9bfd-2219cd71e09e'  # 21 Balzers Coating (SuZhou) Co
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



#===============================================================
root = Tk()
obj=Sortng_App(root)
root.mainloop()
