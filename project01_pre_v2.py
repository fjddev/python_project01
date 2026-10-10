# -*- coding: utf-8 -*-
"""
Created on Mon Jul 28 15:08:25 2025

@author: Stagemaster
"""
from tkinter import ttk
import tkinter as tk
from tkinter import Toplevel
from tkinter import messagebox
from tkinter import ttk
from tkinter import simpledialog, messagebox

import pandas as pd
from datetime import date
from datetime import datetime

import sqlite3

import os as os1

from tkinter import *


import sys



seed_all=''




# Get today's date in YYYY-MM-DD format
today = date.today()

formatted_date = today.isoformat()
#Toplevel
class Project:
    def __init__(self,master):     
        
        #Toplevel.__init__(self)
        super().__init__()
        ##self.geometry("650x750+550+200")
        self.master=master
        self.master.geometry("800x2000+5+5")
        self.master.title("Project on Amazon Data Analysis")
        self.master.resizable(True, True)
        self.sales_data= pd.read_excel('sales_data.xlsx')

        
        # Create a frame inside the main window
        self.resultFrame = tk.Frame(self.master, bg="lightblue",height=300, width=1000, padx=10, pady=10)
        self.resultFrame.pack(fill="x")
        
        
        self.date_entry = tk.StringVar()
        
        self.date_search_entry=tk.StringVar()
        self.imei_search_entry=tk.StringVar()
        self.location_search_entry=tk.StringVar()

        # Create form labels and entry fields




        
        
        #scrollbar = Scrollbar(self.resultFrame,orient="vertical")
        ## Show all
        tk.Label(self.resultFrame,fg='dark blue',bg='light gray',font=("Arial", 14), text="Show results for Amazon Data Analysis",borderwidth=1,relief="raised").grid(row=7, column=1, padx=20, pady=2,sticky='w')
        ''' 
           ##Treview
        self.listbox_all = tk.Listbox()
        self.listbox_all = tk.Listbox(self.resultFrame, width=70, height=10, selectmode=tk.SINGLE, bd=4, bg="light blue")
        self.listbox_all.grid(row=8, column=1, padx=0, pady=10,sticky='w')
        '''
        self.tree = ttk.Treeview(self.resultFrame, show="headings")
        self.tree.grid(row=8, column=1, padx=0, pady=10,sticky='w')
        ## ,command=self.tree.yview
        scroll=ttk.Scrollbar(self.resultFrame,orient=tk.HORIZONTAL)
        #,expand=True
        scroll.grid(row=9,column=1,sticky='we')
        self.tree.config(yscrollcommand=scroll.set)


        # Create a frame inside the main window
        self.buttonFrame = tk.Frame(self.master, bg="blue4",height=100, width=800, padx=10, pady=10)
        self.buttonFrame.pack() 
        self.filter_label=tk.Label(self.buttonFrame, text="Filter Categories", font=("Arial", 11))
        self.filter_label.grid(row=9, column=5, padx=10, pady=10, sticky="n")
        
        self.clean_button = tk.Button(self.buttonFrame, text="Clean Data",bg="orange",command=self.clean_data)
        self.clean_button.grid(row=10, column=2, padx=10, pady=0, sticky='w')   
        
        self.filter_button = tk.Button(self.buttonFrame, text="Filter",bg="Orange",command=self.filter_category)
        self.filter_button.grid(row=10, column=3, padx=10, pady=0,sticky="e")  
        
        self.filter_amount_button = tk.Button(self.buttonFrame, text="Amount",bg="Orange")
        self.filter_amount_button.grid(row=10, column=4, padx=10, pady=0,sticky='e')   

        self.desc_button = tk.Button(self.buttonFrame, text="Describe Data",bg="Orange",command=self.data_desc)
        self.desc_button.grid(row=10, column=5, padx=10, pady=0,sticky='e')   
        
        self.info_button = tk.Button(self.buttonFrame, text="DF  Info",bg="Orange",command=self.display_df_info)
        self.info_button.grid(row=10, column=6, padx=10, pady=0,sticky='e')   
        
        self.high_amount_button = tk.Button(self.buttonFrame, text="Amount/Category",bg="Orange",command=self.filter_by_amount_category)
        self.high_amount_button.grid(row=10, column=7, padx=10, pady=0,sticky='e')   
        
        v=IntVar()
        v.set(1)
        amount_rb = Radiobutton(self.buttonFrame, text="Amount",variable=v, value=1).grid(row=11, column=1)
        amount_cat_rb = Radiobutton(self.buttonFrame, text="Amount/Category",variable=v,value=2).grid(row=11,column=2)
        # Create a frame inside the main window
        self.searchFrame = tk.Frame(self.master, bg="lightblue",height=50, width=1000, padx=10, pady=10)
        self.searchFrame.pack()
        
        self.filter_label=tk.Label(self.searchFrame, text="Filter Selection", font=("Arial", 11))

        self.filter_label.grid(row=0, column=0, padx=10, pady=10, sticky="e")

        '''
        self.clean_button2 = tk.Button(self.searchFrame, text="Clean Data Test",bg="Orange")
        self.clean_button2.grid(row=3, column=2, padx=10, pady=0, sticky='w')   
        '''
        self.amountVar=tk.StringVar()
        self.amountVar.set("0")  # set text

        self.filterText = tk.StringVar()
        self.filterText.set("") # Initial value
        self.amount_label = tk.Label(self.searchFrame, textvariable=self.filterText, font=("Arial", 11))
        self.amount_label.grid(row=5, column=1, padx=10, pady=10, sticky="e")
        
        self.amount_entry = tk.Entry(self.searchFrame,textvariable=self.amountVar, font=('calibre',10,'normal'))
        self.amount_entry.grid(row=5,column=2)
        #self.amount_entry.config(state='disabled') # OR entry['state'] = 'disabled'
        
        
        self.lb_Catlabel = tk.Label(self.searchFrame, text="Select a Category:", font=("Arial", 12, "bold"))
        self.lb_Catlabel.grid(row=6,column=2, padx=10, pady=0, sticky='e') # Adds spacing around the label
        
        # Create Listbox
        self.filterCat = tk.StringVar()
        self.filterCat.set(None)  # Initial value
        self.cat_label = tk.Label(self.searchFrame, textvariable=self.filterCat, font=("Arial", 11))
        self.cat_label.grid(row=6, column=1, padx=10, pady=10, sticky="e")        
        self.listbox = tk.Listbox(self.searchFrame, height=6, width=25)
        self.listbox.grid(row=6,column=2, padx=10, pady=0, sticky='se'
                          )
        self.listbox.bind("<Double-1>", self.on_double_click)
        self.cat_chosen=''

        # Get distinct/unique values from the 'Category' column
        distinct_values = self.sales_data['Category'].unique()

        # Insert unique values into the Tkinter listbox
        for item in distinct_values:
            self.listbox.insert(tk.END, item)
            
            
        self.filter_button = tk.Button(self.searchFrame, text="Filter",bg="Orange")
        self.filter_button.grid(row=8, column=2, padx=10, pady=0, sticky='w')   
        

        self.amount_entry.grid_remove()
        self.listbox.grid_remove()
        self.amount_label.grid_remove()
        self.cat_label.grid_remove()
        self.lb_Catlabel.grid_remove()
        self.filter_button.grid_remove()
        

        
        ##debug self.searchFrame.pack_forget()
        
        
        #self.amount_entry.state(['disabled'])
        #self.amount_in.grid_remove()

        
        #self.searchFrame.forget()
        ### Note: self.destroy is the only way to remove the frame
        #  command=self.destroy,
        ##return_button = tk.Button(self.master, text="Exit", bg="Orange")
        ##return_button.grid(row=9, column=4, padx=10, pady=0, stick='e')    
        

    def hide_show_widgets(self, list_disp):   
        ## Amount
        
        
        self.filter_button.grid()
        #list_disp=['amount', 'category']
        for l in list_disp:
            if l=='amount':
                if self.amount_entry.winfo_viewable():
                    self.amount_entry.grid_remove()
                    self.amount_label.grid_remove()
                    print("remove")
                else:
                    self.amount_entry.grid()
                    self.amount_label.grid()
                    print("show")
            elif l=='category': 
               ## Category listbox
               if self.listbox.winfo_viewable():
                   self.listbox.grid_remove()
                   self.lb_Catlabel.grid_remove()
                   self.cat_label.grid_remove()
               else:
                   self.tree.delete(*self.tree.get_children())
                   self.listbox.grid()
                   self.lb_Catlabel.grid()
                   self.cat_label.grid()
            
            
        
    def helpMe(self, event):
        messagebox.showinfo("Filter ", "Enter a value in any one of the textboxes first.\n Double click on the text box you want to query.\nOnly 1 field can be filtered/queried")
        self.date_search_entry.focus()
        
    def Date_Bnd(self,event): 
        widget = event.widget
        value=event.widget.get()
        print(f"Widget clicked: {widget}")
        print(f"Widget type: {type(widget)}")
        print(value)
        self.filter_seed(value)
        
    def IMEI_Bnd(self,event): 
        widget = event.widget
        value=event.widget.get()
        print(f"Widget clicked: {widget}")
        print(f"Widget type: {type(widget)}")
        print(value)
        self.filter_seed('',value,'')
        
    def Location_Bnd(self,event): 
        widget = event.widget
        value=event.widget.get()
        print(f"Widget clicked: {widget}")
        print(f"Widget type: {type(widget)}")
        print(value)
        self.filter_seed('','',value)


    
        #print(value)

        #print(event.widget._name)
        #print(event.widget['text'])
   
    
    def submit_form(self):
        pass        
    ##
    ## Used to return category selected from the list
    ##
    def on_double_click(self,event):
       print("on_double_click")
       # Get the widget that triggered the event

       # Get the index of the selected item
       self.cat_chosen=""
       category_selection = self.listbox.curselection()
       if category_selection:
          index = category_selection[0]
          self.category_value = self.listbox.get(category_selection[0])
          self.cat_chosen= category_selection[0]
          print(category_selection[0]) 
          print(self.listbox.get(index))




       
    # =============================================================================
    # Exploring the data
    # =============================================================================
    def data_desc(self):
        print("data_desc")
        desc = self.sales_data.describe().reset_index()
        # Define columns based on the new description dataframe
        self.tree["columns"] = list(desc.columns)
        self.tree["show"] = "headings"

        # Configure column headings
        for col in self.tree["columns"]:
          self.tree.heading(col, text=col)
          self.tree.column(col, anchor=tk.CENTER, width=90)

        # Insert rows into Treeview
        for _, row in desc.iterrows():
          self.tree.insert("", "end", values=list(row))

    def display_df_info(self):
        pass
    
    # =============================================================================
    # # 
    # =============================================================================
    # =============================================================================
    # Cleaning the data
    # =============================================================================
    def clean_data(self):
        print("clean_data")
        #Check for missing values in our sales data
 
    
        #drop any rows that has any missing/nan values
        sales_data_dropped = self.sales_data.dropna()
    
        #drop rows with missing amounts based on the amount column
        sales_data_cleaned = sales_data_dropped.dropna(subset = ['Amount'])
    
    
        ## Clear the tree view
        self.tree.delete(*self.tree.get_children())
        no_index_df=sales_data_cleaned[['Category','Amount','ship-city','ship-state' ]]
        
        self.tree["columns"] = list(no_index_df.columns)
    
        for col in no_index_df.columns:
            self.tree.heading(col, text=col)      # Set the column header text
            self.tree.column(col, width=100, anchor="center") # Set column width and alignment
    
        # 5. Insert rows into the Treeview
        # .itertuples(index=False) converts each row into a clean tuple of values
        for row in no_index_df.itertuples(index=False):
            self.tree.insert("", tk.END, values=row)


     # =============================================================================
     # Slicing and Filtering Data
     # =============================================================================

    #Select a subset of our data based on the Category Column
    def filter_category(self,category_type='Top'):
        print("filter_category")
        category_data = self.sales_data[self.sales_data['Category'] == 'Top']
        
        
        ## Clear the tree view
        self.tree.delete(*self.tree.get_children())
        
        self.tree["columns"] = list(category_data.columns) 
        for col in category_data.columns:
            self.tree.heading(col, text=col)      # Set header text
            self.tree.column(col, anchor="center", width=100)  # Adjust alignment and width
        for index, row in category_data.iterrows():
            # Convert row values to a list and insert at the end of the tree
            self.tree.insert("", tk.END, values=list(row))

        
        ###print(type(category_data))
        ###print(category_data.columns)
        ###print(category_data[['Category','Size','Amount']])
        '''treeview
        for index, row in category_data.iterrows():
             # Format the row text using f-strings
             #row_text = f"{row['Category']} | {row['Size']} | {row['Amount']}"
             row_text = f"{ row['Category']:<15}  {row['Size']:<15}    ${row['Amount']:>6.2f}"
             
             self.listbox_all.insert(tk.END, row_text)
        ''' #END treeview
################################################################################################
    def filter_by_amount_category(self,cat="Top", amount_ck=1000):
        print("filter_by_amount_category")
        
        self.tree.delete(*self.tree.get_children())
        self.hide_show_widgets(['amount','category'])
        
        #self.searchFrame.pack(side="right")
        #self.amount_in.grid()
        #self.amount_entry.grid(row=4,column=2)
        #self.amount_entry.config(state=tk.NORMAL) # OR entry['state'] = 'disabled'

        self.filterText.set("Enter an amount.")
        self.filterCat.set("Select a Category")


        while True:
            try:
                #amount_cond = simpledialog.askstring("Filter by amount. Default=1000", "Enter amount to filter: ")
                self.amount_entry.config(state=tk.NORMAL) # OR entry['state'] = 'disabled'
                amount_ck=float( self.amount_entry.get())
                break
            except (ValueError, TypeError):
                messagebox.showwarning("Warning", "You  need to enter a number value or 0 to use default amount: 1000")
                continue
        #cat_cond_arr=self.sales_data['Category'].unique()
        print('cat type')
        print(type(self.cat_chosen))
        '''
        while True:
            try:
                cat_distinct=self.sales_data['Category'].unique()
                print('cat type')
                print(type(cat_distinct))
                sys.exit()
                cat_cond = simpledialog.askstring("Category Filter. Default=Top", "Enter category to filter: ")
                break
            except ():
                messagebox.showwarning("Warning", "You  need to enter a number value or 0 to use default amount: 1000")
                continue
            
            
            
        '''
        if amount_ck==0:
            amount_ck=1000
        print("debug: 01")
                
        ## Clear the tree view
        #self.tree.delete(*self.tree.get_children())
        
#df["address"].str.contains(my_var).any()
#(self.cat_chosen.isin(self.sales_data['Category'].isin(self.cat_chosen))
        filtered_data=self.sales_data[(self.sales_data['Amount'].values >= amount_ck) & (self.sales_data['Category'].str.contains(self.cat_chosen).any())]
        if filtered_data.empty:
            messagebox.showwarning("Warning", "Results not found for your filtered selection")
            return
        print("debug: 02")
        sorted_df = filtered_data.sort_values(by='Amount')
    
        no_index_df=sorted_df[['Category','Amount','ship-city','ship-state' ]]
        print(no_index_df.to_string(index=False))
        self.tree["columns"] = list(no_index_df.columns)
        print("debug: 03")
        for col in no_index_df.columns:
            self.tree.heading(col, text=col)      # Set the column header text
            self.tree.column(col, width=100, anchor="center") # Set column width and alignment
        print("debug: 04")
        ## Clear the tree view
        self.tree.delete(*self.tree.get_children())
        # 5. Insert rows into the Treeview
        # .itertuples(index=False) converts each row into a clean tuple of values
        for row in no_index_df.itertuples(index=False):
            self.tree.insert("", tk.END, values=row)
        print("debug: 05")
            
    def filter_by_category_test(self,cat="Top"):
        
        print("filter_by_category_test")
        self.filterCat.set("Select a Category")


        cat_cond=""
        while True:
            try:
                #amount_cond = simpledialog.askstring("Filter by amount. Default=1000", "Enter amount to filter: ")
                self.amount_entry.config(state=tk.NORMAL) # OR entry['state'] = 'disabled'
                amount_ck=int( self.amount_entry.get())
                break
            except (ValueError, TypeError):
                messagebox.showwarning("Warning", "You  need to enter a number value or 0 to use default amount: 1000")
                continue
        cat_distinct=self.sales_data['Category'].unique()
        print('cat type')
        print(type(cat_distinct))
           
        while True:
            try:
                cat_distinct=self.sales_data['Category'].unique()
                print('cat type')
                print(type(cat_distinct))
                cat_cond = simpledialog.askstring("Category Filter. Default=Top", "Enter category to filter: ")
                break
            except ():
                messagebox.showwarning("Warning", "You  need to enter a number value or 0 to use default amount: 1000")
                continue
            
        
        if amount_ck==0:
            amount_ck=1000
                
        ## Clear the tree view
        self.tree.delete(*self.tree.get_children())
        filtered_data=self.sales_data[(self.sales_data['Amount'] > amount_ck) & (self.sales_data['Category'] == cat)]
        if filtered_data.empty:
            messagebox.showwarning("Warning", "Results not found for your filtered selection")
            return
        
        sorted_df = filtered_data.sort_values(by='Amount')
    
        no_index_df=sorted_df[['Category','Amount','ship-city','ship-state' ]]
        print(no_index_df.to_string(index=False))
        self.tree["columns"] = list(no_index_df.columns)
    
        for col in no_index_df.columns:
            self.tree.heading(col, text=col)      # Set the column header text
            self.tree.column(col, width=100, anchor="center") # Set column width and alignment
    
        # 5. Insert rows into the Treeview
        # .itertuples(index=False) converts each row into a clean tuple of values
        for row in no_index_df.itertuples(index=False):
            self.tree.insert("", tk.END, values=row)            

################################################################################################
    #Select a subset of our data where the Amount > 1000
    def filter_by_amount(self, amount_ck=1000):
        print("filter_by_amount")
        ## Clear the tree view
        self.tree.delete(*self.tree.get_children())
        high_amount_data = self.sales_data[self.sales_data['Amount'] > amount_ck]
        print(high_amount_data[['Category','Amount']])

        #self.tree["columns"] = list(high_amount_data.columns) 
        self.tree["columns"] = ['Category','Amount']
        '''
        for col in high_amount_data.columns:
            self.tree.heading(col, text=col)      # Set header text
            self.tree.column(col, anchor="center", width=100)  # Adjust alignment and width
        '''
        for index, row in high_amount_data.iterrows():
            # Convert row values to a list and insert at the end of the tree
            print(f"row: {row}")
            self.tree.insert("", tk.END, values=list(row))
    

    #Select a subset of data based on multiple conditions
    def filter_multi_conditions(self,sales_data):
        print("filter_multi_conditions")
        filtered_data = sales_data[(sales_data['Category'] == 'Top') & (sales_data['Qty'] == 3)]
        print(filtered_data[['Category','Size','Amount','Fulfilment']])

    # =============================================================================
    # Aggregating Data
    # =============================================================================
    def aggregate_sales(self,sales_data):
        print("aggregate_sales")
        #total sales by category
        category_totals = sales_data.groupby('Category')['Amount'].sum()         ##Groupby Category and then sum
        category_totals = sales_data.groupby('Category', as_index=False)['Amount'].sum()     ##Gives us an index
        category_totals = category_totals.sort_values('Amount', ascending=False)
    
        #calculate the average Amount by Category and Fulfilment
        fulfilment_averages = sales_data.groupby(['Category', 'Fulfilment'], as_index=False)['Amount'].mean()
    
        fulfilment_averages = fulfilment_averages.sort_values('Amount', ascending=False)
    
    
        #calculate the average Amount by Category and Status
        status_averages = sales_data.groupby(['Category', 'Status'], as_index=False)['Amount'].mean()
        ##  Index | Category  | Fulfilment  | Amount
        status_averages = status_averages.sort_values('Amount', ascending=False)
        print(status_averages)
        print(status_averages.columns)
        #print(status_averages[['Category','Size','Amount','Fulfilment']])

    #calculate total sales by shipment and fulfilment
    def total_sales(self,sales_data):
        total_sales_shipandfulfil = sales_data.groupby(['Courier Status', 'Fulfilment'], as_index=False)['Amount'].sum()
        total_sales_shipandfulfil = total_sales_shipandfulfil.sort_values('Amount', ascending=False)
        total_sales_shipandfulfil.rename(columns={'Courier Status' : 'Shipment'}, inplace = True)
        return total_sales_shipandfulfil

    # =============================================================================
    # Exporting the Data
    # =============================================================================
    def export_to_excel(self,status_averages,total_sales_shipandfulfil):
        status_averages.to_excel('average_sales_by_category_and_status.xlsx', index=False)
        total_sales_shipandfulfil.to_excel('total_sales_by_ship_and_fulfil.xlsx', index=False)




    def getCurrentDateTime(self):
        # Get the current date and time
        current_datetime = datetime.now()

        # Convert to string with a specific format
        #formatted_date = current_datetime.strftime("%Y-%m-%d %H:%M:%S")
        formatted_date = current_datetime.strftime("%Y%m%d%H%M%S")
        return formatted_date

def start_project():
    print("Start")
    root = tk.Tk()

    proj=Project(root)
    root.mainloop()


if __name__ == "__main__":
    # This code ONLY runs when you execute math_helper.py directly
    start_project()
  
