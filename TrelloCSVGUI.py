import tkinter as tk
from tkinter import filedialog
from functools import partial
import csv
import sqlite3
import pandas as pd
import sys
import os
from os import path
import subprocess

# excluded lists
guideList = '62e056d02f149d20d4051694'  # '62e056d02f149d20d4051694' = Guide
rbtList = '62e0a66909940c66425f4e59'  # '62e0a66909940c66425f4e59' = Roadmap Building tasks
completeList = '62e05775c97ca95c1f490398'  # '62e05775c97ca95c1f490398' = Complete (Dev Done)
ideas2021List = '6037b11252d9416ed766d30e'  # '6037b11252d9416ed766d30e' = Ideas to Review - from 2021
releasedList = '6318f2b99ece920258808c4c'  # '6318f2b99ece920258808c4c' = Released
released2023List = '63bf1c35ba364f01c442327f' # 2023 Q1 releases
pitOfDespairList = '6037b1889454270830207261' # pit of despair

exclude = [guideList, rbtList, completeList, ideas2021List, releasedList, released2023List, pitOfDespairList]

#excluded cards
separatorCard = '631f8bae6646580132cd4510'

#included lists
activeDevList = '6037b11252d9416ed766d30f' # activeDevList
next3List = '62e059d937cbc61122aa168a' #next3
next46List = '6320da300ff66a00c9d54084' #4-6
laterList = '6037b17835a3218900975486' #later
prioritizationList = '62e09ddcd0afaf297c309bbc' #prioritization
triageList = '62e0a77cf4794110c551efdd'#triage

includeOrderedList = [activeDevList, next3List, next46List, laterList, prioritizationList, triageList]


class Editor(tk.Tk):
    def __init__(self):
        super().__init__()

        self.FONT_SIZE = 12
        self.WINDOW_TITLE = "Excel From CSV"
        self.standard_font = (None, 16)

        self.open_file = ""
        self.title(self.WINDOW_TITLE)
        self.geometry("400x300")
        self.bind("<Control-o>", self.file_open)
        self.main_frame = tk.Frame(self, width=200, height=300, bg="lightgrey")
        self.open_button = tk.Button(self.main_frame, text="Choose File", bg="lightgrey", fg="black", command=self.file_open, font=self.standard_font)
        self.goto_button = tk.Button(self.main_frame, text="Go to File", bg="lightgrey", fg="black", command=self.goto_file, font=self.standard_font, state="disabled")

        self.main_frame.pack(fill=tk.BOTH, expand=1)
        self.open_button.pack(fill=tk.X, padx=50)
        self.goto_button.pack(fill=tk.X, padx=50)


    def goto_file(self, event=None):
        if os.path.exists(self.open_file):
            subprocess.call(["open", os.path.dirname(self.open_file)])


    def file_open(self, event=None):
        file_to_open = filedialog.askopenfilename()

        if file_to_open:
            self.open_file = file_to_open
        else:
            return

        # if getattr(sys, 'frozen', False):
        #     base_dir = os.path.dirname(sys.executable)
        # elif __file__:
        #     base_dir = os.path.dirname(__file__)
        base_dir = os.path.dirname(self.open_file)

        tmpfolder = os.path.join(base_dir, './tmp')
        if (os.path.exists(tmpfolder) == False):
            os.mkdir(tmpfolder)

        # open input CSV file as source
        # open output CSV file as result
        # input = os.path.join(base_dir, './roadmap.csv')
        input = self.open_file
        filteredRows = os.path.join(tmpfolder, './filtered_rows.csv')

        with open(input, "r") as source:
            reader = csv.reader(source)

            with open(filteredRows, "w+") as result:
                writer = csv.writer(result)
                included_line_count = 0
                excluded_line_count = 0
                for r in reader:
                    if (excluded_line_count == 0):
                        excluded_line_count += 1
                        continue
                    elif (r[0] == separatorCard):
                        excluded_line_count += 1
                        continue
                    elif ((r[14]) in exclude):
                        # print(f'Unused Lists: \t{r[1]} is on {r[15]} list')
                        excluded_line_count += 1
                        continue
                    elif (r[18]) == 'FALSE' or (r[18]) == 'false': # not archived
                        included_line_count += 1
                        writer.writerow(r)
                    else:
                        excluded_line_count +=1

                print(f'Included {included_line_count} lines.')
                print(f'Excluded {excluded_line_count} lines.')
                print(f'Total {included_line_count + excluded_line_count} lines.')

        filteredColumns = os.path.join(tmpfolder, './filtered_columns.csv')

        added_rows = 0
        #active dev
        with open(filteredRows, "r") as source2:
            reader2 = csv.reader(source2)
            with open(filteredColumns, "w+") as result:
                writer = csv.writer(result)
                line_count = 0
                for r in reader2:
                    if (r[14] == '6037b11252d9416ed766d30f'):
                        writer.writerow((r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[13],
                            r[14], r[15], r[16], r[17], r[18], r[19], r[20], r[21], r[22],
                            r[23], r[24], r[25], r[26], r[27], r[28], r[29], r[30], r[31]))
                        line_count += 1
                        added_rows += 1

        #next 3 months
        with open(filteredRows, "r") as source2:
            reader2 = csv.reader(source2)
            with open(filteredColumns, "a+") as result:
                writer = csv.writer(result)
                line_count = 0
                for r in reader2:
                    if (r[14] == '62e059d937cbc61122aa168a'):
                        writer.writerow((r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[13],
                            r[14], r[15], r[16], r[17], r[18], r[19], r[20], r[21], r[22],
                            r[23], r[24], r[25], r[26], r[27], r[28], r[29], r[30], r[31]))
                        line_count += 1
                        added_rows += 1

        #4-6 months
        with open(filteredRows, "r") as source2:
            reader2 = csv.reader(source2)
            with open(filteredColumns, "a+") as result:
                writer = csv.writer(result)
                line_count = 0
                for r in reader2:
                    if (r[14] == '6320da300ff66a00c9d54084'):
                        writer.writerow((r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[13],
                            r[14], r[15], r[16], r[17], r[18], r[19], r[20], r[21], r[22],
                            r[23], r[24], r[25], r[26], r[27], r[28], r[29], r[30], r[31]))
                        line_count += 1
                        added_rows += 1

        #later
        with open(filteredRows, "r") as source2:
            reader2 = csv.reader(source2)
            with open(filteredColumns, "a+") as result:
                writer = csv.writer(result)
                line_count = 0
                for r in reader2:
                    if (r[14] == '6037b17835a3218900975486'):
                        writer.writerow((r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[13],
                            r[14], r[15], r[16], r[17], r[18], r[19], r[20], r[21], r[22],
                            r[23], r[24], r[25], r[26], r[27], r[28], r[29], r[30], r[31]))
                        line_count += 1
                        added_rows += 1

        #prioritization
        with open(filteredRows, "r") as source2:
            reader2 = csv.reader(source2)
            with open(filteredColumns, "a+") as result:
                writer = csv.writer(result)
                line_count = 0
                for r in reader2:
                    if (r[14] == '62e09ddcd0afaf297c309bbc'):
                        writer.writerow((r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[13],
                            r[14], r[15], r[16], r[17], r[18], r[19], r[20], r[21], r[22],
                            r[23], r[24], r[25], r[26], r[27], r[28], r[29], r[30], r[31]))
                        line_count += 1
                        added_rows += 1

        #triage
        with open(filteredRows, "r") as source2:
            reader2 = csv.reader(source2)
            with open(filteredColumns, "a+") as result:
                writer = csv.writer(result)
                for r in reader2:
                    if (r[14] == '62e0a77cf4794110c551efdd'):
                        writer.writerow((r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[13],
                            r[14], r[15], r[16], r[17], r[18], r[19], r[20], r[21], r[22],
                            r[23], r[24], r[25], r[26], r[27], r[28], r[29], r[30], r[31]))
                        line_count += 1
                        added_rows += 1
                print(f'total added: {added_rows} lines.')


        dbPath = os.path.join(tmpfolder, './trello.db')
        connection = sqlite3.connect(dbPath)

        # Creating a cursor object to execute
        # SQL queries on a database table
        cursor = connection.cursor()

        try:
            cursor.execute("DROP TABLE roadmap")
        except:
            print('creating table')

        # Table Definition
        create_table = '''CREATE TABLE roadmap(
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        Card_ID TEXT NOT NULL,
                        Card_Name TEXT NOT NULL,
                        Card_URL TEXT NOT NULL,
                        Card_Description TEXT NOT NULL,
                        Labels TEXT NOT NULL,
                        Members TEXT NOT NULL,
                        Due_Date TEXT NOT NULL,
                        Last_Activity_Date TEXT NOT NULL,
                        List_ID TEXT NOT NULL,
                        List_Name TEXT NOT NULL,
                        Board_ID TEXT NOT NULL,
                        Board_Name TEXT NOT NULL,
                        Archived TEXT NOT NULL,
                        Start_Date TEXT NOT NULL,
                        Due_Complete TEXT NOT NULL,
                        Customers TEXT NOT NULL,
                        Desired_Date TEXT NOT NULL,
                        Priority TEXT NOT NULL,
                        Investment_Area TEXT NOT NULL,
                        Inv_Subarea TEXT NOT NULL,
                        Scaling TEXT NOT NULL,
                        Contractual_Obligation TEXT NOT NULL,
                        Sales_Oppy TEXT NOT NULL,
                        Op_Efficiency TEXT NOT NULL,
                        Eng_Effort TEXT NOT NULL,
                        Product_Value TEXT NOT NULL);
                        '''

        # Creating the table into our
        # database
        cursor.execute(create_table)

        # Opening the temp .csv file
        file = open(filteredColumns)

        # Reading the contents of the
        # person-records.csv file
        contents = csv.reader(file)

        # SQL query to insert data into the
        # person table
        insert_records = """INSERT INTO roadmap
            (Card_ID,
            Card_Name,
            Card_URL,
            Card_Description,
            Labels,
            Members,
            Due_Date,
            Last_Activity_Date,
            List_ID,
            List_Name,
            Board_ID,
            Board_Name,
            Archived,
            Start_Date,
            Due_Complete,
            Customers,
            Desired_Date,
            Priority,
            Investment_Area,
            Inv_Subarea,
            Scaling,
            Contractual_Obligation,
            Sales_Oppy,
            Op_Efficiency,
            Eng_Effort,
            Product_Value
            ) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
            ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"""

        # Importing the contents of the file into our table
        cursor.executemany(insert_records, contents)

        # Committing the changes
        connection.commit()

        select_all = """SELECT id, Card_Name, Card_URL, Card_Description,
                        Investment_Area, Inv_Subarea, Scaling, Contractual_Obligation AS Obligation,
                        Sales_Oppy AS Sales, Op_Efficiency AS Efficiency, Eng_Effort AS Effort, Product_Value AS Value,
                        Labels, Members, List_Name, Customers
                        FROM roadmap"""
        cursor.execute(select_all)

        with open(tmpfolder + "/full_roadmap.csv", 'w',newline='') as csv_file:
            csv_writer = csv.writer(csv_file)
            csv_writer.writerow([i[0] for i in cursor.description])
            csv_writer.writerows(cursor)

        healthbridge = """SELECT id, Card_Name, Card_URL,
                        Customers, Scaling, Contractual_Obligation AS Obligation,
                        Sales_Oppy AS Sales, Op_Efficiency AS Efficiency, Eng_Effort AS Effort, Product_Value AS Value
                        FROM roadmap WHERE Labels LIKE '%green%'"""
        cursor.execute(healthbridge)

        with open(tmpfolder + "/healthbridge.csv", 'w',newline='') as csv_file:
            csv_writer = csv.writer(csv_file)
            csv_writer.writerow([i[0] for i in cursor.description])
            csv_writer.writerows(cursor)


        impact_core = """SELECT id, Card_Name, Card_URL,
                        Customers, Scaling, Contractual_Obligation AS Obligation,
                        Sales_Oppy AS Sales, Op_Efficiency AS Efficiency, Eng_Effort AS Effort, Product_Value AS Value
                        FROM roadmap WHERE Labels LIKE '%red%'"""
        cursor.execute(impact_core)

        with open(tmpfolder + "/impact_core.csv", 'w',newline='') as csv_file:
            csv_writer = csv.writer(csv_file)
            csv_writer.writerow([i[0] for i in cursor.description])
            csv_writer.writerows(cursor)

        impact_integrations = """SELECT id, Card_Name, Card_URL,
                        Customers, Scaling, Contractual_Obligation AS Obligation,
                        Sales_Oppy AS Sales, Op_Efficiency AS Efficiency, Eng_Effort AS Effort, Product_Value AS Value
                        FROM roadmap WHERE Labels LIKE '%orange%'"""
        cursor.execute(impact_integrations)

        with open(tmpfolder + "/impact_integrations.csv", 'w',newline='') as csv_file:
            csv_writer = csv.writer(csv_file)
            csv_writer.writerow([i[0] for i in cursor.description])
            csv_writer.writerows(cursor)

        platform = """SELECT id, Card_Name, Card_URL,
                        Customers, Scaling, Contractual_Obligation AS Obligation,
                        Sales_Oppy AS Sales, Op_Efficiency AS Efficiency, Eng_Effort AS Effort, Product_Value AS Value
                        FROM roadmap WHERE Labels LIKE '%blue%'"""
        cursor.execute(platform)

        with open(tmpfolder + "/platform.csv", 'w',newline='') as csv_file:
            csv_writer = csv.writer(csv_file)
            csv_writer.writerow([i[0] for i in cursor.description])
            csv_writer.writerows(cursor)

        all_impact = """SELECT id, Card_Name, Card_URL,
                        Customers, Scaling, Contractual_Obligation AS Obligation,
                        Sales_Oppy AS Sales, Op_Efficiency AS Efficiency, Eng_Effort AS Effort, Product_Value AS Value
                        FROM roadmap WHERE Labels LIKE '%red%' OR Labels LIKE '%orange'"""
        cursor.execute(all_impact)

        with open(tmpfolder + "/all_impact.csv", 'w',newline='') as csv_file:
            csv_writer = csv.writer(csv_file)
            csv_writer.writerow([i[0] for i in cursor.description])
            csv_writer.writerows(cursor)


        # closing the database connection
        connection.close()


        excelPath = os.path.join(base_dir, './roadmap.xlsx')
        # Create a Pandas Excel writer using XlsxWriter
        writer = pd.ExcelWriter(excelPath, engine='xlsxwriter')

        workbook  = writer.book
        cell_format = workbook.add_format()
        cell_format.set_text_wrap()

        df = pd.read_csv(tmpfolder + "/platform.csv")
        df["Customers"].fillna("--", inplace = True)
        df.to_excel(writer, sheet_name='platform', index=False)

        platformSheet = writer.sheets['platform']
        platformSheet.set_column_pixels('A:A', 30)
        platformSheet.set_column_pixels('B:B', 500, cell_format)
        platformSheet.set_column_pixels('C:C', 138)
        platformSheet.set_column_pixels('D:D', 88, cell_format)
        platformSheet.set_column_pixels('E:J', 75)

        df = pd.read_csv(tmpfolder + "/impact_core.csv")
        df["Customers"].fillna("--", inplace = True)
        df.to_excel(writer, sheet_name='impact_core', index=False)
        coreSheet = writer.sheets['impact_core']
        coreSheet.set_column_pixels('A:A', 30)
        coreSheet.set_column_pixels('B:B', 500, cell_format)
        coreSheet.set_column_pixels('C:C', 138)
        coreSheet.set_column_pixels('D:D', 88, cell_format)
        coreSheet.set_column_pixels('E:J', 75)

        df = pd.read_csv(tmpfolder + "/impact_integrations.csv")
        df["Customers"].fillna("--", inplace = True)
        df.to_excel(writer, sheet_name='impact_integrations', index=False)
        integrationsSheet = writer.sheets['impact_integrations']
        integrationsSheet.set_column_pixels('A:A', 30)
        integrationsSheet.set_column_pixels('B:B', 500, cell_format)
        integrationsSheet.set_column_pixels('C:C', 138)
        integrationsSheet.set_column_pixels('D:D', 88, cell_format)
        integrationsSheet.set_column_pixels('E:J', 75)

        df = pd.read_csv(tmpfolder + "/healthbridge.csv")
        df["Customers"].fillna("--", inplace = True)
        df.to_excel(writer, sheet_name='healthbridge', index=False)
        healthbridgeSheet = writer.sheets['healthbridge']
        healthbridgeSheet.set_column_pixels('A:A', 30)
        healthbridgeSheet.set_column_pixels('B:B', 500, cell_format)
        healthbridgeSheet.set_column_pixels('C:C', 138)
        healthbridgeSheet.set_column_pixels('D:D', 88, cell_format)
        healthbridgeSheet.set_column_pixels('E:J', 75)

        df = pd.read_csv(tmpfolder + "/all_impact.csv")
        df["Customers"].fillna("--", inplace = True)
        df.to_excel(writer, sheet_name='all_impact', index=False)
        healthbridgeSheet = writer.sheets['all_impact']
        healthbridgeSheet.set_column_pixels('A:A', 30)
        healthbridgeSheet.set_column_pixels('B:B', 500, cell_format)
        healthbridgeSheet.set_column_pixels('C:C', 138)
        healthbridgeSheet.set_column_pixels('D:D', 88, cell_format)
        healthbridgeSheet.set_column_pixels('E:J', 75)

        # Create ALL Items
        df = pd.read_csv(tmpfolder + "/full_roadmap.csv")
        df["Customers"].fillna("--", inplace = True)
        df["Investment_Area"].fillna("--", inplace = True)
        df["Inv_Subarea"].fillna("--", inplace = True)
        df["Members"].fillna("--", inplace = True)
        df.to_excel(writer, sheet_name='All Items', index=False)

        allItemsSheet = writer.sheets['All Items']
        allItemsSheet.set_column_pixels('A:A', 30)
        allItemsSheet.set_column_pixels('B:B', 230, cell_format)
        allItemsSheet.set_column_pixels('C:C', 138)
        allItemsSheet.set_column_pixels('D:D', 575, cell_format)
        allItemsSheet.set_column_pixels('E:E', 102, cell_format)
        allItemsSheet.set_column_pixels('F:F', 115, cell_format)
        allItemsSheet.set_column_pixels('G:L', 70)
        allItemsSheet.set_column_pixels('M:M', 92, cell_format)
        allItemsSheet.set_column_pixels('N:N', 104, cell_format)
        allItemsSheet.set_column_pixels('O:O', 115, cell_format)
        allItemsSheet.set_column_pixels('P:P', 88, cell_format)

        # Save Data to File
        writer.close()

        self.goto_button.configure(state="normal")




if __name__ == "__main__":
    editor = Editor()
    editor.mainloop()






    #
    #     self.start_button = tk.Button(self.main_frame, text="Start", bg="lightgrey", fg="black", command=self.start, font=self.standard_font)
    #     self.start_button.pack(fill=tk.X, padx=50)
    #
    #
    #
    # def start(self):
    #     if not hasattr(self, "worker"):
    #         self.setup_worker()
    #
    #     self.task_name_entry.configure(state="disabled")
    #     self.start_button.configure(text="Finish", command=self.finish_early)
    #     self.time_remaining_var.set("25:00")
    #     self.pause_button.configure(state="normal")
    #     self.worker.start()
