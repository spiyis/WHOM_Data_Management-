from tkinter import *
import tkinter as tk
from tkinter import ttk
import customtkinter as ctk
import tkinter.messagebox as messagebox
import pandas as pd

def popout_sums():
       popup = Tk()
       popup.wm_title('Sums of Services')
       popup.geometry("500x650+0+0")
       label = Label(popup, text = 'Sums of Services:', font = ('Times', 28, 'bold' ),fg = '#00264D', bg = '#DDCB93' )
       label.grid(row = 1, column = 0)

       month_label = Label(popup, text = 'Input Month:', font = ('Times', 18, 'bold' ),fg = '#00264D')
       month_label.grid(row = 2, column = 0)

       month = Entry(popup, font = ('Times', 16), fg = '#00264D', width = 10)
       month.grid(row = 2, column = 1)

       def get_value():
              e_text=month.get()
              df = pd.read_excel(r"C:\Users\User\Desktop\HopeServices.xlsx")
              diaper_sum = df.loc[df['Month']==month.get()].Diapers.sum()
              sumlabel = Label(popup, text = ('Diapers:', diaper_sum), font = ('Times', 18), fg = '#00264D')
              sumlabel.grid(column = 0, row = 3 )

              wipes_sum = df.loc[df['Month']==month.get()].Wipes.sum()
              sumlabel = Label(popup, text = ('Wipes:', wipes_sum), font = ('Times', 18), fg = '#00264D')
              sumlabel.grid(column = 0, row = 4 )

              preg_sum = df.loc[df['Month']==month.get()].Pregnancy_Test.sum()
              sumlabel = Label(popup, text = ('Pregnancy Tests:', preg_sum), font = ('Times', 18), fg = '#00264D')
              sumlabel.grid(column = 0, row = 5 )

              pregb_sum = df.loc[df['Month']==month.get()].Pregnancy_Book.sum()
              sumlabel = Label(popup, text = ('Pregnancy Books:', pregb_sum), font = ('Times', 18), fg = '#00264D')
              sumlabel.grid(column = 0, row = 6 )

              oneb_sum = df.loc[df['Month']==month.get()].First_Year_Book.sum()
              sumlabel = Label(popup, text = ('1st Year Books:', oneb_sum), font = ('Times', 18), fg = '#00264D')
              sumlabel.grid(column = 0, row = 7 )

              vit_sum = df.loc[df['Month']==month.get()].Vitamins.sum()
              sumlabel = Label(popup, text = ('Vitamins:', vit_sum), font = ('Times', 18), fg = '#00264D')
              sumlabel.grid(column = 0, row = 8 )

              mat_sum = df.loc[df['Month']==month.get()].Maternity_Clothing.sum()
              sumlabel = Label(popup, text = ('Maternity Clothing:', mat_sum), font = ('Times', 18), fg = '#00264D')
              sumlabel.grid(column = 0, row = 9 )

              lay_sum = df.loc[df['Month']==month.get()].Layettes.sum()
              sumlabel = Label(popup, text = ('Layettes:', lay_sum), font = ('Times', 18), fg = '#00264D')
              sumlabel.grid(column = 0, row = 10 )

              fo_sum = df.loc[df['Month']==month.get()].Formula.sum()
              sumlabel = Label(popup, text = ('Formula:', fo_sum), font = ('Times', 18), fg = '#00264D')
              sumlabel.grid(column = 0, row = 11 )

              car_sum = df.loc[df['Month']==month.get()].Car_Seat.sum()
              sumlabel = Label(popup, text = ('Car Seats:', car_sum), font = ('Times', 18), fg = '#00264D')
              sumlabel.grid(column = 0, row = 12 )

              stroll_sum = df.loc[df['Month']==month.get()].Stroller.sum()
              sumlabel = Label(popup, text = ('Strollers:', stroll_sum), font = ('Times', 18), fg = '#00264D')
              sumlabel.grid(column = 0, row = 13 )

              pack_sum = df.loc[df['Month']==month.get()].Pack_N_Play.sum()
              sumlabel = Label(popup, text = ('Pack N Plays:', pack_sum), font = ('Times', 18), fg = '#00264D')
              sumlabel.grid(column = 0, row = 14 )

              ccloth_sum = df.loc[df['Month']==month.get()].Childrens_Clothing.sum()
              sumlabel = Label(popup, text = ("Children's Clothing:", ccloth_sum), font = ('Times', 18), fg = '#00264D')
              sumlabel.grid(column = 0, row = 16 )

              nmg_sum = df.loc[df['Month']==month.get()].New_Mom_Gift.sum()
              sumlabel = Label(popup, text = ('New Mom Gifts:', nmg_sum), font = ('Times', 18), fg = '#00264D')
              sumlabel.grid(column = 0, row = 18 )

       enterm_butt = Button(popup, pady = 1, bd=4, font = ('Times', 12, 'bold'),text = 'Enter', command = get_value, fg = '#00264D', bg = '#D4F0FF' )
       enterm_butt.grid(row = 2, column = 3)


#===================================================================================================================   
               
class HopeServicesDB:
    def __init__(self, root):
        self.root = root
        self.root.title("HOPE Data Management System")
        self.root.geometry("1920x900+0+0") #geometry of resolution 

        #FRAMES
       
        TitleFrame = Frame(self.root, bd = 14, padx = 12, relief = GROOVE, bg = '#D4FAFA')
        TitleFrame.grid(row = 0, column = 0)

        
        MainFrame = Frame(self.root)
        MainFrame.grid(row = 1, column = 0)
        
        TopFrame = Frame(MainFrame, bd = 14, width = 1350, height = 550, padx = 4, relief = RIDGE, bg = '#00264D')
        TopFrame.grid(row = 0, column = 0)
        
        LeftFrameMain = Frame(TopFrame, bd = 10, width = 400, height = 750, relief = RIDGE, bg = '#D4FAFA')
        LeftFrameMain.grid(row = 0, column = 0)
        LeftFrame = Frame(LeftFrameMain, bd = 10, width = 450, height = 500, relief = RIDGE, bg = '#ADD8E6')
        LeftFrame.grid(row = 0, column = 0)
        LeftBottom = Frame(LeftFrameMain, bd = 10, width = 650, height = 190, relief = GROOVE, bg = '#DDCB93')
        LeftBottom.grid(row = 1, column = 0, pady = 1)
        LeftTop = Frame(LeftFrame, bd = 10, width = 650, height = 190, relief = RIDGE, bg = '#D4FAFA')
        LeftTop.grid(row = 0, column = 0)
        
        RightFrame = Frame(TopFrame, bd = 10, width = 1200, height = 550, pady = 6, relief = RIDGE, bg = '#ADD8E6')
        RightFrame.grid(row = 0, column = 1)
        RightBottom = Frame(RightFrame, bd = 10, width = 1350, height = 190, relief = RIDGE, bg = '#DDCB93' )
        RightBottom.grid(row = 1, column = 8)
        #RightTop = Frame(RightFrame, bd = 10, width = 1350, height = 190, relief = RIDGE, bg = '#D4FAFA')
        #RightTop.grid(row = 0, column = 0)

#===============================================================================================================
        #TITLE & ENTER DATA LABELS

        dataTitle = Label(TitleFrame, font = ('Times', 65, 'bold'), padx = 5, text = 'HOPE Services Data Management', fg = '#00264D', bg = '#D4F0FF')   
        dataTitle.grid(row = 0, column = 0)
        subTitle = Label(LeftTop, font = ('Times', 35, 'bold'), padx = 16, text = 'Enter Data Here:', bg = '#D4F0FF', fg = '#00264D')
        subTitle.grid(row=0, column = 0)

#===============================================================================================================
       #LABELS & ENTRIES
       
        client_label = Label(LeftFrame, font = ('Times', 18, 'bold'), text = 'Client Name:', bg = '#ADD8E6', fg = '#00264D')
        client_label.grid(row=1, column = 0)
        client_entry = Entry(LeftFrame, font = ('Times', 16), bg = '#DDCB93' )
        client_entry.grid(row=1, column = 1)

        service_month = Label(LeftFrame, font = ('Times', 18, 'bold'), text = 'Month:', bg = '#ADD8E6', fg = '#00264D')
        service_month.grid(row=2, column = 0)
        service_entry = Entry(LeftFrame, font = ('Times', 16), bg = '#DDCB93')
        service_entry.grid(row = 2, column = 1)
        
        diaper_num_label = Label(LeftFrame, font = ('Times', 18), text = 'Diaper Packs (D):', bg = '#ADD8E6', fg = '#00264D')
        diaper_num_label.grid(row = 3, column = 0)
        diaper_num_entry = Entry(LeftFrame, font = ('Times', 12))
        diaper_num_entry.grid(row = 3, column = 1)
        
        diaper_size_label = Label(LeftFrame, font = ('Times', 18), text = 'Diaper Size (DS):', bg = '#ADD8E6', fg = '#00264D')
        diaper_size_label.grid(row=4, column = 0)
        diaper_size_entry = Entry(LeftFrame, font = ('Times', 12))
        diaper_size_entry.grid(row =4, column = 1)

        wipe_num_label = Label(LeftFrame, font = ('Times', 18), text = 'Wipes (W):', bg = '#ADD8E6', fg = '#00264D')
        wipe_num_label.grid(row=5, column = 0)
        wipe_num_entry = Entry(LeftFrame, font = ('Times', 12))
        wipe_num_entry.grid(row=5, column = 1)

        preg_test_label = Label(LeftFrame, font = ('Times', 18), text = 'Pregnancy Tests (PT):', bg = '#ADD8E6', fg = '#00264D')
        preg_test_label.grid(row=6, column = 0)
        preg_test_entry = Entry(LeftFrame, font = ('Times', 12))
        preg_test_entry.grid(row=6, column = 1)

        pregbook_label = Label(LeftFrame, font = ('Times', 18), text = 'Pregnancy Book (PBook):', bg = '#ADD8E6', fg = '#00264D')
        pregbook_label.grid(row=7, column = 0)
        pregbook_entry = Entry(LeftFrame, font = ('Times', 12))
        pregbook_entry.grid(row=7, column = 1)


        yrbook_label = Label(LeftFrame, font = ('Times', 18), text = 'First-Year Book (1Book):', bg = '#ADD8E6', fg = '#00264D')
        yrbook_label.grid(row=8, column = 0)
        yrbook_entry = Entry(LeftFrame, font = ('Times', 12))
        yrbook_entry.grid(row=8, column = 1)

        vitamins_label = Label(LeftFrame, font = ('Times', 18), text = 'Vitamins (V):', bg = '#ADD8E6', fg = '#00264D')
        vitamins_label.grid(row=9, column = 0)
        vitamins_entry = Entry(LeftFrame, font = ('Times', 12))
        vitamins_entry.grid(row=9, column = 1)

        matcloth_label = Label(LeftFrame, font = ('Times', 18), text = 'Maternity Clothes (MC):', bg = '#ADD8E6', fg = '#00264D')
        matcloth_label.grid(row=10, column = 0)
        matcloth_entry = Entry(LeftFrame, font = ('Times', 12))
        matcloth_entry.grid(row=10, column = 1)

        lay_label = Label(LeftFrame, font = ('Times', 18), text = 'Layettes (L):', bg = '#ADD8E6', fg = '#00264D')
        lay_label.grid(row=11, column = 0)
        lay_entry = Entry(LeftFrame, font = ('Times', 12))
        lay_entry.grid(row=11, column = 1)

        formula_label = Label(LeftFrame, font = ('Times', 18), text = 'Formula (F):', bg = '#ADD8E6', fg = '#00264D')
        formula_label.grid(row=12, column = 0)
        formula_entry = Entry(LeftFrame, font = ('Times', 12))
        formula_entry.grid(row=12, column = 1)

        stroll_label = Label(LeftFrame, font = ('Times', 18), text = 'Stroller (S):', bg = '#ADD8E6', fg = '#00264D')
        stroll_label.grid(row=13, column = 0)
        stroll_entry = Entry(LeftFrame, font = ('Times', 12))
        stroll_entry.grid(row=13, column = 1)

        carseat_label = Label(LeftFrame, font = ('Times', 18), text = 'Car Seat (CS):', bg = '#ADD8E6', fg = '#00264D')
        carseat_label.grid(row=14, column = 0)
        carseat_entry = Entry(LeftFrame, font = ('Times', 12))
        carseat_entry.grid(row=14, column = 1)

        pack_label = Label(LeftFrame, font = ('Times', 18), text = 'Pack N Play (PNP):', bg = '#ADD8E6', fg = '#00264D')
        pack_label.grid(row=15, column = 0)
        pack_entry = Entry(LeftFrame, font = ('Times', 12))
        pack_entry.grid(row=15, column = 1)

        cloth_label = Label(LeftFrame, font = ('Times', 18), text = "Children's Clothes (CC):", bg = '#ADD8E6', fg = '#00264D')
        cloth_label.grid(row=16, column = 0)
        cloth_entry = Entry(LeftFrame, font = ('Times', 12))
        cloth_entry.grid(row=16, column = 1)

        gift_label = Label(LeftFrame, font = ('Times', 18), text = 'New Mom Gift (NMG):', bg = '#ADD8E6', fg = '#00264D')
        gift_label.grid(row=17, column = 0)
        gift_entry = Entry(LeftFrame, font = ('Times', 12))
        gift_entry.grid(row=17, column = 1)

        other_label = Label(LeftFrame, font = ('Times', 18), text = 'Other:', bg = '#ADD8E6', fg = '#00264D')
        other_label.grid(row=18, column = 0)
        other_entry = Entry(LeftFrame, font = ('Times', 12))
        other_entry.grid(row=18, column = 1)

#===============================================================================================================
        
        #BUTTON FUNCTIONS

        def update_data():
                try:
                        df = pd.read_excel(r"C:\Users\User\Desktop\HopeServices.xlsx")
                        new_data = {
                                'Client_Name' : [client_entry.get()],
                                'Month' : [service_entry.get()],
                                'Diapers' : [diaper_num_entry.get()],
                                'Diaper_Size' : [diaper_size_entry.get()],
                                'Wipes' : [wipe_num_entry.get()],
                                'Pregnancy_Test' : [preg_test_entry.get()],
                                'Pregnancy_Book' : [pregbook_entry.get()],
                                'First_Year_Book' : [yrbook_entry.get()],
                                'Vitamins' : [vitamins_entry.get()],
                                'Maternity_Clothing' : [matcloth_entry.get()],
                                'Layettes' : [lay_entry.get()],
                                'Formula' : [formula_entry.get()],
                                'Car_Seat' : [carseat_entry.get()],
                                'Stroller' : [stroll_entry.get()],
                                'Pack_N_Play' : [pack_entry.get()],
                                "Childrens_Clothing" : [cloth_entry.get()],
                                'New_Mom_Gift' : [gift_entry.get()],
                                'Other' : [other_entry.get()]
                        }
                        
                        new_df = pd.DataFrame(new_data)
                        df = pd.concat([df, new_df], ignore_index = True)
                        df.to_excel(r"C:\Users\User\Desktop\HopeServices.xlsx", index = False)
                        messagebox.showinfo('Success', 'Data Updated Successfully')
                        #Clear entries
                        reset_entries()

                        #Refresh Treeview
                        refresh_treeview()

                except Exception as e:
                      messagebox.showerror('Error', str(e))   


        def reset_entries():
                client_entry.delete(0, END)
                service_entry.delete(0, END)
                diaper_num_entry.delete(0, END)
                diaper_size_entry.delete(0, END)
                wipe_num_entry.delete(0, END)
                preg_test_entry.delete(0, END)
                pregbook_entry.delete(0, END)
                yrbook_entry.delete(0, END)
                vitamins_entry.delete(0, END)
                matcloth_entry.delete(0, END)
                lay_entry.delete(0, END)
                formula_entry.delete(0, END)
                stroll_entry.delete(0, END)
                carseat_entry.delete(0, END)
                pack_entry.delete(0, END)
                cloth_entry.delete(0, END)
                gift_entry.delete(0, END)
                other_entry.delete(0, END)

        def refresh_treeview():
                try:
                        df = pd.read_excel(r"C:\Users\User\Desktop\HopeServices.xlsx")
                        treeview.delete(*treeview.get_children())
                        for index, row in df.iterrows():
                                treeview.insert('', 'end', values = (row['Client_Name'], row['Month'], row['Diapers'], row['Diaper_Size'], row['Wipes'], row['Pregnancy_Test'], row['Pregnancy_Book'], row['First_Year_Book'], row['Vitamins'], row['Maternity_Clothing'], row['Layettes'], row['Formula'], row['Car_Seat'], row['Stroller'], row['Pack_N_Play'], row["Childrens_Clothing"], row['New_Mom_Gift'], row['Other']))
                
                except Exception as e:
                        messagebox.showerror('Error', str(e))

        def exit_program():
                result = messagebox.askquestion('Confirm Exit', 'Are you sure you want to exit?')
                if result == 'yes':
                        root.destroy()


#=====================================================================================================================        
        #BUTTONS
        
        input_butt = Button(LeftBottom, pady = 1, bd=4, font = ('Times', 26, 'bold'), width = 11, height = 1, text = 'Input', command = update_data, fg = '#00264D', bg = '#D4F0FF'  )
        input_butt.grid(row = 0, column = 0)

        #delete_butt = Button(RightBottom, pady =1, bd=4, font = ('Times', 26, 'bold'), text = 'Delete Entry', command = delete_entry)
        #delete_butt.grid(row = 0, column = 1)


        reset_butt = Button(RightBottom, pady =1, bd=4, font = ('Times', 26, 'bold'), text = 'Reset Entry', command = reset_entries, fg = '#00264D', bg = '#D4F0FF' )
        reset_butt.grid(row = 0, column = 4)

        exit_butt = Button(RightBottom, pady =1, bd=4, font = ('Times', 26, 'bold'), text = 'Exit Platform', command = exit_program, fg = '#00264D', bg = '#D4F0FF')
        exit_butt.grid(row = 0, column = 6)

        sum_butt = Button(RightBottom, pady =1, bd=4, font = ('Times', 26, 'bold'), text = 'Total Services', command = popout_sums, fg = '#00264D',  bg = '#D4F0FF')
        sum_butt.grid(row = 0, column = 5)
#===============================================================================================================
        #TREEVIEW WIDGET

        style = ttk.Style()
        style.configure('Treeview.Heading', font = ('Times', 9, 'bold'))
        style.configure('Treeview', rowheight = 40, font = ('Times', 9))

        treeview_columns = ('Name', 'Month', 'D', 'DS', 'W', 'PT', 'PBook', '1Book', 'V', 'MC', 'L', 'F', 'CS', 'S', 'PNP', 'CC', 'NMG', 'Other')
        treeview = ttk.Treeview(RightFrame, columns = treeview_columns, show = 'headings', height = 10)
        treeview.grid(row = 0, columnspan = 15, pady = 20)

        for col in treeview_columns:
            treeview.heading(col, text = col)
            treeview.column(col, width = 65)
            treeview.column(col, anchor = 'center')

#Get data from existing excel sheet
        try:
        
                df = pd.read_excel(r"C:\Users\User\Desktop\HopeServices.xlsx")
                for index, row in df.iterrows():
                        treeview.insert('','end', values = (row['Client_Name'], row['Month'], row['Diapers'], row['Diaper_Size'], row['Wipes'], row['Pregnancy_Test'], row['Pregnancy_Book'], row['First_Year_Book'], row['Vitamins'], row['Maternity_Clothing'], row['Layettes'], row['Formula'], row['Car_Seat'], row['Stroller'], row['Pack_N_Play'], row["Childrens_Clothing"], row['New_Mom_Gift'], row['Other']))

        except Exception as e:
             messagebox.showerror('Error', str(e))

#=================================================================================================================

        def on_treeview_select(event):
                selected_item = treeview.focus()
                if selected_item:
                        values = treeview.item(selected_item, 'values')
                        client_entry.delete(0, tk.END)
                        service_entry.delete(0, tk.END)
                        diaper_num_entry.delete(0, tk.END)
                        diaper_size_entry.delete(0, tk.END)
                        wipe_num_entry.delete(0, tk.END)
                        preg_test_entry.delete(0, tk.END)
                        pregbook_entry.delete(0, tk.END)
                        yrbook_entry.delete(0, tk.END)
                        vitamins_entry.delete(0, tk.END)
                        matcloth_entry.delete(0, tk.END)
                        lay_entry.delete(0, tk.END)
                        formula_entry.delete(0, tk.END)
                        stroll_entry.delete(0, tk.END)
                        carseat_entry.delete(0, tk.END)
                        pack_entry.delete(0, tk.END)
                        cloth_entry.delete(0, tk.END)
                        gift_entry.delete(0, tk.END)
                        other_entry.delete(0, tk.END)

                        client_entry.insert(0, values[0])
                        service_entry.insert(0, values[1])
                        diaper_num_entry.insert(0, values[2])
                        diaper_size_entry.insert(0, values[3])
                        wipe_num_entry.insert(0, values[4])
                        preg_test_entry.insert(0, values[5])
                        pregbook_entry.insert(0, values[6])
                        yrbook_entry.insert(0, values[7])
                        vitamins_entry.insert(0, values[8])
                        matcloth_entry.insert(0, values[9])
                        lay_entry.insert(0, values[10])
                        formula_entry.insert(0, values[11])
                        stroll_entry.insert(0, values[12])
                        carseat_entry.insert(0, values[13])
                        pack_entry.insert(0, values[14])
                        cloth_entry.insert(0, values[15])
                        gift_entry.insert(0, values[16])
                        other_entry.insert(0, values[17])
                
        treeview.bind('<<TreeviewSelect>>', on_treeview_select)
        
        #def delete_entry():
                #try:
                        #df = pd.read_excel(r"C:\Users\mpiyis\Desktop\HopeServices.xlsx")
                        #new_data = {
                                #'Client Name' : [None],
                                #'Diapers' : [None],
                                #'Diaper Size' : [None],
                                #'Wipes' : [None],
                                #'Pregnancy Test' : [None],
                                #'Pregnancy Book' : [None],
                                #'1st Year Book' : [None],
                                #'Vitamins' : [None],
                                #'Maternity Clothing' : [None],
                                #'Layettes' : [None],
                                #'Formula' : [None],
                                #'Car Seat' : [None],
                                #'Stroller' : [None],
                                #'Pack N Play' : [None],
                                #"Children's Clothing" : [None],
                                #'New Mom Gift' : [None],
                                #'Other' : [None]
                                #}
                        #new_df = pd.DataFrame(new_data)
                        #new_df = new_df[new_df != None]
                        #new_df = new_df.dropna(subset= ['Client Name','Diapers','Diaper Size','Wipes','Pregnancy Test','Pregnancy Book','1st Year Book','Vitamins','Maternity Clothing','Layettes','Formula','Car Seat','Stroller','Pack N Play',"Children's Clothing",'New Mom Gift','Other'])
                        #df = pd.concat([df, new_df], ignore_index = True)
                        #new_df.to_excel(r"C:\Users\mpiyis\Desktop\HopeServices.xlsx", index = False)
                        #selected_item = treeview.focus()
                        #treeview.delete(selected_item)
                        #messagebox.showinfo('Success', 'Data Deleted Successfully')

                #except Exception as e:
                        #messagebox.showerror('Error', str(e)) 
        
        
        #delete_butt = Button(RightBottom, pady =1, bd=4, font = ('Times', 26, 'bold'), text = 'Delete Entry', command = delete_entry)
        #delete_butt.grid(row = 0, column = 1)

        

if __name__ == '__main__':
    root = ctk.CTk()
    application = HopeServicesDB(root)
    root.mainloop()

