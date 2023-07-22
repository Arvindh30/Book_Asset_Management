from tkinter import *
import tkinter as tk
from tkinter import ttk
import openpyxl

def load_data():
    path="D:\Projects\Book Asset Management\DataEntry.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet=workbook.active

    list_values=list(sheet.values)
    for col_name in list_values[0]:
        XlView.heading(col_name,text=col_name)
    for value_tuple in list_values[1:]:
        XlView.insert('',tk.END,values=value_tuple)

root = tk.Tk()
root.title("Data Entry")
root.geometry("1800x1080")
p1 = PhotoImage(file='d:\Projects\Book Asset Management\icon.png')
root.iconphoto(False,p1)

Style=ttk.Style(root)
root.tk.call("source", r"d:\Projects\Book Asset Management\forest-light.tcl")
Style.theme_use("forest-light")


Frame = Frame(root,pady=20,padx=10)
Frame.pack()

Header = ttk.Label(Frame, text="                                         Data Entry",font=('Times New Roman bold',24))
Header.grid(row=0,column=1,sticky="ew")
DetailsFrame = ttk.LabelFrame(Frame,text="Details")
DetailsFrame.grid(row=1,column=0,padx=20,pady=20,sticky="news")

"""BookId = ttk.Label(DetailsFrame, text="ID : ")
BookId.grid(row=1,padx=5,pady=20,sticky="ew")
BookIdField = ttk.Entry(DetailsFrame)
BookIdField.grid(row=1, column=1,padx=5,pady=20,sticky="ew")"""

Book = ttk.Label(DetailsFrame, text="Name : ")
Book.grid(row=1,column=0,padx=5,pady=20,sticky="ew")
BookField = ttk.Entry(DetailsFrame)
BookField.grid(row=1, column=1,padx=5,pady=20,sticky="ew")

def InsertRow():
    if(URField.get()=="" or PriceField.get()=="" or YearField.get()=="" or BookField.get()=="" or AuthorField.get()=="" or ReviewsField.get()=="" or GenreField.get()==""):
        ErrorLabel=Label(Frame,text="Please Enter All Required Field ***")
        ErrorLabel.grid(row=2,column=1)
        

    #BookID = int(BookIdField.get())
    BookName = str(BookField.get())
    AuthorName = AuthorField.get()
    UserRatingVal = float(URField.get())
    Review = int(ReviewsField.get())
    PrinceVal = float(PriceField.get())
    GenreVal = GenreField.get()
    YearVal = int(YearField.get())
    #GenderVal = r.get()
    
    path="D:\Projects\Book Asset Management\DataEntry.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet=workbook.active
    row_values=[BookName,AuthorName,UserRatingVal,Review,PrinceVal,YearVal,GenreVal]
    sheet.append(row_values)
    workbook.save(path)

    XlView.insert("",tk.END,values=row_values)

    #BookIdField.delete(0,"")
    BookField.delete(0,"")
    AuthorField.delete(0,"")
    PriceField.delete(0,"")
    YearField.delete(0,"")


Author = ttk.Label(DetailsFrame, text="Author : ")
Author.grid(row=2,column=0,padx=5,pady=20,sticky="ew")
AuthorField = ttk.Entry(DetailsFrame)
AuthorField.grid(row=2,column=1,padx=5,pady=20,sticky="ew")

UserRating = ttk.Label(DetailsFrame, text="User Rating : ")
UserRating.grid(row=3,column=0,padx=5,pady=20,sticky="ew")
URField = ttk.Spinbox(DetailsFrame,from_=0.0,to=5.0)
URField.grid(row=3, column=1,padx=5,pady=20,sticky="ew")

Reviews = ttk.Label(DetailsFrame, text="Reviews : ")
Reviews.grid(row=4,column=0,padx=2,pady=20,sticky="ew")
ReviewsField = ttk.Spinbox(DetailsFrame,from_=35,to=1000000)
ReviewsField.grid(row=4, column=1,padx=5,pady=20,sticky="ew")

Price = ttk.Label(DetailsFrame, text="Price : ")
Price.grid(row=5,column=0,padx=5,pady=20,sticky="ew")
PriceField = ttk.Spinbox(DetailsFrame,from_=0,to=110)
PriceField.grid(row=5,column=1,padx=5,pady=20,sticky="ew")

Genre = ttk.Label(DetailsFrame, text="Genre : ")
Genre.grid(row=7,column=0,padx=2,pady=20,sticky="ew")
GenreField = ttk.Combobox(DetailsFrame,values=["Fiction","Non Fiction"])
GenreField.current(0)
GenreField.grid(row=7,column=1,padx=5,pady=20,sticky="ew")

Year = ttk.Label(DetailsFrame, text="Year : ")
Year.grid(row=6,column=0,padx=5,pady=20,sticky="ew")
YearField = ttk.Spinbox(DetailsFrame,from_=2009,to=2019)
YearField.grid(row=6, column=1,padx=5,pady=20,sticky="ew")

"""Gender = ttk.Label(DetailsFrame, text="Gender : ")
Gender.grid(row=5,column=0,padx=5,pady=20,sticky="ew")
r=StringVar()
ttk.Radiobutton(DetailsFrame,text="Male",variable=r,value="Male").grid(row=5,column=1,pady=20,sticky="ew")
ttk.Radiobutton(DetailsFrame,text="Female",variable=r,value="Female").grid(row=5,column=2,pady=20,sticky="ew")
myLabel=Label(root,text=r.get())"""

Insert = ttk.Button(Frame, text="Insert", command=InsertRow,width=20)
Insert.grid(row=2,column=0,padx=50,sticky="nsew")

XlFrame = ttk.Frame(Frame)
XlFrame.grid(row=1, column=1, pady=20)

XlScroll = ttk.Scrollbar(XlFrame)
XlScroll.pack(side="right",fill="y")

cols = ("Name", "Author","User Rating","Reviews","Price","Year","Genre")
XlView = ttk.Treeview(XlFrame, show="headings", column=cols, height=25)
#XlView.column("# 0",anchor=CENTER,stretch=NO,width=1000)
XlView.column("# 1",anchor=CENTER,stretch=NO,width=520)
XlView.column("# 2",anchor=CENTER,stretch=NO,width=230)
XlView.column("# 3",anchor=CENTER,stretch=NO,width=100)
XlView.column("# 4",anchor=CENTER,stretch=NO,width=50)
XlView.column("# 5",anchor=CENTER,stretch=NO,width=50)
XlView.column("# 6",anchor=CENTER,stretch=NO,width=50)
XlView.column("# 7",anchor=CENTER,width=100)

XlView.pack(side="left", fill="both")
XlScroll.config(command=XlView.yview)

load_data()

root.mainloop()