import pandas as pd
from tkinter import filedialog
import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3

# ------------------ COLOR THEME ------------------
BG_COLOR = "#f1f5f9"      # light background
CARD_COLOR = "#ffffff"
ACCENT_COLOR = "#2563eb"
TEXT_COLOR = "#111827"
SUCCESS_COLOR = "#16a34a"

# ------------------ DATABASE ------------------

def create_table():
    conn = sqlite3.connect("institution.db")
    cursor = conn.cursor()

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS students(
        id INTEGER PRIMARY KEY,
        name TEXT,
        age INTEGER,
        course TEXT
    )
    """)

    conn.commit()
    conn.close()


# ------------------ ADD STUDENT ------------------

def add_student():

    student_id = entry_id.get()
    name = entry_name.get()
    age = entry_age.get()
    course = entry_course.get()

    if student_id == "" or name == "" or age == "" or course == "":
        messagebox.showerror("Error","All fields required")
        return

    conn = sqlite3.connect("institution.db")
    cursor = conn.cursor()

    try:
        cursor.execute(
            "INSERT INTO students VALUES(?,?,?,?)",
            (student_id,name,age,course)
        )
        conn.commit()
    except:
        messagebox.showerror("Error","ID already exists")

    conn.close()

    clear_fields()
    show_students()


# ------------------ SHOW STUDENTS ------------------

def show_students():

    conn = sqlite3.connect("institution.db")
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM students")
    rows = cursor.fetchall()

    student_table.delete(*student_table.get_children())

    for row in rows:
        student_table.insert("",tk.END,values=row)

    conn.close()
    update_count()


# ------------------ SELECT ------------------

def select_student(event):

    selected = student_table.focus()
    values = student_table.item(selected,'values')

    if values:

        entry_id.delete(0,tk.END)
        entry_name.delete(0,tk.END)
        entry_age.delete(0,tk.END)
        entry_course.delete(0,tk.END)

        entry_id.insert(0,values[0])
        entry_name.insert(0,values[1])
        entry_age.insert(0,values[2])
        entry_course.insert(0,values[3])


# ------------------ UPDATE ------------------

def update_student():

    student_id = entry_id.get()

    conn = sqlite3.connect("institution.db")
    cursor = conn.cursor()

    cursor.execute("""
    UPDATE students
    SET name=?, age=?, course=?
    WHERE id=?
    """,(entry_name.get(),
         entry_age.get(),
         entry_course.get(),
         student_id))

    conn.commit()
    conn.close()

    show_students()


# ------------------ DELETE ------------------

def delete_student():

    student_id = entry_id.get()

    confirm = messagebox.askyesno("Confirm","Delete student?")

    if confirm:

        conn = sqlite3.connect("institution.db")
        cursor = conn.cursor()

        cursor.execute("DELETE FROM students WHERE id=?",(student_id,))

        conn.commit()
        conn.close()

        clear_fields()
        show_students()


# ------------------ CLEAR ------------------

def clear_fields():

    entry_id.delete(0,tk.END)
    entry_name.delete(0,tk.END)
    entry_age.delete(0,tk.END)
    entry_course.delete(0,tk.END)


# ------------------ SEARCH ------------------

def search_student():

    keyword = search_entry.get()

    conn = sqlite3.connect("institution.db")
    cursor = conn.cursor()

    query="""
    SELECT * FROM students
    WHERE id LIKE ? OR name LIKE ? OR course LIKE ?
    """

    cursor.execute(query,('%'+keyword+'%',
                          '%'+keyword+'%',
                          '%'+keyword+'%'))

    rows=cursor.fetchall()

    student_table.delete(*student_table.get_children())

    for row in rows:
        student_table.insert("",tk.END,values=row)

    conn.close()


# ------------------ STUDENT COUNT ------------------

def update_count():

    conn = sqlite3.connect("institution.db")
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*) FROM students")
    count = cursor.fetchone()[0]

    student_count_label.config(text=f"Total Students : {count}")

    conn.close()


# ------------------ SORT ------------------

def sort_students(column,reverse=False):

    conn=sqlite3.connect("institution.db")
    cursor=conn.cursor()

    query=f"SELECT * FROM students ORDER BY {column} {'DESC' if reverse else 'ASC'}"
    cursor.execute(query)

    rows=cursor.fetchall()

    student_table.delete(*student_table.get_children())

    for row in rows:
        student_table.insert("",tk.END,values=row)

    conn.close()


def apply_sort(event):

    mapping={
        "ID Ascending":("id",False),
        "ID Descending":("id",True),
        "Name Ascending":("name",False),
        "Name Descending":("name",True),
        "Age Ascending":("age",False),
        "Age Descending":("age",True),
        "Course Ascending":("course",False),
        "Course Descending":("course",True)
    }

    choice=sort_dropdown.get()
    column,reverse=mapping.get(choice)

    sort_students(column,reverse)


# ------------------ EXPORT TO EXCEL ------------------

def export_to_excel():

    conn = sqlite3.connect("institution.db")

    query = "SELECT * FROM students"
    df = pd.read_sql_query(query, conn)

    conn.close()

    if df.empty:
        messagebox.showwarning("No Data","No students to export")
        return

    file = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel File","*.xlsx")]
    )

    if file:
        df.to_excel(file, index=False)
        messagebox.showinfo("Success","Data exported to Excel successfully!")


root=tk.Tk()
root.title("Student Management Dashboard")
root.geometry("900x600")
root.configure(bg=BG_COLOR)

# ------------------ TABLE STYLE ------------------

style = ttk.Style()

style.theme_use("default")

style.configure(
    "Treeview",
    background="white",
    foreground="black",
    rowheight=30,
    fieldbackground="white",
    bordercolor="black",
    borderwidth=2
)

style.configure(
    "Treeview.Heading",
    font=("Segoe UI", 11, "bold"),
    anchor="center"
)

style.map(
    "Treeview",
    background=[("selected", "#3b82f6")]
)
# Header
header=tk.Label(root,
                text="Student Management Dashboard",
                font=("Segoe UI",20,"bold"),
                bg=ACCENT_COLOR,
                fg="white",
                pady=10)

header.pack(fill="x")

# Student count
student_count_label=tk.Label(root,
                             text="Total Students : 0",
                             font=("Segoe UI",11,"bold"),
                             bg=BG_COLOR,
                             fg=SUCCESS_COLOR)

student_count_label.pack(pady=5)

# Form
form_frame=tk.Frame(root,bg=CARD_COLOR,padx=20,pady=20)
form_frame.pack(pady=10)

tk.Label(form_frame,text="ID",bg=CARD_COLOR,fg=TEXT_COLOR).grid(row=0,column=0,pady=5)
entry_id=tk.Entry(form_frame,width=25)
entry_id.grid(row=0,column=1)

tk.Label(form_frame,text="NAME",bg=CARD_COLOR,fg=TEXT_COLOR).grid(row=1,column=0,pady=5)
entry_name=tk.Entry(form_frame,width=25)
entry_name.grid(row=1,column=1)

tk.Label(form_frame,text="AGE",bg=CARD_COLOR,fg=TEXT_COLOR).grid(row=2,column=0,pady=5)
entry_age=tk.Entry(form_frame,width=25)
entry_age.grid(row=2,column=1)

tk.Label(form_frame,text="COURSE",bg=CARD_COLOR,fg=TEXT_COLOR).grid(row=3,column=0,pady=5)
entry_course=tk.Entry(form_frame,width=25)
entry_course.grid(row=3,column=1)

# Buttons
btn_frame=tk.Frame(form_frame,bg=CARD_COLOR)
btn_frame.grid(row=4,columnspan=2,pady=10)

btn_style={"bg":ACCENT_COLOR,"fg":"white","width":10,"bd":0}

tk.Button(btn_frame,text="Add",command=add_student,**btn_style).grid(row=0,column=0,padx=5)
tk.Button(btn_frame,text="Update",command=update_student,**btn_style).grid(row=0,column=1,padx=5)
tk.Button(btn_frame,text="Delete",command=delete_student,**btn_style).grid(row=0,column=2,padx=5)
tk.Button(btn_frame,text="Clear",command=clear_fields,**btn_style).grid(row=0,column=3,padx=5)
tk.Button(btn_frame,text="Export Excel",command=export_to_excel,**btn_style).grid(row=0,column=4,padx=5)
# ------------------ KEYBOARD NAVIGATION ------------------

entry_id.bind("<Return>", lambda e: entry_name.focus())
entry_name.bind("<Return>", lambda e: entry_age.focus())
entry_age.bind("<Return>", lambda e: entry_course.focus())
entry_course.bind("<Return>", lambda e: add_student())
# Search
search_frame=tk.Frame(root,bg=BG_COLOR)
search_frame.pack()

search_entry=tk.Entry(search_frame,width=25)
search_entry.grid(row=0,column=0,padx=5)

tk.Button(search_frame,text="Search",command=search_student).grid(row=0,column=1,padx=5)
tk.Button(search_frame,text="Show All",command=show_students).grid(row=0,column=2,padx=5)

# Sort
sort_options=[
"ID Ascending","ID Descending",
"Name Ascending","Name Descending",
"Age Ascending","Age Descending",
"Course Ascending","Course Descending"
]

sort_dropdown=ttk.Combobox(root,values=sort_options,state="readonly",width=20)
sort_dropdown.pack(pady=10)
sort_dropdown.set("Sort Students")
sort_dropdown.bind("<<ComboboxSelected>>",apply_sort)

# Table
columns=("ID","NAME","AGE","COURSE")

student_table=ttk.Treeview(root,columns=columns,show="headings")

for col in columns:
    student_table.heading(col,text=col)
    student_table.column(col,width=200)

student_table.pack(fill="both",expand=True,padx=20,pady=20)

student_table.bind("<ButtonRelease-1>",select_student)

create_table()
show_students()

root.mainloop()
