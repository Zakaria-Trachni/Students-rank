import sqlite3
import pandas as pd
from tkinter.ttk import Treeview, Scrollbar
from customtkinter import *
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class Student:
    def __init__(self, name, cp1, cp2, final_grade, choix1, choix2, choix3, choix4, choix5, choix6, choix7, choix8, choix9):
        self.name = name
        self.cp1 = cp1
        self.cp2 = cp2
        self.final_grade = final_grade
        self.choix = [choix1, choix2, choix3, choix4, choix5, choix6, choix7, choix8, choix9]

def center_window(root, width, height):
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)

    root.geometry(f"{width}x{height}+{x}+{y}")

def main():
    root = CTk()
    root.title("Affectation de choix de filières 2024")
    center_window(root, 1250, 600)
    # root.resizable(False, False)
    set_appearance_mode("light")

    def show_piechart(frame, labels, sizes, colors, figsize=(1, 1)):
        # Create a figure for the pie chart
        fig, ax = plt.subplots(figsize=figsize)

        # Global variables for wedges and other pie chart elements
        wedges, texts, autotexts = None, None, None

        def create_pie_chart(explode=None):
            nonlocal wedges, texts, autotexts  # Access the outer scope variables
            ax.clear()  # Clear the axes for a fresh plot
            wedges, texts, autotexts = ax.pie(sizes, explode=explode, labels=labels, colors=colors,
                                            autopct='%1.1f%%', shadow=True, startangle=140,
                                            textprops={'fontsize': 5})
            ax.axis('equal')  # Ensure the pie chart is drawn as a circle
            canvas.draw()  # Redraw the canvas to update the figure

        # Function to update the chart when hovering over a slice
        def on_hover(event):
            if wedges is not None:
                for i, wedge in enumerate(wedges):
                    if wedge.contains_point([event.x, event.y]):
                        explode = [0] * len(labels)
                        explode[i] = 0.1  # Explode the hovered slice
                        create_pie_chart(explode=explode)
                        return  # Exit after updating the chart to avoid unnecessary redraws

                # If no wedge is hovered, reset to the original chart
                create_pie_chart()

        # Embed the figure in the Tkinter window using FigureCanvasTkAgg
        canvas = FigureCanvasTkAgg(fig, master=frame)
        canvas.draw()

        # Place the canvas on the Tkinter window
        canvas.get_tk_widget().pack(side="top", fill="both", expand=1)

        # Create the initial pie chart without any slice exploded
        create_pie_chart()

        # Bind the hover event
        fig.canvas.mpl_connect("motion_notify_event", on_hover)

    def store_excel_data():
        file = pd.read_excel('students grades.xlsx')

        # Connect to (or create) the SQLite database
        conn = sqlite3.connect('student_grades.db')
        cursor = conn.cursor()

        # Create a table for student grades (if it doesn't already exist)
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS student_table (
            student_id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_name TEXT,
            CP1 REAL,
            CP2 REAL,
            Note_Finale REAL,
            choix1 TEXT,
            choix2 TEXT,
            choix3 TEXT,
            choix4 TEXT,
            choix5 TEXT,
            choix6 TEXT,
            choix7 TEXT,
            choix8 TEXT,
            choix9 TEXT
        )
        ''')
        
        # Insert student grades into the table
        for _, row in file.iterrows():
            # Check if the student already exists in the database
            cursor.execute('SELECT COUNT(*) FROM student_table WHERE student_name = ?', (row['Nom complet'],))
            result = cursor.fetchone()[0]

            if result == 0:
            # If the student does not exist, insert the data
                cursor.execute('''
                INSERT INTO student_table (student_name, cp1, cp2, note_finale, choix1, choix2, choix3, choix4, choix5, choix6, choix7, choix8, choix9)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (row['Nom complet'], round(row['CP 1'],3), round(row['CP 2'],3), round(row['Note Finale'],3), row['Choix 1'], row['Choix 2'], row['Choix 3'], row['Choix 4'], row['Choix 5'], row['Choix 6'], row['Choix 7'], row['Choix 8'], row['Choix 9'])
                )
            else:
                print(f"Student {row['Nom complet']} already exists in the database.")
        # Commit the changes and close the connection
        conn.commit()
        conn.close()

    def load_students():
        conn = sqlite3.connect("student_grades.db")
        cursor = conn.cursor()

        # Modify this query to match your table structure
        cursor.execute("""
            SELECT student_name, cp1, cp2, note_finale, choix1, choix2, choix3, choix4, choix5, choix6, choix7, choix8, choix9
            FROM student_table ORDER BY note_finale DESC;
            """)

        students = []
        for row in cursor.fetchall():
            student = Student(*row)  # Unpack row values using * operator
            students.append(student)

        conn.close()
        return students

    def choosing_major(students):
        major_choosed_list = []
        major_color_list = []
        choices_val = [50, 35, 35, 30, 30, 20, 29, 30, 15]
        choices_str = ['GINF', 'IDSCC', 'SICS', 'GCIV', 'GIND', 'GELC', 'MGSI', 'ITIRC', 'GSEIR']
        colors = ["#ADD8E6", "#90EE90", "#FFB6C1", "#FFA500", "#F0E68C", "#84ADE3", "#AFEEEE", "#FFDAB9", "#FA8072"]
        # Insert data into the Treeview:
        for i, student in enumerate(students):
            # Check conditions of failing:
            if student.cp1<10 or student.cp2<10 or student.final_grade<10:
                continue
            else: # The student passed:
                for choice in student.choix:
                    if choice in choices_str:
                        index = choices_str.index(choice) # get the index of the major on the list
                        place_available = choices_val[index] # Number of students places for each major background
                        if place_available:
                            major_choosed_list.append(choice)
                            major_color_list.append(colors[index]) # Color to use for higlighting the treeview row background
                            choices_val[index]-=1
                            break
                        else:
                            continue
        return (major_choosed_list, major_color_list)

    def update_screen():
        # global treeview_list

        ''' The "update_screen" function will update the treeview and also
                    return values to configure the labels: '''

        ## Those functions are used to cotrol the treeview scrollbar:
        # Function to sync vertical scrolling across all Treeviews
        def sync_scroll(*args):
            main_treeview.yview(*args)
            for treeview in treeview_list:
                treeview.yview(*args)
        # Function to handle mouse wheel scrolling
        def on_mousewheel(event):
            # Scroll all Treeviews together
            main_treeview.yview_scroll(int(-1*(event.delta/120)), "units")
            for treeview in treeview_list:
                treeview.yview_scroll(int(-1*(event.delta/120)), "units")
            return "break"  # To prevent default behavior
        # create scrollbar
        y_scrollbar = Scrollbar(second_frame, orient="vertical")
        # Pack the vertical scrollbar on the right
        y_scrollbar.pack(side="right", fill="y")
        # Configure the scrollbar to control the vertical scrolling for all Treeviews
        y_scrollbar.config(command=sync_scroll)

        ## Fetch data from the database:
        students = load_students()

        ## Create the main Treeview widget:
        main_columns = ("num", "student_name", "cp1", "cp2", "note_finale")
        main_treeview = Treeview(second_frame, columns=main_columns, show="headings", selectmode="none", yscrollcommand=y_scrollbar.set)
        ## Create Treeview widgets:
        treeview_list = []
        columns = ("choix1", "choix2", "choix3", "choix4", "choix5", "choix6", "choix7", "choix8", "choix9")
        for i in range(len(columns)):
            # Create the Treeview widget:
            treeview = Treeview(second_frame, columns=(columns[i],), show="headings", selectmode="none", yscrollcommand=y_scrollbar.set)
            treeview_list.append(treeview)

        ## Fill all the treeview widgets and collect data to return:
        # initialze variables:
        counter=1
        sum_var=0
        min_grade=20 # the maximum grade
        failed = passable = assez_bien = bien = tres_bien = 0
        major_choosed_list, major_color_list = [], []
        # Fetch choosed major data:
        major_choosed_list, major_color_list = choosing_major(students)

        # Insert data into the Treeview:
        # for student, major_choosed, major_color in zip(students, major_choosed_list, major_color_list):
        for i, student in enumerate(students):
            # Check conditions of failing:
            if student.cp1<10 or student.cp2<10 or student.final_grade<10: 
                failed+=1
            else: # The student passed:
                # Collecting piechart data:
                note_finale = student.final_grade
                if 10 <= note_finale < 12:
                    passable+=1
                elif 12 <= note_finale < 14:
                    assez_bien+=1
                elif 14 <= note_finale < 16:
                    bien+=1
                else:
                    tres_bien+=1
                # Main treeview widget:
                tag = 'even' if counter%2==0 else 'odd'
                main_treeview.insert("", "end", values=(counter, student.name, student.cp1, student.cp2, student.final_grade), tags=(tag,))
                
                # Pick the choosen major and color:
                if i < len(major_choosed_list) and i < len(major_color_list):
                    # We used this 'if' statement because: 
                    #   len(major_color_list or major_choosed_list)==274 and len(students)==300
                    major_choosed, major_color = major_choosed_list[i], major_color_list[i]
                # Other treeview widgets:
                for j, treeview in enumerate(treeview_list):
                    if student.choix[j] == major_choosed:
                        tag = major_choosed
                        treeview.insert("", "end", values=student.choix[j], tags=(tag,))
                        # Color the major row background:
                        treeview.tag_configure(tag, background=major_color) # <----
                    else:
                        tag = 'even' if counter%2==0 else 'odd'
                        treeview.insert("", "end", values=student.choix[j], tags=(tag,))
                        # Color the major row background:
                        treeview.tag_configure(tag, background=major_color) # <----
                # Calculate sum and min grade:
                sum_var += note_finale
                if note_finale < min_grade:
                    min_grade = note_finale
                counter+=1

        # Configure row tags with background colors
        main_treeview.tag_configure('even', background='#f0f0f0')
        main_treeview.tag_configure('odd', background='#ffffff')
        # Set heading infos:
        main_treeview.heading("num", text="N°")
        main_treeview.heading("student_name", text="Nom de l'étudiant")
        main_treeview.heading("cp1", text="CP 1")
        main_treeview.heading("cp2", text="CP 2")
        main_treeview.heading("note_finale", text="Note finale")
        # Set column widths and placement:
        main_treeview.column("num", anchor="center", width=40)
        main_treeview.column("student_name", anchor="w", width=250)
        main_treeview.column("cp1", anchor="center", width=90)
        main_treeview.column("cp2", anchor="center", width=90)
        main_treeview.column("note_finale", anchor="center", width=100)
        # Bind mouse wheel event to the main Treeview
        main_treeview.bind("<MouseWheel>", on_mousewheel)
        # Pack the main treeview:
        main_treeview.pack(side="left", expand=True, fill="both")
        
        # Configure the treeview widgets:
        for i, treeview in enumerate(treeview_list):
            # Set widths and placement:
            treeview.heading(columns[i], text=f"Choix {i+1}")
            treeview.column(columns[i], anchor="center", width=70)
            # Configure row tags with background colors
            treeview.tag_configure('even', background='#f0f0f0')
            treeview.tag_configure('odd', background='#ffffff')
            # Bind mouse wheel event to each Treeview
            treeview.bind("<MouseWheel>", on_mousewheel)
            # Pack the treeview:
            treeview.pack(side="left", fill="y", expand=True)
        
        # return labels values:
        total = len(students)
        passed = total - failed
        max_grade = students[0].final_grade
        avg_grade = sum_var/counter
        return (total, passed, failed, passable, assez_bien, bien, tres_bien, max_grade, min_grade, avg_grade)

    # // frames:
    principal_frame = CTkFrame(root)
    principal_frame.pack(expand=True, fill="both", padx=10, pady=10)

    first_frame = CTkFrame(principal_frame)
    first_frame.pack(fill="x", padx=10, pady=(10,5))
    # sub frames:
    infos_frame = CTkFrame(first_frame, fg_color="transparent")
    infos_frame.pack(side="left", fill="both", padx=5, pady=5)
    passed_failed_frame = CTkFrame(infos_frame, height=100)
    students_infos_frame = CTkFrame(infos_frame, height=200)
    passed_failed_frame.pack(padx=5, pady=20)
    students_infos_frame.pack(padx=5, pady=10)
    
    piechart_frame = CTkFrame(first_frame)
    piechart_frame.pack(expand=True, fill="both", padx=5, pady=5)

    second_frame = CTkFrame(principal_frame)
    second_frame.pack(expand=True, fill="both", padx=10, pady=(5,10))
    
    
    # // labels, buttons,.. :

    ## Second_frame:
    # Store data from excel into a database:
    store_excel_data()
    # Update the treeview and the main labels:
    total, passed, failed, passable, assez_bien, bien, tres_bien, max_grade, min_grade, avg_grade = update_screen()

    ## First_frame:
    # -> passed_failed_frame:
    green_label = CTkLabel(passed_failed_frame, text="", width=50, height=30, fg_color="green", corner_radius=5)
    passed_label = CTkLabel(passed_failed_frame, text=f"{passed} Réussi", font=("Helvetica", 18))
    passed_percentage = CTkLabel(passed_failed_frame, text=f"({round(100*passed/total,2)}%)", font=("Helvetica", 18))
    red_label = CTkLabel(passed_failed_frame, text="", width=50, height=30, fg_color="red", corner_radius=5)
    failed_label = CTkLabel(passed_failed_frame, text=f"{failed} Échoué", font=("Helvetica", 18))
    failed_percentage = CTkLabel(passed_failed_frame, text=f"({round(100*failed/total,2)}%)", font=("Helvetica", 18))
    green_label.grid(row=0, column=0, padx=5, pady=5)
    passed_label.grid(row=0, column=1, padx=5, pady=5)
    passed_percentage.grid(row=0, column=2, padx=5, pady=5)
    red_label.grid(row=1, column=0, padx=5, pady=5)
    failed_label.grid(row=1, column=1, padx=5, pady=5)
    failed_percentage.grid(row=1, column=2, padx=5, pady=5)

    # -> infos_frame:
    success_rate_label = CTkLabel(students_infos_frame, text=f"• Taux de réussite: {round(100*passed/total,2)}%", font=("Helvetica", 18))
    moyenne_label = CTkLabel(students_infos_frame, text=f"• Moyenne: {round(avg_grade,2)}", font=("Helvetica", 18))
    note_max_label = CTkLabel(students_infos_frame, text=f"• Note max: {round(max_grade,2)}", font=("Helvetica", 18))
    note_min_label = CTkLabel(students_infos_frame, text=f"• Note min: {round(min_grade,2)}", font=("Helvetica", 18))
    success_rate_label.grid(row=0, column=0, padx=(30,10), pady=(10,5), sticky="w")
    moyenne_label.grid(row=0, column=1, padx=10, pady=(10,5), sticky="w")
    note_max_label.grid(row=1, column=0, padx=(30,10), pady=(5,10), sticky="w")
    note_min_label.grid(row=1, column=1, padx=10, pady=(5,10), sticky="w")

    # -> pie_chart_frame:
    labels = ['Échouer', 'Passable', 'Assez bien', 'Bien', 'Très bien']
    sizes = [failed, passable, assez_bien, bien, tres_bien]
    colors = ['gold', 'yellowgreen', 'lightcoral', 'lightskyblue', 'lightgrey']
    show_piechart(piechart_frame, labels, sizes, colors, figsize=(1.1,1.1))
    
    root.mainloop()

if __name__ == "__main__":
    main()