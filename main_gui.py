# import tkinter so got GUI now
import tkinter
import tkinter.ttk as ttk
from tkinter import messagebox
import os
from services import AttendanceService
from datetime import datetime
import subprocess
import sys
import shutil
import time

# Helper: ensure csv -> xlsx if possible and open the file with Excel (Windows).
def open_in_excel(path):
    """
    Convert CSV->XLSX if possible, then try to open path in real Excel.
    Returns (True, path) on success, (False, message) on failure.
    """
    if not path:
        return False, "No path"

    # If CSV, try to convert to .xlsx using pandas+openpyxl (preferred)
    base, ext = os.path.splitext(path)
    try:
        if ext.lower() == ".csv":
            try:
                import pandas as _pd
                # create unique xlsx name so we don't overwrite a file Excel may have locked
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_xlsx = f"{base}_{ts}.xlsx"
                df = _pd.read_csv(path, encoding="utf-8")
                # write using openpyxl engine (requires openpyxl installed)
                df.to_excel(out_xlsx, index=False, engine="openpyxl")
                path = out_xlsx
            except Exception:
                # fallback: leave CSV as-is (Excel can open CSV too)
                pass

        # Wait briefly for file to be present and stable
        for _ in range(20):
            if os.path.exists(path):
                try:
                    # try opening immediately (Windows)
                    if sys.platform.startswith("win"):
                        try:
                            os.startfile(path)  # most reliable on Windows
                            return True, path
                        except Exception:
                            pass
                    else:
                        # macOS / Linux
                        try:
                            if sys.platform == "darwin":
                                subprocess.Popen(["open", path])
                            else:
                                subprocess.Popen(["xdg-open", path])
                            return True, path
                        except Exception:
                            pass
                except Exception:
                    pass
            time.sleep(0.05)

        # If os.startfile didn't work, try Windows 'start' with shell True and proper quoting
        try:
            if sys.platform.startswith("win"):
                subprocess.Popen(f'start "" "{path}"', shell=True)
                return True, path
        except Exception:
            pass

        # Try PowerShell Start-Process
        try:
            if sys.platform.startswith("win"):
                subprocess.Popen(['powershell', '-NoProfile', '-Command', f'Start-Process -FilePath "{path}"'])
                return True, path
        except Exception:
            pass

        return False, "Failed to open file with system commands"
    except Exception as e:
        return False, str(e)

#main GUI class
class AttendanceApp(tkinter.Tk):
    #runs once when program starts
    def __init__(self):
        #runs parent class (tk.Tk) init method also the window
        super().__init__()

        # Create data directory if it doesn't exist
        data_dir = os.path.join(os.path.dirname(__file__), 'data')
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)

        try:
            # Initialize service
            self.service = AttendanceService()
            print("Service initialized successfully")
        except Exception as e:
            messagebox.showerror("Initialization Error", f"Failed to initialize service: {str(e)}")
            self.destroy()
            return

        # sync local cached views from service
        self.sync_from_service()

        #app setup
        self.title("CheckMeIN Attendance Management System")
        self.geometry("800x600") #default window size

        #styleee
        self.style = ttk.Style(self)
        self.style.theme_use('clam') #use clam theme, maybe nicer i guess

        #store current user info
        self.current_user = None

        #create container for pages
        #we will have multiple pages in the app
        container = ttk.Frame(self)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        #dict to hold pages
        self.frames = {}

        #create/add page to frames dict
        for F in (LoginPage, AdminPage, LecturerPage, StudentPage):
            #create instance of page
            frame = F(parent=container, controller=self)
            self.frames[F] = frame
            #place frame in grid btw stack on top of each other
            frame.grid(row=0, column=0, sticky="nsew")

        #show login page first
        self.show_frame(LoginPage)

    def show_frame(self, page_name):
        frame = self.frames[page_name]
        #line that bring the frame to the front
        frame.tkraise()
        # Call on_show method if it exists
        if hasattr(frame, 'on_show'):
            frame.on_show()

    def login(self, username, password):
        # trim input
        username = (username or "").strip()
        password = (password or "").strip()

        # ensure latest data (in case admin added user)
        try:
            self.service.reload()
        except Exception:
            pass
        self.sync_from_service()

        # check if username exists
        if username in self.service.users and self.service.users[username]['password'] == password:
            self.current_user = username
            role = self.service.users[username]['role']

            #show menu based on role
            if role == 'admin':
                self.show_frame(AdminPage)
            elif role == 'lecturer':
                self.show_frame(LecturerPage)
            elif role == 'student':
                self.show_frame(StudentPage)

            messagebox.showinfo("Login Successful", f"Welcome, {username}!")
        else:
            messagebox.showerror("Login Failed", "Invalid username or password.")

    def logout(self):
        #log out current user and return to login page
        self.current_user = None
        self.show_frame(LoginPage)

    def sync_from_service(self):
        # copy important lists/dicts from service for backward compatibility with UI code
        self.users = dict(self.service.users)
        self.students = list(self.service.students)
        self.lecturers = list(self.service.lecturers)
        self.classes = list(self.service.classes)
        # attendance_today used by UI for quick checks (class -> {student: status})
        self.attendance = self.service.get_attendance_map_for_date()

#login page
class LoginPage(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)

        #give access to all methods and attributes of AttendanceApp
        self.controller = controller

        #login layout
        login_frame = ttk.Frame(self, padding="20")
        login_frame.pack(expand=True) #this for centers the frame

        #widgets
        title = ttk.Label(login_frame, text="Login", font=("Arial", 20, "bold"))
        title.pack(pady=10)

        #username box
        user_label = ttk.Label(login_frame, text="Username")
        user_label.pack(pady=5)
        self.user_entry = ttk.Entry(login_frame, width=30)
        self.user_entry.pack(pady=5, padx=20)

        #password box
        pass_label = ttk.Label(login_frame, text="Password")
        pass_label.pack(pady=5)
        self.pass_entry = ttk.Entry(login_frame, width=30, show="*")
        self.pass_entry.pack(pady=5, padx=20)

        #bind enter key to login
        self.pass_entry.bind("<Return>", self.on_login_click)

        #Login button
        login_button = ttk.Button(
            login_frame,
            text="Login",
            command=self.on_login_click
        )
        login_button.pack(pady=20, padx=10)

    def on_login_click(self, event=None): #event=None to allow both button click and enter key
        username = self.user_entry.get()
        password = self.pass_entry.get()

        if not username or not password:
            messagebox.showwarning("Input Error", "Please enter both username and password.")
            return

        #call login method from controller (AttendanceApp)
        self.controller.login(username, password)

        #clear password field after login attempt
        self.pass_entry.delete(0, 'end')


#admin page
class AdminPage(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        #layout
        main_frame = ttk.Frame(self, padding="20")
        main_frame.pack(expand=True, fill="both") #centers the frame

        title = ttk.Label(main_frame, text="Admin Menu", font=("Arial", 20, "bold"))
        title.pack(pady=10, anchor="w")

        #widget
        #use "notebook" for tabs view for different admin functions
        notebook = ttk.Notebook(main_frame)
        notebook.pack(expand=True, fill="both", pady=10)

        #tab 1: add student
        student_tab = ttk.Frame(notebook, padding="10")
        notebook.add(student_tab, text="Add Students")

        ttk.Label(student_tab, text="New Student Username").pack(pady=5, anchor="w")
        self.student_user_entry = ttk.Entry(student_tab, width=40)
        self.student_user_entry.pack(pady=5, fill="x")

        ttk.Label(student_tab, text="New Student Password").pack(pady=5, anchor="w")
        self.student_pass_entry = ttk.Entry(student_tab, width=40)
        self.student_pass_entry.pack(pady=5, fill="x")

        # New: full name input (optional)
        ttk.Label(student_tab, text="Full Name (optional)").pack(pady=5, anchor="w")
        self.student_fullname_entry = ttk.Entry(student_tab, width=40)
        self.student_fullname_entry.pack(pady=5, fill="x")

        add_student_btn = ttk.Button(student_tab, text="Add Student", command=self.add_student)
        add_student_btn.pack(pady=10)

        #tab 2: add lecturer
        lecturer_tab = ttk.Frame(notebook, padding="10")
        notebook.add(lecturer_tab, text="Add Lecturer")

        ttk.Label(lecturer_tab, text="Lecturer Username").pack(pady=5, anchor="w")
        self.lecturer_user_entry = ttk.Entry(lecturer_tab, width=40)
        self.lecturer_user_entry.pack(pady=5, fill="x")

        ttk.Label(lecturer_tab, text="Lecturer Password").pack(pady=5, anchor="w")
        self.lecturer_pass_entry = ttk.Entry(lecturer_tab, width=40, show="*")
        self.lecturer_pass_entry.pack(pady=5, fill="x")

        add_lecturer_btn = ttk.Button(lecturer_tab, text="Add Lecturer", command=self.add_lecturer)
        add_lecturer_btn.pack(pady=10)

        #tab 3: add class
        class_tab = ttk.Frame(notebook, padding="10")
        notebook.add(class_tab, text="Add Class")

        ttk.Label(class_tab, text="New Class Name (e.g., 'SWE3001: English Extra Program')").pack(pady=5, anchor="w")
        self.class_name_entry = ttk.Entry(class_tab, width=40)
        self.class_name_entry.pack(pady=5, fill="x")

        add_class_btn = ttk.Button(class_tab, text="Add Class", command=self.add_class)
        add_class_btn.pack(pady=10)

        #tab 4: logout button (button)
        logout_button = ttk.Button(main_frame, text="Logout", command=self.controller.logout)
        logout_button.pack(pady=10, anchor="e", side="bottom")

        # --- new: Current data view (users / classes) ---
        view_frame = ttk.Frame(main_frame, padding="5")
        view_frame.pack(fill="x", pady=10)

        # Users list (username : role)
        users_frame = ttk.Frame(view_frame)
        users_frame.pack(side="left", fill="both", expand=True, padx=5)
        ttk.Label(users_frame, text="Users (username (FullName) : role)").pack(anchor="w")
        self.users_listbox = tkinter.Listbox(users_frame, height=8)
        self.users_listbox.pack(fill="both", expand=True)

        # Classes list
        classes_frame = ttk.Frame(view_frame)
        classes_frame.pack(side="left", fill="both", expand=True, padx=5)
        ttk.Label(classes_frame, text="Classes").pack(anchor="w")
        self.classes_listbox = tkinter.Listbox(classes_frame, height=8)
        self.classes_listbox.pack(fill="both", expand=True)

        # Refresh button
        refresh_frame = ttk.Frame(view_frame)
        refresh_frame.pack(side="left", fill="y", padx=5)
        refresh_btn = ttk.Button(refresh_frame, text="Refresh Lists", command=self.populate_lists)
        refresh_btn.pack(pady=4)

        # New: Delete selected user button
        delete_btn = ttk.Button(refresh_frame, text="Delete Selected User", command=self.delete_selected_user)
        delete_btn.pack(pady=4)

        # New: Delete selected class button
        delete_class_btn = ttk.Button(refresh_frame, text="Delete Selected Class", command=self.delete_selected_class)
        delete_class_btn.pack(pady=4)

        # initial populate
        self.populate_lists()

    #admin functions
    def add_student(self):
        username = self.student_user_entry.get()
        password = self.student_pass_entry.get()
        fullname = self.student_fullname_entry.get().strip()

        if not username or not password:
            messagebox.showwarning("Input Error", "Username and password cannot be empty.")
            return

        success, msg = self.controller.service.add_user(username, password, "student")
        if not success:
            messagebox.showerror("Error", msg)
            return

        # if fullname provided, try to save it in service.users metadata
        if fullname:
            try:
                # service.users is expected to be a dict: set display_name
                if username in self.controller.service.users:
                    self.controller.service.users[username]['display_name'] = fullname
                    # try to persist if service provides save()
                    if hasattr(self.controller.service, 'save'):
                        try:
                            self.controller.service.save()
                        except Exception:
                            pass
            except Exception:
                pass

        # refresh controller cache
        self.controller.sync_from_service()
        # update visible lists
        self.populate_lists()
        messagebox.showinfo("Success", f"Student '{username}' added successfully.")
        self.student_user_entry.delete(0, 'end')
        self.student_pass_entry.delete(0, 'end')
        self.student_fullname_entry.delete(0, 'end')

    def add_lecturer(self):
        username = self.lecturer_user_entry.get()
        password = self.lecturer_pass_entry.get()

        if not username or not password:
            messagebox.showwarning("Input Error", "Username and password cannot be empty.")
            return

        success, msg = self.controller.service.add_user(username, password, "lecturer")
        if not success:
            messagebox.showerror("Error", msg)
            return
        self.controller.sync_from_service()
        self.populate_lists()
        messagebox.showinfo("Success", f"Lecturer '{username}' added successfully.")
        self.lecturer_user_entry.delete(0, 'end')
        self.lecturer_pass_entry.delete(0, 'end')

    def add_class(self):
        class_name = self.class_name_entry.get()

        if not class_name:
            messagebox.showwarning("Input Error", "Class name cannot be empty.")
            return

        success, msg = self.controller.service.add_class(class_name, "")
        if not success:
            messagebox.showerror("Error", msg)
            return
        self.controller.sync_from_service()
        self.populate_lists()
        messagebox.showinfo("Success", f"Class '{class_name}' added successfully.")
        self.class_name_entry.delete(0, 'end')
        #update dropdowns
        self.controller.frames[LecturerPage].update_class_list()
        self.controller.frames[StudentPage].update_class_list()

    def delete_selected_user(self):
        sel = None
        try:
            sel_index = self.users_listbox.curselection()
            if not sel_index:
                messagebox.showwarning("Select User", "Please select a user to delete.")
                return
            sel_text = self.users_listbox.get(sel_index)
            # displayed as "username (FullName) : role" or "username : role"
            username = sel_text.split(':')[0].strip()
            username = username.split()[0]  # in case "username (FullName)"
        except Exception:
            messagebox.showerror("Error", "Unable to determine selected user.")
            return

        if not messagebox.askyesno("Confirm Delete", f"Delete user '{username}'?"):
            return

        # Attempt to delete via service API if present, else modify service.users
        try:
            if hasattr(self.controller.service, 'delete_user'):
                ok, msg = self.controller.service.delete_user(username)
                if not ok:
                    messagebox.showerror("Error", msg)
                    return
            else:
                # fallback: remove from service.users dict if present
                if username in getattr(self.controller.service, 'users', {}):
                    try:
                        del self.controller.service.users[username]
                        if hasattr(self.controller.service, 'save'):
                            try:
                                self.controller.service.save()
                            except Exception:
                                pass
                    except Exception as e:
                        messagebox.showerror("Error", f"Failed to delete user: {e}")
                        return
                else:
                    messagebox.showinfo("Info", "User not found in service.")
                    return
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete user: {e}")
            return

        # refresh view
        self.controller.sync_from_service()
        self.populate_lists()
        messagebox.showinfo("Deleted", f"User '{username}' deleted.")

    def delete_selected_class(self):
        # Delete class selected in classes_listbox
        try:
            sel_index = self.classes_listbox.curselection()
            if not sel_index:
                messagebox.showwarning("Select Class", "Please select a class to delete.")
                return
            class_name = self.classes_listbox.get(sel_index).strip()
        except Exception:
            messagebox.showerror("Error", "Unable to determine selected class.")
            return

        if not messagebox.askyesno("Confirm Delete", f"Delete class '{class_name}'?"):
            return

        try:
            # prefer service API
            if hasattr(self.controller.service, 'delete_class'):
                ok, msg = self.controller.service.delete_class(class_name)
                if not ok:
                    messagebox.showerror("Error", msg)
                    return
            else:
                # fallback: remove from service.classes
                classes = getattr(self.controller.service, 'classes', None)
                if classes and class_name in classes:
                    try:
                        classes.remove(class_name)
                        if hasattr(self.controller.service, 'save'):
                            try:
                                self.controller.service.save()
                            except Exception:
                                pass
                    except Exception as e:
                        messagebox.showerror("Error", f"Failed to delete class: {e}")
                        return
                else:
                    messagebox.showinfo("Info", "Class not found in service.")
                    return
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete class: {e}")
            return

        # refresh controller and other UI lists
        self.controller.sync_from_service()
        # update lecturer & student pages' class lists if present
        try:
            self.controller.frames[LecturerPage].update_class_list()
        except Exception:
            pass
        try:
            self.controller.frames[StudentPage].update_class_list()
        except Exception:
            pass
        self.populate_lists()
        messagebox.showinfo("Deleted", f"Class '{class_name}' deleted.")

    # new helpers
    def populate_lists(self):
        """Populate the users and classes listboxes from current service data."""
        # ensure latest data
        self.controller.sync_from_service()

        # users listbox
        self.users_listbox.delete(0, 'end')
        for uname, info in self.controller.service.users.items():
            role = info.get('role', '')
            display = info.get('display_name') or ''
            if display:
                self.users_listbox.insert('end', f"{uname} ({display}) : {role}")
            else:
                self.users_listbox.insert('end', f"{uname} : {role}")

        # classes listbox
        self.classes_listbox.delete(0, 'end')
        for cname in self.controller.service.classes:
            self.classes_listbox.insert('end', cname)

    def on_show(self):
        """Called when Admin page is shown — refresh lists."""
        self.populate_lists()

#lecturer menu
class LecturerPage(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        #layout
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(expand=True, fill="both") #centers the frame

        title = ttk.Label(main_frame, text="Lecturer Menu", font=("Arial", 20, "bold"))
        title.pack(pady=10, anchor="w")

        #widget
        #top selection part
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill="x", pady=5)

        ttk.Label(top_frame, text="Select Class:").pack(side="left", padx=5)

        #dropdown menu (combobox)
        self.class_combobox = ttk.Combobox(
            top_frame,
            values=self.controller.classes,
            state="readonly",
            width=30
        )
        self.class_combobox.pack(side="left", padx=5, fill="x", expand=True)

        # instruction label so buttons are obvious
        instr = ttk.Label(top_frame, text="Select a class, then click: View Attendance | Show Trend", foreground="blue")
        instr.pack(side="left", padx=10)

        # action buttons placed in a separate actions frame (visible)
        actions_frame = ttk.Frame(main_frame)
        actions_frame.pack(fill="x", pady=8, padx=5)

        view_button = ttk.Button(actions_frame, text="View Attendance", command=self.view_attendance, width=16)
        view_button.pack(side="left", padx=6)

        # Show Trend button (clearly visible)
        self.trend_button = ttk.Button(actions_frame, text="Show Trend", command=self.show_trend, width=12)
        self.trend_button.pack(side="left", padx=6)

        # Export Excel button (moved from main_gui_1)
        self.export_btn = ttk.Button(actions_frame, text="Export Excel", command=self.export_class_excel, width=12)
        self.export_btn.pack(side="left", padx=6)

        #attendance display area
        #show in table format

        #create frame for treeview
        table_frame = ttk.Frame(main_frame)
        table_frame.pack(fill="both", expand=True, pady=10)

        #column
        columns = ("student", "status")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")

        #heading
        self.tree.heading("student", text="Student Username")
        self.tree.heading("status", text="Attendance Status")

        #scrollbar
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)

        #pack tree and scrollbar
        scrollbar.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)

        #tags for coloring rows
        self.tree.tag_configure('Present', background='lightgreen')
        self.tree.tag_configure('Absent', background='lightcoral')

        #logout button
        logout_button = ttk.Button(main_frame, text="Logout", command=self.controller.logout)
        logout_button.pack(pady=10, anchor="e", side="bottom")

    def view_attendance(self):
        #get selected class from dropdown (combobox)
        selected_class = self.class_combobox.get()
        if not selected_class:
            messagebox.showwarning("Input Error", "Please select a class.")
            return

        #clear old data from table (treeview)
        for item in self.tree.get_children():
            self.tree.delete(item)

        attendance_map = self.controller.service.get_attendance_map_for_date()
        class_records = attendance_map.get(selected_class, {})

        present_count = 0
        absent_count = 0

        if not self.controller.students:
            self.tree.insert('', 'end', values=("No students found.", ""), tags=())
            return

        # Insert per-student rows showing username and optional full name
        for student in self.controller.students:
            # display as "username (FullName)" if available
            info = self.controller.service.users.get(student, {}) if hasattr(self.controller.service, 'users') else {}
            display_name = info.get('display_name') if info else None
            display_label = f"{student} ({display_name})" if display_name else student

            if student in class_records:
                status = class_records[student]
                tag = "Present" if status == "Present" else ""
                present_count += 1
            else:
                status = "Absent"
                tag = "Absent"
                absent_count += 1
            self.tree.insert('', 'end', values=(display_label, status), tags=(tag,))

        # separator and totals
        self.tree.insert('', 'end', values=("", ""), tags=())
        self.tree.insert('', 'end', values=("Total Present:", present_count), tags=())
        self.tree.insert('', 'end', values=("Total Absent:", absent_count), tags=())
        # Also keep overall stats available; do not erase the student list.

    def export_class_excel(self):
        selected_class = self.class_combobox.get()
        if not selected_class:
            messagebox.showwarning("Input Error", "Please select a class.")
            return
        # call service export (CSV/Excel)
        try:
            if hasattr(self.controller.service, 'export_class_stats_to_excel'):
                ok, out = self.controller.service.export_class_stats_to_excel(selected_class)
            elif hasattr(self.controller.service, 'export_class_stats'):
                ok, out = self.controller.service.export_class_stats(selected_class)
            else:
                messagebox.showinfo("No Export", "Export function not available in service.")
                return
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export: {e}")
            return

        if ok:
            ok_open, info = open_in_excel(out)
            if ok_open:
                messagebox.showinfo("Exported", f"Saved and opened: {info}")
            else:
                messagebox.showinfo("Exported", f"Saved: {out}\nBut failed to open automatically: {info}")
        else:
            messagebox.showinfo("No Data", out)

    def show_trend(self):
        selected_class = self.class_combobox.get()
        if not selected_class:
            messagebox.showwarning("Input Error", "Please select a class.")
            return
        ok, out = self.controller.service.plot_attendance_trend(selected_class)
        if not ok:
            messagebox.showinfo("No Data", out)
            return
        # open image in new window
        plot_window = tkinter.Toplevel(self)
        plot_window.title(f"Attendance Trend - {selected_class}")
        try:
            img = tkinter.PhotoImage(file=out)
            label = ttk.Label(plot_window, image=img)
            label.image = img
            label.pack()
        except Exception:
            ttk.Label(plot_window, text=f"Plot saved to: {out}").pack()

    def on_show(self):
        # refresh class list and enable buttons so they are visible
        self.update_class_list()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.class_combobox.set('')
        # enable/disable action buttons depending on whether classes exist
        has = bool(self.controller.classes)
        state = "normal" if has else "disabled"
        try:
            self.trend_button.config(state=state)
            # export button state
            self.export_btn.config(state=state)
        except Exception:
            pass

    def update_class_list(self):
        #refresh class list in combobox
        # ensure controller has latest data
        self.controller.sync_from_service()
        self.class_combobox['values'] = self.controller.classes

#student menu
class StudentPage(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        #layout
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(expand=True, fill="both") #centers the frame

        self.title_label = ttk.Label(main_frame, text="Student Menu", font=("Arial", 20, "bold"))
        self.title_label.pack(pady=10)

        #widget
        ttk.Label(main_frame, text="Select Class:").pack(padx=10)

        #dropdown menu (combobox)
        self.class_combobox = ttk.Combobox(
            main_frame,
            state="readonly",
            width=40,
            values=self.controller.classes,
        )
        self.class_combobox.pack(pady=5)

        mark_button = ttk.Button(
            main_frame,
            text="Mark As Present",
            command=self.mark_attendance
        )

        mark_button.pack(pady=20, ipadx=10)

        # New: Export History button (moved from main_gui_1)
        export_hist_btn = ttk.Button(main_frame, text="Export History", command=self.export_history)
        export_hist_btn.pack(pady=5)

        #logout button
        logout_button = ttk.Button(main_frame, text="Logout", command=self.controller.logout)
        logout_button.pack(pady=10, side="bottom", anchor="e")

    def mark_attendance(self):
        student_name = self.controller.current_user
        selected_class = self.class_combobox.get()

        if not selected_class:
            messagebox.showwarning("No Class", "Please select a class.")
            return

        success, msg = self.controller.service.mark_attendance(selected_class, student_name, "Present")
        if not success:
            messagebox.showinfo("Info", msg)
            return
        # refresh controller cache
        self.controller.sync_from_service()
        messagebox.showinfo("Success", f"Attendance marked as Present for {selected_class}.")

    def update_class_list(self):
        #refresh class list in combobox
        self.class_combobox['values'] = self.controller.classes

    def export_history(self):
        if not self.controller.current_user:
            messagebox.showwarning("Not logged in", "Please login as a student to export history.")
            return
        try:
            if hasattr(self.controller.service, 'export_student_history_to_excel'):
                ok, out = self.controller.service.export_student_history_to_excel(self.controller.current_user)
            elif hasattr(self.controller.service, 'export_student_history'):
                ok, out = self.controller.service.export_student_history(self.controller.current_user)
            else:
                messagebox.showinfo("No Export", "Export function not available in service.")
                return
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export: {e}")
            return

        if ok:
            ok_open, info = open_in_excel(out)
            if ok_open:
                messagebox.showinfo("Exported", f"Saved and opened: {info}")
            else:
                messagebox.showinfo("Exported", f"Saved: {out}\nBut failed to open automatically: {info}")
        else:
            messagebox.showinfo("No Data", out)

    def on_show(self):
        # Update welcome message
        if self.controller.current_user:
            self.title_label.config(text=f"Welcome, {self.controller.current_user}!")
        # Update class list in combobox
        self.update_class_list()
        # Clear selection
        self.class_combobox.set('')

if __name__ == "__main__":
    try:
        app = AttendanceApp()
        app.mainloop()
    except Exception as e:
        print(f"Fatal error: {str(e)}")



















