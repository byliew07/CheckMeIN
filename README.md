# CheckMeIN: The Swinburne GUI Attendance Tracker 📝
Welcome to **CheckMeIN**! 👋 This is a lightweight and powerful attendance-tracking system that runs entirely as a Python file.

It was built as a **Foundation in Programming project at Swinburne Sarawak** and is perfect for learning about data persistence (using *CSVs*), user roles, and modular programming.

## ✨ Features

### V1.0: The Original CLI 💻

* A robust, command-line interface for all operations.

* Add, edit, and remove students and classes.

* Take attendance directly from your terminal.

* Saves all data to .csv files.

### V2.0: The New GUI ✨

* A full-featured Graphical User Interface (GUI) for a user-friendly experience.

* All the features of V1, but now with buttons, forms, and visual feedback!

* Easily manage your database without touching the command line.

### V3.0 The New Visualize Data System 👀

* Added service layer `services.py` to handle logic.

* Added data visualization.

* Added Excel export capabilities.

* Added advanced management features for admin.


## 🏫 3 Distinct User Roles

### 🚀 The Admin (login: `admin` / `admin123`)
The all-powerful administrator who sets up the system.
* ➕ **Add New Students**: Register new students with a username and password.
* ➕ **Add New Lecturers**: Register new lecturers to the system.
* ➕ **Create New Classes**: Build the class catalog for everyone.
* 👀 **Data Overview**: View real-time lists of all users and classes.
* 🗑️ **Maintenance**: Delete specific users or classes directly from the UI.

### 👩‍🏫 The Lecturer (login: `lecturer1` / `lecturer123`)
The "*eyes*" of the operation.
* 👀 **View Class Roster**: Select any class to see a full attendance report.
* 📊 **Formatted Table**: See who's "Present" and who's "Absent" in a clean, beautiful table format.
* 💻 **Data Visualization**: Generate and view a **14-day Attendance Trend Graph** (Line Chart).
* 📤 **Export Data:** Export class attendance statistics to **Excel (.xlsx)**.

### 🎓 The Student (login: `student1` / `student123`)
The most important user!
* ✋ **Mark Attendance**: Students can log in, pick a class, and mark themselves as 'Present'.
* ✅ **Prevents Double-Marking**: The system is smart! It won't let a student mark attendance more than once.
* 📤 **Personal History**: Export personal attendance history to **Excel (.xlsx)**.

## 🚀 How to Run
Getting started is as easy as 1-2-3!
### 1. **Clone or Download**
* Download the ZIP or clone the repository:
```
git clone https://github.com/byliew07/CheckMeIN.git
```
### 2. **Navigate to the Folder**
* Make sure you have Python 3 installed.
* Run the `main_gui.py` file from your terminal:
```
python main_gui.py
```
**That's it!** The very first time you run it, the program will automatically generate three new files for you to store all your data:
* `users.csv` 🧑‍🤝‍🧑
* `classes.csv` 📚
* `attendance.csv` 📈

All data is persisted in the `data/` folder.

## 🛠️ Project Structure
This project is split into three clean modules to keep things organized:
* `main_gui.py`: **The heart of the program!** ❤️ This file handles the main login logic, loads all the interface and data at startup, and directs users to the correct menu.
* `database.py`: **The "brains" 🧠 behind the data.** This module contains all the functions for reading from and writing to the `.csv` files.
* `services.py`: **The "CPU" 🖥️ of the program.** This service layer processes data, handles plotting (`matplotlib`), and export logic (`openpyxl`/`.csv`). It is heavily used by main_gui.py.

## 💻 Tech Stack
* **A Functionable Computer** 💻
* **Python 3** 🐍
* **Python** `csv` **Module** (built-in) 🗃️
* **Python** `tkinter` **Module** (for the v2.0 GUI) 🖥️
* **Python** `openpyxl` **Module** (for Excel formatting) 📈
* **Python** `matplotlib` **Module** (for graphs) 📊
* **Python** `pandas` **Module** (for CSV-to-Excel conversion) ➡️

Run the following command to install dependencies:
```
pip install matplotlib openpyxl pandas
```
## 🤝 Contributing
This was a fun project! Feel free to fork it, improve it, or suggest new features. Pull requests are always welcome!
