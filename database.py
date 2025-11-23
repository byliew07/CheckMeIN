import csv
import os
from datetime import datetime

BASE_DIR = os.path.join(os.path.dirname(__file__), "data")
USERS_CSV = os.path.join(BASE_DIR, "users.csv")
CLASSES_CSV = os.path.join(BASE_DIR, "classes.csv")
ATTENDANCE_CSV = os.path.join(BASE_DIR, "attendance.csv")

def ensure_data_dir():
    if not os.path.exists(BASE_DIR):
        os.makedirs(BASE_DIR)
    # create files with headers if missing
    if not os.path.exists(USERS_CSV):
        with open(USERS_CSV, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["username", "password", "role"])
            writer.writerow(["admin", "admin123", "admin"])
    if not os.path.exists(CLASSES_CSV):
        with open(CLASSES_CSV, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["class_name", "lecturer_username"])
    if not os.path.exists(ATTENDANCE_CSV):
        with open(ATTENDANCE_CSV, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["date", "class_name", "student_username", "status", "time_in"])

def load_users():
    ensure_data_dir()
    users = {}
    students = []
    lecturers = []
    with open(USERS_CSV, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            users[row["username"]] = {"password": row["password"], "role": row["role"]}
            if row["role"] == "student":
                students.append(row["username"])
            elif row["role"] == "lecturer":
                lecturers.append(row["username"])
    return users, students, lecturers

def load_classes():
    ensure_data_dir()
    classes = {}  # class_name -> lecturer_username (may be empty)
    with open(CLASSES_CSV, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            classes[row["class_name"]] = row.get("lecturer_username", "")
    return classes

def load_attendance_records():
    ensure_data_dir()
    records = []
    with open(ATTENDANCE_CSV, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            records.append(row)
    return records

def append_user(username, password, role):
    ensure_data_dir()
    with open(USERS_CSV, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([username, password, role])

def append_class(class_name, lecturer_username=""):
    ensure_data_dir()
    with open(CLASSES_CSV, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([class_name, lecturer_username])

def append_attendance(date, class_name, student_username, status, time_in):
    ensure_data_dir()
    with open(ATTENDANCE_CSV, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([date, class_name, student_username, status, time_in])

def update_attendance_record(date, class_name, student_username, new_status):
    ensure_data_dir()
    rows = []
    changed = False
    with open(ATTENDANCE_CSV, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row["date"] == date and row["class_name"] == class_name and row["student_username"] == student_username:
                row["status"] = new_status
                changed = True
            rows.append(row)
    if changed:
        with open(ATTENDANCE_CSV, "w", newline="", encoding="utf-8") as f:
            fieldnames = ["date", "class_name", "student_username", "status", "time_in"]
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)
    return changed

