# Keep all your imports as is
import cv2
import numpy as np
import face_recognition
import os
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import Label, Button, messagebox, ttk, Entry
from PIL import Image, ImageTk
import pyttsx3
import threading

# === CONFIGURATION === #
FACE_DIRECTORY = "C:/Users/91935/OneDrive/Desktop/Face_Attendance_System/faces_data"
LOGO_PATH = "C:/Users/91935/OneDrive/Desktop/Face_Attendance_System/image.png"
ADMIN_USERNAME = "GAURAV"
ADMIN_PASSWORD = "BCADA"

# === VOICE SETUP === #
engine = pyttsx3.init()
engine.setProperty('rate', 160)
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

def speak(text):
    threading.Thread(target=engine.say, args=(text,)).start()
    engine.runAndWait()

if not os.path.exists(FACE_DIRECTORY):
    print(" Folder NOT found! Please create 'faces_data' in the project directory.")
    exit()

def load_faces(directory=FACE_DIRECTORY):
    known_face_encodings = []
    known_face_names = []
    for file in os.listdir(directory):
        if file.endswith((".jpg", ".png")):
            img_path = os.path.join(directory, file)
            try:
                image = face_recognition.load_image_file(img_path)
                encodings = face_recognition.face_encodings(image)
                if encodings:
                    known_face_encodings.append(encodings[0])
                    known_face_names.append(os.path.splitext(file)[0])
                else:
                    print(f"\u26a0 No face detected in {file}, skipping.")
            except Exception as e:
                print(f" Error processing {file}: {e}")
    if not known_face_encodings:
        print(" No valid faces found! Add some face images in 'faces_data'.")
        exit()
    return known_face_encodings, known_face_names

known_face_encodings, known_face_names = load_faces()


def mark_attendance(name, course, subject, selected_hour):
    today_date = datetime.now().strftime("%Y-%m-%d")
    current_time = datetime.now().strftime("%H:%M:%S")
    current_hour = datetime.now().strftime("%H")

    course_file = f"C:/Users/91935/OneDrive/Desktop/Face_Attendance_System/{course}.xlsx"

    if not os.path.exists(course_file):
        with pd.ExcelWriter(course_file, engine="openpyxl") as writer:
            pd.DataFrame(columns=["Name", "Date", "Time", "Hour"]).to_excel(writer, sheet_name=subject, index=False)

    xl = pd.ExcelFile(course_file)
    df = xl.parse(subject) if subject in xl.sheet_names else pd.DataFrame(columns=["Name", "Date", "Time", "Hour"])

    already_marked = ((df["Name"] == name) & (df["Date"] == today_date) & (df["Hour"] == selected_hour)).any()

    if already_marked:
        print(f"{name} already marked for {subject} at hour {selected_hour} today.")
        speak(f"{name}, YOU ARE ALREADY MARKED PRESENT , MOVE ASIDE  MY NIGGA")
    else:
        new_entry = pd.DataFrame([[name, today_date, current_time, selected_hour]], columns=["Name", "Date", "Time", "Hour"])
        df = pd.concat([df, new_entry], ignore_index=True)

        with pd.ExcelWriter(course_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name=subject, index=False)

        print(f"Marked {name} for {subject} at hour {selected_hour} in {course}.")
        speak(f"{name}, your attendance has been marked.")


def recognize_faces(course, subject, selected_hour):
    cap = cv2.VideoCapture(0)
    exit_clicked = [False]
    session_marked = set()

    def click_exit(event, x, y, flags, param):
        if 10 <= x <= 110 and 10 <= y <= 50:
            exit_clicked[0] = True

    cv2.namedWindow("Face Recognition")
    cv2.setMouseCallback("Face Recognition", click_exit)

    while True:
        ret, frame = cap.read()
        if not ret:
            print("Camera not accessible!")
            break

        small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)
        rgb_frame = cv2.cvtColor(small_frame, cv2.COLOR_BGR2RGB)

        face_locations = face_recognition.face_locations(rgb_frame)
        face_encodings = face_recognition.face_encodings(rgb_frame, face_locations)

        for face_encoding, face_location in zip(face_encodings, face_locations):
            matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
            name = "WHO TF R YOU"
            face_distances = face_recognition.face_distance(known_face_encodings, face_encoding)
            best_match_index = np.argmin(face_distances)

            if matches[best_match_index]:
                name = known_face_names[best_match_index]

                if name not in session_marked:
                    mark_attendance(name, course, subject, selected_hour)
                    session_marked.add(name)
                else:
                    print(f"{name} already marked this session.")
                    speak(f"{name}, YOU ARE ALREADY MARKED PRESENT , MOVE ASIDE  MY NIGGA")

            top, right, bottom, left = [v * 4 for v in face_location]
            cv2.rectangle(frame, (left, top), (right, bottom), (0, 255, 0), 2)
            cv2.putText(frame, name, (left, top - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 255, 0), 2)

            enc_preview = np.round(face_encoding[:5], 2)
            enc_text = f"Enc: {enc_preview.tolist()}"
            cv2.putText(frame, enc_text, (left, bottom + 20), cv2.FONT_HERSHEY_PLAIN, 1, (255, 0, 255), 1)

        cv2.rectangle(frame, (10, 10), (110, 50), (0, 0, 255), -1)
        cv2.putText(frame, "EXIT", (30, 40), cv2.FONT_HERSHEY_SIMPLEX, 0.9, (255, 255, 255), 2)

        cv2.imshow("Face Recognition", frame)
        if cv2.waitKey(1) & 0xFF == ord('q') or exit_clicked[0]:
            break

    cap.release()
    cv2.destroyAllWindows()
    view_attendance_excel(course)


def view_attendance_excel(course):
    course_file = f"C:/Users/91935/OneDrive/Desktop/Face_Attendance_System/{course}.xlsx"
    try:
        os.startfile(course_file)
    except Exception as e:
        messagebox.showerror("Error", f"Unable to open attendance file for {course}.\n{e}")


COURSE_SUBJECTS = {
    "BCADA": ("MACHINE LEARNING", "POWER BI", "AI & IOT", "HADOOP", "MATHEMATICS"),
    "BCA (REGULAR)": ("JAVA", "C++", "DATABASE"),
    "BSC": ("PHYSICS", "CHEMISTRY", "BIOLOGY"),
    "BBA": ("MARKETING", "HR", "FINANCE"),
    "BCOM": ("ACCOUNTS", "STATISTICS", "ECONOMICS"),
    "BVOC": ("FILM_MAKING", "MOTION_GRAPHICS", "COMMUNICATIVE_ENGLISH"),
    "PSYCHOLOGY": ("INTRO TO PSYCH", "COUNSELING", "COGNITIVE SCIENCE"),
    "ECONOMICS": ("MICROECONOMICS", "MACROECONOMICS", "DEVELOPMENT STUDIES"),
    "BA": ("HISTORY", "POLITICAL SCIENCE", "SOCIOLOGY"),
    "BSC IT": ("NETWORKING", "DATA STRUCTURES", "PROGRAMMING"),
    "BMS": ("MANAGEMENT", "ENTREPRENEURSHIP", "BUSINESS ETHICS")
}

def open_main_interface():
    root = tk.Tk()
    root.title("Attendance System")
    root.geometry("600x500")
    root.configure(bg="lightblue")
    root.attributes("-alpha", 0.0)

    def fade_in():
        alpha = root.attributes("-alpha")
        if alpha < 1.0:
            alpha += 0.05
            root.attributes("-alpha", alpha)
            root.after(30, fade_in)

    fade_in()

    def check_login():
        if username_entry.get() == ADMIN_USERNAME and password_entry.get() == ADMIN_PASSWORD:
            login_frame.pack_forget()
            welcome_frame.pack(pady=20)
            speak(" welcome to advance computing department. kindly select the course and respective subject")
        else:
            messagebox.showerror("Login Failed", "Incorrect username or password")

    def start_attendance():
        selected_course = course_var.get()
        selected_subject = subject_var.get()
        selected_hour = hour_var.get()
        if not selected_course or not selected_subject or not selected_hour:
            messagebox.showwarning("Missing Info", "Please select course, subject and hour")
        else:
            root.destroy()
            recognize_faces(selected_course, selected_subject, selected_hour)

    login_frame = tk.Frame(root, bg="lightblue")
    Label(login_frame, text="Admin Login", font=("Arial", 16, "bold"), bg="lightblue").pack(pady=10)
    Label(login_frame, text="Username", bg="lightblue").pack()
    username_entry = Entry(login_frame)
    username_entry.pack(pady=5)
    Label(login_frame, text="Password", bg="lightblue").pack()
    password_entry = Entry(login_frame, show='*')
    password_entry.pack(pady=5)
    Button(login_frame, text="Login", bg="green", fg="white", command=check_login).pack(pady=20)
    login_frame.pack(pady=30)

    welcome_frame = tk.Frame(root, bg="lightblue")
    Label(welcome_frame, text="WELCOME TO ADVANCED COMPUTING DEPARTMENT ATTENDANCE SYSTEM",
          font=("Arial", 14, "bold"), bg="lightblue", wraplength=550).pack(pady=10)
    Label(welcome_frame, text="Designed by GAURAV NAIK - REG NO. 221BCADA42",
          font=("Arial", 12, "bold"), bg="lightblue").pack(pady=5)

    try:
        img = Image.open(LOGO_PATH)
        img = img.resize((150, 150), Image.LANCZOS)
        logo = ImageTk.PhotoImage(img)
        Label(welcome_frame, image=logo, bg="lightblue").pack(pady=10)
        welcome_frame.image = logo
    except Exception as e:
        print(f"Error loading logo: {e}")

    Label(welcome_frame, text="Select Course", font=("Arial", 12), bg="lightblue").pack(pady=5)
    course_var = tk.StringVar()
    course_dropdown = ttk.Combobox(welcome_frame, textvariable=course_var, font=("Arial", 12))
    course_dropdown['values'] = list(COURSE_SUBJECTS.keys())
    course_dropdown.pack(pady=5)

    Label(welcome_frame, text="Select Subject", font=("Arial", 12), bg="lightblue").pack(pady=5)
    subject_var = tk.StringVar()
    subject_dropdown = ttk.Combobox(welcome_frame, textvariable=subject_var, font=("Arial", 12))
    subject_dropdown.pack(pady=5)

    Label(welcome_frame, text="Select Hour", font=("Arial", 12), bg="lightblue").pack(pady=5)
    hour_var = tk.StringVar()
    hour_dropdown = ttk.Combobox(welcome_frame, textvariable=hour_var, font=("Arial", 12))
    hour_dropdown['values'] = [str(i).zfill(2) for i in range(1, 13)]
    hour_dropdown.pack(pady=5)

    def update_subjects(event):
        selected_course = course_var.get()
        if selected_course in COURSE_SUBJECTS:
            subject_dropdown['values'] = COURSE_SUBJECTS[selected_course]

    course_dropdown.bind("<<ComboboxSelected>>", update_subjects)

    Button(welcome_frame, text="Start Attendance", font=("Arial", 14), bg="black", fg="white", command=start_attendance).pack(pady=20)
    Button(welcome_frame, text="Exit", font=("Arial", 12), bg="red", fg="white", command=root.destroy).pack(pady=5)

    root.mainloop()

open_main_interface()













