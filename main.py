import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import cv2
import pandas as pd
from datetime import datetime
import face_recognition
import os
import json
from PIL import Image, ImageTk
from ttkthemes import ThemedTk

student_data_file = 'student_data.json'

# Initialize students_data as an empty list
students_data = []

# Load student data from the JSON file if it exists and is valid
if os.path.exists(student_data_file):
    try:
        with open(student_data_file, 'r') as f:
            students_data = json.load(f)

        if not isinstance(students_data, list):
            raise ValueError("Invalid data format in the JSON file")

    except (json.JSONDecodeError, ValueError) as e:
        messagebox.showwarning("Data Error", "Error loading student data. Initializing an empty "
                                             "data list.")
        students_data = []

# Create a DataFrame to store student data
students_df = pd.DataFrame(columns=['Student Name', 'Photo Path'])

# Initialize the DataFrame with data from the loaded list
for student_info in students_data:
    students_df = pd.concat([students_df, pd.DataFrame([student_info])], ignore_index=True)

# Create a DataFrame to store attendance data
attendance_df = pd.DataFrame(columns=['Date', 'Time', 'Student Name'])


# Function to add a new student's data (name and photo)
def add_student_data():
    global students_df  # Declare students_df as a global variable
    student_name = tk.simpledialog.askstring("Student Name", "Enter Student Name:")
    if student_name:
        file_path = filedialog.askopenfilename(title="Select Student Photo")
        if file_path:
            student_info = {'Student Name': student_name, 'Photo Path': file_path}
            students_df = pd.concat([students_df, pd.DataFrame([student_info])], ignore_index=True)
            students_data.append(student_info)  # Update the student data list
            save_student_data()  # Save the updated student data to the JSON file
            messagebox.showinfo("Student Data Added", f"Student data for {student_name} has been "
                                                      f"added.")


# Function to save student data to the JSON file
def save_student_data():
    with open(student_data_file, 'w') as f:
        json.dump(students_data, f)

# Create a set to keep track of recognized student names
recognized_students = set()

# Function to start the camera and perform facial recognition
def start_camera():
    # Open the camera
    cap = cv2.VideoCapture(0)

    recognized_students = set()  # Keep track of currently recognized students

    while True:
        ret, frame = cap.read()
        rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)

        face_locations = face_recognition.face_locations(rgb_frame)
        face_encodings = face_recognition.face_encodings(rgb_frame, face_locations)

        currently_recognized_students = set()

        for (top, right, bottom, left), face_encoding in zip(face_locations, face_encodings):
            for index, row in students_df.iterrows():
                student_name = row['Student Name']
                student_photo_path = row['Photo Path']

                # Load the student's photo
                student_image = face_recognition.load_image_file(student_photo_path)

                # Encode the student's face
                student_face_encoding = face_recognition.face_encodings(student_image)[0]

                # Compare the detected face with the student's face
                match = face_recognition.compare_faces([student_face_encoding], face_encoding)

                if match[0]:
                    # Get the current date and time
                    now = datetime.now()
                    current_time = now.strftime("%Y-%m-%d %H:%M:%S")

                    # Log attendance in the DataFrame if not already recognized in this frame
                    if student_name not in recognized_students:
                        attendance_df.loc[len(attendance_df)] = [current_time.split()[0],
                                                                 current_time.split()[1],
                                                                 student_name]
                        recognized_students.add(student_name)

                    # Draw a rectangle around the detected face
                    cv2.rectangle(frame, (left, top), (right, bottom), (0, 255, 0), 2)

                    # Display student name followed by "present" above the box
                    cv2.putText(frame, f"{student_name} present", (left, top - 10),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.5,
                                (0, 255, 0), 2)

                    # Add the recognized student to the set for this frame
                    currently_recognized_students.add(student_name)

        # Update the recognized students with the currently recognized students for this frame
        recognized_students = currently_recognized_students

        # Display the frame with detected faces
        cv2.imshow('Facial Recognition', frame)

        if cv2.waitKey(1) & 0xFF == 27:
            break

    cap.release()
    cv2.destroyAllWindows()


# Function to save attendance data to an Excel sheet
def save_attendance():
    try:
        attendance_df.to_excel('attendance.xlsx', index=False)
        messagebox.showinfo("Attendance Saved",
                            "Attendance data has been saved to 'attendance.xlsx'.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")


# Create the main window
root = ThemedTk(theme="xpnative")
root.title("Facial Recognition Attendance System")
root.geometry("800x400")
root.configure(bg="#E0E9FF")

# Create a title label
title_label = tk.Label(root, text="Facial Recognition Attendance System",
                       font=("Times New Roman", 20), bg="#E0E9FF")
title_label.pack(pady=20)

add_student_button = ttk.Button(root, text="Add Student Data", command=add_student_data, width=20)
add_student_button.pack(pady=20)

camera_icon = Image.open("camera.png")
camera_icon = ImageTk.PhotoImage(camera_icon)
camera_button = ttk.Button(root, text="Start Camera", image=camera_icon, command=start_camera)
camera_button.pack(pady=20)

save_button = ttk.Button(root, text="Save Attendance", command=save_attendance, width=20)
save_button.pack(pady=20)

# Start the GUI main loop
root.mainloop()
