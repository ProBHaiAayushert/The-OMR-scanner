import cv2
import numpy as np
import os
import csv
import matplotlib.pyplot as plt
from tkinter import Tk, Label, Button, filedialog, ttk
from tkinter import messagebox
import tkinter as tk
from PIL import Image, ImageTk

# The answers dictionary for example
the_answers = {
    '1.jpg': ['52', '10', '90', '87', '64', '65', '53', '34', '45', '64', '56', '67', '56', '67', '48', '67', '77', '41', '34', '48'],
    '2.jpg': ['52', '10', '90', '87', '64', '65', '53', '34', '45', '64', '56', '67', '56', '67', '48', '67', '77', '41', '34', '48'],
    '3.jpg': ['52'] * 20,
    '4.jpg': ['52'] * 20,
    '5.jpg': ['52'] * 20,
}

class OMRApp:
    def __init__(self, master):
        self.master = master
        master.title("OMR Scanning Application")
        master.geometry("600x600")  # Increased height for additional fields
        master.configure(bg='lightblue')

        self.label = Label(master, text="OMR Scanning Application", font=("Arial", 24), bg='lightblue')
        self.label.pack(pady=20)

        # Dropdown for selecting total strength
        self.total_strength_label = Label(master, text="Select Total Strength:", font=("Arial", 14), bg='lightblue')
        self.total_strength_label.pack(pady=10)

        self.total_strength_var = tk.IntVar()
        self.total_strength_dropdown = ttk.Combobox(master, textvariable=self.total_strength_var, 
                                                    values=[20, 25, 30, 35, 40, 45, 50], state="readonly")
        self.total_strength_dropdown.current(0)  # Default value
        self.total_strength_dropdown.pack(pady=10)

        # Input field for teacher's name
        self.teacher_name_label = Label(master, text="Enter Teacher's Name:", font=("Arial", 14), bg='lightblue')
        self.teacher_name_label.pack(pady=10)

        self.teacher_name_var = tk.StringVar()
        self.teacher_name_entry = tk.Entry(master, textvariable=self.teacher_name_var, font=("Arial", 14), width=30)
        self.teacher_name_entry.pack(pady=10)

        # School logo
        self.logo = Image.open(r'C:\Users\WORKSPACE\Desktop\ideathonBackend\images.jpg')  # Provide your school logo path here
        self.logo = self.logo.resize((500, 250), Image.LANCZOS)
        self.logo_image = ImageTk.PhotoImage(self.logo)
        self.logo_label = Label(master, image=self.logo_image, bg='lightblue')
        self.logo_label.pack(pady=10)

        self.folder_button = Button(master, text="Select OMR Sheets Folder", command=self.select_folder, 
                                    bg='yellow', width=40, height=1)
        self.folder_button.pack(pady=10)

        self.answer_key_button = Button(master, text="Select Answer Key", command=self.select_answer_key, 
                                        bg='yellow', width=40, height=1)
        self.answer_key_button.pack(pady=10)

        self.process_button = Button(master, text="Process Images", command=self.process_images, 
                                     bg='yellow', width=40, height=1)
        self.process_button.pack(pady=10)

        self.omr_folder = ""
        self.answer_key_path = ""

    def select_folder(self):
        """Select the folder containing OMR sheets."""
        self.omr_folder = filedialog.askdirectory()
        if self.omr_folder:
            messagebox.showinfo("Folder Selected", f"OMR Sheets Folder: {self.omr_folder}")

    def select_answer_key(self):
        """Select the answer key file."""
        self.answer_key_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if self.answer_key_path:
            messagebox.showinfo("Answer Key Selected", f"Answer Key: {self.answer_key_path}")

    def process_images(self):
        """Process images and calculate results."""
        total_strength = self.total_strength_var.get()
        teacher_name = self.teacher_name_var.get()  # Get teacher's name from the input field
        if not self.omr_folder or not self.answer_key_path:
            messagebox.showerror("Error", "Please select OMR Sheets folder and Answer Key file.")
            return
        # Call your processing function here
        process_images(self.omr_folder, self.answer_key_path, total_strength, teacher_name)  # Pass teacher's name

def read_answer_key(answer_key_path):
    with open(answer_key_path, mode='r') as file:
        reader = csv.DictReader(file)
        return {int(row['Question']): row['Answer'] for row in reader}

def save_results(results):
    with open('C:/Users/WORKSPACE/Desktop/ideathonBackend/omr_results.csv', 'w', newline='') as csvfile:
        fieldnames = ['Roll No', 'Total Marks']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for result in results:
            writer.writerow(result)

def save_statistics(statistics, teacher_name):
    with open('C:/Users/WORKSPACE/Desktop/ideathonBackend/Report.csv', 'w', newline='') as csvfile:
        fieldnames = [
            'Teacher Name',  # Added teacher name to the fieldnames
            'Average Marks', 
            'Highest Marks', 
            'Lowest Marks', 
            'A Grade (90-100%)', 
            'B Grade (50-90%)', 
            'C Grade (33-50%)', 
            'Failed Students (<33%)', 
            'Range',
            'Total Strength', 
            'Appeared Students', 
            'Passed Students', 
            'Failed Students', 
            'Passing Percentage',
        ]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerow({'Teacher Name': teacher_name, **statistics})  # Add this line to write teacher's name

    with open('C:/Users/WORKSPACE/Desktop/ideathonBackend/statistics.csv', 'w', newline='') as csvfile:
        fieldnames = [
            'Average Marks', 
            'Highest Marks', 
            'Lowest Marks', 
            'A Grade (90-100%)', 
            'B Grade (50-90%)', 
            'C Grade (33-50%)', 
            'Failed Students (<33%)', 
            'Range',
            'Total Strength', 
            'Appeared Students', 
            'Passed Students', 
            'Failed Students', 
            'Passing Percentage',
            'Teacher Name'  # Added teacher name to the report
        ]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        statistics['Teacher Name'] = teacher_name  # Add teacher name to statistics
        writer.writerow(statistics)

def process_images(folder_path, answer_key_path, total_strength, teacher_name):
    answer_key = read_answer_key(answer_key_path)
    results = []

    for image_file in os.listdir(folder_path):
        if image_file.endswith('.jpg'):
            img_path = os.path.join(folder_path, image_file)
            original_image = cv2.imread(img_path)

            # Process the image (blur, grayscale, etc.)
            blurred_image = cv2.GaussianBlur(original_image, (5, 5), 0)
            gray_image = cv2.cvtColor(blurred_image, cv2.COLOR_BGR2GRAY)
            _, thresh_image = cv2.threshold(gray_image, 200, 255, cv2.THRESH_BINARY_INV)
            kernel = np.ones((5, 5), np.uint8)
            morphed_image = cv2.morphologyEx(thresh_image, cv2.MORPH_CLOSE, kernel)
            contours, _ = cv2.findContours(morphed_image, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

            # Calculate total marks
            roll_number = image_file.split('.')[0]
            user_answers = the_answers.get(image_file, [])
            total_marks = sum(1 for q, answer in enumerate(user_answers, start=1) if answer == answer_key.get(q))

            results.append({'Roll No': roll_number, 'Total Marks': total_marks})

    save_results(results)
    analyze_results(results, total_strength, teacher_name)  # Call analyze with the total strength and teacher's name

def analyze_results(results, total_strength, teacher_name):
    roll_nos = [result['Roll No'] for result in results]
    scores = [result['Total Marks'] for result in results]

    average_score = np.mean(scores)
    highest_score = max(scores)
    lowest_score = min(scores)
    range_score = highest_score - lowest_score

    count_90_100 = sum(1 for score in scores if score >= 18)  # Assuming 20 is the max marks (90%)
    count_50_90 = sum(1 for score in scores if 10 <= score < 18)  # 50% to 90%
    count_33_50 = sum(1 for score in scores if 7 <= score < 10)  # 33% to 50%
    count_below_33 = sum(1 for score in scores if score < 7)  # Below 33%

    # Calculate students appeared, passed, failed, and passing percentage
    appeared_students = len(scores)
    passed_students = count_50_90 + count_33_50 + count_90_100
    failed_students = count_below_33
    passing_percentage = (passed_students / total_strength) * 100 if total_strength else 0

    statistics = {
        'Average Marks': average_score,
        'Highest Marks': highest_score,
        'Lowest Marks': lowest_score,
        'A Grade (90-100%)': count_90_100,
        'B Grade (50-90%)': count_50_90,
        'C Grade (33-50%)': count_33_50,
        'Failed Students (<33%)': count_below_33,
        'Range': range_score,
        'Total Strength': total_strength,
        'Appeared Students': appeared_students,
        'Passed Students': passed_students,
        'Failed Students': failed_students,
        'Passing Percentage': passing_percentage,
    }

    save_statistics(statistics, teacher_name)  # Save with teacher's name
    messagebox.showinfo("Processing Complete", "Results and statistics saved successfully.")

if __name__ == "__main__":
    root = Tk()
    app = OMRApp(root)
    root.mainloop()
