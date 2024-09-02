# ChatBot-APP
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import difflib
import tkinter as tk
from tkinter import messagebox, scrolledtext
from PIL import Image, ImageTk  

# Load the Excel files
file_path1 = 'file1.xlsx' #Add the path to your excel or csv files here
file_path2 = 'file2.xlsx'

try:
    data1 = pd.read_excel(file_path1)
    data2 = pd.read_excel(file_path2)
except Exception as e:
    print(f"Error loading Excel files: {e}")
    exit()

# Define the columns for the first file
questions = data1['Question']
answers = data1['Answer']
descriptions = data1['Document Description']
links = data1['Link']

# Function to get the best matching response
def get_best_match(query):
    vectorizer = TfidfVectorizer().fit_transform(questions.tolist() + [query])
    vectors = vectorizer.toarray()
    cosine_similarities = cosine_similarity([vectors[-1]], vectors[:-1]).flatten()
    best_match_idx = cosine_similarities.argmax()
    return questions[best_match_idx], answers[best_match_idx]

# Function to get the extension numbers based on department with flexible matching
def get_extensions_by_department(department):
    department = department.strip().lower()
    
    if 'Department' not in data2.columns:
        return "The 'Department' column is missing in the extensions Excel file."
    
    data2['Department'] = data2['Department'].str.strip().str.lower()
    exact_matches = data2[data2['Department'] == department]
    
    if exact_matches.empty:
        close_matches = difflib.get_close_matches(department, data2['Department'].unique(), n=3, cutoff=0.6)
        if close_matches:
            suggestions = ', '.join(close_matches)
            return f"No exact match found for '{department}'. Did you mean one of these?: {suggestions}"
        else:
            return f"No match found for the department: {department.title()}"
    
    result = ""
    for index, row in exact_matches.iterrows():
        result += f"\nName: {row['User Name']}\nDesignation: {row['Designation']}\nExtension: {row['Ext.']}\n" + "-"*40
    
    return result

# Function to handle asking a question
def ask_question():
    user_query = entry_question.get()
    if user_query:
        question, answer = get_best_match(user_query)
        response_text_question.delete(1.0, tk.END)
        response_text_question.insert(tk.END, f"Best matching question: {question}\n\nAnswer: {answer}")
    else:
        messagebox.showwarning("Input Required", "Please enter a question.")

# Function to handle showing links
def show_links():
    if messagebox.askyesno("Show Links", "Would you like to see all document descriptions and links?"):
        response_text_links.delete(1.0, tk.END)
        for description, link in zip(descriptions, links):
            if pd.notna(description) and pd.notna(link):
                response_text_links.insert(tk.END, f"\nDocument Description: {description}\nLink: {link}\n" + "-"*40 + "\n")

# Function to handle getting extension numbers
def get_extension_number():
    department_name = entry_department.get()
    if department_name:
        extension_info = get_extensions_by_department(department_name)
        response_text_extension.delete(1.0, tk.END)
        response_text_extension.insert(tk.END, extension_info)
    else:
        messagebox.showwarning("Input Required", "Please enter a department name.")

# Create the main application window
root = tk.Tk()
root.title("OGDCL Chatbot Interface")
root.geometry("800x700")

# Create side borders with a soft gray shade
left_border = tk.Frame(root, bg="#B0BEC5", width=35)
left_border.pack(side=tk.LEFT, fill=tk.Y)

right_border = tk.Frame(root, bg="#B0BEC5", width=35)
right_border.pack(side=tk.RIGHT, fill=tk.Y)

# Create UI elements with neutral and light shades
frame_main = tk.Frame(root, bg="#ECEFF1")
frame_main.pack(expand=True, fill=tk.BOTH, padx=(10, 10))

# Load the texture image
try:
    texture_image = Image.open("image.jpg") # add the path to your image here
    texture_image = texture_image.resize((1900, 160))  # Resize to fit the label size
    texture_photo = ImageTk.PhotoImage(texture_image)
except Exception as e:
    print(f"Error loading texture image: {e}")
    texture_photo = None

# Create a label with the texture background
label_title = tk.Label(frame_main, font=("Times New Roman", 18, "bold"),
                       fg="#37474F", image=texture_photo, compound='center')

# Pack the label with some padding
label_title.pack(pady=10)

frame_questions = tk.Frame(frame_main, bg="#CFD8DC", bd=2, relief=tk.GROOVE)
frame_questions.pack(pady=10, fill=tk.X, expand=True)
response_text_question = scrolledtext.ScrolledText(frame_questions, wrap=tk.WORD, width=120, height=5, font=("Arial", 10), bg="#FFFFFF", fg="#37474F")
response_text_question.pack(pady=5, fill=tk.X, expand=True)

label_question = tk.Label(frame_questions, text="Type your question below:", font=("Arial", 12), bg="#CFD8DC", fg="#37474F")
label_question.pack(pady=5)
entry_question = tk.Entry(frame_questions, width=70, font=("Arial", 10), bg="#FFFFFF", fg="#37474F")
entry_question.pack(pady=5)
button_ask_question = tk.Button(frame_questions, text="Get Answer", command=ask_question, bg="#90A4AE", fg="white", font=("Arial", 10))
button_ask_question.pack(pady=5)
response_text_question = scrolledtext.ScrolledText(frame_questions, wrap=tk.WORD, width=120, height=5, font=("Arial", 10), bg="#FFFFFF", fg="#37474F")
response_text_question.pack(pady=5, fill=tk.X, expand=True)

frame_links = tk.Frame(frame_main, bg="#CFD8DC", bd=2, relief=tk.GROOVE)
frame_links.pack(pady=10, fill=tk.X, expand=True)

label_links = tk.Label(frame_links, text="Click on the button below if you wish to access links!", font=("Arial", 12), bg="#CFD8DC", fg="#37474F")
label_links.pack(pady=5)
button_get_links = tk.Button(frame_links, text="Get Links", command=show_links, bg="#90A4AE", fg="white", font=("Arial", 10))
button_get_links.pack(pady=5)
response_text_links = scrolledtext.ScrolledText(frame_links, wrap=tk.WORD, width=120, height=5, font=("Arial", 10), bg="#FFFFFF", fg="#37474F")
response_text_links.pack(pady=5, fill=tk.X, expand=True)

frame_extension = tk.Frame(frame_main, bg="#CFD8DC", bd=2, relief=tk.GROOVE)
frame_extension.pack(pady=10, fill=tk.X, expand=True)

label_extension = tk.Label(frame_extension, text="Enter the Department name below:", font=("Arial", 12), bg="#CFD8DC", fg="#37474F")
label_extension.pack(pady=5)
entry_department = tk.Entry(frame_extension, width=50, font=("Arial", 10), bg="#FFFFFF", fg="#37474F")
entry_department.pack(pady=5)
button_get_extension = tk.Button(frame_extension, text="Get Extension Number", command=get_extension_number, bg="#90A4AE", fg="white", font=("Arial", 10))
button_get_extension.pack(pady=5)
response_text_extension = scrolledtext.ScrolledText(frame_extension, wrap=tk.WORD, width=120, height=5, font=("Arial", 10), bg="#FFFFFF", fg="#37474F")
response_text_extension.pack(pady=5, fill=tk.X, expand=True)

# Start the application
root.mainloop()

