import openpyxl
from openpyxl import load_workbook
import customtkinter as ctk
from tkinter import messagebox

# Initialize the main application window
ctk.set_appearance_mode("System")  # Modes: "System" (default), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (default), "green", "dark-blue"

app = ctk.CTk()
app.title("Employee Payroll Calculator")
app.geometry("960x740")  # Set window size to 960x740

# Function to read data from Excel and calculate results
def calculate_payroll():
    try:
        # Load the workbook and select the active sheet
        book = load_workbook('data.xlsx')
        sheet = book.active

        # Define the fields and their corresponding cell addresses
        fields = [
            ("Employee Name", "A2"),
            ("Employee ID", "B2"),
            ("Department", "C2"),
            ("Designation", "D2"),
            ("Date of Joining", "E2"),
            ("Date of Birth", "F2"),
            ("UAN", "G2"),
            ("PF No", "H2"),
            ("ESI No", "I2"),
            ("Basic Salary", "J2"),
            ("Conveyance", "K2"),
            ("Special Allowance", "L2"),
            ("PF Deduction", "M2"),
            ("ESI Deduction", "N2"),
            ("PT Deduction", "O2"),
        ]

        # Read the values from the Excel sheet
        data = {}
        for field, cell in fields:
            data[field] = sheet[cell].value

        # Calculate total earnings, deductions, and net pay
        total_earnings = data["Basic Salary"] + data["Conveyance"] + data["Special Allowance"]
        total_deductions = data["PF Deduction"] + data["ESI Deduction"] + data["PT Deduction"]
        net_pay = total_earnings - total_deductions

        # Clear previous results
        for widget in result_frame.winfo_children():
            widget.destroy()

        # Display the results in three columns
        headers = ["Details", "Earnings", "Deductions"]
        for col, header in enumerate(headers):
            ctk.CTkLabel(
                result_frame, 
                text=header, 
                font=("Helvetica", 16, "bold"),  # Bold font for headers
                text_color="black"  # Black color for headers
            ).grid(row=0, column=col, padx=20, pady=10)

        # Column 1: Employee Details
        details = [
            ("Employee Name", data["Employee Name"]),
            ("Employee ID", data["Employee ID"]),
            ("Department", data["Department"]),
            ("Designation", data["Designation"]),
            ("Date of Joining", data["Date of Joining"]),
            ("Date of Birth", data["Date of Birth"]),
            ("UAN", data["UAN"]),
            ("PF No", data["PF No"]),
            ("ESI No", data["ESI No"]),
        ]

        for row_idx, (field, value) in enumerate(details, start=1):
            ctk.CTkLabel(
                result_frame, 
                text=field, 
                font=("Helvetica", 14),  # Regular font for fields
                text_color="black"  # Black color for fields
            ).grid(row=row_idx, column=0, padx=20, pady=5, sticky="w")
            ctk.CTkLabel(
                result_frame, 
                text=value, 
                font=("Helvetica", 14),  # Regular font for values
                text_color="black"  # Black color for values
            ).grid(row=row_idx, column=0, padx=20, pady=5, sticky="e")

        # Column 2: Earnings
        earnings = [
            ("Basic Salary", f"₹{data['Basic Salary']}"),
            ("Conveyance", f"₹{data['Conveyance']}"),
            ("Special Allowance", f"₹{data['Special Allowance']}"),
            ("Total Earnings", f"₹{total_earnings}"),
        ]

        for row_idx, (field, value) in enumerate(earnings, start=1):
            ctk.CTkLabel(
                result_frame, 
                text=field, 
                font=("Helvetica", 14),  # Regular font for fields
                text_color="black"  # Black color for fields
            ).grid(row=row_idx, column=1, padx=20, pady=5, sticky="w")
            ctk.CTkLabel(
                result_frame, 
                text=value, 
                font=("Helvetica", 14),  # Regular font for values
                text_color="black"  # Black color for values
            ).grid(row=row_idx, column=1, padx=20, pady=5, sticky="e")

        # Column 3: Deductions
        deductions = [
            ("PF Deduction", f"₹{data['PF Deduction']}"),
            ("ESI Deduction", f"₹{data['ESI Deduction']}"),
            ("PT Deduction", f"₹{data['PT Deduction']}"),
            ("Total Deductions", f"₹{total_deductions}"),
            ("Net Pay", f"₹{net_pay}"),
        ]

        for row_idx, (field, value) in enumerate(deductions, start=1):
            ctk.CTkLabel(
                result_frame, 
                text=field, 
                font=("Helvetica", 14),  # Regular font for fields
                text_color="black"  # Black color for fields
            ).grid(row=row_idx, column=2, padx=20, pady=5, sticky="w")
            ctk.CTkLabel(
                result_frame, 
                text=value, 
                font=("Helvetica", 14),  # Regular font for values
                text_color="black"  # Black color for values
            ).grid(row=row_idx, column=2, padx=20, pady=5, sticky="e")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# GUI Components
title_label = ctk.CTkLabel(
    app, 
    text="Employee Payroll Calculator", 
    font=("Helvetica", 24, "bold"),  # Larger and bold font for title
    text_color="black"  # Black color for title
)
title_label.pack(pady=30)

calculate_button = ctk.CTkButton(
    app, 
    text="Calculate Payroll", 
    command=calculate_payroll, 
    font=("Helvetica", 16),  # Larger font for button
    fg_color="#2E86C1",  # Blue color for button
    hover_color="#1B4F72",  # Darker blue on hover
    text_color="white"  # White text for button
)
calculate_button.pack(pady=20)

# Frame to display results
result_frame = ctk.CTkFrame(app)
result_frame.pack(pady=20, padx=40, fill="both", expand=True)

# Run the application
app.mainloop()