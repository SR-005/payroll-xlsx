import openpyxl
from openpyxl import load_workbook
import customtkinter as ctk
from tkinter import messagebox

# Initialize the main application window
ctk.set_appearance_mode("Dark")  # Set to Dark mode for white text
ctk.set_default_color_theme("blue")  # Themes: "blue" (default), "green", "dark-blue"

app = ctk.CTk()
app.title("Employee Payroll Calculator")
app.geometry("960x740")  # Set window size to 960x740

# Initialize frames as None (they will be created dynamically)
details_frame = None
earnings_frame = None
deductions_frame = None
summary_frame = None

# Function to read data from Excel and calculate results
def calculate_payroll():
    global details_frame, earnings_frame, deductions_frame, summary_frame

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

        # Clear previous results if frames already exist
        if details_frame:
            for widget in details_frame.winfo_children():
                widget.destroy()
        if earnings_frame:
            for widget in earnings_frame.winfo_children():
                widget.destroy()
        if deductions_frame:
            for widget in deductions_frame.winfo_children():
                widget.destroy()
        if summary_frame:
            for widget in summary_frame.winfo_children():
                widget.destroy()

        # Create frames if they don't exist
        if not details_frame:
            details_frame = ctk.CTkFrame(columns_frame, fg_color="#2E2E2E")  # Dark grey
            details_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        if not earnings_frame:
            earnings_frame = ctk.CTkFrame(columns_frame, fg_color="#2E2E2E")  # Dark grey
            earnings_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        if not deductions_frame:
            deductions_frame = ctk.CTkFrame(columns_frame, fg_color="#2E2E2E")  # Dark grey
            deductions_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        if not summary_frame:
            summary_frame = ctk.CTkFrame(app, fg_color="transparent")
            summary_frame.pack(fill="x", padx=20, pady=10)

        # Display the results in three columns
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

        ctk.CTkLabel(
            details_frame, 
            text="Details", 
            font=("Helvetica", 16, "bold"),  # Bold font for headers
            text_color="white"  # White color for headers
        ).pack(pady=10)

        for field, value in details:
            row_frame = ctk.CTkFrame(details_frame, fg_color="transparent")
            row_frame.pack(fill="x", padx=10, pady=5)
            ctk.CTkLabel(
                row_frame, 
                text=field, 
                font=("Helvetica", 14),  # Regular font for fields
                text_color="white"  # White color for fields
            ).pack(side="left", padx=10)
            ctk.CTkLabel(
                row_frame, 
                text=value, 
                font=("Helvetica", 14),  # Regular font for values
                text_color="white"  # White color for values
            ).pack(side="right", padx=10)

        # Column 2: Earnings
        earnings = [
            ("Basic Salary", f"₹{data['Basic Salary']}"),
            ("Conveyance", f"₹{data['Conveyance']}"),
            ("Special Allowance", f"₹{data['Special Allowance']}"),
        ]

        ctk.CTkLabel(
            earnings_frame, 
            text="Earnings", 
            font=("Helvetica", 16, "bold"),  # Bold font for headers
            text_color="white"  # White color for headers
        ).pack(pady=10)

        for field, value in earnings:
            row_frame = ctk.CTkFrame(earnings_frame, fg_color="transparent")
            row_frame.pack(fill="x", padx=10, pady=5)
            ctk.CTkLabel(
                row_frame, 
                text=field, 
                font=("Helvetica", 14),  # Regular font for fields
                text_color="white"  # White color for fields
            ).pack(side="left", padx=10)
            ctk.CTkLabel(
                row_frame, 
                text=value, 
                font=("Helvetica", 14),  # Regular font for values
                text_color="white"  # White color for values
            ).pack(side="right", padx=10)

        # Column 3: Deductions
        deductions = [
            ("PF Deduction", f"₹{data['PF Deduction']}"),
            ("ESI Deduction", f"₹{data['ESI Deduction']}"),
            ("PT Deduction", f"₹{data['PT Deduction']}"),
        ]

        ctk.CTkLabel(
            deductions_frame, 
            text="Deductions", 
            font=("Helvetica", 16, "bold"),  # Bold font for headers
            text_color="white"  # White color for headers
        ).pack(pady=10)

        for field, value in deductions:
            row_frame = ctk.CTkFrame(deductions_frame, fg_color="transparent")
            row_frame.pack(fill="x", padx=10, pady=5)
            ctk.CTkLabel(
                row_frame, 
                text=field, 
                font=("Helvetica", 14),  # Regular font for fields
                text_color="white"  # White color for fields
            ).pack(side="left", padx=10)
            ctk.CTkLabel(
                row_frame, 
                text=value, 
                font=("Helvetica", 14),  # Regular font for values
                text_color="white"  # White color for values
            ).pack(side="right", padx=10)

        # Summary Section
        ctk.CTkLabel(
            summary_frame, 
            text="Summary", 
            font=("Helvetica", 18, "bold"),  # Larger and bold font for summary header
            text_color="white"  # White color for summary header
        ).pack(pady=10)

        summary_data = [
            ("Total Earnings", f"₹{total_earnings}"),
            ("Total Deductions", f"₹{total_deductions}"),
            ("Net Pay", f"₹{net_pay}"),
        ]

        for field, value in summary_data:
            row_frame = ctk.CTkFrame(summary_frame, fg_color="transparent")
            row_frame.pack(fill="x", padx=20, pady=5)  # Adjusted padding for better spacing
            ctk.CTkLabel(
                row_frame, 
                text=field, 
                font=("Helvetica", 16, "bold"),  # Bold font for summary fields
                text_color="white"  # White color for summary fields
            ).pack(side="left", padx=10)
            ctk.CTkLabel(
                row_frame, 
                text=value, 
                font=("Helvetica", 16, "bold"),  # Bold font for summary values
                text_color="white"  # White color for summary values
            ).pack(side="right", padx=10)

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# GUI Components
title_label = ctk.CTkLabel(
    app, 
    text="Employee Payroll Calculator", 
    font=("Helvetica", 24, "bold"),  # Larger and bold font for title
    text_color="white"  # White color for title
)
title_label.pack(pady=20)

# Frame for buttons
button_frame = ctk.CTkFrame(app, fg_color="transparent")
button_frame.pack(pady=10)

# Calculate Payroll Button
calculate_button = ctk.CTkButton(
    button_frame, 
    text="Calculate Payroll", 
    command=calculate_payroll, 
    font=("Helvetica", 16),  # Larger font for button
    fg_color="#2E86C1",  # Blue color for button
    hover_color="#1B4F72",  # Darker blue on hover
    text_color="white"  # White text for button
)
calculate_button.pack(side="left", padx=10)

# Quit Button
quit_button = ctk.CTkButton(
    button_frame, 
    text="Quit", 
    command=app.quit,  # Close the application
    font=("Helvetica", 16),  # Larger font for button
    fg_color="#E74C3C",  # Red color for button
    hover_color="#943126",  # Darker red on hover
    text_color="white"  # White text for button
)
quit_button.pack(side="left", padx=10)

# Frame for columns (Details, Earnings, Deductions)
columns_frame = ctk.CTkFrame(app, fg_color="transparent")
columns_frame.pack(fill="both", expand=True, padx=20, pady=10)

# Run the application
app.mainloop()