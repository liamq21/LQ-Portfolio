import pandas as pd
import matplotlib.pyplot as plt
from tkinter import Tk, filedialog, Button, Label, Frame
from tkinter import messagebox
import tkinter.font as font
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib.ticker import ScalarFormatter

# Setting colours scheme
colors = {
    'background': '#f0f4f7',    # Soft light grey background
    'button': '#4CAF50',        # Modern green for the buttons
    'button_text': 'white',     # White text for buttons
    'text': '#333333',          # Dark grey text
    'title_color': '#30475e',   # Blue-grey for plot titles
    'label_color': '#222831',   # Darker color for plot labels
    'plot1': '#17a2b8',         # Blue color for plot 1
    'plot2': '#17a2b8',         
    'plot3': '#17a2b8',         
    'plot4': '#17a2b8',         
}

# Function to process the Excel file and show all graphs on the same page
def process_file():
    # Open a file dialog to select the Excel file
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    
    if not file_path:
        messagebox.showerror("Error", "No file selected!")
        return
    
    try:
        # Read the selected Excel file
        df = pd.read_excel(file_path)
        
        # Convert the settlement_date column to datetime format
        df['settlement_date'] = pd.to_datetime(df['settlement_date'])
        
        # Average property price trend over time
        df['year_month'] = df['settlement_date'].dt.to_period('M')
        avg_price_trend = df.groupby('year_month')['property_price'].mean()
        
        # Property type distribution
        property_type_distribution = df['property_type'].value_counts()
        
        # Average buyer age by property type
        avg_age_by_type = df.groupby('property_type')['buyer_age'].mean()
        
        # Refinancing trends over time
        refinancing_trend = df.groupby('year_month')['refinancing'].sum()

        # Clear the GUI before creating new graphs
        for widget in frame.winfo_children():
            widget.destroy()

        # Create the figure and subplots
        fig = Figure(figsize=(14, 10), dpi=100)
        
        # Plot 1: Average property price trend over time
        ax1 = fig.add_subplot(221)
        avg_price_trend.plot(kind='line', ax=ax1, color=colors['plot1'])
        ax1.set_title("Average Property Price Trend", fontsize=14, fontweight='bold', color=colors['title_color'])
        ax1.set_xlabel("Year-Month", fontsize=12, color=colors['label_color'])
        ax1.set_ylabel("Average Price", fontsize=12, color=colors['label_color'])
        ax1.yaxis.set_major_formatter(ScalarFormatter())
        ax1.grid(True)

        # Plot 2: Property type distribution
        ax2 = fig.add_subplot(222)
        property_type_distribution.plot(kind='bar', ax=ax2, color=colors['plot2'])
        ax2.set_title("Property Type Distribution", fontsize=14, fontweight='bold', color=colors['title_color'])
        ax2.set_xlabel("Property Type", fontsize=12, color=colors['label_color'])
        ax2.set_ylabel("Number of Properties", fontsize=12, color=colors['label_color'])

        # Plot 3: Average buyer age by property type
        ax3 = fig.add_subplot(223)
        avg_age_by_type.plot(kind='bar', ax=ax3, color=colors['plot3'])
        ax3.set_title("Average Buyer Age by Property Type", fontsize=14, fontweight='bold', color=colors['title_color'])
        ax3.set_xlabel("Property Type", fontsize=12, color=colors['label_color'])
        ax3.set_ylabel("Average Age", fontsize=12, color=colors['label_color'])

        # Plot 4: Refinancing trends over time
        ax4 = fig.add_subplot(224)
        refinancing_trend.plot(kind='line', ax=ax4, color=colors['plot4'])
        ax4.set_title("Refinancing Trends Over Time", fontsize=14, fontweight='bold', color=colors['title_color'])
        ax4.set_xlabel("Year-Month", fontsize=12, color=colors['label_color'])
        ax4.set_ylabel("Number of Refinances", fontsize=12, color=colors['label_color'])
        ax4.grid(True)

        # Adjust layout spacing
        fig.tight_layout(pad=4.0)  # Increase padding to avoid overlap

        # Add the figure to the GUI window
        canvas = FigureCanvasTkAgg(fig, master=frame)
        canvas.draw()
        canvas.get_tk_widget().pack()

    except Exception as e:
        messagebox.showerror("Error", f"Failed to process file: {str(e)}")

# Create the GUI 
root = Tk()
root.title("Property Data Insights")
root.geometry("1200x900")  # Resize to accommodate graphs
root.config(bg=colors['background'])

# Editing fonts and styles
title_font = font.Font(family="Helvetica", size=16, weight="bold")
button_font = font.Font(family="Helvetica", size=12)

# Title label
title_label = Label(root, text="Upload Property Settlement Data", bg=colors['background'], fg=colors['text'], font=title_font)
title_label.pack(pady=20)

# Button to select the Excel file
button = Button(root, text="Select Excel File", command=process_file, bg=colors['button'], fg=colors['button_text'], font=button_font, padx=10, pady=5)
button.pack(pady=10)

# Frame for the graphs
frame = Frame(root, bg=colors['background'])
frame.pack(fill='both', expand=True, padx=10, pady=10)

# Start the GUI loop
root.mainloop()
