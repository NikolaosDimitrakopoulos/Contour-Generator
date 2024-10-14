import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from scipy.interpolate import griddata
from scipy.spatial import ConvexHull
from matplotlib import path as mplPath
import matplotlib.cm as cm  # To get access to Matplotlib colormaps

# Global variables to store the loaded data
data = None
sheets = None
excel_file_path = None

# Function to load the Excel file and populate the sheet dropdown
def load_file():
    global sheets, data, excel_file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            # Load the Excel file to get sheet names
            sheets = pd.ExcelFile(file_path).sheet_names
            sheet_var.set(sheets[0])
            sheet_menu['menu'].delete(0, 'end')
            for sheet in sheets:
                sheet_menu['menu'].add_command(label=sheet, command=tk._setit(sheet_var, sheet))
            excel_file_path = file_path
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load the file: {str(e)}")

# Function to load the selected sheet and update column dropdowns
def load_sheet():
    global data
    selected_sheet = sheet_var.get()
    try:
        # Load the selected sheet into the DataFrame, specifying the header is in the second row (index 1)
        data = pd.read_excel(excel_file_path, sheet_name=selected_sheet, header=1)
        column_names = data.columns
        # Update dropdowns with column names
        x_axis_var.set(column_names[0])
        y_axis_var.set(column_names[1])
        z_axis_var.set(column_names[2])
        x_axis_menu['menu'].delete(0, 'end')
        y_axis_menu['menu'].delete(0, 'end')
        z_axis_menu['menu'].delete(0, 'end')
        for col in column_names:
            x_axis_menu['menu'].add_command(label=col, command=tk._setit(x_axis_var, col))
            y_axis_menu['menu'].add_command(label=col, command=tk._setit(y_axis_var, col))
            z_axis_menu['menu'].add_command(label=col, command=tk._setit(z_axis_var, col))
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load the sheet: {str(e)}")

def generate_contour():
    try:
        # Extract selected columns for X, Y, Z
        X = data[x_axis_var.get()].values
        Y = data[y_axis_var.get()].values
        Z = data[z_axis_var.get()].values

        # Check if X, Y, Z arrays have consistent lengths
        if len(X) != len(Y) or len(X) != len(Z):
            raise ValueError("X, Y, and Z columns must have the same length.")

        # Plot the raw data to allow the user to select the boundary points
        fig, ax = plt.subplots()
        scatter = ax.scatter(X, Y, c=Z, cmap='viridis', marker='o', label='Data Points')
        plt.colorbar(scatter, ax=ax, label=z_axis_var.get())
        plt.xlabel(x_axis_var.get())
        plt.ylabel(y_axis_var.get())
        plt.title('Click to define the boundary for interpolation (Press Enter when done)')

        # Let user click on points to define the hull boundary
        boundary_points = plt.ginput(n=-1, timeout=0, show_clicks=True)  # User can click indefinitely until Enter
        plt.close(fig)  # Close the interactive plot after the selection

        # Convert the list of boundary points to a NumPy array
        boundary_points = np.array(boundary_points)

        # Ensure the user clicked at least 3 points
        if len(boundary_points) < 3:
            raise ValueError("You must define at least three points to form a boundary.")

        # Create a ConvexHull based on the manually selected boundary points
        hull = ConvexHull(boundary_points)

        # Create a regular grid to interpolate data onto
        x_unique = np.linspace(min(boundary_points[:, 0]), max(boundary_points[:, 0]), 500)
        y_unique = np.linspace(min(boundary_points[:, 1]), max(boundary_points[:, 1]), 500)

        X_grid, Y_grid = np.meshgrid(x_unique, y_unique)

        # Create a mask to exclude points outside the convex hull
        boundary_path = mplPath.Path(boundary_points)
        grid_points = np.vstack((X_grid.ravel(), Y_grid.ravel())).T
        mask = boundary_path.contains_points(grid_points).reshape(X_grid.shape)

        # Get the selected interpolation method and colormap
        method = interpolation_var.get()
        colormap = colormap_var.get()  # Get the selected colormap

        # Interpolate Z values onto the grid using griddata
        Z_grid = griddata((X, Y), Z, (X_grid, Y_grid), method=method)

        # Apply the mask to ensure Z values outside the user-defined hull are NaN
        Z_grid[~mask] = np.nan  # Set Z values outside the boundary to NaN

        # Plotting the filled contour only inside the manually defined boundary
        plt.figure()
        contour_plot = plt.contourf(X_grid, Y_grid, Z_grid, levels=50, cmap=colormap, extend='both')
        plt.colorbar(contour_plot, label=z_axis_var.get())

        # Plot the boundary points in red for debugging
        plt.scatter(boundary_points[:, 0], boundary_points[:, 1], c='red', marker='x', label='Selected Boundary Points')

        # Plot the original data points colored by their Z values
        plt.scatter(X, Y, c=Z, cmap=colormap, edgecolor='slategrey', marker='o', label='Data Points (colored by Z value)')

        # Labels and title
        plt.xlabel(x_axis_var.get())
        plt.ylabel(y_axis_var.get())
        plt.legend()

        plt.show()

    except Exception as e:
        messagebox.showerror("Error", f"Failed to generate contour plot: {str(e)}")


# Initialize tkinter window
root = tk.Tk()
root.title("Contour Plot Generator")

# Set the window size
root.geometry("450x550")

# Variables to store column and sheet selections
x_axis_var = tk.StringVar()
y_axis_var = tk.StringVar()
z_axis_var = tk.StringVar()
sheet_var = tk.StringVar()
interpolation_var = tk.StringVar(value='linear')  # Default to linear interpolation in GUI
colormap_var = tk.StringVar(value='viridis')  # Default colormap to viridis

# List of common colormaps
common_colormaps = ['viridis', 'jet', 'plasma', 'inferno', 'magma', 'cividis', 'Greys', 'Purples', 'Blues', 'Greens', 'Oranges']

# Load Excel File Button
load_button = tk.Button(root, text="Load Excel File", command=load_file)
load_button.pack(pady=10)

# Dropdown for sheet selection
tk.Label(root, text="Select Sheet:").pack()
sheet_menu = tk.OptionMenu(root, sheet_var, ())
sheet_menu.pack(pady=5)

# Load Sheet Button
load_sheet_button = tk.Button(root, text="Load Sheet", command=load_sheet)
load_sheet_button.pack(pady=10)

# Dropdowns for X, Y, and Z columns
tk.Label(root, text="Select X Axis Column:").pack()
x_axis_menu = tk.OptionMenu(root, x_axis_var, ())
x_axis_menu.pack(pady=5)

tk.Label(root, text="Select Y Axis Column:").pack()
y_axis_menu = tk.OptionMenu(root, y_axis_var, ())
y_axis_menu.pack(pady=5)

tk.Label(root, text="Select Z Axis Column:").pack()
z_axis_menu = tk.OptionMenu(root, z_axis_var, ())
z_axis_menu.pack(pady=5)

# Add Dropdown for interpolation method selection
tk.Label(root, text="Select Interpolation Method:").pack()
interpolation_menu = tk.OptionMenu(root, interpolation_var, 'linear', 'cubic')
interpolation_menu.pack(pady=10)

# Dropdown for colormap selection
tk.Label(root, text="Select Colormap:").pack()
colormap_menu = tk.OptionMenu(root, colormap_var, *common_colormaps)
colormap_menu.pack(pady=10)

# Run Button
run_button = tk.Button(root, text="Run", command=generate_contour)
run_button.pack(pady=20)

# Main loop
root.mainloop()
