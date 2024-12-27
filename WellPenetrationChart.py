import pandas as pd  # Importing pandas for data manipulation
import matplotlib.pyplot as plt  # Importing matplotlib for plotting
import numpy as np  # Importing numpy for numerical operations

# Function to determine color and fraction based on the state of a category
def get_color_and_fraction(state, color):
    if state == "Present":
        return color, 1.0  # Full color if present
    elif state == "Ambiguous":
        return [color, 'white'], [0.5, 0.5]  # Half color and half white if ambiguous
    elif state == "Not Present":
        return 'white', 1.0  # White if not present
    elif state == "Unevaluated":
        return 'gray', 1.0  # Gray if unevaluated
    else:
        return color, 1.0  # Default to full color

# Function to create a pie chart in the specified axis with given category states and diameter
def create_pie_chart(ax, category_states, diameter):
    categories = ["Reservoir", "Seal", "Charge", "Trap"]  # Define the categories
    colors = {
        "Reservoir": "yellow",
        "Seal": "brown",
        "Charge": "red",
        "Trap": "green"
    }  # Assign colors to each category
    
    labels = []
    wedge_colors = []
    wedge_fractions = []

    # Loop through each category and determine its color and fraction
    for category in categories:
        state = category_states.get(category, "Unevaluated")  # Get the state or default to unevaluated
        color, fraction = get_color_and_fraction(state, colors[category])  # Get color and fraction based on state
        if isinstance(fraction, list):
            labels.extend([category] * len(fraction))  # Extend labels for ambiguous states
            wedge_colors.extend(color)  # Extend colors for ambiguous states
            wedge_fractions.extend(fraction)  # Extend fractions for ambiguous states
        else:
            labels.append(category)  # Append label for non-ambiguous states
            wedge_colors.append(color)  # Append color for non-ambiguous states
            wedge_fractions.append(fraction)  # Append fraction for non-ambiguous states

    # Create the pie chart with specified properties
    ax.pie(wedge_fractions, colors=wedge_colors, startangle=90, 
           wedgeprops={"edgecolor": "black", 'linewidth': 0.5, 'antialiased': True}, radius=diameter/2)
    ax.axis('equal')  # Ensure the pie chart is a circle
    ax.set_frame_on(True)  # Set frame on for better visibility

    # Initialize a flag to check if all values are "Present"
    all_present = True

    # Loop through the dictionary values
    for value in category_states.values():
        if value != 'Present':
            all_present = False
            break  # Exit the loop early if any value is not "Present"

    # Check the flag and set facecolor if all values are "Present"
    if all_present:
        ax.set_facecolor('#8CCD60') # light green

# Main function to generate the well penetration chart
def main(input_file,
         well_col='Well Name',
         reservoir_col='Reservoir',
         seal_col='Seal',
         charge_col='Charge',
         trap_col='Trap',
         play_col='Play',
         age_col='Age',
         cell_width=2.0,  # Width of each cell
         cell_height=1.0,  # Height of each cell
         header_cell_width=1.5,  # Width of header cells
         header_cell_height=1.5,  # Height of header cells
         vertical_header_width=0.8,  # Fixed width for vertical headers
         grid_line_thickness=0.5,  # Thickness of grid lines
         grid_line_color='black'):  # Color of grid lines
    
    # Read the Excel file into a DataFrame
    df = pd.read_excel(input_file)
    
    # Print column names for debugging purposes
    print("Column Names in DataFrame:")
    print(df.columns.tolist())
    
    # Check if required columns exist in the DataFrame
    required_columns = [well_col, reservoir_col, seal_col, charge_col, trap_col, play_col, age_col]
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        raise KeyError(f"Missing columns: {missing_columns}")  # Raise an error if any required column is missing
    
    # Check for empty cells in specific columns and fill them with 'Unevaluated'
    empty_cells = df[(df[reservoir_col].isnull()) | (df[seal_col].isnull()) | (df[charge_col].isnull()) | (df[trap_col].isnull())]
    if len(empty_cells) != 0:
        print(f"Warning: Empty cells found in columns {reservoir_col}, {seal_col}, {charge_col}, {trap_col}. Assigning 'Unevaluated' to these cells.")
        df[reservoir_col] = df[reservoir_col].fillna('Unevaluated')
        df[seal_col] = df[seal_col].fillna('Unevaluated')
        df[charge_col] = df[charge_col].fillna('Unevaluated')
        df[trap_col] = df[trap_col].fillna('Unevaluated')
    
    # Check for empty cells in other columns and fill them with 'Empty'
    empty_cells_other = df[df.isnull().any(axis=1)]
    if len(empty_cells_other) != 0:
        print(f"Warning: Empty cells found in other columns. Assigning 'Empty' string to these cells.")
        df.fillna('Empty', inplace=True)
    
    # Sort the DataFrame by Age in ascending order
    df = df.sort_values(by=age_col, ascending=True)

    # Set colors from a column named Color if it exists; otherwise, use grey
    color_col='Color'
    if color_col not in df.columns:  # Check if the column is in the input file
        color_col='grey'  # If not present, then color the plays grey
    else:
        color_col = df[color_col].unique()  # Get unique colors from the Color column

    # Extract unique plays and wells from the DataFrame
    unique_plays = df[play_col].unique()
    unique_wells = df[well_col].unique()
    
    num_wells = len(unique_wells)
    num_plays = len(unique_plays)
    
    # Determine figure size dynamically based on number of plays and wells
    fig_width = cell_width * (num_wells + 1) + header_cell_width + vertical_header_width
    fig_height = cell_height * num_plays + header_cell_height
    
    # Create subplots with dynamic figure size
    fig, axes = plt.subplots(nrows=num_plays, ncols=num_wells + 1, figsize=(fig_width, fig_height),
                              gridspec_kw={'wspace': 0, 'hspace': 0})
    
    # Handle cases where there is only one play or well to ensure axes is always a 2D array
    if num_wells == 1:
        axes = np.array([axes]).reshape(1, -1)
    if num_plays == 1:
        axes = np.array([axes]).reshape(-1, 1)
    
    # Set vertical header column with fixed width and add play names
    for i in range(num_plays):
        play_value = unique_plays[i]
        # Add text with rotation and vertical padding
        axes[i, 0].text(0.1, 0.5, play_value, fontsize=12, fontweight='bold', rotation='horizontal', va='center', ha='left')
    
    # Set main headers for wells
    for j in range(num_wells):
        well_name = unique_wells[j]
        axes[0, j + 1].set_title(well_name, fontsize=10, pad=header_cell_height * 0.5, fontweight='bold', rotation=0)
    
    # Create pie charts for each cell based on the data
    for i in range(num_plays):
        axes[i, 0].set_facecolor(f'{color_col[i]}')  # Add color to the Plays labels from the input file
        for j in range(num_wells):
            well_data = df[(df[well_col] == unique_wells[j]) & (df[play_col] == unique_plays[i])]
            if len(well_data) != 0:
                category_states = {
                    "Reservoir": well_data[reservoir_col].values[0],
                    "Seal": well_data[seal_col].values[0],
                    "Charge": well_data[charge_col].values[0],
                    "Trap": well_data[trap_col].values[0]
                }
                diameter = min(cell_width, cell_height)  # Determine the diameter of the pie chart
                create_pie_chart(axes[i, j + 1], category_states, diameter)
            else:
                axes[i, j + 1].axis('on')  # Ensure axis is on if no data
    
    # Remove subplot tick labels for cleaner look
    for ax in axes.flatten():
        ax.set_xticks([])
        ax.set_yticks([])
    
    # Add grid lines to display boundaries between subplots
    for ax in axes.flatten():
        ax.spines['top'].set_color(grid_line_color)
        ax.spines['bottom'].set_color(grid_line_color)
        ax.spines['left'].set_color(grid_line_color)
        ax.spines['right'].set_color(grid_line_color)
        ax.spines['top'].set_linewidth(grid_line_thickness)
        ax.spines['bottom'].set_linewidth(grid_line_thickness)
        ax.spines['left'].set_linewidth(grid_line_thickness)
        ax.spines['right'].set_linewidth(grid_line_thickness)


    # Add a title to the entire chart
    plt.suptitle('Well Penetration Chart', fontsize=16, fontweight='bold')

    # Display the plot
    plt.show()

# Example usage of the main function with specified parameters
Commented_code_above_heavily = True  # This line is just for reference and not part of the code

# Uncomment the following lines to run the main function with your specific input file
main(input_file=r'C:\Users\Ed\Desktop\plays.xlsm',
     well_col='Well Name',
     reservoir_col='Reservoir',
     seal_col='Seal',
     charge_col='Charge',
     trap_col='Trap',
     play_col='Play',
     age_col='Age',
     cell_width=2.0,
     cell_height=1.0,
     header_cell_width=1.5,
     header_cell_height=1.5,
     vertical_header_width=0.8,
     grid_line_thickness=0.5,
     grid_line_color='#DCDCDC') # light grey
