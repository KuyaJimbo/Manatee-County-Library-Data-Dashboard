# Import Helper Functions From Python Libraries
from openpyxl import load_workbook, Workbook
import os, re
import shutil

def create_branch_template():
    # Ask if they are sure they want to create a new branch template (y/n)
    confirmation = input("Are you sure you want to create a new branch template? (y/n): ").lower().strip()
    
    if confirmation != 'y':
        print("Operation cancelled.")
        return
    
    # Ask for the name of the branch (excluding the word branch)
    branch_name = input("Enter the name of the branch (excluding the word 'branch'): ").strip()
    
    if not branch_name:
        print("Branch name cannot be empty.")
        return
    
    # Keep track of the name with a variable
    branch_filename = f"{branch_name} Branch.xlsx"
    
    try:
        # Go to parent directory
        parent_dir = os.path.dirname(os.getcwd())
        
        # Search for folder named 'Template'
        template_folder = os.path.join(parent_dir, 'Template')
        
        if not os.path.exists(template_folder):
            print(f"Error: Template folder not found at {template_folder}")
            return
        
        # Path to the original template file
        original_template = os.path.join(template_folder, "Template for Branches.xlsx")
        
        if not os.path.exists(original_template):
            print(f"Error: 'Template for Branches.xlsx' not found in {template_folder}")
            return
        
        # Create a copy of "Template for Branches.xlsx" and rename it
        new_template_path = os.path.join(template_folder, branch_filename)
        
        # Check if file already exists
        if os.path.exists(new_template_path):
            overwrite = input(f"File '{branch_filename}' already exists. Overwrite? (y/n): ").lower().strip()
            if overwrite != 'y':
                print("Operation cancelled.")
                return
        
        # Copy the file
        shutil.copy2(original_template, new_template_path)
        
        # Load the workbook and modify the specified cells
        workbook = load_workbook(new_template_path)
                
        # Change the following cells so they have the value of the name provided
        cells_to_update = ['B4', 'B38', 'B47', 'B56', 'B65']
        
        for worksheet in workbook.worksheets:
            for cell in cells_to_update:
                worksheet[cell] = branch_name
        
        # Save the workbook
        workbook.save(new_template_path)
        workbook.close()
        
        # Print that it was successful
        print(f"Success! Branch template '{branch_filename}' has been created in the Template folder.")
        input("Press Enter to exit...")
        
    except FileNotFoundError as e:
        print(f"Error: File not found - {e}")
    except PermissionError as e:
        print(f"Error: Permission denied - {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    create_branch_template()