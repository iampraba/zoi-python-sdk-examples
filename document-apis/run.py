import os
import subprocess
import sys


def run_all_python_files_in_folder(folder_path):
    # List all files in the given folder
    files = os.listdir(folder_path)

    # Filter out files that have a .py extension
    python_files = [f for f in files if f.endswith('.py')]

    # Run each Python file
    for python_file in python_files:
        file_path = os.path.join(folder_path, python_file)
        try:
            result = subprocess.run(['python3', file_path], check=True, capture_output=True, text=True)
            print(f'Output of {python_file}:\n{result.stdout}')
        except subprocess.CalledProcessError as e:
            print(f'Error running {python_file}:\n{e.stderr}')


if __name__ == "__main__":
    if len(sys.argv) == 2:
        folder_path = sys.argv[1]
    else:
        folder_path = os.getcwd()  # Use current directory as default

    run_all_python_files_in_folder(folder_path)
