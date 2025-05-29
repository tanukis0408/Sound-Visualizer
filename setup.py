import os
import sys
import subprocess

def install_dependencies():
    """Installs dependencies from requirements.txt using pip."""
    print("Detecting operating system...")
    platform = sys.platform

    if platform.startswith('win'):
        print("Detected Windows.")
        command = [sys.executable, '-m', 'pip', 'install', '-r', 'requirements.txt']
    else:
        print(f"Detected {platform} (assuming Unix-like).")
        command = [sys.executable, '-m', 'pip', 'install', '-r', 'requirements.txt']

    print("Installing dependencies...")
    try:
        # Use sys.executable to ensure the correct python environment's pip is used
        result = subprocess.run(command, check=True, capture_output=True, text=True)
        print("Installation successful.")
        print("\n--- Pip Output ---")
        print(result.stdout)
        if result.stderr:
            print("\n--- Pip Errors (if any) ---")
            print(result.stderr)
        print("------------------")

    except subprocess.CalledProcessError as e:
        print(f"Error during installation: {e}")
        print("\n--- Pip Output (if any) ---")
        print(e.stdout)
        if e.stderr:
            print("\n--- Pip Errors ---")
            print(e.stderr)
        sys.exit(1)
    except FileNotFoundError:
        print("Error: Could not find pip. Make sure Python is installed and in your PATH.")
        sys.exit(1)

if __name__ == "__main__":
    install_dependencies() 