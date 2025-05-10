import subprocess

# Run the first script
try:
    print("Running script1.py...")
    subprocess.run(["python3", "data_analasys.py"], check=True)
    print("Finished script1.py")
except subprocess.CalledProcessError as e:
    print(f"Error running script1.py: {e}")
    exit(1)

# Run the second script
try:
    print("Running script2.py...")
    subprocess.run(["python3", "summery_with_error_bars.py"], check=True)
    print("Finished script2.py")
except subprocess.CalledProcessError as e:
    print(f"Error running script2.py: {e}")
    exit(1)
