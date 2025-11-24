import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.integrate import cumulative_trapezoid
import os

# ================= CONFIGURATION =================
DATA_FOLDER = "Datasets"
OUTPUT_FOLDER = "Output_Advanced"
TRIM_SECONDS = 1.0  

if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

# ================= UTILITIES =================
def load_phyphox_xls(filepath):
    try:
        df_acc = pd.read_excel(filepath, sheet_name='Accelerometer')
        df_gyro = pd.read_excel(filepath, sheet_name='Gyroscope')
        
        # Standardize Time
        acc_t = [c for c in df_acc.columns if "Time" in c][0]
        gyro_t = [c for c in df_gyro.columns if "Time" in c][0]
        df_acc.rename(columns={acc_t: "Time"}, inplace=True)
        df_gyro.rename(columns={gyro_t: "Time"}, inplace=True)
        
        # Merge
        df_merged = df_acc.copy()
        for col in df_gyro.columns:
            if col != "Time":
                df_merged[col] = np.interp(df_acc["Time"], df_gyro["Time"], df_gyro[col])
        
        # Trim
        t_min = df_merged["Time"].min() + TRIM_SECONDS
        t_max = df_merged["Time"].max() - TRIM_SECONDS
        return df_merged[(df_merged["Time"] >= t_min) & (df_merged["Time"] <= t_max)].copy()
    except Exception as e:
        print(f"Error loading {filepath}: {e}")
        return None

def get_calibration_bias():
    """Finds stationary.xls and calculates Zero-G offset."""
    # Look for file with "stationary" in name
    files = [f for f in os.listdir(DATA_FOLDER) if "stationary" in f.lower() and f.endswith(('.xls', '.xlsx'))]
    
    if not files:
        print("WARNING: 'stationary.xls' not found! Using 0 bias (Correction will fail).")
        return {}
    
    path = os.path.join(DATA_FOLDER, files[0])
    print(f"Calibrating with: {files[0]}...")
    df = load_phyphox_xls(path)
    
    bias = {}
    for col in df.columns:
        if "Time" not in col:
            bias[col] = df[col].mean()
            print(f"  > {col} Bias: {bias[col]:.4f}")
    return bias

# ================= ANALYSIS =================
def analyze_drift(df, filename, bias):
    """Plots Raw Velocity (Drift) vs Corrected Velocity."""
    # Identify Y-axis (Forward)
    acc_col = [c for c in df.columns if "Acceleration y" in c][0]
    
    time = df["Time"].values
    acc_raw = df[acc_col].values
    
    # 1. Raw Velocity (Problem)
    vel_raw = cumulative_trapezoid(acc_raw, time, initial=0)
    
    # 2. Corrected Velocity (Solution)
    acc_corr = acc_raw - bias.get(acc_col, 0)
    vel_corr = cumulative_trapezoid(acc_corr, time, initial=0)
    
    plt.figure(figsize=(10, 6))
    plt.plot(time, vel_raw, label="Raw (Drift)", color='gray', linestyle='--')
    plt.plot(time, vel_corr, label="Corrected (Calibrated)", color='green', linewidth=2)
    plt.title(f"Drift Correction Analysis - {filename}")
    plt.xlabel("Time (s)")
    plt.ylabel("Velocity (m/s)")
    plt.legend()
    plt.grid(True)
    
    plt.savefig(os.path.join(OUTPUT_FOLDER, f"{filename}_Drift_Compare.png"))
    plt.close()

def analyze_heading(df, filename, bias):
    """Plots Steering Angle from Gyro Z."""
    gyro_col = [c for c in df.columns if "Gyroscope z" in c][0]
    
    time = df["Time"].values
    gyro_raw = df[gyro_col].values
    
    # Remove bias
    gyro_corr = gyro_raw - bias.get(gyro_col, 0)
    
    # Integrate
    angle_rad = cumulative_trapezoid(gyro_corr, time, initial=0)
    angle_deg = np.degrees(angle_rad)
    
    plt.figure(figsize=(10, 6))
    plt.plot(time, angle_deg, label="Heading Angle", color='purple', linewidth=2)
    plt.title(f"Steering Logic - {filename}")
    plt.xlabel("Time (s)")
    plt.ylabel("Degrees")
    plt.axhline(0, color='black', linestyle='--')
    plt.grid(True)
    plt.savefig(os.path.join(OUTPUT_FOLDER, f"{filename}_Heading.png"))
    plt.close()

# ================= MAIN =================
# 1. Get Bias
bias = get_calibration_bias()

# 2. Process Files
files = [f for f in os.listdir(DATA_FOLDER) if f.endswith(('.xls', '.xlsx'))]
print(f"Starting Advanced Analysis on {len(files)} files...")

for f in files:
    name = os.path.splitext(f)[0].lower()
    if "stationary" in name: continue # Skip calibration file
        
    df = load_phyphox_xls(os.path.join(DATA_FOLDER, f))
    if df is None: continue
    
    # Logic Router
    if "straight" in name or "reverse" in name:
        analyze_drift(df, name, bias)
        print(f"  > Generated Velocity plot for {name}")
        
    elif "lane" in name:
        analyze_heading(df, name, bias)
        print(f"  > Generated Heading plot for {name}")

print("\nDone! Check 'Output_Advanced' folder.")