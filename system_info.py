import os
import platform
import subprocess
import re
import speedtest
from win32api import GetSystemMetrics
import math

def run_command(command):
    try:
        # Run a command and return the stripped stdout
        result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        return result.stdout.strip()
    except Exception as e:
        return f"Error: {e}"

def get_installed_software():
    # Get a list of installed software using WMIC command
    return run_command(['wmic', 'product', 'get', 'name'])

def get_internet_speed():
    # Get internet speed using the speedtest library
    st = speedtest.Speedtest()
    download_speed = st.download()
    upload_speed = st.upload()
    return f"Download Speed: {download_speed / 1024 / 1024:.2f} Mbps, Upload Speed: {upload_speed / 1024 / 1024:.2f} Mbps"

def get_screen_resolution():
    # Get screen resolution using WMIC command
    resolution_info = run_command(['wmic', 'desktopmonitor', 'get', 'screenwidth,screenheight'])
    match = re.search(r'(\d+)\s+(\d+)', resolution_info)
    if match:
        return f"Screen Resolution: {match.group(2)}x{match.group(1)}"  # width x height
    return "Screen Resolution information not available"

def calculate_dpi(width_pixels, height_pixels, diagonal_inches):
    # Calculate DPI using the Pythagorean theorem
    diagonal_pixels = math.sqrt(width_pixels**2 + height_pixels**2)
    dpi = diagonal_pixels / diagonal_inches
    return dpi

def get_cpu_info():
    # Get CPU information using WMIC command
    cpu_info = run_command(['wmic', 'cpu', 'get', 'caption,NumberOfCores,NumberOfLogicalProcessors'])
    match = re.search(r'(.+?)\s+(\d+)\s+(\d+)', cpu_info)
    if match:
        return match.group(1), int(match.group(2)), int(match.group(3))
    return "CPU information not available", 0, 0

def get_gpu_info():
    # Get GPU information using WMIC command
    return run_command(['wmic', 'path', 'win32_videocontroller', 'get', 'caption'])

def get_ram_size():
    # Get RAM size using WMIC command
    ram_info = run_command(['wmic', 'memorychip', 'get', 'capacity'])
    ram_sizes = [int(size) for size in ram_info.split('\n') if size.strip().isdigit()]
    total_ram_gb = sum(ram_sizes) // (1024 ** 3)
    return f"{total_ram_gb} GB"

def get_screen_size():
    # Get screen size, diagonal size, and DPI
    width = GetSystemMetrics(0)
    height = GetSystemMetrics(1)
    diagonal_inches = 15  # Replace with the actual diagonal size of your screen in inches
    dpi = calculate_dpi(width, height, diagonal_inches)
    return f"Screen Size: {width}x{height} pixels, Diagonal Size: {diagonal_inches} inches, DPI: {dpi:.2f}"

def get_mac_address(interface='Wi-Fi'):
    # Get MAC address using the getmac command
    mac_info = run_command(['getmac', '/FO', 'CSV', '/V'])
    mac_addresses = re.findall(r'([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})', mac_info)
    return ''.join(mac_addresses[0]) if mac_addresses else "MAC address not found"

def get_public_ip():
    # Get public IP address using curl
    return run_command(['curl', 'ifconfig.me'])

def get_windows_version():
    # Get Windows version using the platform module
    return platform.version()

if __name__ == "__main__":
    print("Installed Software:")
    print(get_installed_software())

    print("\nInternet Speed:")
    print(get_internet_speed())

    print("\nScreen Resolution:")
    print(get_screen_resolution())

    print("\nCPU Information:")
    cpu_model, cores, threads = get_cpu_info()
    print(f"CPU Model: {cpu_model}")
    print(f"Number of Cores: {cores}")
    print(f"Number of Threads: {threads}")

    print("\nGPU Information:")
    print(get_gpu_info())

    print("\nRAM Size:")
    print(get_ram_size())

    print("\nScreen Size:")
    print(get_screen_size())

    print("\nMAC Address (Wi-Fi):")
    print(get_mac_address())

    print("\nPublic IP Address:")
    print(get_public_ip())

    print("\nWindows Version:")
    print(get_windows_version())
