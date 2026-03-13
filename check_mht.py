import os
import subprocess
import glob
import win32com.client

def test_method(name, func):
    print(f"--- Testing Method: {name} ---")
    try:
        content = func()
        if content and ("html" in content.lower() or "<body" in content.lower()):
            print(f"✅ Success! Found HTML tags. (Length: {len(content)})")
            print(f"Sample: {content[:200]}...")
            return True
        else:
            print("❌ Failed: No HTML content found or content is empty.")
    except Exception as e:
        print(f"❌ Error: {e}")
    return False

def check_file(file_path):
    abs_path = os.path.abspath(file_path)
    print(f"\nChecking file: {abs_path}")

    # Method 1: CMD 'type' command
    def via_type():
        result = subprocess.check_output(f'type "{abs_path}"', shell=True, stderr=subprocess.STDOUT)
        return result.decode('utf-8', errors='ignore')

    # Method 2: PowerShell 'Get-Content'
    def via_powershell():
        cmd = f'powershell -Command "Get-Content -Path \'{abs_path}\' -Raw"'
        result = subprocess.check_output(cmd, shell=True, stderr=subprocess.STDOUT)
        return result.decode('utf-8', errors='ignore')

    # Method 3: Scripting.FileSystemObject (COM)
    def via_fso():
        fso = win32com.client.Dispatch("Scripting.FileSystemObject")
        file = fso.OpenTextFile(abs_path, 1) # 1 = ForReading
        content = file.ReadAll()
        file.Close()
        return content

    results = []
    results.append(test_method("CMD Type", via_type))
    results.append(test_method("PowerShell Get-Content", via_powershell))
    results.append(test_method("FileSystemObject (COM)", via_fso))
    
    if not any(results):
        print("\n[Summary] All headless methods failed. DRM might be blocking these processes too.")
    else:
        print("\n[Summary] Headless reading is possible!")

if __name__ == "__main__":
    files = glob.glob("inputs/*.mht")
    if not files:
        print("No .mht files found in 'inputs/' folder.")
    else:
        check_file(files[0])
