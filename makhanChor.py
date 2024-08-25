import zipfile
import os
import shutil
import re
import argparse
from colorama import Fore, Back, Style

def display_msg():
    print(Fore.GREEN)
    print("""
                               █████      █████                          
                          ░░███      ░░███                           
 █████████████    ██████   ░███ █████ ░███████    ██████   ████████  
░░███░░███░░███  ░░░░░███  ░███░░███  ░███░░███  ░░░░░███ ░░███░░███ 
 ░███ ░███ ░███   ███████  ░██████░   ░███ ░███   ███████  ░███ ░███ 
 ░███ ░███ ░███  ███░░███  ░███░░███  ░███ ░███  ███░░███  ░███ ░███ 
 █████░███ █████░░████████ ████ █████ ████ █████░░████████ ████ █████
░░░░░ ░░░ ░░░░░  ░░░░░░░░ ░░░░ ░░░░░ ░░░░ ░░░░░  ░░░░░░░░ ░░░░ ░░░░░                                                     
                                                                     
   █████████  █████                                                  
  ███░░░░░███░░███                                                   
 ███     ░░░  ░███████    ██████  ████████                           
░███          ░███░░███  ███░░███░░███░░███                          
░███          ░███ ░███ ░███ ░███ ░███ ░░░                           
░░███     ███ ░███ ░███ ░███ ░███ ░███                               
 ░░█████████  ████ █████░░██████  █████                              
  ░░░░░░░░░  ░░░░ ░░░░░  ░░░░░░  ░░░░░                               
    """)
    print(Style.RESET_ALL)
    print(Fore.GREEN+Back.WHITE+"\nBy: Asim Tara Pathak | A tool to unlock password protected pptx file.\n"+Style.RESET_ALL)

def process_pptx_file(pptx_file_path):
    print(Fore.GREEN)
    
    # Extract filename and extension
    base_name, ext = os.path.splitext(pptx_file_path)
    
    if ext.lower() != '.pptx':
        print("Error: The file is not a PowerPoint (.pptx) file.")
        return

    # Define paths
    zip_file_path = base_name + '.zip'
    temp_dir = base_name + '_temp'

    # Rename the PPTX file to ZIP
    os.rename(pptx_file_path, zip_file_path)
    print(f"Renamed {pptx_file_path} to {zip_file_path}")

    # Extract the ZIP file
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        os.makedirs(temp_dir, exist_ok=True)
        zip_ref.extractall(temp_dir)
        print(f"Extracted files to {temp_dir}")

    # Define the path to the presentation.xml file
    presentation_xml_path = None
    for root, dirs, files in os.walk(temp_dir):
        for file in files:
            if file == 'presentation.xml':
                presentation_xml_path = os.path.join(root, file)
                break
        if presentation_xml_path:
            break

    # Check if the presentation.xml file was found
    if not presentation_xml_path:
        print("Error: presentation.xml not found in the extracted files.")
        shutil.rmtree(temp_dir)
        os.rename(zip_file_path, pptx_file_path)  # Restore original file
        return

    print(f"Found presentation.xml at {presentation_xml_path}")

    # Read the content of the presentation.xml file
    with open(presentation_xml_path, 'r', encoding='utf-8') as file:
        content = file.read()

    # Define the pattern to remove
    pattern = re.compile(r'<p:modifyVerifier.*?</p:extLst>', re.DOTALL)
    modified_content = pattern.sub('', content)

    # Write the modified content to a temporary file with .txt extension
    base_name_xml, ext_xml = os.path.splitext(presentation_xml_path)
    temp_txt_path = base_name_xml + '.txt'
    with open(temp_txt_path, 'w', encoding='utf-8') as file:
        file.write(modified_content)

    print(f"Modified XML content and saved to {temp_txt_path}")

    # # Remove the existing presentation.xml file
    if os.path.isfile(presentation_xml_path):
        os.remove(presentation_xml_path)

    print(f"Removed existing {presentation_xml_path}")

    # Rename the temporary .txt file back to .xml
    new_xml_path = presentation_xml_path  # The original presentation.xml path
    if os.path.isfile(temp_txt_path):
        os.rename(temp_txt_path, new_xml_path)
        print(f"Renamed {temp_txt_path} back to {new_xml_path}")

    # Recreate the ZIP file with the updated content
    new_pptx_file_path = base_name + ext

    # Use a context manager to handle ZIP file creation
    with zipfile.ZipFile(new_pptx_file_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, temp_dir)
                zip_ref.write(file_path, arcname)

    print(f"Recreated ZIP file as {new_pptx_file_path}")

    # Clean up temporary files and directories
    shutil.rmtree(temp_dir)
    os.remove(zip_file_path)
    print(f"Cleaned up temporary files and removed {zip_file_path}")

    print(Style.RESET_ALL)

if __name__ == "__main__":

    parser = argparse.ArgumentParser(description='Process a PowerPoint file by modifying its XML content.')
    parser.add_argument('pptx_file', help='Path to the PowerPoint file to process')
    args = parser.parse_args()

    display_msg()
    try:
        process_pptx_file(args.pptx_file)
    except:
        print(Fore.RED,"Error!! File type not supported by tool",Style.RESET_ALL)
