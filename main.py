import os
from pathlib import Path
import win32com.client
import subprocess
import shutil

download_folder = str(os.path.join(Path.home(), "Downloads")).replace("\\","/").replace("\\","/")

while (True):
    i = input("\nContinue ? (y/n) : ")
    if i.lower() == 'n':
        break
    elif i.lower():
        path = Path(os.getcwd()) / "temp"
        path.mkdir(parents= True, exist_ok= True)
        proc = subprocess.Popen(f"explorer {path}")
        path_in = str(path)
        print("\n*Drag and drop your .pptx files in the latest file exlorer folder*")
        path_out = input("\nOutput path to save the final merged ppt (leave empty to save to 'Downloads') : ")
        
        def merge_presentations(presentations, path):
            ppt_instance = win32com.client.Dispatch('PowerPoint.Application')
            prs = ppt_instance.Presentations.open(os.path.abspath(presentations[0]), True, False, False)
            
            for i in range(1, len(presentations)):
                try:
                    prs.Slides.InsertFromFile(os.path.abspath(presentations[i]), prs.Slides.Count)
                except:
                    pass
        
            prs.SaveAs(os.path.abspath(path))
            prs.Close()
        
        lst = []
        liste = os.listdir(path_in)
        
        for file in liste:
            if file[-4:] == 'pptx':
                lst.append(f"{path_in}/{file}")
                print(file)
                
        if path_out == '':
            output_path = f"{download_folder}/MergedPpt.pptx"
        else:
            output_path = f"{path_out}/MergedPpt.pptx"
            
        try:
            merge_presentations(lst, output_path)
        except IndexError:
            print(f"\nNo file has been added to {path_in}")
            
        shutil.rmtree(path_in, ignore_errors=False, onerror=None)
        break
    else:
        pass
