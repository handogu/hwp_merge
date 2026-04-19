import os
from pyhwpx import Hwp

def merge_hwp_files():
    hwp = Hwp()
    current_path = os.getcwd()
    files = [f for f in os.listdir(current_path) 
             if f.lower().endswith(('.hwp', '.hwpx')) and not f.startswith('~$')]
    files.sort()

    if not files:
        print("합칠 한글 파일이 폴더에 없습니다.")
        return

    hwp.open(os.path.join(current_path, files[0]))
    for file in files[1:]:
        hwp.insert_file(os.path.join(current_path, file), "Section")

    hwp.save_as(os.path.join(current_path, "★전체_부서_합본.hwp"))
    hwp.quit()

if __name__ == "__main__":
    merge_hwp_files()
