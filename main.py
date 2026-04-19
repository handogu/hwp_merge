import os
from pathlib import Path
import traceback

def merge_hwp_files():
    try:
        from pyhwpx import Hwp
    except ImportError:
        print("❌ pyhwpx 라이브러리가 설치되지 않았습니다.")
        print("설치 명령: pip install python-hwp")
        return
    
    hwp = None
    try:
        current_path = os.getcwd()
        
        # 한글 파일 찾기 (숨김 파일 제외)
        files = sorted([
            f for f in os.listdir(current_path) 
            if f.lower().endswith(('.hwp', '.hwpx')) 
            and not f.startswith('~$')
            and os.path.isfile(os.path.join(current_path, f))
        ])
        
        if not files:
            print("⚠️ 합칠 한글 파일이 폴더에 없습니다.")
            return
        
        print(f"📄 발견된 파일 ({len(files)}개): {files}")
        
        # 첫 번째 파일 열기
        hwp = Hwp()
        first_file = os.path.join(current_path, files[0])
        print(f"📖 기본 파일 열기: {files[0]}")
        hwp.open(first_file)
        
        # 나머지 파일 삽입
        for file in files[1:]:
            file_path = os.path.join(current_path, file)
            print(f"➕ 파일 병합 중: {file}")
            try:
                # "Section" 대신 "NextPage" 옵션도 시도 가능
                hwp.insert_file(file_path, "Section")
            except Exception as e:
                print(f"⚠️ {file} 삽입 실패: {str(e)}")
                continue
        
        # 병합 파일 저장
        output_file = os.path.join(current_path, "★전체_부서_합본.hwp")
        print(f"💾 병합 파일 저장 중: {output_file}")
        hwp.save_as(output_file)
        
        print("✅ 병합 완료!")
        
    except Exception as e:
        print(f"❌ 오류 발생: {str(e)}")
        traceback.print_exc()
    
    finally:
        # 확실하게 프로세스 종료
        if hwp is not None:
            try:
                hwp.quit()
            except:
                pass

if __name__ == "__main__":
    merge_hwp_files()
