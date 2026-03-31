import os
from pathlib import Path

def find_real_md_files():
    # 루트 경로를 사용자가 알려준 정확한 notebooks 폴더로 지정합니다.
    root_dir = Path(r"C:\Users\H12018\AppData\Roaming\UpNote\UpNote Backup\F5p9KpV016SPRSpBH6y3uJ8NcFm2\Markdown\Synapse OS\notebooks")
    
    print(f"[{root_dir.name}] 폴더 및 모든 하위 트리의 마크다운 파일 탐색을 시작합니다...\n")
    
    if not root_dir.exists():
        print(f"오류: {root_dir} 경로가 존재하지 않습니다.")
        return
        
    # rglob을 통해 하위의 모든 뎁스를 재귀적으로 긁어옵니다.
    md_files = list(root_dir.rglob("*.md"))
    
    if not md_files:
        print("오류: 해당 경로 및 하위 폴더에 마크다운(.md) 파일이 전혀 없습니다.")
        return
        
    print(f"탐색 완료: 총 {len(md_files)}개의 마크다운 파일을 발견했습니다.\n")
    
    # 폴더별 파일 개수 집계
    parent_dirs = {}
    for f in md_files:
        parent = str(f.parent.relative_to(root_dir))
        parent_dirs[parent] = parent_dirs.get(parent, 0) + 1
        
    print("[하위 폴더별 파일 분포 현황]")
    for p, count in parent_dirs.items():
        folder_name = "최상위" if p == "." else p
        print(f" - {folder_name} (파일 수: {count}개)")
        
    print("\n[실제 파일명 샘플 5개 (UUID 형태인지 실제 제목인지 확인용)]")
    for f in md_files[:5]:
        print(f" - {f.name}")

if __name__ == "__main__":
    find_real_md_files()