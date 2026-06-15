import os
import json
import math
import tkinter as tk
from tkinter import messagebox, ttk

# 설정 및 데이터 저장 파일 경로 (로컬 PC 저장)
DB_FILE = "admin_db.json"

# 1. 관리자 데이터베이스 로드/저장 함수
def load_admin_db():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    # 기본 예시 데이터
    return {
        "MTRL-001": "P0804",
        "MTRL-002": "P1201",
        "MTRL-003": "P0802"
    }

def save_admin_db(data):
    try:
        with open(DB_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
    except Exception as e:
        messagebox.showerror("에러", f"관리자 DB 저장 중 오류가 발생했습니다: {e}")

class MaterialCheckerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📋 사내 보안 우회용 자재 청구 자동 점검 시스템 (Copy & Paste)")
        self.root.geometry("950x750")
        self.root.minsize(900, 650)

        self.admin_db = load_admin_db()

        # UI 스타일 설정
        self.style = ttk.Style()
        self.style.theme_use("clam")

        # 메인 탭 구성
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.tab_check = ttk.Frame(self.notebook)
        self.tab_admin = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_check, text=" 📊 자재 청구 복사 점검 ")
        self.notebook.add(self.tab_admin, text=" ⚙️ 관리자 모드 (포장사양 대량 관리) ")

        self.create_check_tab()
        self.create_admin_tab()

    # -------------------------------------------------------------------------
    # TAB 1: 자재 청구 점검 화면 구성 (Ctrl+V 복사 점검)
    # -------------------------------------------------------------------------
    def create_check_tab(self):
        # 상단 안내 영역
        info_frame = ttk.Frame(self.tab_check)
        info_frame.pack(fill="x", padx=15, pady=10)
        
        info_lbl = ttk.Label(
            info_frame, 
            text="💡 [사용방법]\n1. 엑셀 시트에서 A열부터 R열까지 드래그(최대 1,000개 행)하여 복사(Ctrl+C)합니다.\n2. 아래 입력창에 클릭 후 붙여넣기(Ctrl+V)하고 아래 [▶️ 점검 시작] 버튼을 클릭하세요.",
            font=("Arial", 10, "bold"),
            foreground="#005A9E"
        )
        info_lbl.pack(anchor="w")

        # 붙여넣기 입력창 영역 (A열부터 R열용)
        paste_frame = ttk.LabelFrame(self.tab_check, text="엑셀 데이터 붙여넣기 칸 (A열 ~ R열 데이터, 최대 1000행)")
        paste_frame.pack(fill="x", padx=15, pady=5)

        # 가로/세로 스크롤바가 장착된 텍스트박스
        x_scroll = ttk.Scrollbar(paste_frame, orient="horizontal")
        y_scroll = ttk.Scrollbar(paste_frame, orient="vertical")
        
        self.txt_paste = tk.Text(
            paste_frame, 
            height=8, 
            wrap="none", 
            xscrollcommand=x_scroll.set, 
            yscrollcommand=y_scroll.set,
            font=("Courier New", 9)
        )
        self.txt_paste.pack(side="left", fill="both", expand=True, padx=(5, 0), pady=5)
        
        y_scroll.config(command=self.txt_paste.yview)
        y_scroll.pack(side="right", fill="y", pady=5)
        x_scroll.config(command=self.txt_paste.xview)
        x_scroll.pack(side="bottom", fill="x")

        # 🚨 [점검 시작] 버튼 영역 (직관적이고 확실한 크기로 배치)
        btn_frame = ttk.Frame(self.tab_check)
        btn_frame.pack(fill="x", padx=15, pady=10)

        self.btn_start = tk.Button(
            btn_frame, 
            text="▶️ 점검 시작 (Start Verification)", 
            font=("Arial", 13, "bold"),
            bg="#0078D7", 
            fg="white", 
            activebackground="#005A9E",
            activeforeground="white",
            relief="raised",
            bd=3,
            command=self.run_verification
        )
        self.btn_start.pack(fill="x", ipady=10)

        # 점검 결과 통계 영역
        stats_frame = ttk.LabelFrame(self.tab_check, text="점검 결과 요약")
        stats_frame.pack(fill="x", padx=15, pady=5)

        self.lbl_stat_total = ttk.Label(stats_frame, text="총 점검 수: - 건", font=("Arial", 10, "bold"))
        self.lbl_stat_total.pack(side="left", padx=30, pady=8)

        self.lbl_stat_success = ttk.Label(stats_frame, text="정상 수량: - 건", foreground="green", font=("Arial", 10, "bold"))
        self.lbl_stat_success.pack(side="left", padx=30, pady=8)

        self.lbl_stat_error = ttk.Label(stats_frame, text="오류/불일치: - 건", foreground="red", font=("Arial", 10, "bold"))
        self.lbl_stat_error.pack(side="left", padx=30, pady=8)

        # 결과 테이블 목록 영역
        list_frame = ttk.LabelFrame(self.tab_check, text="상세 분석 리스트")
        list_frame.pack(fill="both", expand=True, padx=15, pady=10)

        # 결과 테이블 트리뷰
        columns = ("row_num", "material_code", "pack_spec", "qty_mass", "qty_common", "qty_req", "result", "desc")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings")

        self.tree.heading("row_num", text="행")
        self.tree.heading("material_code", text="자재코드(A)")
        self.tree.heading("pack_spec", text="포장사양(H)")
        self.tree.heading("qty_mass", text="양산청구수량(J)")
        self.tree.heading("qty_common", text="공용청구수량(L)")
        self.tree.heading("qty_req", text="필요수량(O)")
        self.tree.heading("result", text="판정 결과")
        self.tree.heading("desc", text="상세오류 정보")

        self.tree.column("row_num", width=40, anchor="center")
        self.tree.column("material_code", width=110, anchor="center")
        self.tree.column("pack_spec", width=90, anchor="center")
        self.tree.column("qty_mass", width=110, anchor="e")
        self.tree.column("qty_common", width=110, anchor="e")
        self.tree.column("qty_req", width=90, anchor="e")
        self.tree.column("result", width=90, anchor="center")
        self.tree.column("desc", width=250, anchor="w")

        # 스크롤바 연결
        scrollbar_y = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar_y.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar_y.pack(side="right", fill="y")

    def run_verification(self):
        raw_text = self.txt_paste.get("1.0", "end-1c").strip()
        if not raw_text:
            messagebox.showwarning("경고", "붙여넣기 칸에 엑셀 데이터를 먼저 입력(Ctrl+V)해 주세요!")
            return

        # 테이블 초기화
        for item in self.tree.get_children():
            self.tree.delete(item)

        lines = raw_text.split('\n')
        # 빈 줄 제외 및 최대 1000개 행 제한
        valid_lines = [line for line in lines if line.strip()]
        valid_lines = valid_lines[:1000]

        total_count = 0
        success_count = 0
        error_count = 0
        error_rows_info = []

        for idx, line in enumerate(valid_lines):
            excel_row_num = idx + 1  # 상대 행 번호 기록
            total_count += 1

            # 탭 문자 분해
            cols = line.split('\t')
            
            # Index Error를 막기 위해 최대 R열(18개 열) 크기로 강제 패딩 처리
            while len(cols) < 18:
                cols.append("")

            # 엑셀 열 매핑 매뉴얼 정보 바탕 추출
            # A열(Index 0): 자재코드, H열(Index 7): 포장사양, J열(Index 9): 양산청구수량, L열(Index 11): 공용청구수량, O열(Index 14): 필요수량
            material_code = cols[0].strip() if cols[0] else "N/A"
            packing_spec = cols[7].strip() if cols[7] else ""
            qty_mass_raw = cols[9].strip() if cols[9] else "0"
            qty_common_raw = cols[11].strip() if cols[11] else "0"
            qty_required_raw = cols[14].strip() if cols[14] else "0"

            is_db_mapped = False
            row_status = "정상"
            reasons = []

            # 💡 [보완기능] H열 포장사양이 비어있거나 누락되었을 경우 -> 관리자 DB 자동 매핑
            if packing_spec == "" or packing_spec.lower() in ["nan", "none"]:
                if material_code in self.admin_db:
                    packing_spec = self.admin_db[material_code]
                    is_db_mapped = True
                else:
                    row_status = "에러"
                    reasons.append("H열 포장사양 누락 및 관리자 DB 매핑 정보 없음")

            # 수량 데이터 정수 및 실수 형변환
            try:
                qty_required = float(qty_required_raw.replace(",", ""))
                qty_mass = float(qty_mass_raw.replace(",", ""))
                qty_common = float(qty_common_raw.replace(",", ""))
            except ValueError:
                row_status = "에러"
                reasons.append("수량 데이터가 올바른 숫자가 아닙니다")
                qty_required, qty_mass, qty_common = 0, 0, 0

            if row_status != "에러":
                # [규칙 1] 포장사양 코드 슬라이싱 (P0804 -> '08', P1201 -> '12')
                parsed_spec = ""
                if len(packing_spec) >= 3:
                    parsed_spec = packing_spec[1:3]
                else:
                    row_status = "에러"
                    reasons.append(f"포장사양 자릿수 부족(형식오류: '{packing_spec}')")

                # [규칙 2] 양산청구수량(J) == 필요수량(O) 검증
                if qty_mass != qty_required:
                    row_status = "에러"
                    reasons.append(f"양산수량 불일치(필요:{int(qty_required)} / 청구:{int(qty_mass)})")

                # [규칙 3] 공용청구수량(L) 조건 분기 적용
                if parsed_spec == "08":
                    expected_common = 30 if qty_required <= 500 else math.ceil(qty_required * 0.05)
                    if qty_common != expected_common:
                        row_status = "에러"
                        reasons.append(f"공용수량 오류(기준 {expected_common} / 청구 {int(qty_common)})")
                elif parsed_spec == "12":
                    expected_common = 3 # 12사양은 3개 고정 추가
                    if qty_common != expected_common:
                        row_status = "에러"
                        reasons.append(f"공용수량 오류('12'사양은 3개 고정 / 청구 {int(qty_common)})")

            # 판정 저장 및 UI 트리뷰 출력
            if row_status == "정상":
                success_count += 1
                status_text = "정상"
                desc_text = "검증 통과"
                if is_db_mapped:
                    desc_text += " (관리자 DB 자동보완)"
            else:
                error_count += 1
                status_text = "❌ 오류"
                desc_text = ", ".join(reasons)
                error_rows_info.append(f"• 붙여넣기 {excel_row_num}번째 행 ({material_code}): {desc_text}")

            self.tree.insert("", "end", values=(
                excel_row_num, material_code, packing_spec, 
                int(qty_mass) if qty_mass.is_integer() else qty_mass, 
                int(qty_common) if qty_common.is_integer() else qty_common, 
                int(qty_required) if qty_required.is_integer() else qty_required, 
                status_text, desc_text
            ))

        # 통계판 업데이트
        self.lbl_stat_total.config(text=f"총 점검 수: {total_count} 건")
        self.lbl_stat_success.config(text=f"정상 수량: {success_count} 건")
        self.lbl_stat_error.config(text=f"오류/불일치: {error_count} 건")

        # 결과 알림창 트리거
        if error_count == 0:
            messagebox.showinfo("🎉 검증 완료", f"축하합니다! 붙여넣은 {total_count}건의 데이터에 오류가 전혀 발견되지 않았습니다.")
        else:
            summary_msg = f"🚨 총 {error_count}개의 항목에서 수량 불일치 또는 사양 오류가 감지되었습니다.\n\n"
            summary_msg += "\n".join(error_rows_info[:10])  # 최대 10행 요약 출력
            if len(error_rows_info) > 10:
                summary_msg += f"\n\n외 {len(error_rows_info)-10}건의 불일치 오류가 추가 감지되었습니다."
            messagebox.showerror("🚨 수량 불일치 감지", summary_msg)


    # -------------------------------------------------------------------------
    # TAB 2: 관리자 모드 (자재 포장사양 대량 복사 등록 관리)
    # -------------------------------------------------------------------------
    def create_admin_tab(self):
        # 상단 설명 영역
        desc_lbl = ttk.Label(
            self.tab_admin, 
            text="⚙️ 관리자용 포장사양 일괄 데이터베이스 등록창\n\n여기에 자재별 포장사양 기준을 대량으로 복사해서 붙여넣고 아래 버튼을 누르면\n사양 정보가 없는 자재를 분석할 때 로컬 DB에서 대조하여 자동으로 분석해줍니다. (최대 20,000행 지원)",
            font=("Arial", 10)
        )
        desc_lbl.pack(fill="x", padx=15, pady=10)

        # 대량 일괄 붙여넣기 폼 영역
        bulk_frame = ttk.LabelFrame(self.tab_admin, text="포장사양 대량 붙여넣기 영역 (형식: 자재코드  포장사양)")
        bulk_frame.pack(fill="both", expand=True, padx=15, pady=5)

        x_scroll_admin = ttk.Scrollbar(bulk_frame, orient="horizontal")
        y_scroll_admin = ttk.Scrollbar(bulk_frame, orient="vertical")
        
        self.txt_admin_paste = tk.Text(
            bulk_frame, 
            height=12, 
            wrap="none", 
            xscrollcommand=x_scroll_admin.set, 
            yscrollcommand=y_scroll_admin.set,
            font=("Courier New", 9)
        )
        self.txt_admin_paste.pack(side="left", fill="both", expand=True, padx=(5, 0), pady=5)
        
        y_scroll_admin.config(command=self.txt_admin_paste.yview)
        y_scroll_admin.pack(side="right", fill="y", pady=5)
        x_scroll_admin.config(command=self.txt_admin_paste.xview)
        x_scroll_admin.pack(side="bottom", fill="x")

        # 대량 업데이트용 실행 버튼
        btn_db_update = tk.Button(
            self.tab_admin,
            text="💾 관리자 DB 일괄 업데이트 및 저장",
            font=("Arial", 11, "bold"),
            bg="#28A745",
            fg="white",
            activebackground="#218838",
            activeforeground="white",
            relief="raised",
            bd=3,
            command=self.run_admin_db_bulk_update
        )
        btn_db_update.pack(fill="x", padx=15, pady=10, ipady=6)

        # 목록 조회 및 삭제 프레임
        list_frame = ttk.LabelFrame(self.tab_admin, text="현재 로컬 DB에 등록 및 등록 예정인 자재 코드 수량 현황")
        list_frame.pack(fill="x", padx=15, pady=5)

        self.lbl_db_count = ttk.Label(list_frame, text="현재 등록된 포장사양 개수: 0 개", font=("Arial", 10, "bold"))
        self.lbl_db_count.pack(padx=20, pady=10, side="left")

        btn_clear_db = ttk.Button(list_frame, text="🗑️ 전체 DB 초기화", command=self.clear_admin_db)
        btn_clear_db.pack(padx=20, pady=10, side="right")

        self.update_db_count_label()

    def update_db_count_label(self):
        self.lbl_db_count.config(text=f"현재 로컬 DB에 등록된 자재 개수: {len(self.admin_db)} 개")

    def run_admin_db_bulk_update(self):
        raw_text = self.txt_admin_paste.get("1.0", "end-1c").strip()
        if not raw_text:
            messagebox.showwarning("경고", "붙여넣기 창에 등록할 대용량 데이터를 입력해 주세요.")
            return

        lines = raw_text.split('\n')
        valid_lines = [line for line in lines if line.strip()]
        valid_lines = valid_lines[:20000] # 최대 20,000행 제한

        new_entries_count = 0
        updated_entries_count = 0

        for line in valid_lines:
            # 탭 혹은 공백 문자로 스플릿 처리
            cols = line.split('\t')
            if len(cols) < 2:
                cols = line.split() # 탭이 없으면 공백 기준으로 스플릿 시도
            
            if len(cols) >= 2:
                mtrl_code = cols[0].strip()
                pack_spec = cols[1].strip()

                if mtrl_code and pack_spec:
                    if mtrl_code in self.admin_db:
                        updated_entries_count += 1
                    else:
                        new_entries_count += 1
                    self.admin_db[mtrl_code] = pack_spec

        # 파일에 영구 보존
        save_admin_db(self.admin_db)
        self.update_db_count_label()
        
        # 텍스트박스 리셋
        self.txt_admin_paste.delete("1.0", tk.END)

        messagebox.showinfo(
            "💾 DB 동기화 완료", 
            f"로컬 포장사양 데이터베이스를 성공적으로 업데이트하였습니다!\n\n"
            f"• 신규 등록 건수: {new_entries_count}건\n"
            f"• 기존 덮어쓰기 건수: {updated_entries_count}건\n"
            f"• 총 DB 누적 수량: {len(self.admin_db)}건"
        )

    def clear_admin_db(self):
        confirm = messagebox.askyesno("⚠️ DB 전체 삭제 확인", "정말로 저장되어 있는 모든 자재 포장사양 DB 데이터를 초기화하시겠습니까?")
        if confirm:
            self.admin_db = {}
            save_admin_db(self.admin_db)
            self.update_db_count_label()
            messagebox.showinfo("성공", "관리자 DB가 모두 안전하게 비워졌습니다.")

if __name__ == "__main__":
    root = tk.Tk()
    app = MaterialCheckerApp(root)
    root.mainloop()
