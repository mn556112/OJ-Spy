import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import requests
from bs4 import BeautifulSoup
import re
from openpyxl import Workbook
import time
from datetime import timedelta
import threading
import sys, os

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

cancel_flag = False


def log(msg):
    log_box.insert(tk.END, msg + "\n")
    log_box.see(tk.END)


def update_progress(done, total, start_time):
    percent = done / total if total > 0 else 0
    progress_var.set(percent * 100)

    elapsed = time.time() - start_time
    eta = (elapsed / done * (total - done)) if done > 0 else 0
    eta_str = str(timedelta(seconds=int(eta)))

    progress_label.config(text=f"{percent*100:5.1f}% | ETA {eta_str}")
    root.update_idletasks()


def extract_problems(html, student_id):
    soup = BeautifulSoup(html, "html.parser")
    problems = []

    for tr in soup.select("tbody tr"):
        tds = tr.find_all("td")
        if len(tds) < 2:
            continue

        name_tag = tds[1].find("a")
        if not name_tag:
            continue

        problem_name = name_tag.get_text(strip=True)

        status_a = tds[-1].find("a", href=True)
        if not status_a:
            continue

        href = status_a["href"]
        if "judge/status" not in href:
            continue

        m = re.search(r"/(\d+)\?uid=", href)
        if not m:
            continue

        base = href.split("?")[0]
        status_url = f"https://ex-oj.sejong.ac.kr{base}?uid={student_id}"

        problems.append((problem_name, status_url))

    return problems


def get_max_score(session, url):
    m = re.search(r"(.*?/status/\d+/\d+/\d+)(/\d+)?(\?.*)", url)
    if not m:
        raise ValueError("URL 구조 인식 실패: " + url)

    base = m.group(1)
    suffix = m.group(3)

    max_score = 0
    page = 1

    while True:
        if page == 1:
            page_url = f"{base}{suffix}"
        else:
            page_suffix = 10 * (page - 1)
            page_url = f"{base}/{page_suffix}{suffix}"

        res = session.get(page_url)
        soup = BeautifulSoup(res.text, "html.parser")
        spans = soup.select("td span")

        found = False
        for sp in spans:
            text = sp.get_text(strip=True)
            m2 = re.search(r"(\d+)", text)
            if m2:
                score = int(m2.group(1))
                found = True
                if score > max_score:
                    max_score = score

        if not found:
            break

        page += 1

    return max_score

def calc_group_score(score1, score2):
    score1_half = score1 * 0.5

    if score1 == 0 and score2 == 0:
        return 0

    if score1_half >= score2:
        return score1_half

    return score2


def cancel_process():
    global cancel_flag
    cancel_flag = True
    log("⛔ 작업 취소 요청됨...")


def start_process():
    global cancel_flag
    cancel_flag = False
    start_button.config(text="취소", command=cancel_process)
    threading.Thread(target=run_program).start()



def assign_grade_with_ratio(rank_data, grade_ratio):
    n = len(rank_data)

    cutA = int(n * grade_ratio["A"])
    cutB = cutA + int(n * grade_ratio["B"])
    cutC = cutB + int(n * grade_ratio["C"])
    cutD = cutC + int(n * grade_ratio["D"])

    graded = []

    for i, (sid, score) in enumerate(rank_data, start=1):

        if score == 0:
            graded.append((sid, score, "F"))
            continue

        if i <= cutA:
            base = "A"; start, end = 1, cutA
        elif i <= cutB:
            base = "B"; start, end = cutA + 1, cutB
        elif i <= cutC:
            base = "C"; start, end = cutB + 1, cutC
        elif i <= cutD:
            base = "D"; start, end = cutC + 1, cutD
        else:
            graded.append((sid, score, "F"))
            continue

        mid = (start + end) / 2
        grade = base + ("+" if i <= mid else "0")
        graded.append((sid, score, grade))

    return graded



def run_program():
    global cancel_flag

    login_id = entry_id.get().strip()
    login_pw = entry_pw.get().strip()
    problem_list_url = entry_url.get().strip()

    student_ids = [s.strip() for s in text_students.get("1.0", tk.END).split("\n") if s.strip()]

    save_path = save_path_var.get().strip()
    if not save_path:
        save_path = "scores.xlsx"

    use_ratio = chk_use_ratio_var.get() == 1

    # 등급 비율 체크
    if use_ratio:
        try:
            grade_ratio = {
                "A": float(entry_ratio_A.get()),
                "B": float(entry_ratio_B.get()),
                "C": float(entry_ratio_C.get()),
                "D": float(entry_ratio_D.get()),
                "F": float(entry_ratio_F.get())
            }
        except:
            messagebox.showerror("오류", "비율은 숫자여야 합니다.")
            return

        if abs(sum(grade_ratio.values()) - 1.0) > 0.01:
            messagebox.showerror("오류", "비율 합이 1.0이어야 합니다.")
            return

    # 체크박스 규칙
    if use_individual_var.get() + use_group_var.get() != 1:
        messagebox.showerror("오류", "개별 문제 또는 그룹 문제 중 하나만 선택하세요.")
        return


    log("▶ 로그인 중...")

    session = requests.Session()
    LOGIN_URL = "https://ex-oj.sejong.ac.kr/index.php/auth/authentication?returnURL="

    resp = session.post(LOGIN_URL, data={"id": login_id, "password": login_pw}, allow_redirects=False)
    location = resp.headers.get("Location", "")

    if resp.status_code == 303 and "index.php/judge" in location:
        log("✅ 로그인 성공!")
    else:
        log(f"❌ 로그인 실패 (status={resp.status_code}, location={location})")
        start_button.config(text="시작", command=start_process)
        return


    # 문제 목록 파싱
    try:
        response = session.get(problem_list_url)
        response.raise_for_status()
        html = response.text

        problems = extract_problems(html, student_ids[0])
        if not problems:
            raise ValueError("문제 목록을 읽을 수 없습니다.")
    except Exception as e:
        log(f"❌ 문제 페이지 오류: {e}")
        messagebox.showerror("오류", str(e))
        start_button.config(text="시작", command=start_process)
        return


    total_tasks = len(student_ids) * len(problems)
    done = 0
    start_time = time.time()
    rank_data = []

    for sid in student_ids:
        if cancel_flag:
            log("⛔ 취소됨")
            start_button.config(text="시작", command=start_process)
            return

        log(f"\n▶ {sid} 계산 중...")

        problem_scores = []

        for name, url_template in problems:

            if cancel_flag:
                break

            url = re.sub(r"uid=\d+", f"uid={sid}", url_template)
            score = get_max_score(session, url)

            done += 1
            update_progress(done, total_tasks, start_time)
            log(f" - {name}: {score}")

            problem_scores.append((name, score))

        # -----------------------------
        # 계산 모드에 따라 총점 계산
        # -----------------------------

        # ⭐ 1) 개별 문제 계산 모드
        if use_individual_var.get() == 1:
            total_score = sum(score for _, score in problem_scores)

        # ⭐ 2) 그룹 문제 계산 모드
        else:
            groups = {}
            individual = []

            for name, score in problem_scores:

                m = re.search(r"(\d+)\s*[-_]\s*(\d+)", name)
                if m:
                    g = m.group(1)
                    s = m.group(2).lstrip("0") or "0"

                    # ➤ 그룹은 s=1 또는 s=2 일 때만 인정
                    if s in ["1", "2"]:
                        if g not in groups:
                            groups[g] = {"1": None, "2": None}
                        groups[g][s] = score
                    else:
                        individual.append(score)
                else:
                    individual.append(score)

            # 그룹 방식 총점 계산
            total_score = sum(individual)

            for g, subs in groups.items():
                filled = [v for v in subs.values() if v is not None]

                if len(filled) == 1:
                    total_score += filled[0]

                elif len(filled) == 2:
                    s1 = subs["1"] or 0
                    s2 = subs["2"] or 0
                    total_score += calc_group_score(s1, s2)


        log(f"🎯 총점 = {total_score}")
        rank_data.append((sid, total_score))


    # 순위 정렬
    rank_data.sort(key=lambda x: x[1], reverse=True)

    if use_ratio:
        graded_list = assign_grade_with_ratio(rank_data, grade_ratio)
    else:
        graded_list = [(sid, score, ("F" if score == 0 else "")) for sid, score in rank_data]

    # -----------------------------
    # 엑셀 저장
    # -----------------------------
    wb = Workbook()
    ws = wb.active
    ws.append(["순위", "학번", "총점", "등급"])

    for i, (sid, total, grade) in enumerate(graded_list, start=1):
        ws.append([i, sid, total, grade])

    try:
        wb.save(save_path)
        log(f"📄 저장 완료: {save_path}")
        messagebox.showinfo("완료", "작업이 완료되었습니다.")
    except Exception as e:
        log(f"❌ 저장 실패: {e}")
        messagebox.showerror("오류", str(e))

    start_button.config(text="시작", command=start_process)



# ---------------- GUI ----------------

root = tk.Tk()
root.title("OJ SPY v1.0.2")
root.geometry("650x1000")
root.iconbitmap(resource_path("icon.ico"))

# 저장 경로
ttk.Label(root, text="결과 저장 위치").pack()
save_path_var = tk.StringVar()
entry_save_path = ttk.Entry(root, textvariable=save_path_var)
entry_save_path.pack(fill="x")

def choose_save_path():
    path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel 파일", "*.xlsx")],
        title="저장 위치 선택"
    )
    if path:
        save_path_var.set(path)

ttk.Button(root, text="저장 위치 선택", command=choose_save_path).pack(pady=5)

# 로그인 입력
ttk.Label(root, text="OJ 아이디").pack()
entry_id = ttk.Entry(root)
entry_id.pack(fill="x")

ttk.Label(root, text="OJ 비밀번호").pack()
entry_pw = ttk.Entry(root, show="*")
entry_pw.pack(fill="x")

ttk.Label(root, text="문제 리스트 URL").pack()
entry_url = ttk.Entry(root)
entry_url.pack(fill="x")

# 학번 입력
ttk.Label(root, text="학번 리스트").pack()
text_students = tk.Text(root, height=8)
text_students.pack(fill="both")

# 체크박스 추가: 개별 / 그룹 계산 선택
use_individual_var = tk.IntVar()
use_group_var = tk.IntVar()

chk_individual = ttk.Checkbutton(root, text="개별 문제로 계산", variable=use_individual_var)
chk_group = ttk.Checkbutton(root, text="그룹 문제로 계산", variable=use_group_var)

chk_individual.pack()
chk_group.pack()

# 등급 비율
chk_use_ratio_var = tk.IntVar()
chk_use_ratio = ttk.Checkbutton(root, text="등급 비율 사용 (A/B/C/D/F)", variable=chk_use_ratio_var)
chk_use_ratio.pack(pady=5)

frame_ratio = ttk.LabelFrame(root, text="등급 비율 (합계=1.0)")
frame_ratio.pack(fill="x", pady=10)

labels = ["A", "B", "C", "D", "F"]
entries = {}

for i, grade in enumerate(labels):
    ttk.Label(frame_ratio, text=f"{grade} 비율").grid(row=i, column=0)
    ent = ttk.Entry(frame_ratio)
    ent.insert(0, "0.20" if grade in ["A","B","C","D"] else "0.00")
    ent.grid(row=i, column=1)
    entries[grade] = ent

entry_ratio_A = entries["A"]
entry_ratio_B = entries["B"]
entry_ratio_C = entries["C"]
entry_ratio_D = entries["D"]
entry_ratio_F = entries["F"]

progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, maximum=100, variable=progress_var)
progress_bar.pack(fill="x", pady=10)

progress_label = ttk.Label(root, text="0.0%")
progress_label.pack()

log_box = tk.Text(root, height=15)
log_box.pack(fill="both", pady=10)

start_button = ttk.Button(root, text="시작", command=start_process)
start_button.pack(pady=10)

root.mainloop()
