import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import requests
from bs4 import BeautifulSoup
import re
from openpyxl import Workbook
import time
from datetime import timedelta
import threading

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

        problem_id = int(m.group(1))
        base = href.split("?")[0]
        status_url = f"https://ex-oj.sejong.ac.kr{base}?uid={student_id}"

        problems.append((problem_name, problem_id, status_url))

    return problems


def get_max_score(session, url):
    res = session.get(url)
    soup = BeautifulSoup(res.text, "html.parser")

    for td in soup.find_all("td"):
        text = td.get_text(strip=True)
        m = re.match(r"(\d+)\s*/\s*(\d+)", text)
        if m:
            return int(m.group(1))

    return 0


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
    student_ids = text_students.get("1.0", tk.END).strip().split("\n")
    save_path = save_path_var.get().strip()

    if not login_id or not login_pw or not problem_list_url:
        messagebox.showerror("오류", "모든 입력값을 입력하세요.")
        start_button.config(text="시작", command=start_process)
        return

    if not save_path:
        save_path = "scores.xlsx"

    use_ratio = chk_use_ratio_var.get() == 1

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
    else:
        grade_ratio = None

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

    try:
        response = session.get(problem_list_url)
        response.raise_for_status()
        html = response.text

        problems = extract_problems(html, student_ids[0])
        if not problems:
            raise ValueError("문제 목록을 읽을 수 없습니다.")

    except Exception as e:
        log(f"❌ 문제 페이지 오류: {e}")
        messagebox.showerror("오류", f"문제 페이지 파싱 실패:\n{e}")
        start_button.config(text="시작", command=start_process)
        return

    log(f"📌 문제 {len(problems)}개 확인됨")

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

        scores = {}

        for name, pid, url_template in problems:

            if cancel_flag:
                break

            url = re.sub(r"uid=\d+", f"uid={sid}", url_template)
            score = get_max_score(session, url)

            done += 1
            update_progress(done, total_tasks, start_time)
            log(f" - {name}: {score}")

            m = re.search(r"문제\s*(\d+)-(\d+)", name)
            if m:
                g = m.group(1)
                s = m.group(2)

                if g not in scores:
                    scores[g] = {"1": 0, "2": 0}

                scores[g][s] = score

        total_score = sum(calc_group_score(v["1"], v["2"]) for v in scores.values())
        log(f"🎯 총점 = {total_score}")
        rank_data.append((sid, total_score))

    rank_data.sort(key=lambda x: x[1], reverse=True)

    if use_ratio:
        graded_list = assign_grade_with_ratio(rank_data, grade_ratio)
    else:
        graded_list = [(sid, score, ("F" if score == 0 else "")) for sid, score in rank_data]

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
        messagebox.showerror("오류", f"저장 실패\n{e}")

    start_button.config(text="시작", command=start_process)


# ---------------- GUI ----------------

root = tk.Tk()
root.title("OJ 자동 채점 프로그램")
root.geometry("650x950")


# 저장 경로 선택
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
