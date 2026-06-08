import argparse
import json
import os
import re
import subprocess
import sys
import threading
import traceback
from contextlib import redirect_stderr, redirect_stdout
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk


APP_NAME = "GmarketCrawler"
TASK_NAME = "GmarketDailyCrawler"
DEFAULT_RUN_TIME = "14:00"


def get_app_data_dir():
    base = os.environ.get("APPDATA")
    if base:
        return Path(base) / APP_NAME
    return Path.home() / f".{APP_NAME}"


def get_config_path():
    return get_app_data_dir() / "config.json"


def get_log_dir():
    return get_app_data_dir() / "logs"


def get_log_path():
    return get_log_dir() / f"{datetime.now():%Y-%m-%d}.log"


def default_config():
    return {
        "workbook_path": "",
        "run_time": DEFAULT_RUN_TIME,
    }


def load_config(config_path=None):
    path = Path(config_path) if config_path else get_config_path()
    if not path.exists():
        return default_config()

    data = json.loads(path.read_text(encoding="utf-8"))
    config = default_config()
    config.update(data)
    return config


def save_config(config, config_path=None):
    path = Path(config_path) if config_path else get_config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(config, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def validate_run_time(value):
    if not isinstance(value, str) or not re.fullmatch(r"\d{2}:\d{2}", value):
        raise ValueError("시간은 HH:MM 형식으로 입력해 주세요. 예: 14:00")

    hour, minute = map(int, value.split(":"))
    if hour > 23 or minute > 59:
        raise ValueError("시간은 00:00부터 23:59 사이여야 합니다.")

    return value


def get_run_command():
    if getattr(sys, "frozen", False):
        return [sys.executable, "--run"]
    return [sys.executable, str(Path(__file__).resolve()), "--run"]


def build_schtasks_args(run_command=None, run_time=DEFAULT_RUN_TIME):
    command = run_command if run_command is not None else get_run_command()
    return [
        "schtasks",
        "/Create",
        "/TN",
        TASK_NAME,
        "/TR",
        subprocess.list2cmdline(command),
        "/SC",
        "DAILY",
        "/ST",
        validate_run_time(run_time),
        "/F",
    ]


def register_daily_task(run_time, run_command=None):
    args = build_schtasks_args(run_command=run_command, run_time=run_time)
    return subprocess.run(args, check=True, capture_output=True, text=True)


def validate_workbook_path(workbook_path):
    path = Path(workbook_path)
    if not path.exists():
        raise ValueError("엑셀 파일을 찾을 수 없습니다.")
    if path.suffix.lower() != ".xlsx":
        raise ValueError("xlsx 엑셀 파일을 선택해 주세요.")
    return path


def run_crawler_from_config(config_path=None):
    log_path = get_log_path()
    log_path.parent.mkdir(parents=True, exist_ok=True)

    with log_path.open("a", encoding="utf-8") as log_file:
        with redirect_stdout(log_file), redirect_stderr(log_file):
            print(f"\n[{datetime.now():%Y-%m-%d %H:%M:%S}] 실행 시작")
            try:
                config = load_config(config_path)
                workbook_path = validate_workbook_path(config.get("workbook_path", ""))
                import gmarket

                gmarket.main(workbook_path=workbook_path)
                print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] 실행 완료")
            except Exception:
                print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] 실행 실패")
                traceback.print_exc()
                raise

    return log_path


class GmarketCrawlerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gmarket 자동 수집 설정")
        self.root.geometry("400x240")
        self.root.resizable(False, False)

        config = load_config()
        run_time = config.get("run_time", DEFAULT_RUN_TIME)
        if not re.fullmatch(r"\d{2}:\d{2}", run_time):
            run_time = DEFAULT_RUN_TIME

        self.workbook_var = tk.StringVar(value=config.get("workbook_path", ""))
        self.hour_var = tk.StringVar(value=run_time.split(":")[0])
        self.minute_var = tk.StringVar(value=run_time.split(":")[1])
        self.status_var = tk.StringVar(value="엑셀 파일과 실행 시간을 선택해 주세요.")

        self.build_ui()

    def build_ui(self):
        frame = ttk.Frame(self.root, padding=18)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="엑셀 파일").grid(row=0, column=0, sticky="w")
        workbook_entry = ttk.Entry(frame, textvariable=self.workbook_var, width=52)
        workbook_entry.grid(row=1, column=0, sticky="ew", pady=(4, 12))
        ttk.Button(frame, text="찾기", command=self.browse_workbook).grid(
            row=1,
            column=1,
            padx=(8, 0),
            pady=(4, 12),
        )

        ttk.Label(frame, text="매일 실행 시간").grid(row=2, column=0, sticky="w")
        time_frame = ttk.Frame(frame)
        time_frame.grid(row=3, column=0, sticky="w", pady=(4, 16))
        ttk.Spinbox(
            time_frame,
            from_=0,
            to=23,
            wrap=True,
            width=4,
            textvariable=self.hour_var,
            format="%02.0f",
        ).pack(side="left")
        ttk.Label(time_frame, text="시").pack(side="left", padx=(4, 10))
        ttk.Spinbox(
            time_frame,
            from_=0,
            to=59,
            wrap=True,
            width=4,
            textvariable=self.minute_var,
            format="%02.0f",
        ).pack(side="left")
        ttk.Label(time_frame, text="분").pack(side="left", padx=(4, 0))

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=4, column=0, columnspan=2, sticky="w")
        ttk.Button(button_frame, text="저장하고 예약 등록", command=self.save_and_register).pack(
            side="left"
        )
        ttk.Button(button_frame, text="지금 실행", command=self.run_now).pack(
            side="left",
            padx=(8, 0),
        )
        ttk.Button(button_frame, text="로그 폴더 열기", command=self.open_log_folder).pack(
            side="left",
            padx=(8, 0),
        )

        ttk.Label(frame, textvariable=self.status_var, foreground="#444").grid(
            row=5,
            column=0,
            columnspan=2,
            sticky="w",
            pady=(18, 0),
        )

        frame.columnconfigure(0, weight=1)

    def browse_workbook(self):
        path = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if path:
            self.workbook_var.set(path)

    def get_selected_config(self):
        hour = int(self.hour_var.get())
        minute = int(self.minute_var.get())
        run_time = validate_run_time(f"{hour:02d}:{minute:02d}")
        workbook_path = validate_workbook_path(self.workbook_var.get())
        return {
            "workbook_path": str(workbook_path),
            "run_time": run_time,
        }

    def save_and_register(self):
        try:
            config = self.get_selected_config()
            save_config(config)
            register_daily_task(config["run_time"])
        except Exception as exc:
            messagebox.showerror("예약 등록 실패", str(exc))
            self.status_var.set("예약 등록 실패. 로그 또는 메시지를 확인해 주세요.")
            return

        self.status_var.set(f"매일 {config['run_time']}에 자동 실행되도록 등록했습니다.")
        messagebox.showinfo("예약 등록 완료", "자동 실행 예약을 등록했습니다.")

    def run_now(self):
        try:
            config = self.get_selected_config()
            save_config(config)
        except Exception as exc:
            messagebox.showerror("실행 실패", str(exc))
            return

        self.status_var.set("크롤러를 실행 중입니다. 브라우저가 열릴 수 있습니다.")
        thread = threading.Thread(target=self.run_now_in_background, daemon=True)
        thread.start()

    def run_now_in_background(self):
        try:
            subprocess.Popen(get_run_command())
        except Exception as exc:
            self.root.after(
                0,
                lambda: messagebox.showerror("실행 실패", str(exc)),
            )
            self.root.after(0, lambda: self.status_var.set("실행 시작에 실패했습니다."))
            return

        self.root.after(
            0,
            lambda: self.status_var.set("실행을 시작했습니다. 결과는 로그 폴더에서 확인하세요."),
        )

    def open_log_folder(self):
        log_dir = get_log_dir()
        log_dir.mkdir(parents=True, exist_ok=True)
        os.startfile(log_dir)


def show_ui():
    root = tk.Tk()
    GmarketCrawlerApp(root)
    root.mainloop()


def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument("--run", action="store_true", help="저장된 설정으로 크롤러 실행")
    args = parser.parse_args(argv)

    if args.run:
        try:
            run_crawler_from_config()
        except Exception:
            return 1
        return 0

    show_ui()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
