# -- coding: utf-8 --
###############################################################
# AsahiSingleOrderLauncher_DnD.py
# Windows / Python 3.10以上 / pip install pywin32
# 既存Cmdに必要なopenpyxl・マスター・テンプレートは従来どおり配置します。
###############################################################
"""用途を選んで既存Cmdを呼び出す発注作業の入口。

数量比較・履歴管理・帳票作成は既存Cmdに委譲します。
外部処理とExcel終了待機はワーカースレッドで行い、GUIを停止させません。
最終帳票への赤字反映はこのプログラムの対象外です。
"""

import ctypes
import hashlib
import os
import queue
import re
import subprocess
import sys
import threading
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

HIGHLIGHTER_CMD = "AsahiSingleOrderQuantityChangeHighlighter_Cmd.py"
SELECTOR_CMD = "AsahiSingleOrderProductCodeSelector_Cmd.py"
PREPARED_ORDER_ENV = "ASAHI_ORDER_HIGHLIGHTER_PREPARED_ORDER"
EXCEL_PATH_ENV = "ASAHI_LAUNCHER_EXCEL_PATH"
WINDOW_TITLE = "Asahi 発注作業"
MODE_HIGHLIGHT = 1
MODE_REPORT = 2
MODE_BOTH = 3
REVIEW_STATES = ("review_ready", "review_changed")


class WorkflowError(Exception):
    """後続処理へ進まず、利用者に案内する必要がある状態。"""


@dataclass(frozen=True)
class FileStamp:
    size: int
    modified_ns: int
    sha256: str


@dataclass(frozen=True)
class CommandResult:
    returncode: int
    stdout: str = ""
    stderr: str = ""

    @property
    def details(self) -> str:
        return "\n".join(text.strip() for text in (self.stdout, self.stderr) if text.strip())


def validate_input(path: Path) -> Path:
    path = path.expanduser().resolve()
    if path.suffix.lower() != ".xlsx" or not path.is_file():
        raise WorkflowError("存在するXLSXファイルを1つ指定してください。")
    return path


def file_stamp(path: Path) -> FileStamp:
    """再保存も検出。読取り途中に変化したファイルは確定済みと扱いません。"""
    before = path.stat()
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    after = path.stat()
    if (before.st_size, before.st_mtime_ns) != (after.st_size, after.st_mtime_ns):
        raise WorkflowError("確認中に入力Excelが更新されました。保存して閉じ、再度確認してください。")
    return FileStamp(after.st_size, after.st_mtime_ns, digest.hexdigest())


class Services:
    """同じフォルダーのCmdと、WindowsのExcelだけを操作します。"""

    def __init__(self, script_directory: Path | None = None):
        self.script_directory = (script_directory or Path(__file__).resolve().parent).resolve()

    def check_commands(self, mode: int) -> None:
        names = []
        if mode in (MODE_HIGHLIGHT, MODE_BOTH):
            names.append(HIGHLIGHTER_CMD)
        if mode in (MODE_REPORT, MODE_BOTH):
            names.append(SELECTOR_CMD)
        for name in names:
            if not (self.script_directory / name).is_file():
                raise WorkflowError(name + " が見つかりません。統合プログラムと同じフォルダーに配置してください。")

    def run_cmd(self, name: str, path: Path, prepare: bool = False,
                prepared_order: str | None = None) -> CommandResult:
        command = [sys.executable, str(self.script_directory / name)]
        if prepare:
            command.append("--prepare")
        command.append(str(path))
        environment = {**os.environ, "PYTHONIOENCODING": "utf-8"}
        environment.pop(PREPARED_ORDER_ENV, None)
        if prepared_order is not None:
            environment[PREPARED_ORDER_ENV] = prepared_order
        result = subprocess.run(
            command, cwd=self.script_directory, env=environment,
            check=False, capture_output=True, text=True, encoding="utf-8", errors="replace",
        )
        return CommandResult(result.returncode, result.stdout, result.stderr)

    @contextmanager
    def input_guard(self, path: Path):
        """Excelの書込みハンドルを検出し、帳票作成中の上書き・置換を防ぎます。

        読取り共有だけを許可するため、既存Cmdは元のパスを読み取れます。
        コピーを渡さず、従来どおりの出力先・ファイル名を維持します。
        """
        if path.with_name("~$" + path.name).exists():
            raise WorkflowError("入力Excelが開かれています。上書き保存してExcelを閉じてください。")
        if os.name != "nt":
            raise WorkflowError("Excelの終了確認にはWindowsが必要です。")
        from ctypes import wintypes
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        create_file = kernel32.CreateFileW
        create_file.argtypes = (wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                                ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE)
        create_file.restype = wintypes.HANDLE
        close_handle = kernel32.CloseHandle
        close_handle.argtypes = (wintypes.HANDLE,)
        close_handle.restype = wintypes.BOOL
        # GENERIC_READ / FILE_SHARE_READ / OPEN_EXISTING / FILE_ATTRIBUTE_NORMAL
        handle = create_file(str(path), 0x80000000, 1, None, 3, 0x80, None)
        if handle == ctypes.c_void_p(-1).value:
            error = ctypes.get_last_error()
            raise WorkflowError(
                "入力Excelを読取り用に確保できません。保存してExcelを閉じ、アクセス権も確認してください。"
                + "\nWindowsエラー: " + str(error)
            )
        try:
            yield
        finally:
            close_handle(handle)

    def ensure_closed(self, path: Path) -> None:
        with self.input_guard(path):
            pass

    def stamp(self, path: Path) -> FileStamp:
        return file_stamp(path)

    def open_excel(self, path: Path, wait: bool) -> None:
        environment = {**os.environ, EXCEL_PATH_ENV: str(path)}
        script = "$ErrorActionPreference = 'Stop'; $excelPath = $env:" + EXCEL_PATH_ENV + "; "
        if wait:
            # /xは既存Excelへの受渡しを避け、待機対象を専用プロセスにします。
            script += (
                "$excelExe = $null; "
                "foreach ($key in @('Registry::HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\excel.exe', "
                "'Registry::HKEY_LOCAL_MACHINE\\Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\excel.exe', "
                "'Registry::HKEY_LOCAL_MACHINE\\Software\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\App Paths\\excel.exe')) { "
                "if (Test-Path -LiteralPath $key) { $excelExe = (Get-Item -LiteralPath $key).GetValue(''); break } }; "
                "if (-not $excelExe) { $excelExe = (Get-Command excel.exe -ErrorAction Stop).Source }; "
                "$quotedPath = [char]34 + $excelPath + [char]34; "
                "$process = Start-Process -FilePath $excelExe -ArgumentList @('/x', $quotedPath) -PassThru; "
                "$process.WaitForExit(); exit $process.ExitCode"
            )
        else:
            script += "Start-Process -FilePath $excelPath"
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-Command", script], env=environment,
            check=False, capture_output=True, text=True, errors="replace",
        )
        if result.returncode != 0:
            raise WorkflowError("Excelの起動または終了待機に失敗しました。\n" + result.stderr.strip())


class Workflow:
    """GUIとは独立した処理順序。notifyはGUIのキューへ通知するだけです。"""

    def __init__(self, services: Services, notify: Callable[[str, object], None]):
        self.services = services
        self.notify = notify
        self.state = "idle"
        self.path: Path | None = None
        self.mode = 0
        self.review_stamp: FileStamp | None = None

    @property
    def engaged(self) -> bool:
        return self.state == "running" or self.state in REVIEW_STATES

    def set_state(self, state: str) -> None:
        self.state = state
        self.notify("state", state)

    def status(self, message: str) -> None:
        self.notify("status", message)

    def result(self, state: str, message: str) -> None:
        self.set_state(state)
        self.notify("result", (state, message))

    def execute(self, action: Callable[[], None]) -> None:
        try:
            action()
        except Exception as exception:
            self.result("failed", str(exception))

    def begin(self, path: Path, mode: int) -> None:
        if self.engaged:
            return
        self.set_state("running")
        self.execute(lambda: self._begin(path, mode))

    def _begin(self, path: Path, mode: int) -> None:
        if mode not in (MODE_HIGHLIGHT, MODE_REPORT, MODE_BOTH):
            raise WorkflowError("実行する用途を選んでください。")
        self.path = validate_input(path)
        self.mode = mode
        self.review_stamp = None
        self.services.check_commands(mode)
        if mode == MODE_REPORT:
            self.create_reports()
        else:
            self.highlight()

    def checked_highlighter(self, prepare: bool, token: str | None = None) -> CommandResult:
        result = self.services.run_cmd(HIGHLIGHTER_CMD, self.path, prepare, token)
        # 注文識別トークンは連携用の内部情報なので、利用者向け結果欄には出しません。
        details = "\n".join(line for line in result.details.splitlines()
                            if not line.startswith("Prepared order: "))
        if details:
            self.notify("log", details)
        if result.returncode != 0:
            raise WorkflowError("数量処理に失敗しました。帳票作成には進みません。\n" + result.details)
        return result

    def highlight(self) -> None:
        self.status("編集前の履歴を確認しています。")
        self.services.ensure_closed(self.path)
        prepared = self.checked_highlighter(True)
        tokens = re.findall(r"^Prepared order: ([0-9a-f]{64})\s*$", prepared.stdout, re.MULTILINE)
        if len(tokens) != 1:
            raise WorkflowError("編集前の注文識別情報を確認できません。処理を中止します。")
        self.status("Excelで数量を編集し、上書き保存してExcelを閉じてください。")
        self.services.open_excel(self.path, wait=True)
        self.services.ensure_closed(self.path)
        self.status("数量を比較し、赤字と履歴を保存しています。")
        self.checked_highlighter(False, tokens[0])
        if self.mode == MODE_HIGHLIGHT:
            self.status("赤字化が完了しました。結果確認用のExcelを開いています。")
            try:
                self.services.open_excel(self.path, wait=False)
            except Exception as exception:
                raise WorkflowError("数量の赤字化・履歴保存は完了しましたが、結果のExcelを開けませんでした。"
                                    "\nExcelを手動で開いてください。\n" + str(exception)) from exception
            self.result("complete", "数量の比較・赤字化が完了しました。\n"
                        "さらに修正した場合は、保存して閉じた後、再度数量修正を実行してください。")
            return
        with self.services.input_guard(self.path):
            self.review_stamp = self.services.stamp(self.path)
        self.status("赤字化したExcelを確認し、Excelを閉じてください。追加編集した場合は再比較が必要です。")
        self.services.open_excel(self.path, wait=True)
        self.review_status()

    def review_status(self) -> None:
        with self.services.input_guard(self.path):
            changed = self.services.stamp(self.path) != self.review_stamp
        if changed:
            self.set_state("review_changed")
            self.status("確認中にExcelが更新されました。［数量修正をやり直す］で再比較してください。")
        else:
            self.set_state("review_ready")
            self.status("内容確認が完了しました。［この内容で帳票を作成］を選べます。")

    def confirm_reports(self) -> None:
        if self.state != "review_ready":
            return
        self.set_state("running")
        self.execute(lambda: self.create_reports(self.review_stamp))

    def restart(self) -> None:
        if self.state not in REVIEW_STATES:
            return
        self.set_state("running")
        self.execute(self.highlight)

    def finish_review(self) -> None:
        if self.state in REVIEW_STATES:
            self.result("complete", "帳票を作成せず終了しました。数量修正・赤字・履歴は保存済みです。")

    def create_reports(self, expected: FileStamp | None = None) -> None:
        self.services.check_commands(MODE_REPORT)
        # 検査後も読取り共有のハンドルを維持し、確認後の外部更新を防ぎます。
        with self.services.input_guard(self.path):
            if expected is not None and self.services.stamp(self.path) != expected:
                self.set_state("review_changed")
                self.status("確認後にExcelが更新されました。数量修正をやり直してください。")
                return
            self.status("帳票を作成しています。商品コード選択画面が表示されたら商品を選択してください。")
            result = self.services.run_cmd(SELECTOR_CMD, self.path)
        if result.details:
            self.notify("log", result.details)
        if result.returncode == 0:
            self.result("complete", "帳票の作成が完了しました。\n" + result.details)
        elif result.returncode == 2:
            self.result("cancelled", "商品コードの選択がキャンセルされました。保存済みの数量修正・履歴は残ります。\n"
                        + result.details)
        elif result.returncode == 3:
            self.result("not_found", "該当商品がありません。保存済みの数量修正・履歴は残ります。\n" + result.details)
        else:
            raise WorkflowError("帳票作成に失敗しました。保存済みの数量修正・履歴は残ります。\n" + result.details)


class WindowsDropTarget:
    """Tkの最上位WindowsウィンドウでWM_DROPFILESを受け付けます。"""

    def __init__(self, root, accept: Callable[[list[str]], None]):
        import win32api
        import win32con
        import win32gui
        self.gui = win32gui
        self.con = win32con
        self.api = win32api
        self.accept = accept
        self.closed = False
        root.update_idletasks()
        self.handle = win32gui.GetAncestor(root.winfo_id(), 2)  # GA_ROOT
        self.previous = win32gui.SetWindowLong(self.handle, win32con.GWL_WNDPROC, self.window_proc)
        win32gui.DragAcceptFiles(self.handle, True)

    def window_proc(self, handle, message, wparam, lparam):
        if message == self.con.WM_DROPFILES:
            try:
                count = self.api.DragQueryFile(wparam, -1)
                self.accept([self.api.DragQueryFile(wparam, i) for i in range(count)])
            finally:
                self.api.DragFinish(wparam)
            return 0
        return self.gui.CallWindowProc(self.previous, handle, message, wparam, lparam)

    def restore_window_proc(self):
        """pywin32が返した元のWNDPROCポインターを、ネイティブAPIで戻します。"""
        if callable(self.previous) or isinstance(self.previous, dict):
            self.gui.SetWindowLong(self.handle, self.con.GWL_WNDPROC, self.previous)
            return
        from ctypes import wintypes
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        if ctypes.sizeof(ctypes.c_void_p) == 8:
            set_window_long = user32.SetWindowLongPtrW
            value_type = ctypes.c_ssize_t
        else:
            set_window_long = user32.SetWindowLongW
            value_type = wintypes.LONG
        set_window_long.argtypes = (wintypes.HWND, ctypes.c_int, value_type)
        set_window_long.restype = value_type
        ctypes.set_last_error(0)
        previous = set_window_long(self.handle, self.con.GWL_WNDPROC, self.previous)
        error = ctypes.get_last_error()
        if previous == 0 and error != 0:
            raise OSError(error, "元のWindowsウィンドウ処理を復元できませんでした。")

    def close(self):
        if self.closed:
            return
        if not self.gui.IsWindow(self.handle):
            self.closed = True
            return
        try:
            self.gui.DragAcceptFiles(self.handle, False)
        finally:
            # 受付停止に失敗しても、元のウィンドウ処理の復元は試みます。
            self.restore_window_proc()
            self.closed = True


class LauncherApp:
    """Tkメインスレッドだけが画面を操作し、外部処理とはキューで連携。"""

    def __init__(self, root, services: Services | None = None, install_drop: bool = True):
        import tkinter as tk
        from tkinter import messagebox, ttk
        self.root = root
        self.dialogs = messagebox
        self.events = queue.Queue()
        self.workflow = Workflow(services or Services(), lambda kind, value: self.events.put((kind, value)))
        self.worker_busy = False
        self.selected_path: Path | None = None
        self.closed = False
        self.closing = False
        self.drop_target = None
        self.mode_var = tk.IntVar(value=0)
        self.file_var = tk.StringVar(value="ここに入力Excelを1つドラッグ＆ドロップしてください。")
        self.status_var = tk.StringVar(value="入力Excelと用途を選んでください。")
        root.title(WINDOW_TITLE)
        root.geometry("790x690")
        root.minsize(690, 620)
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        frame = ttk.Frame(root, padding=18)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=WINDOW_TITLE, font=("Yu Gothic UI", 18, "bold")).pack(anchor="w")
        ttk.Label(frame, text="Excelをドロップし、用途を選んで［実行］を押してください。", wraplength=720).pack(anchor="w", pady=(4, 14))
        drop_frame = ttk.LabelFrame(frame, text="入力Excel（XLSX・1ファイル）", padding=12)
        drop_frame.pack(fill="x")
        ttk.Label(drop_frame, textvariable=self.file_var, wraplength=700).pack(anchor="w")
        purposes = ttk.LabelFrame(frame, text="用途", padding=10)
        purposes.pack(fill="x", pady=12)
        self.radios = []
        for value, label in (
            (MODE_HIGHLIGHT, "数量を修正して赤字にする"),
            (MODE_REPORT, "現在の内容で帳票を作成する（Excelを保存して閉じてください）"),
            (MODE_BOTH, "数量を修正してから帳票を作成する"),
        ):
            button = ttk.Radiobutton(purposes, text=label, variable=self.mode_var, value=value,
                                     command=self.update_controls)
            button.pack(anchor="w", pady=4)
            self.radios.append(button)
        buttons = ttk.Frame(frame)
        buttons.pack(fill="x")
        self.run_button = ttk.Button(buttons, text="実行", command=self.start)
        self.run_button.pack(side="left")
        self.clear_button = ttk.Button(buttons, text="入力を解除", command=self.clear_input)
        self.clear_button.pack(side="left", padx=8)
        self.review_frame = ttk.LabelFrame(frame, text="数量の比較・赤字化が完了しました", padding=10)
        self.confirm_button = ttk.Button(self.review_frame, text="この内容で帳票を作成", command=self.confirm)
        self.confirm_button.pack(side="left")
        self.restart_button = ttk.Button(self.review_frame, text="数量修正をやり直す", command=self.restart)
        self.restart_button.pack(side="left", padx=8)
        self.finish_button = ttk.Button(self.review_frame, text="ここで終了", command=self.finish_review)
        self.finish_button.pack(side="left")
        self.status_label = ttk.Label(frame, textvariable=self.status_var, wraplength=720)
        self.status_label.pack(fill="x", pady=14)
        log_frame = ttk.LabelFrame(frame, text="処理結果", padding=6)
        log_frame.pack(fill="both", expand=True)
        scrollbar = ttk.Scrollbar(log_frame)
        scrollbar.pack(side="right", fill="y")
        self.log = tk.Text(log_frame, height=11, wrap="word", state="disabled", yscrollcommand=scrollbar.set)
        self.log.pack(fill="both", expand=True)
        scrollbar.configure(command=self.log.yview)
        ttk.Label(frame, text="最終帳票への赤字反映は、この統合プログラムの対象外です。", wraplength=720).pack(anchor="w", pady=(10, 0))
        if install_drop:
            self.drop_target = WindowsDropTarget(root, self.accept_files)
        self.update_controls()
        self.timer_id = root.after(80, self.drain_events)

    def append_log(self, message: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", message.rstrip() + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def update_controls(self) -> None:
        active = self.worker_busy or self.workflow.engaged
        review = self.workflow.state in REVIEW_STATES
        for radio in self.radios:
            radio.configure(state="disabled" if active else "normal")
        self.run_button.configure(state="normal" if not active and self.selected_path is not None
                                  and self.mode_var.get() in (1, 2, 3) else "disabled")
        self.clear_button.configure(state="normal" if not active and self.selected_path is not None else "disabled")
        if review:
            if not self.review_frame.winfo_manager():
                self.review_frame.pack(fill="x", pady=(12, 0), before=self.status_label)
        else:
            self.review_frame.pack_forget()
        self.confirm_button.configure(state="normal" if not self.worker_busy and self.workflow.state == "review_ready" else "disabled")
        for button in (self.restart_button, self.finish_button):
            button.configure(state="normal" if review and not self.worker_busy else "disabled")

    def accept_files(self, files: list[str]) -> None:
        if self.worker_busy or self.workflow.engaged:
            self.status_var.set("処理・内容確認中です。現在の作業を終えてから入力を変更してください。")
            return
        if len(files) != 1:
            self.dialogs.showerror(WINDOW_TITLE, "入力Excelを1つだけドロップしてください。", parent=self.root)
            return
        try:
            path = validate_input(Path(files[0]))
        except Exception as exception:
            self.dialogs.showerror(WINDOW_TITLE, str(exception), parent=self.root)
            return
        self.selected_path = path
        self.file_var.set(str(path))
        self.mode_var.set(0)
        self.status_var.set("用途を選んで［実行］を押してください。まだ処理は開始していません。")
        self.update_controls()

    def clear_input(self) -> None:
        if self.worker_busy or self.workflow.engaged:
            return
        self.selected_path = None
        self.mode_var.set(0)
        self.file_var.set("ここに入力Excelを1つドラッグ＆ドロップしてください。")
        self.status_var.set("入力Excelと用途を選んでください。")
        self.update_controls()

    def submit(self, action: Callable[[], None]) -> None:
        if self.worker_busy:
            return
        self.worker_busy = True
        self.update_controls()
        def work():
            try:
                action()
            except Exception as exception:
                self.workflow.result("failed", str(exception))
            finally:
                self.events.put(("worker_finished", None))
        # 強制終了を避けるため、プロセス終了で切られるdaemonにはしません。
        threading.Thread(target=work, daemon=False).start()

    def start(self) -> None:
        if self.worker_busy or self.workflow.engaged or self.selected_path is None or self.mode_var.get() not in (1, 2, 3):
            return
        path, mode = self.selected_path, self.mode_var.get()
        self.submit(lambda: self.workflow.begin(path, mode))

    def confirm(self) -> None:
        if self.workflow.state == "review_ready":
            self.submit(self.workflow.confirm_reports)

    def restart(self) -> None:
        if self.workflow.state in REVIEW_STATES:
            self.submit(self.workflow.restart)

    def finish_review(self) -> None:
        if not self.worker_busy:
            self.workflow.finish_review()

    def drain_events(self) -> None:
        if self.closed:
            return
        while True:
            try:
                kind, value = self.events.get_nowait()
            except queue.Empty:
                break
            if kind == "worker_finished":
                self.worker_busy = False
            elif kind == "status":
                self.status_var.set(value)
                self.append_log(value)
            elif kind == "log":
                self.append_log(value)
            elif kind == "result":
                state, message = value
                self.status_var.set({"complete": "完了しました。", "failed": "処理に失敗しました。",
                                     "cancelled": "キャンセルされました。", "not_found": "該当商品がありません。"}[state])
                self.append_log(message)
                # 長い出力先一覧は結果欄へ残し、ダイアログは要点だけを表示します。
                dialog = self.dialogs.showerror if state == "failed" else self.dialogs.showinfo
                dialog(WINDOW_TITLE, message[:900], parent=self.root)
        self.update_controls()
        self.timer_id = self.root.after(80, self.drain_events)

    def on_close(self) -> None:
        if self.closed or self.closing:
            return
        self.closing = True
        try:
            if self.worker_busy or self.workflow.state == "running":
                self.dialogs.showinfo(WINDOW_TITLE, "処理中は終了できません。Excelや商品選択画面の操作を終え、処理の完了を待ってください。",
                                      parent=self.root)
                return
            if self.workflow.state in REVIEW_STATES:
                if not self.dialogs.askyesno(WINDOW_TITLE, "帳票を作成せず終了しますか？数量修正・赤字・履歴は保存済みです。", parent=self.root):
                    return
            self.closed = True
            errors = []
            try:
                if self.drop_target is not None:
                    self.drop_target.close()
            except Exception as exception:
                errors.append(exception)
            try:
                if self.timer_id is not None:
                    self.root.after_cancel(self.timer_id)
                    self.timer_id = None
            except Exception as exception:
                errors.append(exception)
            finally:
                # 後片付けの一部が失敗しても、操作不能の画面を残しません。
                try:
                    self.root.destroy()
                finally:
                    self.root.quit()
                    for exception in errors:
                        print("終了時の後片付けで問題が発生しました: " + str(exception), file=sys.stderr)
        finally:
            self.closing = False


def main() -> None:
    if os.name != "nt":
        print("このプログラムはWindowsのExcel用です。", file=sys.stderr)
        return
    import tkinter as tk
    from tkinter import messagebox
    root = tk.Tk()
    try:
        LauncherApp(root)
    except Exception as exception:
        messagebox.showerror(WINDOW_TITLE, "起動できません。pywin32の導入と配置を確認してください。\n" + str(exception), parent=root)
        root.destroy()
        return
    root.mainloop()


if __name__ == "__main__":
    main()
