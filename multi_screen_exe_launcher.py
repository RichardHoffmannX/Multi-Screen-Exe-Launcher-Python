import sys
import os
import json
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import win32api
    import win32con
    import win32gui
    import win32process
except ImportError:
    print("ERROR: pywin32 is not installed.", flush=True)
    print("Install it with: pip install pywin32", flush=True)
    sys.exit(1)


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
HISTORY_FILE = os.path.join(SCRIPT_DIR, "exe_launcher_history.json")


class ExeLauncherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Multi-Screen EXE Launcher")
        self.root.geometry("700x700")
        self.root.minsize(640, 620)

        # Bring launcher to front on startup
        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.after(1000, lambda: self.root.attributes("-topmost", False))

        self.history = self.load_history()

        self.process = None
        self.search_attempts = 0
        self.max_search_attempts = 60

        self.monitors = []

        self.create_widgets()
        self.refresh_monitors()
        self.populate_history_dropdown()

    # ----------------------------------------------------------
    # UI
    # ----------------------------------------------------------

    def create_widgets(self):
        # Executable
        file_frame = ttk.LabelFrame(
            self.root,
            text=" Select Executable ",
            padding=10
        )
        file_frame.pack(fill="x", padx=15, pady=8)

        self.path_entry = ttk.Entry(file_frame)
        self.path_entry.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(0, 5)
        )

        ttk.Button(
            file_frame,
            text="Browse...",
            command=self.browse_file
        ).pack(side="right")

        # Window title
        title_frame = ttk.LabelFrame(
            self.root,
            text=" Target Window Title / Search Text ",
            padding=10
        )
        title_frame.pack(fill="x", padx=15, pady=5)

        self.title_entry = ttk.Entry(title_frame)
        self.title_entry.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(0, 5)
        )

        ttk.Button(
            title_frame,
            text="Get Running Window Title",
            command=self.get_running_window_title
        ).pack(side="right")

        ttk.Label(
            self.root,
            text="The launcher primarily detects the launched window by process ID. "
                 "The title can also be used as an additional hint.",
            wraplength=650
        ).pack(anchor="w", padx=20, pady=(0, 5))

        # Monitor selection
        monitor_frame = ttk.LabelFrame(
            self.root,
            text=" Target Monitor ",
            padding=10
        )
        monitor_frame.pack(fill="x", padx=15, pady=5)

        self.monitor_cb = ttk.Combobox(
            monitor_frame,
            state="readonly"
        )
        self.monitor_cb.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(0, 5)
        )

        ttk.Button(
            monitor_frame,
            text="Refresh",
            command=self.refresh_monitors
        ).pack(side="right")

        # Display mode
        options_frame = ttk.LabelFrame(
            self.root,
            text=" Window Display Mode ",
            padding=10
        )
        options_frame.pack(fill="x", padx=15, pady=5)

        self.display_mode = tk.StringVar(value="maximized")

        ttk.Radiobutton(
            options_frame,
            text="Standard Windowed",
            variable=self.display_mode,
            value="windowed"
        ).pack(anchor="w", pady=2)

        ttk.Radiobutton(
            options_frame,
            text="Fixed Size Window - 1280 x 720",
            variable=self.display_mode,
            value="fixed"
        ).pack(anchor="w", pady=2)

        ttk.Radiobutton(
            options_frame,
            text="Maximized Window",
            variable=self.display_mode,
            value="maximized"
        ).pack(anchor="w", pady=2)

        ttk.Radiobutton(
            options_frame,
            text="Borderless Fullscreen",
            variable=self.display_mode,
            value="fullscreen"
        ).pack(anchor="w", pady=2)

        # History
        history_frame = ttk.LabelFrame(
            self.root,
            text=" History / Previous Apps ",
            padding=10
        )
        history_frame.pack(fill="x", padx=15, pady=8)

        self.history_cb = ttk.Combobox(
            history_frame,
            state="readonly"
        )
        self.history_cb.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(0, 5)
        )

        self.history_cb.bind(
            "<<ComboboxSelected>>",
            self.on_history_select
        )

        ttk.Button(
            history_frame,
            text="Delete Entry",
            command=self.delete_current_history
        ).pack(side="right")

        # Launch
        btn_frame = ttk.Frame(self.root, padding=10)
        btn_frame.pack(fill="x", padx=15, pady=5)

        self.launch_btn = ttk.Button(
            btn_frame,
            text="Launch on Selected Screen",
            command=self.launch_app
        )
        self.launch_btn.pack(fill="x", ipady=7)

        # Save location
        link_label = tk.Label(
            self.root,
            text="Show Save File Location",
            fg="blue",
            cursor="hand2",
            font=("TkDefaultFont", 9, "underline")
        )

        link_label.pack(pady=5)

        link_label.bind(
            "<Button-1>",
            lambda event: self.show_file_location()
        )

        # Status
        self.status_label = ttk.Label(
            self.root,
            text="Ready",
            relief="sunken",
            anchor="w"
        )
        self.status_label.pack(
            side="bottom",
            fill="x"
        )

    # ----------------------------------------------------------
    # RUNNING WINDOW TITLE PICKER
    # ----------------------------------------------------------

    def get_running_window_title(self):
        windows = []

        def enum_windows_callback(hwnd, result):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return

                title = win32gui.GetWindowText(hwnd).strip()

                if not title:
                    return

                # Ignore our own launcher
                if title == self.root.title():
                    return

                rect = win32gui.GetWindowRect(hwnd)

                width = rect[2] - rect[0]
                height = rect[3] - rect[1]

                # Ignore tiny/helper windows
                if width < 100 or height < 100:
                    return

                _, pid = win32process.GetWindowThreadProcessId(hwnd)

                result.append({
                    "hwnd": hwnd,
                    "title": title,
                    "pid": pid
                })

            except Exception:
                pass

        win32gui.EnumWindows(
            enum_windows_callback,
            windows
        )

        windows.sort(key=lambda item: item["title"].lower())

        if not windows:
            messagebox.showinfo(
                "Windows",
                "No running application windows were found."
            )
            return

        self.show_window_picker(windows)

    def show_window_picker(self, windows):
        picker = tk.Toplevel(self.root)

        picker.title("Select Running Window")
        picker.geometry("720x450")
        picker.minsize(520, 320)

        picker.transient(self.root)
        picker.grab_set()

        ttk.Label(
            picker,
            text="Select a running application window:"
        ).pack(
            anchor="w",
            padx=10,
            pady=(10, 5)
        )

        listbox_frame = ttk.Frame(picker)
        listbox_frame.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=5
        )

        scrollbar = ttk.Scrollbar(
            listbox_frame,
            orient="vertical"
        )

        listbox = tk.Listbox(
            listbox_frame,
            yscrollcommand=scrollbar.set
        )

        scrollbar.config(
            command=listbox.yview
        )

        scrollbar.pack(
            side="right",
            fill="y"
        )

        listbox.pack(
            side="left",
            fill="both",
            expand=True
        )

        for window in windows:
            listbox.insert(
                tk.END,
                f'{window["title"]}    [PID {window["pid"]}]'
            )

        def select_window():
            selection = listbox.curselection()

            if not selection:
                return

            index = selection[0]

            selected_window = windows[index]

            title = selected_window["title"]

            self.title_entry.delete(
                0,
                tk.END
            )

            self.title_entry.insert(
                0,
                title
            )

            self.status_label.config(
                text=f'Window title selected: "{title}"'
            )

            picker.destroy()

        button_frame = ttk.Frame(
            picker
        )

        button_frame.pack(
            fill="x",
            padx=10,
            pady=10
        )

        ttk.Button(
            button_frame,
            text="Use Selected Window",
            command=select_window
        ).pack(
            side="left"
        )

        ttk.Button(
            button_frame,
            text="Cancel",
            command=picker.destroy
        ).pack(
            side="right"
        )

        listbox.bind(
            "<Double-Button-1>",
            lambda event: select_window()
        )

        if windows:
            listbox.selection_set(0)
            listbox.focus_set()

    # ----------------------------------------------------------
    # MONITOR HANDLING
    # ----------------------------------------------------------

    def refresh_monitors(self):
        self.monitors = []

        try:
            raw_monitors = win32api.EnumDisplayMonitors()

            for monitor_data in raw_monitors:
                monitor_handle = monitor_data[0]

                info = win32api.GetMonitorInfo(monitor_handle)

                monitor_rect = info["Monitor"]
                work_rect = info["Work"]

                is_primary = bool(
                    info.get("Flags", 0) & win32con.MONITORINFOF_PRIMARY
                )

                monitor = {
                    "handle": monitor_handle,
                    "monitor": monitor_rect,
                    "work": work_rect,
                    "primary": is_primary
                }

                self.monitors.append(monitor)

        except Exception as e:
            messagebox.showerror(
                "Monitor Error",
                f"Could not enumerate monitors:\n\n{e}"
            )
            return

        values = []

        for index, monitor in enumerate(self.monitors):
            left, top, right, bottom = monitor["monitor"]

            width = right - left
            height = bottom - top

            primary_text = " - Primary" if monitor["primary"] else ""

            values.append(
                f"Monitor {index + 1}: "
                f"{width}x{height} "
                f"({left},{top})"
                f"{primary_text}"
            )

        self.monitor_cb["values"] = values

        if len(values) >= 2:
            self.monitor_cb.current(1)
        elif values:
            self.monitor_cb.current(0)

        self.status_label.config(
            text=f"Detected {len(values)} monitor(s)."
        )

    def get_selected_monitor(self):
        if not self.monitors:
            return None

        index = self.monitor_cb.current()

        if index < 0:
            index = 0

        if index >= len(self.monitors):
            index = 0

        return self.monitors[index]

    # ----------------------------------------------------------
    # FILE SELECTION
    # ----------------------------------------------------------

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Executable File",
            filetypes=[
                ("Executable Files", "*.exe"),
                ("All Files", "*.*")
            ]
        )

        if not file_path:
            return

        self.path_entry.delete(0, tk.END)
        self.path_entry.insert(0, file_path)

        filename = os.path.splitext(
            os.path.basename(file_path)
        )[0]

        self.title_entry.delete(0, tk.END)
        self.title_entry.insert(0, filename)

    # ----------------------------------------------------------
    # HISTORY
    # ----------------------------------------------------------

    def load_history(self):
        if not os.path.exists(HISTORY_FILE):
            return {}

        try:
            with open(
                HISTORY_FILE,
                "r",
                encoding="utf-8"
            ) as f:
                data = json.load(f)

                if isinstance(data, dict):
                    return data

        except Exception as e:
            print(
                f"History load error: {e}",
                flush=True
            )

        return {}

    def save_history(self):
        try:
            with open(
                HISTORY_FILE,
                "w",
                encoding="utf-8"
            ) as f:
                json.dump(
                    self.history,
                    f,
                    indent=4,
                    ensure_ascii=False
                )

        except Exception as e:
            print(
                f"Failed to save history: {e}",
                flush=True
            )

    def populate_history_dropdown(self):
        choices = []

        for title, path in self.history.items():
            choices.append(
                f"{title} ({path})"
            )

        self.history_cb["values"] = choices

        if choices:
            self.history_cb.current(0)
        else:
            self.history_cb.set("")

    def on_history_select(self, event=None):
        selected = self.history_cb.get()

        if not selected:
            return

        try:
            title, path = selected.rsplit(" (", 1)

            path = path[:-1] if path.endswith(")") else path

            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, path)

            self.title_entry.delete(0, tk.END)
            self.title_entry.insert(0, title)

        except Exception as e:
            print(
                f"History selection error: {e}",
                flush=True
            )

    def delete_current_history(self):
        selected = self.history_cb.get()

        if not selected:
            messagebox.showinfo(
                "Delete Entry",
                "No history item selected."
            )
            return

        try:
            title, _ = selected.rsplit(" (", 1)

            if title in self.history:
                del self.history[title]

                self.save_history()
                self.populate_history_dropdown()

                self.status_label.config(
                    text=f"Removed '{title}' from history."
                )

        except Exception as e:
            messagebox.showerror(
                "Error",
                str(e)
            )

    # ----------------------------------------------------------
    # CONFIG FILE
    # ----------------------------------------------------------

    def show_file_location(self):
        if not os.path.exists(HISTORY_FILE):
            self.save_history()

        try:
            subprocess.Popen(
                [
                    "explorer.exe",
                    "/select,",
                    os.path.normpath(HISTORY_FILE)
                ]
            )

        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Could not open directory:\n\n{e}"
            )

    # ----------------------------------------------------------
    # LAUNCH
    # ----------------------------------------------------------

    def launch_app(self):
        exe_path = self.path_entry.get().strip()
        title_hint = self.title_entry.get().strip()

        if not exe_path:
            messagebox.showerror(
                "Error",
                "Please select an executable."
            )
            return

        if not os.path.isfile(exe_path):
            messagebox.showerror(
                "Error",
                "The selected executable does not exist."
            )
            return

        monitor = self.get_selected_monitor()

        if monitor is None:
            messagebox.showerror(
                "Error",
                "No monitor was detected."
            )
            return

        history_name = title_hint

        if not history_name:
            history_name = os.path.splitext(
                os.path.basename(exe_path)
            )[0]

        self.history[history_name] = exe_path
        self.save_history()
        self.populate_history_dropdown()

        self.status_label.config(
            text=f"Launching {os.path.basename(exe_path)}..."
        )

        self.launch_btn.config(state="disabled")

        try:
            working_directory = os.path.dirname(exe_path)

            self.process = subprocess.Popen(
                [exe_path],
                cwd=working_directory
            )

        except Exception as e:
            self.launch_btn.config(state="normal")

            messagebox.showerror(
                "Launch Error",
                f"Could not launch application:\n\n{e}"
            )

            return

        self.search_attempts = 0

        # Don't block Tkinter with time.sleep()
        self.root.after(
            500,
            self.wait_for_window
        )

    # ----------------------------------------------------------
    # WINDOW SEARCHING
    # ----------------------------------------------------------

    def find_process_window(self, process_id):
        candidates = []

        title_hint = self.title_entry.get().strip().lower()

        def enum_handler(hwnd, results):
            if not win32gui.IsWindow(hwnd):
                return

            if not win32gui.IsWindowVisible(hwnd):
                return

            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)

                if pid != process_id:
                    return

                title = win32gui.GetWindowText(hwnd)

                rect = win32gui.GetWindowRect(hwnd)

                width = rect[2] - rect[0]
                height = rect[3] - rect[1]

                # Ignore tiny/helper windows
                if width < 100 or height < 100:
                    return

                results.append(
                    {
                        "hwnd": hwnd,
                        "title": title,
                        "width": width,
                        "height": height
                    }
                )

            except Exception:
                pass

        win32gui.EnumWindows(
            enum_handler,
            candidates
        )

        if not candidates:
            return None

        # Prefer a window whose title matches our hint
        if title_hint:
            for candidate in candidates:
                if title_hint in candidate["title"].lower():
                    return candidate["hwnd"]

        # Otherwise return the largest visible window
        candidates.sort(
            key=lambda item: item["width"] * item["height"],
            reverse=True
        )

        return candidates[0]["hwnd"]

    def wait_for_window(self):
        if self.process is None:
            self.launch_btn.config(state="normal")
            return

        # Process terminated
        if self.process.poll() is not None:
            self.launch_btn.config(state="normal")

            self.status_label.config(
                text="Application exited before a window was detected."
            )

            return

        hwnd = self.find_process_window(
            self.process.pid
        )

        if hwnd:
            self.configure_window(hwnd)
            self.launch_btn.config(state="normal")
            return

        self.search_attempts += 1

        if self.search_attempts >= self.max_search_attempts:
            self.launch_btn.config(state="normal")

            self.status_label.config(
                text="Could not detect the application's main window."
            )

            messagebox.showwarning(
                "Window Not Found",
                "The application started, but its main window "
                "could not be detected.\n\n"
                "Some games launch another process. "
                "If so, additional child-process detection may be needed."
            )

            return

        self.status_label.config(
            text=(
                "Waiting for application window..."
                f" {self.search_attempts}"
            )
        )

        self.root.after(
            250,
            self.wait_for_window
        )

    # ----------------------------------------------------------
    # WINDOW MANAGEMENT
    # ----------------------------------------------------------

    def configure_window(self, hwnd):
        monitor = self.get_selected_monitor()

        if monitor is None:
            return

        mode = self.display_mode.get()

        monitor_left, monitor_top, monitor_right, monitor_bottom = (
            monitor["monitor"]
        )

        work_left, work_top, work_right, work_bottom = (
            monitor["work"]
        )

        monitor_width = monitor_right - monitor_left
        monitor_height = monitor_bottom - monitor_top

        work_width = work_right - work_left
        work_height = work_bottom - work_top

        try:
            win32gui.ShowWindow(
                hwnd,
                win32con.SW_RESTORE
            )

            if mode == "windowed":
                self.restore_window_style(hwnd)

                left, top, right, bottom = win32gui.GetWindowRect(hwnd)

                width = max(640, right - left)
                height = max(480, bottom - top)

                width = min(width, work_width)
                height = min(height, work_height)

                x = work_left + 50
                y = work_top + 50

                win32gui.SetWindowPos(
                    hwnd,
                    win32con.HWND_TOP,
                    x,
                    y,
                    width,
                    height,
                    win32con.SWP_SHOWWINDOW
                )

                self.status_label.config(
                    text="Success: Window moved to selected monitor."
                )

            elif mode == "fixed":
                self.restore_window_style(hwnd)

                width = min(1280, work_width)
                height = min(720, work_height)

                x = work_left + ((work_width - width) // 2)
                y = work_top + ((work_height - height) // 2)

                win32gui.SetWindowPos(
                    hwnd,
                    win32con.HWND_TOP,
                    x,
                    y,
                    width,
                    height,
                    win32con.SWP_SHOWWINDOW
                )

                self.status_label.config(
                    text=(
                        f"Success: {width}x{height} centered "
                        "on selected monitor."
                    )
                )

            elif mode == "maximized":
                self.restore_window_style(hwnd)

                win32gui.SetWindowPos(
                    hwnd,
                    win32con.HWND_TOP,
                    work_left + 50,
                    work_top + 50,
                    800,
                    600,
                    win32con.SWP_SHOWWINDOW
                )

                win32gui.ShowWindow(
                    hwnd,
                    win32con.SW_MAXIMIZE
                )

                self.status_label.config(
                    text="Success: Maximized on selected monitor."
                )

            elif mode == "fullscreen":
                self.make_borderless(hwnd)

                win32gui.ShowWindow(
                    hwnd,
                    win32con.SW_RESTORE
                )

                win32gui.SetWindowPos(
                    hwnd,
                    win32con.HWND_TOP,
                    monitor_left,
                    monitor_top,
                    monitor_width,
                    monitor_height,
                    (
                        win32con.SWP_FRAMECHANGED |
                        win32con.SWP_SHOWWINDOW
                    )
                )

                self.status_label.config(
                    text=(
                        f"Success: Borderless fullscreen "
                        f"{monitor_width}x{monitor_height}."
                    )
                )

            try:
                win32gui.BringWindowToTop(hwnd)
                win32gui.SetForegroundWindow(hwnd)
            except Exception:
                pass

        except Exception as e:
            self.status_label.config(
                text="Failed to configure window."
            )

            messagebox.showerror(
                "Window Error",
                f"Could not configure application window:\n\n{e}"
            )

    # ----------------------------------------------------------
    # WINDOW STYLES
    # ----------------------------------------------------------

    def restore_window_style(self, hwnd):
        style = win32gui.GetWindowLong(
            hwnd,
            win32con.GWL_STYLE
        )

        style |= win32con.WS_CAPTION
        style |= win32con.WS_THICKFRAME
        style |= win32con.WS_MINIMIZEBOX
        style |= win32con.WS_MAXIMIZEBOX
        style |= win32con.WS_SYSMENU

        win32gui.SetWindowLong(
            hwnd,
            win32con.GWL_STYLE,
            style
        )

        win32gui.SetWindowPos(
            hwnd,
            None,
            0,
            0,
            0,
            0,
            (
                win32con.SWP_NOMOVE |
                win32con.SWP_NOSIZE |
                win32con.SWP_NOZORDER |
                win32con.SWP_FRAMECHANGED
            )
        )

    def make_borderless(self, hwnd):
        style = win32gui.GetWindowLong(
            hwnd,
            win32con.GWL_STYLE
        )

        style &= ~win32con.WS_CAPTION
        style &= ~win32con.WS_THICKFRAME
        style &= ~win32con.WS_MINIMIZEBOX
        style &= ~win32con.WS_MAXIMIZEBOX
        style &= ~win32con.WS_SYSMENU

        win32gui.SetWindowLong(
            hwnd,
            win32con.GWL_STYLE,
            style
        )

        win32gui.SetWindowPos(
            hwnd,
            None,
            0,
            0,
            0,
            0,
            (
                win32con.SWP_NOMOVE |
                win32con.SWP_NOSIZE |
                win32con.SWP_NOZORDER |
                win32con.SWP_FRAMECHANGED
            )
        )


# --------------------------------------------------------------
# MAIN
# --------------------------------------------------------------

if __name__ == "__main__":
    root = tk.Tk()
    app = ExeLauncherApp(root)
    root.mainloop()
