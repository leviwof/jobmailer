from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from brevo_mailer import BrevoMailer, MailerOptions
from config import BASE_DIR, AppConfig, load_app_config, save_json_config
from excel_reader import (
    load_recipients,
    load_recipients_from_report,
    search_recipients,
)
from logger import setup_app_logging


class JobMailerApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.app_config = load_app_config()
        self.sent_logger, self.failed_logger = setup_app_logging(
            self.app_config.path("logs_dir")
        )
        self.pause_event = threading.Event()
        self.stop_event = threading.Event()
        self.worker: threading.Thread | None = None
        self.events: queue.Queue[dict[str, object]] = queue.Queue()

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.title("Brevo Job Mailer")
        self.geometry("1120x760")
        self.minsize(1000, 680)

        self.excel_path = tk.StringVar(
            value=self._display_config_path("excel_path")
        )
        self.resume_path = tk.StringVar(value=str(self.app_config.path("resume_path")))
        self.logo_path = tk.StringVar(value=str(self.app_config.values.get("logo_path", "")))
        self.subject = tk.StringVar(value=str(self.app_config.values.get("subject", "")))
        self.search_query = tk.StringVar(value="")
        self.schedule_time = tk.StringVar(
            value=str(self.app_config.values.get("schedule_start_time", ""))
        )
        self.daily_limit = tk.StringVar(
            value=(
                "Unlimited"
                if self.app_config.values.get("daily_limit") is None
                else str(self.app_config.values.get("daily_limit", 300))
            )
        )
        self.html_enabled = tk.BooleanVar(
            value=self.app_config.get_bool("html_email", True)
        )
        self.status_text = tk.StringVar(value="Ready")
        self.send_policy_text = tk.StringVar(value=self._send_policy_text())
        self.current_name = tk.StringVar(value="-")
        self.current_company = tk.StringVar(value="-")
        self.current_email = tk.StringVar(value="-")
        self.success_count = tk.StringVar(value="0")
        self.failed_count = tk.StringVar(value="0")
        self.skipped_count = tk.StringVar(value="0")
        self.eta = tk.StringVar(value="-")

        self._build_ui()
        self.after(250, self._poll_events)

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        sidebar = ctk.CTkFrame(self, width=280, corner_radius=0)
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)

        title = ctk.CTkLabel(
            sidebar,
            text="Brevo Job Mailer",
            font=ctk.CTkFont(size=24, weight="bold"),
        )
        title.pack(padx=22, pady=(26, 6), anchor="w")
        subtitle = ctk.CTkLabel(
            sidebar,
            text="Personalized job applications with safe retry and reporting.",
            wraplength=230,
            justify="left",
            text_color=("gray25", "gray75"),
        )
        subtitle.pack(padx=22, pady=(0, 24), anchor="w")
        ctk.CTkLabel(
            sidebar,
            textvariable=self.send_policy_text,
            wraplength=230,
            justify="left",
            text_color=("gray25", "gray72"),
        ).pack(padx=22, pady=(0, 18), anchor="w")

        self.start_button = ctk.CTkButton(
            sidebar,
            text="Start",
            height=42,
            command=self.start_sending,
        )
        self.start_button.pack(padx=22, pady=(0, 10), fill="x")

        self.pause_button = ctk.CTkButton(
            sidebar,
            text="Pause",
            height=38,
            command=self.pause_sending,
            state="disabled",
        )
        self.pause_button.pack(padx=22, pady=6, fill="x")

        self.resume_button = ctk.CTkButton(
            sidebar,
            text="Resume",
            height=38,
            command=self.resume_sending,
            state="disabled",
        )
        self.resume_button.pack(padx=22, pady=6, fill="x")

        self.stop_button = ctk.CTkButton(
            sidebar,
            text="Stop",
            height=38,
            fg_color="#8b1e2d",
            hover_color="#a3283a",
            command=self.stop_sending,
            state="disabled",
        )
        self.stop_button.pack(padx=22, pady=6, fill="x")

        ctk.CTkButton(
            sidebar,
            text="Preview Email",
            height=38,
            command=self.preview_email,
        ).pack(padx=22, pady=(26, 6), fill="x")

        ctk.CTkButton(
            sidebar,
            text="Preview Resume",
            height=38,
            command=self.preview_resume,
        ).pack(padx=22, pady=6, fill="x")

        ctk.CTkButton(
            sidebar,
            text="Retry Failed Emails",
            height=38,
            command=self.retry_failed_emails,
        ).pack(padx=22, pady=6, fill="x")

        ctk.CTkButton(
            sidebar,
            text="Export Logs",
            height=38,
            command=self.export_logs,
        ).pack(padx=22, pady=6, fill="x")

        status = ctk.CTkLabel(
            sidebar,
            textvariable=self.status_text,
            wraplength=230,
            justify="left",
            text_color=("gray20", "gray80"),
        )
        status.pack(padx=22, pady=(28, 0), anchor="w")

        main = ctk.CTkFrame(self, fg_color="transparent")
        main.grid(row=0, column=1, sticky="nsew", padx=24, pady=22)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(2, weight=1)

        inputs = ctk.CTkFrame(main)
        inputs.grid(row=0, column=0, sticky="ew")
        inputs.grid_columnconfigure(1, weight=1)

        self._file_row(inputs, 0, "Excel / CSV", self.excel_path, self.browse_excel)
        self._file_row(inputs, 1, "Resume PDF", self.resume_path, self.browse_resume)
        self._file_row(inputs, 2, "Logo (optional)", self.logo_path, self.browse_logo)

        ctk.CTkLabel(inputs, text="Subject").grid(
            row=3, column=0, sticky="w", padx=18, pady=(12, 12)
        )
        ctk.CTkEntry(inputs, textvariable=self.subject).grid(
            row=3, column=1, sticky="ew", padx=10, pady=(12, 12)
        )

        ctk.CTkLabel(inputs, text="Daily Limit").grid(
            row=4, column=0, sticky="w", padx=18, pady=(0, 18)
        )
        limit_row = ctk.CTkFrame(inputs, fg_color="transparent")
        limit_row.grid(row=4, column=1, sticky="ew", padx=10, pady=(0, 18))
        limit_row.grid_columnconfigure(0, weight=0)
        limit_row.grid_columnconfigure(1, weight=1)
        ctk.CTkComboBox(
            limit_row,
            values=["200", "500", "1000", "Unlimited"],
            variable=self.daily_limit,
            width=160,
        ).grid(row=0, column=0, sticky="w")
        ctk.CTkSwitch(
            limit_row,
            text="HTML Email",
            variable=self.html_enabled,
        ).grid(row=0, column=1, sticky="w", padx=18)

        ctk.CTkLabel(inputs, text="Start At").grid(
            row=5, column=0, sticky="w", padx=18, pady=(0, 18)
        )
        schedule_row = ctk.CTkFrame(inputs, fg_color="transparent")
        schedule_row.grid(row=5, column=1, columnspan=2, sticky="ew", padx=(10, 18), pady=(0, 18))
        schedule_row.grid_columnconfigure(0, weight=0)
        schedule_row.grid_columnconfigure(1, weight=1)
        ctk.CTkEntry(
            schedule_row,
            textvariable=self.schedule_time,
            placeholder_text="HH:MM optional",
            width=160,
        ).grid(row=0, column=0, sticky="w")
        ctk.CTkEntry(
            schedule_row,
            textvariable=self.search_query,
            placeholder_text="Search HR by name, company, or email",
        ).grid(row=0, column=1, sticky="ew", padx=12)
        ctk.CTkButton(
            schedule_row,
            text="Search",
            width=90,
            command=self.search_hr,
        ).grid(row=0, column=2, sticky="e")

        stats = ctk.CTkFrame(main)
        stats.grid(row=1, column=0, sticky="ew", pady=18)
        for column in range(7):
            stats.grid_columnconfigure(column, weight=1)
        self._stat(stats, 0, "Current HR", self.current_name)
        self._stat(stats, 1, "Company", self.current_company)
        self._stat(stats, 2, "Current Email", self.current_email)
        self._stat(stats, 3, "Emails Sent", self.success_count)
        self._stat(stats, 4, "Failed", self.failed_count)
        self._stat(stats, 5, "Skipped", self.skipped_count)
        self._stat(stats, 6, "ETA", self.eta)

        preview = ctk.CTkFrame(main)
        preview.grid(row=2, column=0, sticky="nsew")
        preview.grid_columnconfigure(0, weight=1)
        preview.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(
            preview,
            text="Email Preview",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).grid(row=0, column=0, padx=18, pady=(16, 8), sticky="w")
        self.preview_box = ctk.CTkTextbox(preview, wrap="word")
        self.preview_box.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 18))

        self.progress = ctk.CTkProgressBar(main)
        self.progress.grid(row=3, column=0, sticky="ew", pady=(18, 0))
        self.progress.set(0)

    def _file_row(
        self,
        parent: ctk.CTkFrame,
        row: int,
        label: str,
        variable: tk.StringVar,
        command,
    ) -> None:
        ctk.CTkLabel(parent, text=label).grid(
            row=row, column=0, sticky="w", padx=18, pady=(18 if row == 0 else 6, 6)
        )
        ctk.CTkEntry(parent, textvariable=variable).grid(
            row=row,
            column=1,
            sticky="ew",
            padx=10,
            pady=(18 if row == 0 else 6, 6),
        )
        ctk.CTkButton(parent, text="Browse", width=96, command=command).grid(
            row=row,
            column=2,
            padx=(0, 18),
            pady=(18 if row == 0 else 6, 6),
        )

    def _stat(
        self,
        parent: ctk.CTkFrame,
        column: int,
        label: str,
        variable: tk.StringVar,
    ) -> None:
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.grid(row=0, column=column, sticky="ew", padx=8, pady=14)
        ctk.CTkLabel(frame, text=label, text_color=("gray25", "gray72")).pack(anchor="w")
        ctk.CTkLabel(frame, textvariable=variable, font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w")

    def browse_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="Select HR Excel or CSV",
            filetypes=[("Excel/CSV", "*.xlsx *.csv"), ("Excel", "*.xlsx"), ("CSV", "*.csv")],
        )
        if path:
            self.excel_path.set(path)

    def browse_resume(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Resume PDF",
            filetypes=[("PDF", "*.pdf")],
        )
        if path:
            self.resume_path.set(path)

    def browse_logo(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Company Logo",
            filetypes=[("Images", "*.png *.jpg *.jpeg *.gif"), ("All files", "*.*")],
        )
        if path:
            self.logo_path.set(path)

    def preview_email(self) -> None:
        try:
            mailer = self._mailer()
            subject, body = mailer.render_preview(self._mailer_options())
            self.preview_box.delete("1.0", "end")
            self.preview_box.insert("end", f"Subject: {subject}\n\n{body}")
            self.status_text.set("Preview rendered with sample HR and company values.")
        except Exception as exc:
            messagebox.showerror("Preview failed", str(exc))

    def preview_resume(self) -> None:
        path = self._entry_path(self.resume_path.get())
        if not path.exists():
            messagebox.showerror("Resume missing", f"Resume not found:\n{path}")
            return
        os.startfile(path)

    def search_hr(self) -> None:
        if not self.excel_path.get().strip():
            messagebox.showwarning("Excel required", "Please select the HR Excel file first.")
            return
        try:
            matches = search_recipients(
                self.excel_path.get(),
                self.search_query.get(),
                str(self.app_config.values.get("excel_sheet", "HR Contacts")),
            )
            self.preview_box.delete("1.0", "end")
            if not matches:
                self.preview_box.insert("end", "No HR contacts matched your search.")
                return
            lines = [
                f"{item.name or 'Team'} | {item.company or '-'} | {item.email}"
                for item in matches
            ]
            self.preview_box.insert("end", "\n".join(lines))
            self.status_text.set(f"Found {len(matches)} matching HR contact(s).")
        except Exception as exc:
            messagebox.showerror("Search failed", str(exc))

    def retry_failed_emails(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        failed_report = self.app_config.path("reports_dir") / "failed_report.xlsx"
        if not failed_report.exists():
            messagebox.showwarning("No failed report", f"Could not find:\n{failed_report}")
            return
        self._save_current_config()
        self._begin_worker("", self._mailer_options(), retry_failed=True)

    def start_sending(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        if not self.excel_path.get().strip():
            messagebox.showwarning("Excel required", "Please select an Excel or CSV file.")
            return

        self._save_current_config()
        excel_path = self.excel_path.get().strip()
        options = self._mailer_options()
        self._begin_worker(excel_path, options)

    def _begin_worker(
        self,
        excel_path: str,
        options: MailerOptions,
        retry_failed: bool = False,
    ) -> None:
        self.pause_event.clear()
        self.stop_event.clear()
        self._set_running(True)
        self.progress.set(0)
        self.status_text.set("Loading recipients...")
        self.current_name.set("-")
        self.current_company.set("-")
        self.current_email.set("-")
        self.success_count.set("0")
        self.failed_count.set("0")
        self.skipped_count.set("0")
        self.eta.set("-")

        self.worker = threading.Thread(
            target=self._run_sender,
            args=(excel_path, options, retry_failed),
            daemon=True,
        )
        self.worker.start()

    def pause_sending(self) -> None:
        self.pause_event.set()
        self.status_text.set("Paused")
        self.pause_button.configure(state="disabled")
        self.resume_button.configure(state="normal")

    def resume_sending(self) -> None:
        self.pause_event.clear()
        self.status_text.set("Resumed")
        self.pause_button.configure(state="normal")
        self.resume_button.configure(state="disabled")

    def stop_sending(self) -> None:
        self.stop_event.set()
        self.status_text.set("Stopping after the current safe point...")
        self.stop_button.configure(state="disabled")

    def export_logs(self) -> None:
        try:
            archive = self._mailer().export_logs()
            messagebox.showinfo("Logs exported", f"Saved to:\n{archive}")
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc))

    def _run_sender(
        self,
        excel_path: str,
        options: MailerOptions,
        retry_failed: bool = False,
    ) -> None:
        try:
            if not self._wait_for_schedule(options.schedule_time):
                self.events.put({"event": "worker_done", "message": "Sending cancelled before schedule."})
                return

            if retry_failed:
                failed_report = self.app_config.path("reports_dir") / "failed_report.xlsx"
                recipients = load_recipients_from_report(failed_report)
                skipped = []
                loaded_message = f"Loaded {len(recipients)} failed recipient(s) for retry."
            else:
                recipients, skipped = load_recipients(
                    excel_path,
                    str(self.app_config.values.get("excel_sheet", "HR Contacts")),
                )
                loaded_message = (
                    f"Loaded {len(recipients)} valid recipients from HR Contacts; "
                    f"{len(skipped)} skipped before send."
                )
            self.events.put(
                {
                    "event": "loaded",
                    "message": loaded_message,
                }
            )
            result = self._mailer().send_bulk(
                recipients,
                skipped,
                options,
                progress_callback=self.events.put,
            )
            self.events.put(
                {
                    "event": "worker_done",
                    "message": (
                        f"Done. Sent {len(result.sent)}, failed {len(result.failed)}, "
                        f"skipped {len(result.skipped)}."
                    ),
                }
            )
        except Exception as exc:
            self.failed_logger.error("Worker failed: %s", exc)
            self.events.put({"event": "worker_error", "message": str(exc)})

    def _wait_for_schedule(self, schedule_time: str) -> bool:
        value = schedule_time.strip()
        if not value:
            return True
        try:
            hour, minute = [int(part) for part in value.split(":", 1)]
            scheduled = datetime.now().replace(
                hour=hour,
                minute=minute,
                second=0,
                microsecond=0,
            )
        except ValueError:
            raise ValueError("Schedule must use HH:MM format, for example 21:30.")

        if scheduled <= datetime.now():
            scheduled += timedelta(days=1)

        while datetime.now() < scheduled:
            if self.stop_event.is_set():
                return False
            remaining = scheduled - datetime.now()
            minutes = max(0, int(remaining.total_seconds() // 60))
            seconds = max(0, int(remaining.total_seconds() % 60))
            self.events.put(
                {
                    "event": "scheduled_wait",
                    "message": f"Scheduled for {scheduled:%Y-%m-%d %H:%M}; starts in {minutes}m {seconds}s.",
                }
            )
            time_to_sleep = min(5, max(0.5, remaining.total_seconds()))
            self.stop_event.wait(time_to_sleep)
        return not self.stop_event.is_set()

    def _poll_events(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                self._handle_event(event)
        except queue.Empty:
            pass
        self.after(250, self._poll_events)

    def _handle_event(self, event: dict[str, object]) -> None:
        name = str(event.get("event", ""))
        if name in {"worker_done", "worker_error"}:
            self.status_text.set(str(event.get("message", "")))
            self._set_running(False)
            if name == "worker_error":
                messagebox.showerror("Sending failed", str(event.get("message", "")))
            return

        if name == "loaded":
            self.status_text.set(str(event.get("message", "")))
            return

        if "message" in event:
            self.status_text.set(str(event["message"]))
        if "current_name" in event and event["current_name"]:
            self.current_name.set(str(event["current_name"]))
        if "current_company" in event and event["current_company"]:
            self.current_company.set(str(event["current_company"]))
        if "current_email" in event and event["current_email"]:
            self.current_email.set(str(event["current_email"]))
        if "success_count" in event:
            self.success_count.set(str(event["success_count"]))
        if "failed_count" in event:
            self.failed_count.set(str(event["failed_count"]))
        if "skipped_count" in event:
            self.skipped_count.set(str(event["skipped_count"]))
        if "eta" in event:
            self.eta.set(str(event["eta"]))
        total = int(event.get("total", 0) or 0)
        processed = int(event.get("processed", 0) or 0)
        if total > 0:
            self.progress.set(min(processed / total, 1.0))
        if name in {"complete", "stopped"}:
            self._set_running(False)

    def _set_running(self, running: bool) -> None:
        self.start_button.configure(state="disabled" if running else "normal")
        self.pause_button.configure(state="normal" if running else "disabled")
        self.resume_button.configure(state="disabled")
        self.stop_button.configure(state="normal" if running else "disabled")

    def _mailer_options(self) -> MailerOptions:
        limit = self.daily_limit.get().strip()
        daily_limit = None if limit.lower() == "unlimited" else int(limit)
        logo = self.logo_path.get().strip()
        return MailerOptions(
            subject=self.subject.get().strip(),
            resume_path=self._entry_path(self.resume_path.get()),
            template_path=self.app_config.path("template_path"),
            daily_limit=daily_limit,
            html_email=bool(self.html_enabled.get()),
            logo_path=self._entry_path(logo) if logo else None,
            schedule_time=self.schedule_time.get().strip(),
        )

    def _entry_path(self, value: str) -> Path:
        path = Path(value.strip()).expanduser()
        return path if path.is_absolute() else BASE_DIR / path

    def _display_config_path(self, key: str) -> str:
        value = str(self.app_config.values.get(key, "")).strip()
        if not value:
            return ""
        path = Path(value).expanduser()
        return str(path if path.is_absolute() else BASE_DIR / path)

    def _send_policy_text(self) -> str:
        limit = self.app_config.values.get("daily_limit", 300)
        start = self.app_config.values.get("send_window_start", "08:00")
        end = self.app_config.values.get("send_window_end", "12:00")
        return f"Daily cap: {limit} emails\nSend window: {start}-{end}"

    def _mailer(self) -> BrevoMailer:
        self.app_config = load_app_config()
        return BrevoMailer(
            self.app_config,
            self.sent_logger,
            self.failed_logger,
            self.pause_event,
            self.stop_event,
        )

    def _save_current_config(self) -> None:
        daily_limit = self.daily_limit.get().strip()
        values = self.app_config.values | {
            "subject": self.subject.get().strip(),
            "excel_path": self.excel_path.get().strip(),
            "resume_path": self.resume_path.get().strip(),
            "logo_path": self.logo_path.get().strip(),
            "html_email": bool(self.html_enabled.get()),
            "schedule_start_time": self.schedule_time.get().strip(),
            "daily_limit": None if daily_limit.lower() == "unlimited" else int(daily_limit),
        }
        save_json_config(values)
        self.app_config = load_app_config()


def run() -> None:
    app = JobMailerApp()
    app.mainloop()
