from __future__ import annotations

import tkinter as tk


class LogPanel(tk.Frame):
	def __init__(self, parent, *, bg: str = "#1e1e1e", fg: str = "#d4d4d4"):
		super().__init__(parent, bg=bg)
		self.text = tk.Text(
			self,
			height=7,
			font=("Consolas", 9),
			bg=bg,
			fg=fg,
			relief="flat",
			state="disabled",
			padx=12,
			pady=10,
			wrap="word",
		)
		self.text.pack(fill="both", expand=True)
		self.text.tag_config("ok", foreground="#4ec9b0")
		self.text.tag_config("error", foreground="#f44747")
		self.text.tag_config("info", foreground="#9cdcfe")
		self.text.tag_config("warn", foreground="#dcdcaa")
		self.text.tag_config("muted", foreground="#6a9955")

	def write(self, message: str, tag: str = "info") -> None:
		self.text.config(state="normal")
		self.text.insert("end", f"{message}\n", tag)
		self.text.see("end")
		self.text.config(state="disabled")

	def clear(self) -> None:
		self.text.config(state="normal")
		self.text.delete("1.0", "end")
		self.text.config(state="disabled")
