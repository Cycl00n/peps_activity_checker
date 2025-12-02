import customtkinter as ctk
from tkinter import filedialog, messagebox
import json
import os
import re
from analysis import analyze_excel, load_employees, is_activity_cancelled
from mail_sender import send_email_outlook
import unicodedata


ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


# --------------------------------------------------------
#  FONCTION de nettoyage 100% fiable
# --------------------------------------------------------
def normalize(text):
    if not text:
        return ""

    # Remove all weird whitespaces
    text = text.replace("\n", " ").replace("\r", " ").replace("\t", " ")

    # Replace ALL unicode spaces by normal spaces
    text = re.sub(r"\s+", " ", text)

    # Replace ALL hyphens by a normal hyphen
    text = re.sub(r"[‐-‒–—−]", "-", text)  # all unicode hyphens

    # Lowercase
    text = text.lower()

    # Remove accents
    text = "".join(
        c for c in unicodedata.normalize("NFD", text) if unicodedata.category(c) != "Mn"
    )

    # Strip
    return text.strip()


def remove_educators_from_activity(activity, educators):
    activity_norm = normalize(activity)
    tokens = set()

    for emp in educators:
        emp_norm = normalize(emp)
        tokens.add(emp_norm)
        for p in emp_norm.split():
            tokens.add(p)

    cleaned_words = []
    for word in re.split(r"\s+", activity):
        if normalize(word) not in tokens:
            cleaned_words.append(word)

    result = " ".join(cleaned_words)
    return " ".join(result.split())


# --------------------------------------------------------
#  GUI CLASS
# --------------------------------------------------------
class PepsGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("PEPS Activity Checker")
        self.geometry("1500x800")

        self.mode = ctk.StringVar(value="hard")
        self.activities = []
        self.current_cc = ""
        self.current_act = None
        self.selected_educator = None
        self.cc_manually_set = False
        self.mouse_over_list = False  # Track if mouse is over activity list
        self.selected_button = None  # Track currently selected activity button
        self.include_corrections = True  # Toggle for correction messages

        self.build_layout()

        # Bind mousewheel to main window
        self.bind("<MouseWheel>", self._on_global_mousewheel)
        self.bind("<Button-4>", self._on_global_mousewheel)
        self.bind("<Button-5>", self._on_global_mousewheel)

    # --------------------------------------------------------
    #   LAYOUT
    # --------------------------------------------------------
    def build_layout(self):
        # LEFT PANEL
        left = ctk.CTkFrame(self)
        left.pack(side="left", fill="y", padx=5, pady=10)

        ctk.CTkLabel(left, text="Paramètres", font=("Arial", 18, "bold")).pack(pady=10)

        ctk.CTkRadioButton(
            left, text="Minimum Check (présences)", variable=self.mode, value="soft"
        ).pack(anchor="w", padx=10)
        ctk.CTkRadioButton(
            left, text="Full Check (+ descriptions)", variable=self.mode, value="hard"
        ).pack(anchor="w", padx=10)

        ctk.CTkButton(left, text="📁 Charger Excel", command=self.load_excel).pack(
            pady=15, padx=5, fill="x"
        )

        self.stats_incomplete = ctk.CTkLabel(
            left, text="Incomplètes : 0", font=("Arial", 11)
        )
        self.stats_cancelled = ctk.CTkLabel(
            left, text="Annulées : 0", font=("Arial", 11)
        )
        self.stats_total = ctk.CTkLabel(left, text="Total : 0", font=("Arial", 11))

        self.stats_incomplete.pack(anchor="w", padx=10, pady=2)
        self.stats_cancelled.pack(anchor="w", padx=10, pady=2)
        self.stats_total.pack(anchor="w", padx=10, pady=2)

        # Correction toggle button
        self.correction_button = ctk.CTkButton(
            left,
            text="✓ Raison du rappel on",
            command=self.toggle_corrections,
            fg_color="#2d5a2d",
        )
        self.correction_button.pack(pady=10, padx=5, fill="x")

        # Footer in left panel
        footer = ctk.CTkLabel(
            left,
            text="Version 2.5\nDeveloped by Clément N. 2025.",
            text_color="#666666",
            font=("Arial", 9),
            justify="left",
        )
        footer.pack(side="bottom", anchor="sw", padx=5, pady=5)

        # Edit buttons at bottom of left panel
        button_frame = ctk.CTkFrame(left)
        button_frame.pack(side="bottom", pady=10, padx=2, fill="x")

        ctk.CTkButton(
            button_frame,
            text="⚙ Employees",
            command=self.edit_employees,
            font=("Arial", 10),
        ).pack(side="left", padx=2, fill="x", expand=True)

        ctk.CTkButton(
            button_frame,
            text="👥 Residents",
            command=self.edit_residents,
            font=("Arial", 10),
        ).pack(side="left", padx=2, fill="x", expand=True)

        # CENTER PANEL
        center = ctk.CTkFrame(self)
        center.pack(side="left", fill="both", expand=True, padx=5, pady=10)

        # Mail header
        ctk.CTkLabel(center, text="Courrier", font=("Arial", 14, "bold")).pack(pady=5)

        # Educator selector
        self.educator_selector = ctk.CTkComboBox(
            center, values=[], command=self.on_educator_select
        )
        self.educator_selector.pack(fill="x", pady=2)

        # Mail fields frame (compact layout)
        fields_frame = ctk.CTkFrame(center)
        fields_frame.pack(fill="x", pady=5)

        ctk.CTkLabel(fields_frame, text="À:", font=("Arial", 10)).grid(
            row=0, column=0, sticky="w", padx=5
        )
        self.entry_to = ctk.CTkEntry(fields_frame, height=26)
        self.entry_to.grid(row=0, column=1, sticky="ew", padx=5, pady=2)

        ctk.CTkLabel(fields_frame, text="CC:", font=("Arial", 10)).grid(
            row=1, column=0, sticky="w", padx=5
        )
        self.entry_cc = ctk.CTkEntry(fields_frame, height=26)
        self.entry_cc.grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        ctk.CTkLabel(fields_frame, text="Objet:", font=("Arial", 10)).grid(
            row=2, column=0, sticky="w", padx=5
        )
        self.entry_subject = ctk.CTkEntry(fields_frame, height=26)
        self.entry_subject.grid(row=2, column=1, sticky="ew", padx=5, pady=2)

        fields_frame.grid_columnconfigure(1, weight=1)

        # Bind to track manual edits
        self.entry_cc.bind("<KeyRelease>", self.on_cc_edited)

        # Mail text
        self.mail_text = ctk.CTkTextbox(center, font=("Consolas", 14), wrap="word")
        self.mail_text.pack(fill="both", expand=True, pady=5)

        # Send button
        self.send_button = ctk.CTkButton(
            center, text="📧 Envoyer", command=self.send_mail
        )
        self.send_button.pack(pady=10, fill="x")

        # RIGHT PANEL
        right = ctk.CTkFrame(self, width=500)
        right.pack(side="right", fill="y", padx=10, pady=10)
        right.pack_propagate(False)

        self.act_list = ctk.CTkScrollableFrame(right, width=320)
        self.act_list.pack(fill="both", expand=True, pady=(0, 10))

        # Bind mousewheel scroll to the scrollable frame
        self.act_list.bind("<MouseWheel>", self._on_mousewheel)
        self.act_list.bind("<Button-4>", self._on_mousewheel)
        self.act_list.bind("<Button-5>", self._on_mousewheel)

        # Track mouse enter/leave
        self.act_list.bind("<Enter>", self._on_mouse_enter)
        self.act_list.bind("<Leave>", self._on_mouse_leave)

        self.activity_details = ctk.CTkTextbox(
            right, width=320, height=200, font=("Consolas", 13), wrap="word"
        )
        self.activity_details.pack(fill="both", expand=True, pady=0)
        self.activity_details.configure(state="normal")

    def _on_mouse_enter(self, event):
        """Mouse entered activity list"""
        self.mouse_over_list = True

    def _on_mouse_leave(self, event):
        """Mouse left activity list"""
        self.mouse_over_list = False

    def _on_global_mousewheel(self, event):
        """Route mousewheel events only if mouse is over the activity list"""
        if not self.mouse_over_list:
            return

        try:
            if hasattr(self.act_list, "_parent_canvas"):
                canvas = self.act_list._parent_canvas
                if event.num == 5 or event.delta < 0:
                    canvas.yview_scroll(3, "units")
                elif event.num == 4 or event.delta > 0:
                    canvas.yview_scroll(-3, "units")
        except Exception:
            pass

    def _on_mousewheel(self, event):
        """Handle mousewheel scrolling (direct events)"""
        try:
            if hasattr(self.act_list, "_parent_canvas"):
                canvas = self.act_list._parent_canvas
                if event.num == 5 or event.delta < 0:
                    canvas.yview_scroll(3, "units")
                elif event.num == 4 or event.delta > 0:
                    canvas.yview_scroll(-3, "units")
        except Exception:
            pass

    # --------------------------------------------------------
    #   LOAD EXCEL
    # --------------------------------------------------------
    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not path:
            return

        activities, _ = analyze_excel(path, self.mode.get())
        self.activities = activities

        # Count incomplete and cancelled
        incomplete_count = 0
        cancelled_count = 0

        for act in activities:
            # Check if cancelled (has no errors key = it's a cancelled activity)
            is_cancelled = "errors" not in act

            if is_cancelled:
                cancelled_count += 1
            else:
                incomplete_count += 1

        self.stats_incomplete.configure(text=f"Incomplètes : {incomplete_count}")
        self.stats_cancelled.configure(text=f"Annulées : {cancelled_count}")
        self.stats_total.configure(text=f"Total : {len(self.activities)}")

        self.populate_activity_list()

    # --------------------------------------------------------
    #   POPULATE ACTIVITIES
    # --------------------------------------------------------
    def populate_activity_list(self):
        for widget in self.act_list.winfo_children():
            widget.destroy()

        for act in self.activities:
            # Remove educators from line 1
            activity_clean = remove_educators_from_activity(
                act["activity"], act["educators"]
            )
            educators_line = ", ".join(act["educators"])

            button_text = f"{activity_clean}\n{educators_line}\n{act['date']}"

            # Determine tag and color based on activity type
            tag_text = ""
            tag_color = "#303030"

            # CANCELLED HAS PRIORITY: if no errors key, it's cancelled
            if "errors" not in act:
                tag_text = "Annulée"
                tag_color = "#4a4a4a"
            elif act.get("errors"):
                # Get tag based on error type
                if "Aucun" in act["errors"][0]:
                    tag_text = "Présences"
                    tag_color = "#6b1a1a"
                elif "note" in act["errors"][0].lower():
                    tag_text = "Notes"
                    tag_color = "#5a4a1a"
                else:
                    tag_text = "Incomplet"
                    tag_color = "#6b1a1a"
            else:
                # Soft mode: incomplete because no participation
                tag_text = "Présences"
                tag_color = "#6b1a1a"

            # Create frame for button + tag
            frame = ctk.CTkFrame(self.act_list, fg_color="transparent")
            frame.pack(fill="x", pady=4)

            # Bind mouse tracking to frame
            frame.bind("<Enter>", self._on_mouse_enter)
            frame.bind("<Leave>", self._on_mouse_leave)

            # Main activity button
            b = ctk.CTkButton(
                frame,
                text=button_text,
                fg_color="#303030",
                hover_color="#505050",
                command=lambda a=act, btn=None: self.show_activity(a, btn),
            )
            b.pack(side="left", fill="both", expand=True, padx=(0, 5))

            # Store reference to button in lambda to capture it
            b.configure(command=lambda a=act, btn=b: self.show_activity(a, btn))

            # Bind mouse tracking to button
            b.bind("<Enter>", self._on_mouse_enter)
            b.bind("<Leave>", self._on_mouse_leave)

            # Tag button
            if tag_text:
                tag_btn = ctk.CTkButton(
                    frame,
                    text=tag_text,
                    width=70,
                    height=60,
                    fg_color=tag_color,
                    hover_color=tag_color,
                    font=("Arial", 10, "bold"),
                )
                tag_btn.pack(side="left", padx=0)

                # Bind mouse tracking to tag button
                tag_btn.bind("<Enter>", self._on_mouse_enter)
                tag_btn.bind("<Leave>", self._on_mouse_leave)

        # Trigger layout update
        self.act_list.update()

    # --------------------------------------------------------
    #   SHOW ACTIVITY DETAILS + MAIL
    # --------------------------------------------------------
    def show_activity(self, act, btn=None):
        # Reset previous button to normal color (if it still exists)
        if self.selected_button is not None:
            try:
                self.selected_button.configure(fg_color="#303030")
            except:
                pass

        # Highlight new button
        if btn is not None:
            btn.configure(fg_color="#505050")
            self.selected_button = btn

        self.current_act = act  # IMPORTANT : avant on_educator_select()

        # DETAILS
        self.activity_details.configure(state="normal")
        self.activity_details.delete("1.0", "end")

        txt = f"{act['activity']}\nDate : {act['date']}\n\n"
        txt += "Description générale:\n"
        txt += (act.get("desc") or "—") + "\n\n"

        # Only show residents if they exist
        if act.get("residents"):
            txt += "Résidents :\n"
            seen = set()
            for r in act.get("residents", []):
                if r["name"] not in seen:
                    seen.add(r["name"])
                    line = f"\n• {r['name']}"
                    if r["status"] == "a participé":
                        line += " (a participé)"
                        if r["note"]:
                            line += f" : {r['note']}"
                    txt += line + "\n"

        self.activity_details.insert("end", txt)
        self.activity_details.configure(state="disabled")

        # MAIL
        educators = act.get("educators", [])
        if educators:
            self.educator_selector.configure(values=educators)
            self.educator_selector.set(educators[0])
            self.on_educator_select(educators[0])
        else:
            # No educators: clear the selector and show placeholder message
            self.educator_selector.configure(values=[])
            self.educator_selector.set("Aucun éducateur trouvé")
            self.entry_to.delete(0, "end")
            self.entry_cc.delete(0, "end")
            self.entry_subject.delete(0, "end")
            self.mail_text.delete("1.0", "end")
            self.mail_text.insert(
                "end",
                "⚠️ Aucun employé n'est assigné à cette activité.",
            )
            self.send_button.configure(text="📧 Envoyer")

    # --------------------------------------------------------
    #   TRACK CC EDITS
    # --------------------------------------------------------
    def on_cc_edited(self, event=None):
        """Track when user manually edits the CC field"""
        if self.entry_cc.get().strip():
            self.cc_manually_set = True
            self.current_cc = self.entry_cc.get()
        else:
            self.cc_manually_set = False

    # --------------------------------------------------------
    #   TOGGLE CORRECTIONS
    # --------------------------------------------------------
    def toggle_corrections(self):
        """Toggle correction messages on/off"""
        self.include_corrections = not self.include_corrections

        if self.include_corrections:
            self.correction_button.configure(
                text="✓ Raison du rappel on", fg_color="#2d5a2d"
            )
        else:
            self.correction_button.configure(
                text="✕ Raison du rappel off", fg_color="#5a2d2d"
            )

        # Refresh the current email body
        if self.current_act is not None:
            self.on_educator_select(self.selected_educator)

    # --------------------------------------------------------
    #   EDUCATOR SELECTED
    # --------------------------------------------------------
    def on_educator_select(self, name):
        # Skip if no educator found
        if name == "Aucun éducateur trouvé" or not name:
            return

        self.selected_educator = name
        employees = load_employees()
        email = employees.get(name, "inconnu@jardinarlon.be")

        self.entry_to.delete(0, "end")
        self.entry_to.insert(0, email)

        # Only update CC if user hasn't manually set it
        if not self.cc_manually_set:
            self.entry_cc.delete(0, "end")
            self.entry_cc.insert(0, self.current_cc)

        self.entry_subject.delete(0, "end")
        self.entry_subject.insert(0, f"Rappel encodage — {self.current_act['date']}")

        first_name = name.split()[-1]

        # Determine correction message based on error type (only if enabled)
        correction_msg = ""
        if self.include_corrections:
            if self.current_act.get("errors"):
                if "Aucun" in self.current_act["errors"][0]:
                    correction_msg = "Il faut corriger la participation."
                elif "note" in self.current_act["errors"][0].lower():
                    correction_msg = "Il faut corriger les descriptions générales et / ou individuelles."
            else:
                # Soft mode: no participation
                correction_msg = "Il faut corriger la participation des résidents."

        # Build body with context-specific message
        body = f"""Salut {first_name},

Moyen que tu complètes tes encodages stp:

- {self.current_act['date']} : {self.current_act['activity']}"""

        if correction_msg:
            body += f"\n{correction_msg}"

        body += """\n\nN'hésite pas si tu as des questions.
Bien à toi,"""

        self.mail_text.delete("1.0", "end")
        self.mail_text.insert("end", body)

        self.send_button.configure(text="📧 Envoyer")  # reset bouton

    # --------------------------------------------------------
    #   SEND MAIL (avec ✔)
    # --------------------------------------------------------
    def send_mail(self):
        # Validate required fields
        to = self.entry_to.get().strip()
        subject = self.entry_subject.get().strip()
        body = self.mail_text.get("1.0", "end-1c").strip()
        cc = self.entry_cc.get().strip()

        if not to:
            messagebox.showerror("Erreur", "Veuillez spécifier un destinataire")
            return

        if not subject:
            messagebox.showerror("Erreur", "Veuillez spécifier un objet")
            return

        if not body:
            messagebox.showerror("Erreur", "Veuillez rédiger un message")
            return

        # Send email via Outlook
        success, message = send_email_outlook(to, cc, subject, body)

        if success:
            self.send_button.configure(text="✔ Envoyé")
            messagebox.showinfo("Succès", message)
        else:
            messagebox.showerror("Erreur", message)

    # --------------------------------------------------------
    #   EDIT EMPLOYEES JSON WINDOW
    # --------------------------------------------------------
    def edit_employees(self):
        # Create top-level window
        editor_window = ctk.CTkToplevel(self)
        editor_window.title("Modifier employees.json")
        editor_window.geometry("600x500")
        editor_window.resizable(True, True)

        # Load current JSON
        try:
            with open("employees.json", "r", encoding="utf-8") as f:
                json_content = json.dumps(json.load(f), indent=2, ensure_ascii=False)
        except Exception as e:
            json_content = f"Erreur: {str(e)}"

        # Text editor
        ctk.CTkLabel(
            editor_window, text="employees.json", font=("Arial", 14, "bold")
        ).pack(pady=5)
        text_editor = ctk.CTkTextbox(editor_window, font=("Consolas", 12))
        text_editor.pack(fill="both", expand=True, padx=10, pady=5)
        text_editor.insert("1.0", json_content)

        # Buttons frame
        button_frame = ctk.CTkFrame(editor_window)
        button_frame.pack(fill="x", padx=10, pady=10)

        def save_json():
            try:
                content = text_editor.get("1.0", "end-1c")
                data = json.loads(content)
                with open("employees.json", "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=2, ensure_ascii=False)
                editor_window.destroy()
            except json.JSONDecodeError as e:
                ctk.CTkLabel(
                    editor_window, text=f"Erreur JSON: {str(e)}", text_color="red"
                ).pack()

        ctk.CTkButton(button_frame, text="💾 Sauvegarder", command=save_json).pack(
            side="left", padx=5
        )
        ctk.CTkButton(
            button_frame, text="✕ Annuler", command=editor_window.destroy
        ).pack(side="left", padx=5)

    # --------------------------------------------------------
    #   EDIT RESIDENTS JSON WINDOW
    # --------------------------------------------------------
    def edit_residents(self):
        # Create top-level window
        editor_window = ctk.CTkToplevel(self)
        editor_window.title("Modifier residents.json")
        editor_window.geometry("600x500")
        editor_window.resizable(True, True)

        # Load current JSON
        try:
            with open("residents.json", "r", encoding="utf-8") as f:
                json_content = json.dumps(json.load(f), indent=2, ensure_ascii=False)
        except Exception as e:
            json_content = f"Erreur: {str(e)}"

        # Text editor
        ctk.CTkLabel(
            editor_window, text="residents.json", font=("Arial", 14, "bold")
        ).pack(pady=5)
        text_editor = ctk.CTkTextbox(editor_window, font=("Consolas", 12))
        text_editor.pack(fill="both", expand=True, padx=10, pady=5)
        text_editor.insert("1.0", json_content)

        # Buttons frame
        button_frame = ctk.CTkFrame(editor_window)
        button_frame.pack(fill="x", padx=10, pady=10)

        def save_json():
            try:
                content = text_editor.get("1.0", "end-1c")
                data = json.loads(content)
                with open("residents.json", "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=2, ensure_ascii=False)
                editor_window.destroy()
            except json.JSONDecodeError as e:
                ctk.CTkLabel(
                    editor_window, text=f"Erreur JSON: {str(e)}", text_color="red"
                ).pack()

        ctk.CTkButton(button_frame, text="💾 Sauvegarder", command=save_json).pack(
            side="left", padx=5
        )
        ctk.CTkButton(
            button_frame, text="✕ Annuler", command=editor_window.destroy
        ).pack(side="left", padx=5)

    # --------------------------------------------------------
    def edit_json_old(self):
        os.system("xdg-open employees.json")


if __name__ == "__main__":
    app = PepsGUI()
    app.mainloop()
