import json
import os
import random
import time
import zipfile
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, messagebox

APP_BG = "#070b16"
CARD = "#10182b"
CARD2 = "#151f36"
CARD3 = "#1b2742"
TEXT = "#f4f7ff"
MUTED = "#7f8ba8"
PURPLE = "#8b5cf6"
CYAN = "#22d3ee"
GREEN = "#34d399"
YELLOW = "#fbbf24"
RED = "#fb7185"
ORANGE = "#fb923c"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "quizverse_profile.json")
QUESTIONS_FILE = os.path.join(BASE_DIR, "quizverse_questions.xlsx")



#  Question bank loader — reads the .xlsx with the standard library only
#  (an .xlsx is just a zip of XML files), so no pip install is needed.

_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PKG = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def _col_index(ref):
    """'C5' -> 2 (zero-based column number)."""
    n = 0
    for ch in ref:
        if not ch.isalpha():
            break
        n = n * 26 + (ord(ch.upper()) - 64)
    return n - 1


def read_xlsx(path):
    """Return {sheet_name: [[cell, cell, ...], ...]} with every cell as text."""
    sheets = {}
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.iter(_MAIN + "si"):
                shared.append("".join(t.text or "" for t in si.iter(_MAIN + "t")))

        rels = {}
        rel_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        for r in rel_root.iter(_PKG + "Relationship"):
            target = r.get("Target").lstrip("/")
            rels[r.get("Id")] = target if target.startswith("xl/") else "xl/" + target

        wb_root = ET.fromstring(z.read("xl/workbook.xml"))
        for sh in wb_root.iter(_MAIN + "sheet"):
            part = rels[sh.get(_REL + "id")]
            rows = []
            for row in ET.fromstring(z.read(part)).iter(_MAIN + "row"):
                cells = []
                for c in row.iter(_MAIN + "c"):
                    idx = _col_index(c.get("r"))
                    while len(cells) <= idx:
                        cells.append("")
                    kind = c.get("t")
                    if kind == "inlineStr":
                        val = "".join(t.text or "" for t in c.iter(_MAIN + "t"))
                    else:
                        v = c.find(_MAIN + "v")
                        val = (v.text or "") if v is not None else ""
                        if kind == "s" and val != "":
                            val = shared[int(val)]
                    cells[idx] = val.strip()
                rows.append(cells)
            sheets[sh.get("name")] = rows
    return sheets


def load_topics(path=QUESTIONS_FILE):
    """Build the TOPICS structure the game uses from the Excel file."""
    sheets = read_xlsx(path)
    if "Questions" not in sheets:
        raise ValueError("The workbook needs a sheet named 'Questions'.")

    meta = {}
    for row in sheets.get("Topics", [])[1:]:
        row = row + [""] * (3 - len(row))
        if row[0]:
            meta[row[0]] = (row[1] or "❓", row[2] or "#8b5cf6")

    topics = {name: {"icon": icon, "color": color, "questions": []}
              for name, (icon, color) in meta.items()}

    for row in sheets["Questions"][1:]:
        row = row + [""] * (8 - len(row))
        topic, question, a, b, c, d, letter, explanation = row[:8]
        options = [o for o in (a, b, c, d) if o]
        letter = letter.upper()
        if not (topic and question and len(options) >= 2 and letter in "ABCD" and letter):
            continue                      # skip blank / malformed rows
        correct = "ABCD".index(letter)
        if correct >= len(options):
            continue
        topics.setdefault(topic, {"icon": "❓", "color": "#8b5cf6", "questions": []})
        topics[topic]["questions"].append((question, options, correct, explanation))

    topics = {k: v for k, v in topics.items() if v["questions"]}
    if not topics:
        raise ValueError("No valid questions found in the workbook.")
    return topics


# Filled in main() before the game starts
TOPICS = {}

DIFFICULTY = {
    "Easy": {"time": 30, "mult": 1.0},
    "Medium": {"time": 20, "mult": 1.5},
    "Hard": {"time": 12, "mult": 2.0}
}

MODES = {
    "Classic": ("🎯", "Normal quiz with a relaxed run"),
    "Blitz": ("⚡", "Fast questions and aggressive timers"),
    "Survival": ("☠", "Three lives. One mistake hurts."),
    "Streak Rush": ("🔥", "Build the longest combo possible")
}

ACHIEVEMENTS = {
    "First Run": ("🚀", "Complete your first quiz"),
    "Perfect": ("💎", "Get 100% in a quiz"),
    "Combo 5": ("🔥", "Reach a 5-question streak"),
    "Combo 10": ("🌋", "Reach a 10-question streak"),
    "Scholar": ("🧠", "Earn 2,000 total XP"),
    "Explorer": ("🗺", "Play 5 different topics"),
    "Speed Demon": ("⚡", "Finish a quiz with 90%+ accuracy"),
    "Survivor": ("☠", "Finish Survival mode")
}

class QuizVerse:
    def __init__(self, root):
        self.root = root
        self.root.title("QUIZVERSE — Brain Arena")
        self.root.geometry("1180x760")
        self.root.minsize(1000, 680)
        self.root.configure(bg=APP_BG)

        self.profile = self.load_profile()
        self.topic = "Python"
        self.difficulty = "Medium"
        self.mode = "Classic"
        self.count = 10
        self.questions = []
        self.answers = []
        self.current = 0
        self.streak = 0
        self.best_streak = 0
        self.correct = 0
        self.xp_run = 0
        self.lives = 3
        self.time_left = 20
        self.timer_id = None
        self.answer_locked = False
        self.session_topic_set = set()

        self.configure_styles()
        self.show_home()

    def load_profile(self):
        default = {
            "name": "",
            "xp": 0,
            "level": 1,
            "games": 0,
            "correct": 0,
            "questions": 0,
            "best_streak": 0,
            "topics": [],
            "achievements": [],
            "history": []
        }
        try:
            if os.path.exists(DATA_FILE):
                with open(DATA_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                default.update(data)
        except Exception:
            pass
        return default

    def save_profile(self):
        try:
            with open(DATA_FILE, "w", encoding="utf-8") as f:
                json.dump(self.profile, f, indent=2)
        except Exception:
            pass

    def configure_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Game.Horizontal.TProgressbar",
                        troughcolor="#202a43", background=PURPLE,
                        bordercolor="#202a43", lightcolor=PURPLE, darkcolor=PURPLE)

    def clear(self):
        if self.timer_id:
            self.root.after_cancel(self.timer_id)
            self.timer_id = None
        for w in self.root.winfo_children():
            w.destroy()

    def lbl(self, parent, text, size=12, color=TEXT, bold=False, **kwargs):
        return tk.Label(parent, text=text, font=("Segoe UI", size, "bold" if bold else "normal"),
                        fg=color, bg=parent.cget("bg"), **kwargs)

    def btn(self, parent, text, command, bg=CARD3, fg=TEXT, size=10, **kwargs):
        return tk.Button(parent, text=text, command=command,
                         font=("Segoe UI", size, "bold"), bg=bg, fg=fg,
                         activebackground=bg, activeforeground=fg,
                         bd=0, relief="flat", cursor="hand2",
                         padx=15, pady=9, **kwargs)

    def level_info(self):
        level = self.profile["level"]
        current = self.profile["xp"]
        needed = level * 1000
        while current >= needed:
            current -= needed
            level += 1
            needed = level * 1000
        self.profile["level"] = level
        return level, current, needed

    def award_achievement(self, name):
        if name not in self.profile["achievements"]:
            self.profile["achievements"].append(name)
            return True
        return False

    def home_header(self, parent):
        top = tk.Frame(parent, bg=APP_BG)
        top.pack(fill="x", padx=36, pady=(24, 10))
        brand = tk.Frame(top, bg=APP_BG)
        brand.pack(side="left")
        self.lbl(brand, "QUIZ", 17, PURPLE, True).pack(side="left")
        self.lbl(brand, "VERSE", 17, CYAN, True).pack(side="left")
        self.lbl(top, f"LVL {self.profile['level']}   ✦ {self.profile['xp']} XP",
                 10, "#b8c1da", True).pack(side="right", pady=5)
        return top

    def show_home(self):
        self.clear()
        self.root.configure(bg=APP_BG)
        self.home_header(self.root)

        hero = tk.Frame(self.root, bg=CARD)
        hero.pack(fill="x", padx=36, pady=(5, 16))

        left = tk.Frame(hero, bg=CARD)
        left.pack(side="left", fill="both", expand=True, padx=28, pady=24)
        self.lbl(left, f"Welcome back, {self.profile['name'] or 'Player'} 👋", 12, CYAN, True).pack(anchor="w")
        self.lbl(left, "ENTER THE BRAIN ARENA.", 28, TEXT, True).pack(anchor="w", pady=(5, 0))
        self.lbl(left, "Pick your battlefield. Build a combo. Earn XP. Master every topic.", 11, MUTED).pack(anchor="w", pady=(4, 15))
        self.btn(left, "🚀  START A NEW RUN", self.open_setup, PURPLE, size=11).pack(anchor="w")

        stats = tk.Frame(hero, bg=CARD2)
        stats.pack(side="right", padx=22, pady=22, fill="y")
        self.stat_box(stats, "LEVEL", str(self.profile["level"]), PURPLE)
        self.stat_box(stats, "QUIZZES", str(self.profile["games"]), CYAN)
        self.stat_box(stats, "BEST", str(self.profile["best_streak"]), YELLOW)
        self.stat_box(stats, "TOPICS", str(len(self.profile["topics"])), GREEN)

        body = tk.Frame(self.root, bg=APP_BG)
        body.pack(fill="both", expand=True, padx=36)

        left_col = tk.Frame(body, bg=APP_BG)
        left_col.pack(side="left", fill="both", expand=True, padx=(0, 10))
        right_col = tk.Frame(body, bg=APP_BG, width=280)
        right_col.pack(side="right", fill="y", padx=(10, 0))
        right_col.pack_propagate(False)

        self.lbl(left_col, "CHOOSE YOUR ARENA", 10, MUTED, True).pack(anchor="w", pady=(0, 8))
        grid = tk.Frame(left_col, bg=APP_BG)
        grid.pack(fill="both", expand=True)

        items = list(TOPICS.items())
        for i, (name, info) in enumerate(items):
            r, c = divmod(i, 3)
            card = tk.Frame(grid, bg=CARD, cursor="hand2", highlightthickness=1, highlightbackground="#202b45")
            card.grid(row=r, column=c, padx=(0, 8), pady=5, sticky="nsew")
            grid.grid_columnconfigure(c, weight=1)
            grid.grid_rowconfigure(r, weight=1)
            self.lbl(card, f"{info['icon']}  {name}", 10, TEXT, True).pack(anchor="w", padx=13, pady=(11, 2))
            qcount = len(info["questions"])
            mastery = self.topic_mastery(name)
            self.lbl(card, f"{qcount} questions  •  {mastery}% mastered", 8, MUTED).pack(anchor="w", padx=13)
            bar = ttk.Progressbar(card, style="Game.Horizontal.TProgressbar", maximum=100, value=mastery)
            bar.pack(fill="x", padx=13, pady=(7, 11))
            card.bind("<Button-1>", lambda e, t=name: self.select_topic(t))
            for child in card.winfo_children():
                child.bind("<Button-1>", lambda e, t=name: self.select_topic(t))

        self.lbl(right_col, "PLAYER HUB", 10, MUTED, True).pack(anchor="w", pady=(0, 8))
        panel = tk.Frame(right_col, bg=CARD)
        panel.pack(fill="both", expand=True)

        self.lbl(panel, "⚡ QUICK ACTIONS", 9, PURPLE, True).pack(anchor="w", padx=18, pady=(18, 10))
        self.btn(panel, "🎲 Random Topic", self.random_topic, CARD2).pack(fill="x", padx=16, pady=4)
        self.btn(panel, "🏆 Achievements", self.show_achievements, CARD2).pack(fill="x", padx=16, pady=4)
        self.btn(panel, "📊 My Statistics", self.show_statistics, CARD2).pack(fill="x", padx=16, pady=4)
        self.btn(panel, "⚙ Profile", self.profile_dialog, CARD2).pack(fill="x", padx=16, pady=4)

        self.lbl(panel, "RECENT RUNS", 9, CYAN, True).pack(anchor="w", padx=18, pady=(20, 8))
        if not self.profile["history"]:
            self.lbl(panel, "No runs yet.\nYour first record starts here.", 9, MUTED, wraplength=220).pack(anchor="w", padx=18)
        else:
            for h in self.profile["history"][-4:][::-1]:
                self.lbl(panel, f"{h['topic']}  •  {h['score']}%  •  +{h['xp']} XP", 8, "#b5bfd7").pack(anchor="w", padx=18, pady=3)

    def stat_box(self, parent, title, value, color):
        f = tk.Frame(parent, bg=CARD2)
        f.pack(side="left", padx=10, pady=14)
        self.lbl(f, title, 7, MUTED, True).pack()
        self.lbl(f, value, 15, color, True).pack(pady=(2, 0))

    def topic_mastery(self, topic):
        played = [h for h in self.profile["history"] if h["topic"] == topic]
        if not played:
            return 0
        return min(100, round(sum(h["score"] for h in played) / len(played)))

    def select_topic(self, topic):
        self.topic = topic
        self.open_setup()

    def random_topic(self):
        self.topic = random.choice(list(TOPICS))
        self.open_setup()

    def open_setup(self):
        self.clear()
        self.lbl(self.root, "BUILD YOUR RUN", 27, TEXT, True).pack(anchor="w", padx=60, pady=(35, 3))
        self.lbl(self.root, "Everything here changes how the challenge feels.", 10, MUTED).pack(anchor="w", padx=60)

        main = tk.Frame(self.root, bg=APP_BG)
        main.pack(fill="both", expand=True, padx=60, pady=28)

        left = tk.Frame(main, bg=CARD)
        left.pack(side="left", fill="both", expand=True, padx=(0, 12), ipadx=10)

        self.lbl(left, "TOPIC", 9, MUTED, True).pack(anchor="w", padx=25, pady=(25, 10))
        self.setup_topic = tk.StringVar(value=self.topic)
        topic_box = ttk.Combobox(left, textvariable=self.setup_topic, values=list(TOPICS), state="readonly", font=("Segoe UI", 11))
        topic_box.pack(fill="x", padx=25, ipady=7)

        self.lbl(left, "DIFFICULTY", 9, MUTED, True).pack(anchor="w", padx=25, pady=(20, 10))
        self.setup_diff = tk.StringVar(value=self.difficulty)
        for d in DIFFICULTY:
            self.radio(left, d, self.setup_diff, d == self.difficulty).pack(fill="x", padx=25, pady=3)

        self.lbl(left, "QUESTIONS", 9, MUTED, True).pack(anchor="w", padx=25, pady=(20, 10))
        self.setup_count = tk.StringVar(value=str(self.count))
        count_box = ttk.Combobox(left, textvariable=self.setup_count, values=["5", "10", "15", "20"], state="readonly", font=("Segoe UI", 11))
        count_box.pack(fill="x", padx=25, ipady=7)

        right = tk.Frame(main, bg=CARD)
        right.pack(side="right", fill="both", expand=True, padx=(12, 0))
        self.lbl(right, "GAME MODE", 9, MUTED, True).pack(anchor="w", padx=25, pady=(25, 10))

        self.setup_mode = tk.StringVar(value=self.mode)
        for mode, (icon, desc) in MODES.items():
            f = tk.Frame(right, bg=CARD2, cursor="hand2")
            f.pack(fill="x", padx=25, pady=5)
            rb = tk.Radiobutton(f, text=f"{icon}  {mode}", variable=self.setup_mode, value=mode,
                                bg=CARD2, fg=TEXT, selectcolor=CARD3, activebackground=CARD2,
                                activeforeground=TEXT, font=("Segoe UI", 10, "bold"), anchor="w")
            rb.pack(anchor="w", padx=12, pady=(9, 0))
            self.lbl(f, desc, 8, MUTED).pack(anchor="w", padx=37, pady=(0, 9))
            f.bind("<Button-1>", lambda e, m=mode: self.setup_mode.set(m))

        self.btn(self.root, "← Back", self.show_home, CARD3).pack(side="left", padx=60, pady=18)
        self.btn(self.root, "🔥  ENTER ARENA", self.launch_run, PURPLE, size=12).pack(side="right", padx=60, pady=18)

    def radio(self, parent, text, variable, selected):
        return tk.Radiobutton(parent, text=text, variable=variable, value=text,
                              bg=CARD, fg=TEXT, selectcolor=CARD3, activebackground=CARD,
                              activeforeground=TEXT, font=("Segoe UI", 10), anchor="w")

    def launch_run(self):
        self.topic = self.setup_topic.get()
        self.difficulty = self.setup_diff.get()
        self.mode = self.setup_mode.get()
        self.count = int(self.setup_count.get())
        self.profile["name"] = self.profile["name"] or "Player"

        bank = TOPICS[self.topic]["questions"]
        self.questions = random.sample(bank, min(self.count, len(bank)))
        self.answers = [None] * len(self.questions)
        self.current = 0
        self.streak = 0
        self.best_streak = 0
        self.correct = 0
        self.xp_run = 0
        self.lives = 3
        self.session_topic_set.add(self.topic)
        self.show_game()

    def show_game(self):
        self.clear()
        self.answer_locked = False

        top = tk.Frame(self.root, bg=CARD)
        top.pack(fill="x")
        self.lbl(top, f"{TOPICS[self.topic]['icon']}  {self.topic.upper()}", 13, TEXT, True).pack(side="left", padx=30, pady=20)
        self.game_stats = self.lbl(top, "", 9, GREEN, True)
        self.game_stats.pack(side="right", padx=30)

        content = tk.Frame(self.root, bg=APP_BG)
        content.pack(fill="both", expand=True, padx=60, pady=25)

        self.question_no = self.lbl(content, "", 9, MUTED, True)
        self.question_no.pack(anchor="w")
        self.progress = ttk.Progressbar(content, style="Game.Horizontal.TProgressbar", maximum=len(self.questions))
        self.progress.pack(fill="x", pady=(6, 22))

        card = tk.Frame(content, bg=CARD)
        card.pack(fill="both", expand=True)

        top_line = tk.Frame(card, bg=CARD)
        top_line.pack(fill="x", padx=28, pady=(22, 0))
        self.mode_label = self.lbl(top_line, "", 8, PURPLE, True)
        self.mode_label.pack(side="left")
        self.timer_label = self.lbl(top_line, "", 12, YELLOW, True)
        self.timer_label.pack(side="right")

        self.streak_label = self.lbl(card, "", 11, ORANGE, True)
        self.streak_label.pack(anchor="center", pady=(15, 0))

        self.question = self.lbl(card, "", 21, TEXT, True, wraplength=850, justify="center")
        self.question.pack(fill="x", padx=50, pady=(12, 25))

        self.answers_frame = tk.Frame(card, bg=CARD)
        self.answers_frame.pack(fill="x", padx=50)

        self.answer_buttons = []
        for i in range(4):
            b = tk.Button(self.answers_frame, text="", font=("Segoe UI", 11, "bold"),
                          bg=CARD2, fg=TEXT, activebackground=CARD3,
                          activeforeground=TEXT, bd=0, relief="flat",
                          cursor="hand2", anchor="w", padx=18, pady=13,
                          command=lambda n=i: self.answer(n))
            b.pack(fill="x", pady=5)
            self.answer_buttons.append(b)

        bottom = tk.Frame(content, bg=APP_BG)
        bottom.pack(fill="x", pady=(16, 0))
        self.btn(bottom, "🏠 Quit", self.quit_run, CARD3, fg="#d5dced").pack(side="left")
        self.life_label = self.lbl(bottom, "", 9, RED, True)
        self.life_label.pack(side="left", padx=25)
        self.btn(bottom, "💡 50/50", self.use_fifty, CARD3, CYAN).pack(side="right", padx=5)
        self.btn(bottom, "⏭ Skip", self.skip_question, CARD3, MUTED).pack(side="right", padx=5)

        self.load_question()

    def load_question(self):
        self.answer_locked = False
        q = self.questions[self.current]
        self.selected_index = None
        self.question_no.config(text=f"QUESTION {self.current + 1} / {len(self.questions)}")
        self.progress["value"] = self.current + 1
        self.question.config(text=q[0])
        self.mode_label.config(text=f"{MODES[self.mode][0]}  {self.mode.upper()}  •  {self.difficulty.upper()}")
        self.streak_label.config(text=f"🔥 {self.streak} COMBO")
        self.life_label.config(text=f"❤ {self.lives}" if self.mode == "Survival" else "")
        for i, b in enumerate(self.answer_buttons):
            b.config(text=f"{chr(65+i)}   {q[1][i]}", state="normal", bg=CARD2, fg=TEXT)
        if self.mode == "Blitz":
            self.time_left = max(6, DIFFICULTY[self.difficulty]["time"] - 5)
        else:
            self.time_left = DIFFICULTY[self.difficulty]["time"]
        self.tick()

    def tick(self):
        self.timer_label.config(text=f"⏱ {self.time_left}s")
        if self.time_left <= 0:
            self.answer(None, timed_out=True)
            return
        self.time_left -= 1
        self.timer_id = self.root.after(1000, self.tick)

    def stop_timer(self):
        if self.timer_id:
            self.root.after_cancel(self.timer_id)
            self.timer_id = None

    def answer(self, index, timed_out=False):
        if self.answer_locked:
            return
        self.answer_locked = True
        self.stop_timer()

        q = self.questions[self.current]
        correct = q[2]
        self.answers[self.current] = index

        if index == correct:
            self.correct += 1
            self.streak += 1
            self.best_streak = max(self.best_streak, self.streak)
            base = int(100 * DIFFICULTY[self.difficulty]["mult"])
            speed_bonus = self.time_left * 3
            combo_bonus = max(0, self.streak - 1) * 25
            mode_bonus = 40 if self.mode == "Streak Rush" else 0
            gained = base + speed_bonus + combo_bonus + mode_bonus
            self.xp_run += gained
            self.answer_buttons[index].config(bg="#145c4b", fg=GREEN)
        else:
            self.streak = 0
            if index is not None:
                self.answer_buttons[index].config(bg="#65283a", fg=RED)
            if self.mode == "Survival":
                self.lives -= 1
                if self.lives <= 0:
                    self.root.after(550, self.finish_run)
                    return

        self.answer_buttons[correct].config(bg="#145c4b", fg=GREEN)
        self.game_stats.config(text=f"✦ +{self.xp_run} XP     ✓ {self.correct}/{self.current + 1}")
        self.streak_label.config(text=f"🔥 {self.streak} COMBO")
        self.life_label.config(text=f"❤ {self.lives}" if self.mode == "Survival" else "")
        for b in self.answer_buttons:
            b.config(state="disabled")
        self.root.after(850, self.next_question)

    def next_question(self):
        if self.current < len(self.questions) - 1:
            self.current += 1
            self.load_question()
        else:
            self.finish_run()

    def skip_question(self):
        if self.answer_locked:
            return
        self.answer(None)

    def use_fifty(self):
        if self.answer_locked:
            return
        q = self.questions[self.current]
        wrong = [i for i in range(4) if i != q[2]]
        remove = random.sample(wrong, 2)
        for i in remove:
            self.answer_buttons[i].config(text="—", state="disabled", fg=MUTED)

    def quit_run(self):
        if messagebox.askyesno("Leave arena?", "Your current run will not be saved. Leave?"):
            self.show_home()

    def finish_run(self):
        self.stop_timer()
        total = len(self.questions)
        accuracy = round((self.correct / total) * 100)
        old_level = self.profile["level"]

        self.profile["xp"] += self.xp_run
        self.profile["games"] += 1
        self.profile["correct"] += self.correct
        self.profile["questions"] += total
        self.profile["best_streak"] = max(self.profile["best_streak"], self.best_streak)
        if self.topic not in self.profile["topics"]:
            self.profile["topics"].append(self.topic)

        unlocked = []
        if self.profile["games"] == 1 and self.award_achievement("First Run"):
            unlocked.append("🚀 First Run")
        if accuracy == 100 and self.award_achievement("Perfect"):
            unlocked.append("💎 Perfect")
        if self.best_streak >= 5 and self.award_achievement("Combo 5"):
            unlocked.append("🔥 Combo 5")
        if self.best_streak >= 10 and self.award_achievement("Combo 10"):
            unlocked.append("🌋 Combo 10")
        if self.profile["xp"] >= 2000 and self.award_achievement("Scholar"):
            unlocked.append("🧠 Scholar")
        if len(self.profile["topics"]) >= 5 and self.award_achievement("Explorer"):
            unlocked.append("🗺 Explorer")
        if accuracy >= 90 and self.award_achievement("Speed Demon"):
            unlocked.append("⚡ Speed Demon")
        if self.mode == "Survival" and self.award_achievement("Survivor"):
            unlocked.append("☠ Survivor")

        level, _, _ = self.level_info()
        level_up = level > old_level

        self.profile["history"].append({
            "topic": self.topic, "score": accuracy, "xp": self.xp_run,
            "mode": self.mode, "date": time.strftime("%Y-%m-%d")
        })
        self.profile["history"] = self.profile["history"][-20:]
        self.save_profile()

        self.show_results(accuracy, unlocked, level_up)

    def show_results(self, accuracy, unlocked, level_up):
        self.clear()
        if accuracy >= 90:
            headline = "ABSOLUTE BEAST. 🏆"
            sub = "That run was seriously clean."
        elif accuracy >= 75:
            headline = "YOU COOKED. 🔥"
            sub = "Strong run. Your brain is leveling up."
        elif accuracy >= 50:
            headline = "SOLID RUN. ⚡"
            sub = "You're building the foundation. Keep going."
        else:
            headline = "RUN COMPLETE. 💪"
            sub = "Now you know exactly what to improve."

        self.lbl(self.root, headline, 29, TEXT, True).pack(pady=(35, 3))
        self.lbl(self.root, sub, 11, MUTED).pack()

        card = tk.Frame(self.root, bg=CARD)
        card.pack(fill="x", padx=100, pady=24)
        results = [
            ("SCORE", f"{self.correct}/{len(self.questions)}", PURPLE),
            ("ACCURACY", f"{accuracy}%", CYAN),
            ("BEST COMBO", str(self.best_streak), ORANGE),
            ("XP", f"+{self.xp_run}", GREEN)
        ]
        for i, (title, val, col) in enumerate(results):
            card.grid_columnconfigure(i, weight=1)
            f = tk.Frame(card, bg=CARD)
            f.grid(row=0, column=i, sticky="nsew", padx=10, pady=24)
            self.lbl(f, title, 8, MUTED, True).pack()
            self.lbl(f, val, 22, col, True).pack(pady=4)

        if level_up:
            self.lbl(self.root, f"🎉 LEVEL UP! You are now LEVEL {self.profile['level']}", 13, YELLOW, True).pack(pady=(0, 8))

        if unlocked:
            self.lbl(self.root, "NEW ACHIEVEMENTS", 9, PURPLE, True).pack()
            self.lbl(self.root, "   ".join(unlocked), 10, TEXT, True).pack(pady=5)

        self.lbl(self.root, "RUN BREAKDOWN", 9, MUTED, True).pack(anchor="w", padx=100, pady=(16, 7))
        review = tk.Frame(self.root, bg=CARD)
        review.pack(fill="both", expand=True, padx=100)

        txt = tk.Text(review, bg=CARD, fg=TEXT, font=("Segoe UI", 10), bd=0, padx=16, pady=12, wrap="word")
        sb = tk.Scrollbar(review, command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        txt.pack(side="left", fill="both", expand=True)

        for i, (q, options, correct, explanation) in enumerate(self.questions):
            ans = self.answers[i]
            if ans == correct:
                txt.insert("end", f"✓ Q{i+1}  {q}\n", "good")
            else:
                your = "Skipped" if ans is None else options[ans]
                txt.insert("end", f"✗ Q{i+1}  Your answer: {your}  |  Correct: {options[correct]}\n", "bad")
            txt.insert("end", f"   {explanation}\n\n", "muted")

        txt.tag_config("good", foreground=GREEN, font=("Segoe UI", 10, "bold"))
        txt.tag_config("bad", foreground=RED, font=("Segoe UI", 10, "bold"))
        txt.tag_config("muted", foreground=MUTED)
        txt.config(state="disabled")

        bottom = tk.Frame(self.root, bg=APP_BG)
        bottom.pack(pady=18)
        self.btn(bottom, "🏠 Home", self.show_home, CARD3).pack(side="left", padx=5)
        self.btn(bottom, "🔄 Play Again", self.replay, PURPLE).pack(side="left", padx=5)

    def replay(self):
        bank = TOPICS[self.topic]["questions"]
        self.questions = random.sample(bank, min(self.count, len(bank)))
        self.answers = [None] * len(self.questions)
        self.current = 0
        self.streak = 0
        self.best_streak = 0
        self.correct = 0
        self.xp_run = 0
        self.lives = 3
        self.show_game()

    def show_achievements(self):
        self.clear()
        self.lbl(self.root, "ACHIEVEMENT VAULT", 26, TEXT, True).pack(anchor="w", padx=55, pady=(35, 4))
        self.lbl(self.root, "Collect them all. Some are easy. Some are nasty.", 10, MUTED).pack(anchor="w", padx=55)
        grid = tk.Frame(self.root, bg=APP_BG)
        grid.pack(fill="both", expand=True, padx=55, pady=25)
        for i, (name, (icon, desc)) in enumerate(ACHIEVEMENTS.items()):
            r, c = divmod(i, 2)
            unlocked = name in self.profile["achievements"]
            card = tk.Frame(grid, bg=CARD2 if unlocked else CARD)
            card.grid(row=r, column=c, padx=8, pady=8, sticky="nsew")
            grid.grid_columnconfigure(c, weight=1)
            grid.grid_rowconfigure(r, weight=1)
            self.lbl(card, f"{icon}  {name}", 13, YELLOW if unlocked else MUTED, True).pack(anchor="w", padx=18, pady=(16, 5))
            self.lbl(card, desc, 9, TEXT if unlocked else "#56617a").pack(anchor="w", padx=18)
            self.lbl(card, "UNLOCKED" if unlocked else "LOCKED", 7, GREEN if unlocked else "#56617a", True).pack(anchor="w", padx=18, pady=(10, 16))
        self.btn(self.root, "← Back", self.show_home, CARD3).pack(pady=18)

    def show_statistics(self):
        self.clear()
        self.lbl(self.root, "PLAYER STATISTICS", 26, TEXT, True).pack(anchor="w", padx=55, pady=(35, 4))
        self.lbl(self.root, "Your progress across every run.", 10, MUTED).pack(anchor="w", padx=55)
        panel = tk.Frame(self.root, bg=CARD)
        panel.pack(fill="x", padx=55, pady=25)

        total = self.profile["questions"]
        accuracy = round(self.profile["correct"] / total * 100) if total else 0
        stats = [
            ("TOTAL XP", self.profile["xp"], PURPLE),
            ("QUIZZES", self.profile["games"], CYAN),
            ("QUESTIONS", total, TEXT),
            ("GLOBAL ACCURACY", f"{accuracy}%", GREEN),
            ("BEST COMBO", self.profile["best_streak"], ORANGE),
            ("TOPICS PLAYED", len(self.profile["topics"]), YELLOW)
        ]
        for i, (a, b, col) in enumerate(stats):
            r, c = divmod(i, 3)
            f = tk.Frame(panel, bg=CARD)
            f.grid(row=r, column=c, padx=20, pady=18, sticky="w")
            self.lbl(f, a, 8, MUTED, True).pack(anchor="w")
            self.lbl(f, str(b), 20, col, True).pack(anchor="w")

        self.lbl(self.root, "TOPIC MASTERY", 10, PURPLE, True).pack(anchor="w", padx=55)
        mastery = tk.Frame(self.root, bg=CARD)
        mastery.pack(fill="x", padx=55, pady=10)
        for i, topic in enumerate(TOPICS):
            col = i % 2
            row = i // 2
            f = tk.Frame(mastery, bg=CARD)
            f.grid(row=row, column=col, padx=18, pady=8, sticky="ew")
            mastery.grid_columnconfigure(col, weight=1)
            self.lbl(f, f"{TOPICS[topic]['icon']} {topic}", 9, TEXT, True).pack(side="left")
            self.lbl(f, f"{self.topic_mastery(topic)}%", 9, GREEN, True).pack(side="right")
        self.btn(self.root, "← Back", self.show_home, CARD3).pack(pady=20)

    def profile_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("Profile")
        win.geometry("390x230")
        win.configure(bg=CARD)
        win.resizable(False, False)
        self.lbl(win, "YOUR PROFILE", 17, TEXT, True).pack(pady=(25, 15))
        self.lbl(win, "Display name", 9, MUTED).pack()
        entry = tk.Entry(win, bg=CARD2, fg=TEXT, insertbackground=TEXT, bd=0, font=("Segoe UI", 11))
        entry.pack(fill="x", padx=45, pady=7, ipady=8)
        entry.insert(0, self.profile["name"])
        def save():
            name = entry.get().strip()
            if name:
                self.profile["name"] = name[:25]
                self.save_profile()
                win.destroy()
                self.show_home()
        self.btn(win, "SAVE PROFILE", save, PURPLE).pack(pady=15)

def main():
    root = tk.Tk()
    try:
        TOPICS.update(load_topics())
    except Exception as e:
        messagebox.showerror("QuizVerse",
                             f"Could not load questions from:\n{QUESTIONS_FILE}\n\n{e}")
        root.destroy()
        return
    QuizVerse(root)
    root.mainloop()

if __name__ == "__main__":
    main()
