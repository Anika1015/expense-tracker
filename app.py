import tkinter as tk
from tkinter import messagebox, ttk
import os
import sys
import csv
from datetime import datetime
from collections import defaultdict
from itertools import groupby

# Always save CSV next to app.py, no matter where terminal is opened from
BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
FILE     = os.path.join(BASE_DIR, "expenses.csv")
HEADER   = ["NAME", "DATE", "AMOUNT", "MONTHLY TOTALS", "AMOUNT"]

_last_date = datetime.today().strftime("%d-%m-%Y")

# ── date string parser: accepts DD-MM-YYYY and DD-MM-YY ──────────────────────
def parse_date_str(date_str):
    for fmt in ("%d-%m-%Y", "%d-%m-%y"):
        try:
            d = datetime.strptime(date_str.strip(), fmt)
            # always normalise to 4-digit year DD-MM-YYYY
            return d.strftime("%d-%m-%Y")
        except ValueError:
            continue
    return None

# ── row validator: Filters out valid expenses from total summary rows ─────────
def is_real_expense(row):
    if not row or len(row) < 3:
        return False
    name_col   = str(row[0]).strip()
    date_col   = str(row[1]).strip()
    amount_col = str(row[2]).strip()
    if not name_col or not date_col or not amount_col or name_col == "None":
        return False
    if name_col == "NAME" or name_col.startswith(("-", "=")) or "Total" in name_col:
        return False
    try:
        float(amount_col)
        return parse_date_str(date_col) is not None
    except (ValueError, IndexError):
        return False

# ── load: Reads raw expense data cleanly from the CSV file ────────────────────
def load_expenses():
    if not os.path.exists(FILE):
        return []
    rows = []
    try:
        with open(FILE, mode="r", newline="", encoding="utf-8") as f:
            reader = csv.reader(f)
            next(reader, None)  # Skip the header row
            for row in reader:
                if not row:
                    continue
                if not is_real_expense(row):
                    continue
                normalised_date = parse_date_str(str(row[1]).strip()) or str(row[1]).strip()
                rows.append([str(row[0]).strip(), normalised_date, str(row[2]).strip()])
    except Exception:
        pass
    return rows

# ── date sort helper ──────────────────────────────────────────────────────────
def parse_date(row):
    d_str = parse_date_str(row[1])
    if d_str:
        return datetime.strptime(d_str, "%d-%m-%Y")
    return datetime.min

# ── save: Writes daily & monthly data side-by-side without decimals ───────────
def save_expenses(data):
    try:
        sorted_data = sorted(data, key=parse_date, reverse=True)
        monthly = defaultdict(int)

        for row in sorted_data:
            try:
                d = datetime.strptime(row[1], "%d-%m-%Y")
                monthly[d.strftime("%B %Y")] += int(float(row[2]))
            except Exception:
                pass

        sorted_months = sorted(monthly.items(),
                               key=lambda x: datetime.strptime(x[0], "%B %Y"),
                               reverse=True)

        excel_rows = []
        for date_val, grp in groupby(sorted_data, key=lambda r: r[1]):
            grp = list(grp)
            day_total = int(sum(float(r[2]) for r in grp))
            for r in grp:
                excel_rows.append([r[0], r[1], int(float(r[2]))])
            excel_rows.append([f"-- Total ({date_val}) --", "", day_total])

        max_len = max(len(excel_rows), len(sorted_months)) if (excel_rows or sorted_months) else 0

        with open(FILE, mode="w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(HEADER)

            for i in range(max_len):
                row_to_write = ["", "", "", "", ""]
                if i < len(excel_rows):
                    row_to_write[0] = excel_rows[i][0]
                    row_to_write[1] = excel_rows[i][1]
                    row_to_write[2] = excel_rows[i][2]
                if i < len(sorted_months):
                    month_name, month_sum = sorted_months[i]
                    row_to_write[3] = f"Total for {month_name}"
                    row_to_write[4] = month_sum
                writer.writerow(row_to_write)

    except PermissionError:
        messagebox.showerror("File is Open",
            f"Cannot save!\n\nPlease close '{FILE}' in Excel/Text Editor and try again.")

# ── startup load ──────────────────────────────────────────────────────────────
expenses = load_expenses()

# ── refresh tree ──────────────────────────────────────────────────────────────
def refresh_list(data=None):
    if data is None:
        data = expenses
    for row in tree.get_children():
        tree.delete(row)

    sorted_data = sorted(data, key=parse_date, reverse=True)

    for date_val, grp in groupby(sorted_data, key=lambda r: r[1]):
        grp       = list(grp)
        day_total = 0
        for row in grp:
            try:
                amt      = int(float(row[2]))
                orig_idx = next((i for i, e in enumerate(expenses)
                                 if e[0]==row[0] and e[1]==row[1] and e[2]==row[2]), -1)
                tree.insert("", tk.END, iid=str(orig_idx),
                            values=(orig_idx+1, row[0], row[1], f"৳{amt}"),
                            tags=("expense",))
                day_total += amt
            except ValueError:
                continue

        tree.insert("", tk.END,
                    values=("", "", f"Total({date_val})", f"৳{day_total}"),
                    tags=("subtotal",))


# ── add ───────────────────────────────────────────────────────────────────────
def add_expense():
    global _last_date, expenses
    name   = name_entry.get().strip()
    date   = date_entry.get().strip()
    amount = amount_entry.get().strip()

    if not name or not date or not amount:
        messagebox.showwarning("Missing Info", "Please fill in all three fields.")
        return
    try:
        datetime.strptime(date, "%d-%m-%Y")
    except ValueError:
        messagebox.showerror("Invalid Date", "Use DD-MM-YYYY format.\nExample: 17-06-2026")
        return
    try:
        float(amount)
    except ValueError:
        messagebox.showerror("Invalid Amount", "Amount must be a number.")
        return

    try:
        if os.path.exists(FILE):
            with open(FILE, "r+"): pass
    except PermissionError:
        messagebox.showerror("File Open",
            f"Cannot add expense because '{FILE}' is open elsewhere!\n\nPlease close it and try again.")
        return

    expenses = load_expenses()
    expenses.append([name, date, str(int(float(amount)))])
    save_expenses(expenses)

    _last_date = date
    name_entry.delete(0, tk.END)
    amount_entry.delete(0, tk.END)
    name_entry.focus()
    refresh_list()

# ── delete (multi-select) ─────────────────────────────────────────────────────
def delete_expense():
    global expenses
    selected = tree.selection()
    valid    = [s for s in selected if "expense" in tree.item(s)["tags"]]
    if not valid:
        messagebox.showwarning("Nothing Selected",
            "Please select one or more expense rows.\n(Ctrl+click for multiple)")
        return
    names   = [tree.item(s)["values"][1] for s in valid]
    preview = "\n".join(f"• {n}" for n in names[:5])
    if len(names) > 5:
        preview += f"\n  ...and {len(names)-5} more"
    if not messagebox.askyesno("Confirm Delete",
                                f"Delete {len(valid)} item(s)?\n\n{preview}"):
        return

    try:
        if os.path.exists(FILE):
            with open(FILE, "r+"): pass
    except PermissionError:
        messagebox.showerror("File Locked", f"Close '{FILE}' before deleting items.")
        return

    expenses = load_expenses()
    for idx in sorted({int(s) for s in valid}, reverse=True):
        if 0 <= idx < len(expenses):
            expenses.pop(idx)
    save_expenses(expenses)
    refresh_list()

# ── filter ────────────────────────────────────────────────────────────────────
def filter_expenses():
    q = search_entry.get().strip().lower()
    filtered = ([r for r in expenses if q in r[1].lower() or q in r[0].lower()]
                if q else expenses)
    refresh_list(filtered)

# ══════════════════════════════════════════════════════════════════════════════
#  UI
# ══════════════════════════════════════════════════════════════════════════════
root = tk.Tk()
root.title("💰 Expense Tracker")
root.geometry("720x640")
root.resizable(False, False)
root.configure(bg="#f0f4f8")

tk.Label(root, text="💰 Daily Expense Tracker",
         font=("Arial", 17, "bold"), bg="#f0f4f8", fg="#1a3c6e").pack(pady=10)

inp = tk.LabelFrame(root, text="  Add New Expense  ",
                    font=("Arial", 10, "bold"),
                    bg="#f0f4f8", fg="#2e75b6", padx=12, pady=8)
inp.pack(padx=20, fill="x")

for i, lbl in enumerate(["Expense Name:", "Date (DD-MM-YYYY):", "Amount (৳):"]):
    tk.Label(inp, text=lbl, bg="#f0f4f8", font=("Arial", 10),
             anchor="e", width=18).grid(row=i, column=0, pady=4, sticky="e")

name_entry   = tk.Entry(inp, width=26, font=("Arial", 11))
date_entry   = tk.Entry(inp, width=26, font=("Arial", 11))
amount_entry = tk.Entry(inp, width=26, font=("Arial", 11))
name_entry.grid(  row=0, column=1, padx=10, pady=4, sticky="w")
date_entry.grid(  row=1, column=1, padx=10, pady=4, sticky="w")
amount_entry.grid(row=2, column=1, padx=10, pady=4, sticky="w")

tk.Label(inp, text="💡 Date kept from last entry",
         bg="#f0f4f8", fg="#888", font=("Arial", 8)).grid(row=1, column=2, sticky="w")
date_entry.insert(0, _last_date)

tk.Button(inp, text="➕  Add Expense", command=add_expense,
          bg="#2e75b6", fg="white", font=("Arial", 11, "bold"),
          padx=14, pady=6, relief="flat", cursor="hand2").grid(
              row=3, column=0, columnspan=2, pady=8)

sf = tk.Frame(root, bg="#f0f4f8")
sf.pack(padx=20, fill="x", pady=(2, 0))
tk.Label(sf, text="🔍 Filter (name / date):", bg="#f0f4f8",
         font=("Arial", 10)).pack(side="left")
search_entry = tk.Entry(sf, width=18, font=("Arial", 10))
search_entry.pack(side="left", padx=6)
tk.Button(sf, text="Search",   command=filter_expenses,
          bg="#5a9fd4", fg="white", font=("Arial", 9),
          relief="flat", padx=8, cursor="hand2").pack(side="left")
tk.Button(sf, text="Show All", command=lambda: refresh_list(),
          bg="#888",    fg="white", font=("Arial", 9),
          relief="flat", padx=8, cursor="hand2").pack(side="left", padx=4)

cols = ("#", "Expense Name", "Date", "Amount")
style = ttk.Style()
style.theme_use("clam")
style.configure("Treeview", font=("Arial", 10), rowheight=24,
                background="#ffffff", fieldbackground="#ffffff")
style.configure("Treeview.Heading", font=("Arial", 10, "bold"),
                background="#1a3c6e", foreground="white")
style.map("Treeview", background=[("selected", "#2e75b6")])

tf = tk.Frame(root)
tf.pack(padx=20, pady=6, fill="both")

tree = ttk.Treeview(tf, columns=cols, show="headings", height=12,
                    selectmode="extended")
for col, w, anc in zip(cols,
                        [36, 220, 130, 120],
                        ["center","w","center","center"]):
    tree.heading(col, text=col)
    tree.column(col, width=w, anchor=anc)

tree.tag_configure("expense",  background="#ffffff")
tree.tag_configure("subtotal", background="#dbeeff", font=("Arial", 9, "bold"))
tree.tag_configure("blank",    background="#f0f4f8")

sb = ttk.Scrollbar(tf, orient="vertical", command=tree.yview)
tree.configure(yscrollcommand=sb.set)
tree.pack(side="left", fill="both")
sb.pack(side="right", fill="y")

bot = tk.Frame(root, bg="#f0f4f8")
bot.pack(padx=20, fill="x", pady=6)
tk.Button(bot, text="🗑️  Delete Selected  (Ctrl+click for multi)",
          command=delete_expense,
          bg="#c0392b", fg="white", font=("Arial", 10),
          padx=10, pady=5, relief="flat", cursor="hand2").pack(side="left")

try:
    refresh_list()
except Exception:
    pass
root.mainloop()