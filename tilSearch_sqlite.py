import sqlite3, openpyxl, sys, unicodedata, os

DB_FILE = "phones.db"
EXCEL_FILE = "sample_data.xlsx"


def normalize_greek(text):
    # μετατρέπει σε πεζά string χωρίς κενά
    text = str(text).strip().casefold()
    # κάθε γράμμα με τόνο ή διαλυτικά σπάει σε δύο χαρακτήρες
    text = unicodedata.normalize("NFD", text)
    # κρατάει μόνο τα γράμματα
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = "".join(text.split())
    return text


def build_db():
    conn = sqlite3.connect(DB_FILE, isolation_level=None)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS phones (
            id      INTEGER PRIMARY KEY,
            sheet   TEXT NOT NULL,
            unit    TEXT NOT NULL,
            unit_norm TEXT NOT NULL,
            office  TEXT NOT NULL,
            office_norm TEXT NOT NULL,
            phone   TEXT NOT NULL
        )
    """)
    # Αδειάζει τον πίνακα ώστε να μην υπάρχουν διπλές εγγραφές
    conn.execute("DELETE FROM phones")

    wb = openpyxl.load_workbook(EXCEL_FILE)
    for sheet in wb.worksheets:
        # Βήμα 2 αφού οι στήλες πάνε ανά ζεύγη (όνομα|τηλέφωνο)
        for col in range(1, sheet.max_column + 1, 2):
            unit = sheet.cell(row=1, column=col).value
            unit_norm = normalize_greek(unit)
            # Οι γραμμές ξεκινάνε από 2 γιατί η γραμμή 1 είναι τίτλος της μονάδας και όχι γραφέιο
            for row in range(2, sheet.max_row + 1):
                office = sheet.cell(row=row, column=col).value
                phone  = sheet.cell(row=row, column=col + 1).value
                conn.execute(
                    "INSERT INTO phones (sheet, unit, unit_norm, office, office_norm, phone) VALUES (?,?,?,?,?,?)",
                    [sheet.title, str(unit).strip(), unit_norm,
                     str(office).strip(), normalize_greek(office),
                     str(phone).strip()]
                )
    conn.close()
    print("Βάση δεδομένων δημιουργήθηκε/ενημερώθηκε.")


def needs_rebuild():
    # Αν δεν υπάρχει db
    if not os.path.exists(DB_FILE):
        return True
    # Αν δεν υπάρχει excel
    if not os.path.exists(EXCEL_FILE):
        return False
    # Συγκρίνει την ώρα της τελευταίας τροποποίησης
    return os.path.getmtime(EXCEL_FILE) > os.path.getmtime(DB_FILE)


def word_conditions(field, words):
    sql = []
    params = []
    for w in words:
        sql.append(f"{field} LIKE ?") # π.χ. "unit_norm LIKE ?"
        params.append(f"%{w}%") # π.χ. "%μοναδα%"
    sql = " AND ".join(sql)
    return sql, params


def search(conn, category, query=None, unit_query=None, office_query=None):
    found = False
    total = 0

    if category == "1":
        words = []
        for w in query.split():
            words.append(normalize_greek(w))
        # Δημιουργία SQL συνθήκης για όλες τις λέξεις
        cond, params = word_conditions("unit_norm", words)
        # Εύρεση μονάδων που ταιριάζουν με την αναζήτηση
        units = conn.execute(
            f"SELECT DISTINCT unit, unit_norm FROM phones WHERE {cond} ORDER BY sheet, unit",
            params
        ).fetchall()
        for unit, unit_norm in units:
            print(f"\n=== {unit} ===")
            for office, phone in conn.execute(
                "SELECT office, phone FROM phones WHERE unit_norm=? ORDER BY id", [unit_norm]
            ).fetchall():
                print(f"{office:<50} {phone}")
                found = True
            total += 1

    elif category == "2":
        words = []
        for w in query.split():
            words.append(normalize_greek(w))
        # Δημιουργία SQL συνθήκης για όλες τις λέξεις
        cond, params = word_conditions("office_norm", words)
        # Εύρεση γραφείων που ταιριάζουν με την αναζήτηση
        rows = conn.execute(
            f"SELECT sheet, unit, office, phone FROM phones WHERE {cond} ORDER BY sheet, unit",
            params
        ).fetchall()
        for sheet, unit, office, phone in rows:
            print("=" * 30)
            print(f"{sheet} | {unit} -> {office} -> {phone}")
            total += 1
            found = True

    elif category == "3":
        words = []
        for w in query.split():
            words.append(normalize_greek(w))
        # Δημιουργία SQL συνθήκης για μονάδες και γραφεία
        u_cond, u_params = word_conditions("unit_norm", words)
        o_cond, o_params = word_conditions("office_norm", words)

        # Εύρεση μονάδων που ταιριάζουν με την αναζήτηση
        matching_units = conn.execute(
            f"SELECT DISTINCT unit, unit_norm FROM phones WHERE {u_cond}", u_params
        ).fetchall()
        # Δεν επιτρέπει διπλότυπα
        shown_unit_norms = set()
        for unit, unit_norm in matching_units:
            shown_unit_norms.add(unit_norm)
            print(f"\n=== {unit} ===")
            for office, phone in conn.execute(
                "SELECT office, phone FROM phones WHERE unit_norm=? ORDER BY id", [unit_norm]
            ).fetchall():
                print(f"{office:<50} {phone}")
                found = True
            total += 1
        
        # Αναζήτηση στα γραφεία — αγνοεί μονάδες που εμφανίστηκαν ήδη ολόκληρες παραπάνω. Δεν θέλουμε να εμφανιστούν τα γραφεία της ξανά μεμονωμένα στο δεύτερο.
        for sheet, unit, unit_norm, office, phone in conn.execute(
            f"SELECT sheet, unit, unit_norm, office, phone FROM phones WHERE {o_cond}", o_params
        ).fetchall():
            if unit_norm not in shown_unit_norms:
                print("=" * 30)
                print(f"{sheet} | {unit} -> {office} -> {phone}")
                total += 1
                found = True

    elif category == "4":
        u_words = []
        for w in unit_query.split():
            u_words.append(normalize_greek(w))
        o_words = []
        for w in office_query.split():
            o_words.append(normalize_greek(w))
        # Δημιουργία SQL συνθήκης για μονάδα και γραφείο
        u_cond, u_params = word_conditions("unit_norm", u_words)
        o_cond, o_params = word_conditions("office_norm", o_words)
        # Εύρεση εγγραφών που ταιριάζουν και στη μονάδα και στο γραφείο
        rows = conn.execute(
            f"SELECT unit, office, phone FROM phones WHERE {u_cond} AND {o_cond} ORDER BY unit",
            u_params + o_params
        ).fetchall()
        for unit, office, phone in rows:
            print("=" * 30)
            print(f"{unit} -> {office} -> {phone}")
            total += 1
            found = True

    elif category == "5":
        # Αναζήτηση με ακριβή αριθμό τηλεφώνου γι' αυτό χρησιμοποιούμε = και όχι LIKE
        rows = conn.execute(
            "SELECT sheet, unit, office FROM phones WHERE phone=? ORDER BY sheet, unit",
            [query.strip()]
        ).fetchall()
        for sheet, unit, office in rows:
            print("=" * 30)
            print(f"{sheet} | {unit} -> {office}")
            total += 1
            found = True

    return found, total


def print_menu():
    print("""
Επιλέξτε
1. Αναζήτηση μονάδας
2. Αναζήτηση γραφείου
3. Γενική αναζήτηση
4. Αναζήτηση μονάδας + γραφείου
5. Αναζήτηση τηλεφώνου
""")


def get_category():
    while True:
        category = input().strip()
        if category == "exit":
            sys.exit()
        if category in ("1", "2", "3", "4", "5"):
            return category
        print("Μη έγκυρη επιλογή.")


def main():
    # Αν το Excel είναι νεότερο από τη βάση ή η βάση δεν υπάρχει
    if needs_rebuild():
        print("Φόρτωση δεδομένων από Excel...")
        build_db()

    conn = sqlite3.connect(DB_FILE, isolation_level=None)
    print("Πληκτρολογήστε 'exit' για τερματισμό.")

    while True:
        print_menu()
        category = get_category()

        if category == "4":
            unit_query   = input("Μονάδα: ").strip()
            office_query = input("Γραφείο: ").strip()
            found, total = search(conn, category, unit_query=unit_query, office_query=office_query)
        else:
            query = input("Αναζήτηση: ").strip()
            found, total = search(conn, category, query=query)

        if not found:
            print("Δεν βρέθηκαν αποτελέσματα")
        else:
            print(f"\nΣύνολο αποτελεσμάτων: {total}")


if __name__ == "__main__":
    main()
