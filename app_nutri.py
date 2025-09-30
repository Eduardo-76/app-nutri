import os
import sys
import csv
import sqlite3
import datetime as dt
from pathlib import Path

# --------- WhatsApp (pywhatkit) ----------
# Em ambientes sem navegador padrão, você pode precisar setar o 'tab' manualmente.
try:
    import pywhatkit as kit
except Exception:
    kit = None

# --------- PySide6 (GUI) ----------
from PySide6 import QtCore, QtGui, QtWidgets
from PySide6.QtCore import Qt, QDate

# --------- PDF (ReportLab) ----------
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle


APP_NAME = "Nutricionista - Daniela V. C."
APP_DIR = Path.home() / "NutriCalendar"
DB_PATH = APP_DIR / "nutri_calendar.db"

# >>>>>>>> CONFIGURE AQUI SEU WHATSAPP PESSOAL <<<<<<<<
# Formato internacional com DDI: ex: +55DDDNUMERO
OWNER_WHATS = "+5586999999999"

# Envia automaticamente ao abrir (pode desativar aqui)
AUTO_SEND_ON_START = True

# ---- Helpers de data
def qdate_to_str(qd: QDate) -> str:
    if not qd or not qd.isValid():
        return ""
    return qd.toString("yyyy-MM-dd")

def str_to_qdate(s: str) -> QDate:
    try:
        y, m, d = map(int, s.split("-"))
        return QDate(y, m, d)
    except Exception:
        return QDate()

def human_date(s: str) -> str:
    try:
        y, m, d = map(int, s.split("-"))
        return f"{d:02d}/{m:02d}/{y}"
    except Exception:
        return s

# ---- Banco de dados
def ensure_db():
    APP_DIR.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(DB_PATH) as con:
        cur = con.cursor()
        cur.execute("""
        CREATE TABLE IF NOT EXISTS clients (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            age INTEGER,
            whatsapp TEXT,
            first_consult TEXT,
            plan_months INTEGER DEFAULT 5,     -- 1, 2, 3, 6, 12 
            paid_returns INTEGER DEFAULT 0,    -- quantos retornos pagos
            rescheduled INTEGER DEFAULT 0,     -- 0/1
            notes TEXT,
            hidden INTEGER DEFAULT 0           -- 0/1 (oculto)
        )""")
        cur.execute("""
        CREATE TABLE IF NOT EXISTS appointments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            client_id INTEGER NOT NULL,
            date TEXT NOT NULL,                -- yyyy-mm-dd
            kind TEXT DEFAULT 'retorno',       -- 'primeira' | 'retorno'
            status TEXT DEFAULT 'agendado',    -- 'agendado' | 'feito' | 'cancelado'
            FOREIGN KEY (client_id) REFERENCES clients(id) ON DELETE CASCADE
        )""")
        con.commit()

def connect_db():
    return sqlite3.connect(DB_PATH)

# ---- Modelo da Tabela (clientes visíveis)
class ClientsTableModel(QtCore.QAbstractTableModel):
    HEADERS = ["ID", "Cliente", "Idade", "WhatsApp", "1ª Consulta", "Plano (meses)", "Retornos pagos", "Reagendado", "Oculto"]

    def __init__(self, parent=None):
        super().__init__(parent)
        self.rows = []
        self.load()

    def load(self):
        self.beginResetModel()
        with connect_db() as con:
            cur = con.cursor()
            cur.execute("""
                SELECT id, name, age, whatsapp, first_consult, plan_months, paid_returns, rescheduled, hidden
                  FROM clients
                 ORDER BY hidden ASC, name COLLATE NOCASE ASC
            """)
            self.rows = cur.fetchall()
        self.endResetModel()

    def rowCount(self, parent=None):
        return len(self.rows)

    def columnCount(self, parent=None):
        return len(self.HEADERS)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        value = self.rows[index.row()][index.column()]
        if role in (Qt.DisplayRole, Qt.EditRole):
            if index.column() == 4 and value:
                return human_date(value)
            if index.column() in (7, 8):
                return "Sim" if value else "Não"
            return value
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole and orientation == Qt.Horizontal:
            return self.HEADERS[section]
        return super().headerData(section, orientation, role)

    def get_row(self, row):
        return self.rows[row] if 0 <= row < len(self.rows) else None

# ---- Diálogo de confirmação de remoção
def ask_yes_no(parent, title, text):
    box = QtWidgets.QMessageBox(parent)
    box.setWindowTitle(title)
    box.setText(text)
    box.setIcon(QtWidgets.QMessageBox.Question)
    box.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
    return box.exec() == QtWidgets.QMessageBox.Yes

# ---- Janela principal
class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(1120, 680)
        self.setWindowIcon(QtGui.QIcon())  # você pode setar um .ico aqui

        self.model = ClientsTableModel()
        self.build_ui()
        self.apply_pink_theme()
        self.load_table()

        if AUTO_SEND_ON_START:
            QtCore.QTimer.singleShot(2000, self.send_week_whatsapp_auto)

    # ---------- UI ----------
    def build_ui(self):
        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        layout = QtWidgets.QVBoxLayout(central)

        # Barra de botões
        top = QtWidgets.QHBoxLayout()
        self.btn_new = QtWidgets.QPushButton("Novo")
        self.btn_save = QtWidgets.QPushButton("Salvar")
        self.btn_pdf = QtWidgets.QPushButton("Relatório PDF")
        self.btn_csv = QtWidgets.QPushButton("Importar / Exportar CSV")
        self.btn_send_whats = QtWidgets.QPushButton("Enviar WhatsApp (semana)")
        top.addWidget(self.btn_new)
        top.addWidget(self.btn_save)
        top.addWidget(self.btn_pdf)
        top.addWidget(self.btn_csv)
        top.addStretch(1)
        top.addWidget(self.btn_send_whats)
        layout.addLayout(top)

        # Formulário + Consultas
        form_area = QtWidgets.QHBoxLayout()
        layout.addLayout(form_area, 1)

        # ---- formulário do cliente
        form = QtWidgets.QFormLayout()
        form.setLabelAlignment(Qt.AlignLeft)
        form_area.addLayout(form, 1)

        self.ed_id = QtWidgets.QLineEdit()
        self.ed_id.setReadOnly(True)
        self.ed_name = QtWidgets.QLineEdit()
        self.ed_age = QtWidgets.QSpinBox()
        self.ed_age.setRange(0, 120)
        self.ed_whats = QtWidgets.QLineEdit()
        self.dt_first = QtWidgets.QDateEdit(calendarPopup=True)
        self.dt_first.setDisplayFormat("dd/MM/yyyy")
        self.dt_first.setDate(QDate.currentDate())

        self.cb_plan = QtWidgets.QComboBox()
        self.cb_plan.addItems(["1", "2", "3", "6", "12"])
        self.sp_paid = QtWidgets.QSpinBox()
        self.sp_paid.setRange(0, 12)
        self.chk_resched = QtWidgets.QCheckBox("Reagendado")
        self.chk_hidden = QtWidgets.QCheckBox("Ocultar cliente")
        self.ed_notes = QtWidgets.QPlainTextEdit()
        self.ed_notes.setPlaceholderText("Observações sobre a consulta...")

        form.addRow("ID:", self.ed_id)
        form.addRow("Nome do cliente:", self.ed_name)
        form.addRow("Idade:", self.ed_age)
        form.addRow("WhatsApp:", self.ed_whats)
        form.addRow("Primeira consulta:", self.dt_first)
        form.addRow("Plano (meses):", self.cb_plan)
        form.addRow("Retornos pagos:", self.sp_paid)
        form.addRow("", self.chk_resched)
        form.addRow("", self.chk_hidden)
        form.addRow("Observações:", self.ed_notes)

        # ---- lista de próximas consultas
        right = QtWidgets.QVBoxLayout()
        form_area.addLayout(right, 1)

        lbl_next = QtWidgets.QLabel("Próximas consultas")
        lbl_next.setAlignment(Qt.AlignCenter)
        self.list_next = QtWidgets.QListWidget()
        btns_next = QtWidgets.QHBoxLayout()
        self.dt_next = QtWidgets.QDateEdit(calendarPopup=True)
        self.dt_next.setDisplayFormat("dd/MM/yyyy")
        self.dt_next.setDate(QDate.currentDate().addDays(7))
        self.btn_add_next = QtWidgets.QPushButton("Adicionar")
        self.btn_del_next = QtWidgets.QPushButton("Remover")
        btns_next.addWidget(self.dt_next)
        btns_next.addWidget(self.btn_add_next)
        btns_next.addWidget(self.btn_del_next)
        right.addWidget(lbl_next)
        right.addWidget(self.list_next, 1)
        right.addLayout(btns_next)

        # ---- tabela de clientes
        self.table = QtWidgets.QTableView()
        self.table.setModel(self.model)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setAlternatingRowColors(True)
        layout.addWidget(self.table, 2)

        # ---- barra inferior
        bottom = QtWidgets.QHBoxLayout()
        self.btn_update = QtWidgets.QPushButton("Atualizar Selecionado")
        self.btn_remove = QtWidgets.QPushButton("Remover Selecionado")
        bottom.addWidget(self.btn_update)
        bottom.addWidget(self.btn_remove)
        layout.addLayout(bottom)

        # sinais
        self.btn_new.clicked.connect(self.clear_form)
        self.btn_add_next.clicked.connect(self.add_next_date)
        self.btn_del_next.clicked.connect(self.remove_selected_next)
        self.btn_save.clicked.connect(self.save_client)
        self.btn_update.clicked.connect(self.update_selected)
        self.btn_remove.clicked.connect(self.remove_selected)
        self.btn_pdf.clicked.connect(self.export_pdf)
        self.btn_csv.clicked.connect(self.csv_dialog)
        self.btn_send_whats.clicked.connect(self.send_week_whatsapp)
        self.table.doubleClicked.connect(self.table_double_clicked)

    def apply_pink_theme(self):
        # paleta base
        palette = self.palette()
        palette.setColor(QtGui.QPalette.Window, QtGui.QColor("#ffc0cb"))       # fundo principal
        palette.setColor(QtGui.QPalette.Base, QtGui.QColor("#ffe0ea"))         # campos
        palette.setColor(QtGui.QPalette.AlternateBase, QtGui.QColor("#ffd3e0"))
        palette.setColor(QtGui.QPalette.Button, QtGui.QColor("#ff8fb1"))
        palette.setColor(QtGui.QPalette.Highlight, QtGui.QColor("#ff4f86"))
        palette.setColor(QtGui.QPalette.ButtonText, QtGui.QColor("#2b0020"))
        palette.setColor(QtGui.QPalette.Text, QtGui.QColor("#2b0020"))
        palette.setColor(QtGui.QPalette.WindowText, QtGui.QColor("#2b0020"))
        self.setPalette(palette)

        # stylesheet (botões com cantos arredondados)
        self.setStyleSheet("""
            QWidget { font-family: 'Segoe UI', Arial; font-size: 12pt; }
            QGroupBox, QLabel { color: #2b0020; }
            QPushButton {
                background: #ff77a9; color:#2b0020; border: 2px solid #ff4f86;
                border-radius: 12px; padding: 8px 14px; font-weight: 600;
            }
            QPushButton:hover { background:#ff95bb; }
            QLineEdit, QSpinBox, QDateEdit, QPlainTextEdit, QComboBox, QListWidget {
                border: 2px solid #ff4f86; border-radius: 10px; padding:6px; background:#ffe0ea;
            }
            QTableView { gridline-color:#ff4f86; }
            QHeaderView::section { background:#ff95bb; padding:6px; border:1px solid #ff4f86; }
            QCheckBox { font-weight:600; }
        """)

    # ---------- Ações ----------
    def load_table(self):
        self.model.load()
        self.table.resizeColumnsToContents()

    def clear_form(self):
        self.ed_id.clear()
        self.ed_name.clear()
        self.ed_age.setValue(0)
        self.ed_whats.clear()
        self.dt_first.setDate(QDate.currentDate())
        self.cb_plan.setCurrentIndex(0)
        self.sp_paid.setValue(0)
        self.chk_resched.setChecked(False)
        self.chk_hidden.setChecked(False)
        self.ed_notes.clear()
        self.list_next.clear()

    def add_next_date(self):
        d = self.dt_next.date()
        self.list_next.addItem(d.toString("yyyy-MM-dd"))

    def remove_selected_next(self):
        for it in self.list_next.selectedItems():
            self.list_next.takeItem(self.list_next.row(it))

    def table_double_clicked(self, index: QtCore.QModelIndex):
        row = self.model.get_row(index.row())
        if not row:
            return
        client_id = row[0]
        self.fill_form_from_db(client_id)

    def fill_form_from_db(self, client_id: int):
        with connect_db() as con:
            cur = con.cursor()
            cur.execute("SELECT id,name,age,whatsapp,first_consult,plan_months,paid_returns,rescheduled,notes,hidden FROM clients WHERE id=?", (client_id,))
            row = cur.fetchone()
            if not row:
                return
            self.ed_id.setText(str(row[0]))
            self.ed_name.setText(row[1] or "")
            self.ed_age.setValue(row[2] or 0)
            self.ed_whats.setText(row[3] or "")
            self.dt_first.setDate(str_to_qdate(row[4] or ""))
            self.cb_plan.setCurrentText(str(row[5] or 2))
            self.sp_paid.setValue(row[6] or 0)
            self.chk_resched.setChecked(bool(row[7]))
            self.ed_notes.setPlainText(row[8] or "")
            self.chk_hidden.setChecked(bool(row[9]))
            # carregar próximas
            self.list_next.clear()
            cur.execute("SELECT date FROM appointments WHERE client_id=? AND status='agendado' ORDER BY date ASC", (client_id,))
            for (d,) in cur.fetchall():
                self.list_next.addItem(d)

    def save_client(self):
        name = self.ed_name.text().strip()
        if not name:
            QtWidgets.QMessageBox.warning(self, "Validação", "Informe o nome do cliente.")
            return
        with connect_db() as con:
            cur = con.cursor()
            cur.execute("""
                INSERT INTO clients (name, age, whatsapp, first_consult, plan_months, paid_returns, rescheduled, notes, hidden)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                name,
                int(self.ed_age.value()),
                self.ed_whats.text().strip(),
                qdate_to_str(self.dt_first.date()),
                int(self.cb_plan.currentText()),
                int(self.sp_paid.value()),
                1 if self.chk_resched.isChecked() else 0,
                self.ed_notes.toPlainText().strip(),
                1 if self.chk_hidden.isChecked() else 0
            ))
            client_id = cur.lastrowid
            # salvar próximas consultas
            dates = [self.list_next.item(i).text() for i in range(self.list_next.count())]
            for d in dates:
                cur.execute("INSERT INTO appointments (client_id, date, kind, status) VALUES (?, ?, 'retorno', 'agendado')",
                            (client_id, d))
            con.commit()
        self.clear_form()
        self.load_table()
        QtWidgets.QMessageBox.information(self, "Salvo", "Cliente salvo com sucesso.")

    def current_selected_client_id(self):
        idx = self.table.currentIndex()
        if not idx.isValid():
            return None
        row = self.model.get_row(idx.row())
        return row[0] if row else None

    def update_selected(self):
        client_id = self.current_selected_client_id()
        if not client_id:
            QtWidgets.QMessageBox.warning(self, "Seleção", "Selecione um cliente na tabela.")
            return
        with connect_db() as con:
            cur = con.cursor()
            cur.execute("""
                UPDATE clients
                   SET name=?, age=?, whatsapp=?, first_consult=?, plan_months=?, paid_returns=?, rescheduled=?, notes=?, hidden=?
                 WHERE id=?
            """, (
                self.ed_name.text().strip(),
                int(self.ed_age.value()),
                self.ed_whats.text().strip(),
                qdate_to_str(self.dt_first.date()),
                int(self.cb_plan.currentText()),
                int(self.sp_paid.value()),
                1 if self.chk_resched.isChecked() else 0,
                self.ed_notes.toPlainText().strip(),
                1 if self.chk_hidden.isChecked() else 0,
                client_id
            ))
            # atualizar próximas: estratégia simples -> apagar agendadas e recriar pela lista
            cur.execute("DELETE FROM appointments WHERE client_id=? AND status='agendado'", (client_id,))
            for i in range(self.list_next.count()):
                d = self.list_next.item(i).text()
                cur.execute("INSERT INTO appointments (client_id, date, kind, status) VALUES (?, ?, 'retorno', 'agendado')",
                            (client_id, d))
            con.commit()
        self.load_table()
        QtWidgets.QMessageBox.information(self, "Atualizado", "Registro atualizado.")

    def remove_selected(self):
        client_id = self.current_selected_client_id()
        if not client_id:
            QtWidgets.QMessageBox.warning(self, "Seleção", "Selecione um cliente.")
            return
        if not ask_yes_no(self, "Remover", "Deseja remover DEFINITIVAMENTE o cliente selecionado?"):
            return
        with connect_db() as con:
            cur = con.cursor()
            # Apagar consultas dependentes e cliente
            cur.execute("DELETE FROM appointments WHERE client_id=?", (client_id,))
            cur.execute("DELETE FROM clients WHERE id=?", (client_id,))
            con.commit()
        self.load_table()
        self.clear_form()

    # ---------- CSV ----------
    def csv_dialog(self):
        menu = QtWidgets.QMenu()
        a1 = menu.addAction("Exportar CSV…")
        a2 = menu.addAction("Importar CSV…")
        act = menu.exec(QtGui.QCursor.pos())
        if act == a1:
            self.export_csv()
        elif act == a2:
            self.import_csv()

    def export_csv(self):
        fn_clients, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Salvar clientes.csv", str(APP_DIR / "clientes.csv"), "CSV (*.csv)")
        if not fn_clients:
            return
        fn_appts = fn_clients.replace(".csv", "_consultas.csv")

        with connect_db() as con:
            cur = con.cursor()
            # clientes
            cur.execute("SELECT id,name,age,whatsapp,first_consult,plan_months,paid_returns,rescheduled,notes,hidden FROM clients")
            rows = cur.fetchall()
            with open(fn_clients, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f, delimiter=";")
                w.writerow(["id","name","age","whatsapp","first_consult","plan_months","paid_returns","rescheduled","notes","hidden"])
                w.writerows(rows)
            # consultas
            cur.execute("SELECT client_id,date,kind,status FROM appointments")
            rows = cur.fetchall()
            with open(fn_appts, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f, delimiter=";")
                w.writerow(["client_id","date","kind","status"])
                w.writerows(rows)
        QtWidgets.QMessageBox.information(self, "CSV", f"Exportado:\n{fn_clients}\n{fn_appts}")

    def import_csv(self):
        fn_clients, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Abrir clientes.csv", str(APP_DIR), "CSV (*.csv)")
        if not fn_clients:
            return
        fn_appts = fn_clients.replace(".csv", "_consultas.csv")
        if not os.path.exists(fn_appts):
            QtWidgets.QMessageBox.warning(self, "CSV", "Arquivo de consultas não encontrado (deveria ser *_consultas.csv).")
            return

        if not ask_yes_no(self, "Importar", "Isto substituirá os dados existentes. Continuar?"):
            return

        with connect_db() as con:
            cur = con.cursor()
            cur.execute("DELETE FROM appointments")
            cur.execute("DELETE FROM clients")
            # importar clientes
            with open(fn_clients, "r", encoding="utf-8") as f:
                r = csv.DictReader(f, delimiter=";")
                for row in r:
                    cur.execute("""
                        INSERT INTO clients (id,name,age,whatsapp,first_consult,plan_months,paid_returns,rescheduled,notes,hidden)
                        VALUES (?,?,?,?,?,?,?,?,?,?)
                    """, (
                        int(row["id"]),
                        row["name"],
                        int(row["age"]) if row["age"] else None,
                        row["whatsapp"],
                        row["first_consult"] or None,
                        int(row["plan_months"] or 2),
                        int(row["paid_returns"] or 0),
                        int(row["rescheduled"] or 0),
                        row["notes"],
                        int(row["hidden"] or 0),
                    ))
            # importar consultas
            with open(fn_appts, "r", encoding="utf-8") as f:
                r = csv.DictReader(f, delimiter=";")
                for row in r:
                    cur.execute("""
                        INSERT INTO appointments (client_id,date,kind,status)
                        VALUES (?,?,?,?)
                    """, (
                        int(row["client_id"]),
                        row["date"],
                        row["kind"],
                        row["status"],
                    ))
            con.commit()
        self.load_table()
        QtWidgets.QMessageBox.information(self, "CSV", "Importação concluída.")

    # ---------- PDF ----------
    def export_pdf(self):
        fn, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Salvar relatório.pdf",
            str(APP_DIR / "relatorio_clientes.pdf"),
            "PDF (*.pdf)"
        )
        if not fn:
            return

        # Documento com margens
        doc = SimpleDocTemplate(fn, pagesize=A4,
                                leftMargin=28, rightMargin=28,
                                topMargin=28, bottomMargin=28)
        styles = getSampleStyleSheet()
        story = []

        # Título centralizado
        title = Paragraph(
            f"<para align='center'><font size=16 color='#b0004f'><b>{APP_NAME}</b></font></para>",
            styles["Normal"]
        )
        story += [title, Spacer(1, 12)]

        # Cabeçalho + dados
        data = [["Cliente", "Idade", "WhatsApp", "Primeira Consulta",
                 "Próximas", "Plano", "Retornos pagos", "Obs"]]

        with connect_db() as con:
            cur = con.cursor()
            cur.execute("""
                SELECT id,name,age,whatsapp,first_consult,plan_months,
                       paid_returns,rescheduled,notes,hidden
                  FROM clients ORDER BY hidden ASC, name COLLATE NOCASE ASC
            """)
            clients = cur.fetchall()
            for cid, name, age, whats, firstc, plan, paid, resch, notes, hidden in clients:
                cur.execute("SELECT date FROM appointments WHERE client_id=? AND status='agendado' ORDER BY date ASC", (cid,))
                dates = [human_date(d) for (d,) in cur.fetchall()]
                data.append([
                    Paragraph(name + (" (OCULTO)" if hidden else ""), styles["Normal"]),
                    str(age or ""),
                    whats or "",
                    human_date(firstc or ""),
                    Paragraph(", ".join(dates) if dates else "-", styles["Normal"]),
                    str(plan),
                    str(paid),
                    Paragraph(notes or "-", styles["Normal"])
                ])

        # Ajustar larguras (somar <= 539 para caber no A4)
        col_widths = [90, 30, 70, 65, 70, 30, 70, 100]

        tabela = Table(data, colWidths=col_widths, repeatRows=1)
        tabela.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.Color(1, 0.74, 0.84)),
            ("TEXTCOLOR", (0,0), (-1,0), colors.Color(0.16, 0, 0.13)),
            ("GRID", (0,0), (-1,-1), 0.5, colors.Color(1, 0.31, 0.53)),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("ALIGN", (0,0), (-1,0), "CENTER"),
            ("ROWBACKGROUNDS", (0,1), (-1,-1),
                [colors.Color(1, 0.91, 0.94), colors.Color(1, 0.85, 0.90)]),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("FONTSIZE", (0,0), (-1,-1), 8),
            ("LEFTPADDING", (0,0), (-1,-1), 4),
            ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ]))

        story.append(tabela)
        doc.build(story)

        QtWidgets.QMessageBox.information(self, "PDF", f"Relatório salvo em:\n{fn}")

    # ---------- WhatsApp ----------
    def compose_week_message(self):
        # consultas entre hoje e +7 dias
        today = dt.date.today()
        limit = today + dt.timedelta(days=7)
        lines = ["[ === CONSULTAS DA SEMANA === ]"]

        with connect_db() as con:
            cur = con.cursor()
            cur.execute("""
                SELECT c.id, c.name, c.age, c.whatsapp, c.first_consult, c.plan_months, c.paid_returns, c.notes
                  FROM clients c
                 WHERE c.hidden=0
                 ORDER BY c.name COLLATE NOCASE ASC
            """)
            clients = cur.fetchall()
            for cid, name, age, whats, firstc, plan, paid, notes in clients:
                cur.execute("""
                    SELECT date FROM appointments
                     WHERE client_id=? AND status='agendado'
                  ORDER BY date ASC
                """, (cid,))
                appts = [row[0] for row in cur.fetchall()]
                appts_in_week = []
                for d in appts:
                    try:
                        y, m, d2 = map(int, d.split("-"))
                        date_obj = dt.date(y, m, d2)
                        if today <= date_obj <= limit:
                            appts_in_week.append(human_date(d))
                        # guardamos todas depois para "próximas" também
                    except Exception:
                        pass
                if appts_in_week:
                    lines += [
                        "",
                        f"Nome: {name}",
                        f"Idade: {age or '-'}",
                        f"WhatsApp: {whats or '-'}",
                        f"Primeira consulta: {human_date(firstc) if firstc else '-'}",
                        f"Próximas (7 dias): {', '.join(appts_in_week)}",
                        f"Plano: {plan} meses | Retornos pagos: {paid}",
                        f"Obs: {notes or '-'}",
                        "[ ------------------------- ]"
                    ]
        if len(lines) == 1:
            lines.append("Sem consultas nos próximos 7 dias.")
        return "\n".join(lines)

    def send_week_whatsapp(self):
        if kit is None:
            QtWidgets.QMessageBox.warning(self, "WhatsApp", "Biblioteca pywhatkit não disponível.")
            return
        if not OWNER_WHATS.startswith("+"):
            QtWidgets.QMessageBox.warning(self, "WhatsApp", "Configure seu número em OWNER_WHATS no código (formato +55DDDNUMERO).")
            return
        msg = self.compose_week_message()
        try:
            # envia instantaneamente (abre WhatsApp Web e envia)
            kit.sendwhatmsg_instantly(OWNER_WHATS, msg, wait_time=20, tab_close=True)
            QtWidgets.QMessageBox.information(self, "WhatsApp", "Mensagem enviada para seu WhatsApp.")
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "WhatsApp", f"Falha ao enviar: {e}")

    def send_week_whatsapp_auto(self):
        # silencioso (sem messagebox) na abertura do app
        if kit is None or not OWNER_WHATS.startswith("+"):
            return
        msg = self.compose_week_message()
        try:
            kit.sendwhatmsg_instantly(OWNER_WHATS, msg, wait_time=20, tab_close=True)
        except Exception:
            pass  # silencioso

# ---- main
def main():
    ensure_db()
    app = QtWidgets.QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
