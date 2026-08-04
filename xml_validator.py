import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xml.etree.ElementTree as ET
import openpyxl
import os

BG      = "#1e1e2e"
SURFACE = "#2a2a3d"
ACCENT  = "#7c6af7"
ACCENT2 = "#5a4fcf"
TEXT    = "#e0e0f0"
MUTED   = "#8888aa"
SUCCESS = "#4caf82"
ERROR   = "#f06060"
WARNING = "#f0b060"
BORDER  = "#3a3a55"

class XMLValidatorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("XML Validator  v1.0")
        self.geometry("1000x680")
        self.minsize(800, 560)
        self.configure(bg=BG)
        self.excel_path = tk.StringVar(value="Nenhum arquivo selecionado")
        self.xml_paths  = []
        self.results    = []
        self._build_ui()

    def _build_ui(self):
        hdr = tk.Frame(self, bg=ACCENT, height=56)
        hdr.pack(fill="x")
        tk.Label(hdr, text="XML Validator", font=("Segoe UI", 16, "bold"),
                 bg=ACCENT, fg="white").pack(side="left", padx=20, pady=12)
        tk.Label(hdr, text="Comparacao Excel x XML  -  100% offline",
                 font=("Segoe UI", 9), bg=ACCENT, fg="#d0ccff").pack(side="left")

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=20, pady=16)

        left = tk.Frame(body, bg=BG, width=300)
        left.pack(side="left", fill="y", padx=(0, 16))
        left.pack_propagate(False)

        self._section(left, "1. Gabarito Excel")
        self._btn(left, "Selecionar Excel", self._load_excel).pack(fill="x", pady=(0,4))
        tk.Label(left, textvariable=self.excel_path, bg=BG, fg=MUTED,
                 font=("Segoe UI", 8), wraplength=270, justify="left").pack(anchor="w")

        self._sep(left)
        self._section(left, "2. Arquivos XML")
        self._btn(left, "Adicionar XMLs", self._load_xmls).pack(fill="x", pady=(0,4))
        self.xml_listbox = tk.Listbox(left, bg=SURFACE, fg=TEXT, selectbackground=ACCENT,
                                      relief="flat", bd=0, font=("Segoe UI", 8),
                                      height=8, highlightthickness=1, highlightbackground=BORDER)
        self.xml_listbox.pack(fill="x", pady=(0,4))
        self._btn(left, "Remover selecionado", self._remove_xml, color=SURFACE).pack(fill="x")

        self._sep(left)
        self._section(left, "3. Executar")
        self._btn(left, "Validar XMLs", self._run, color=ACCENT).pack(fill="x", pady=(0,4))
        self._btn(left, "Exportar relatorio", self._export, color=ACCENT2).pack(fill="x", pady=(0,4))
        self._btn(left, "Limpar tudo", self._clear, color=SURFACE).pack(fill="x")

        right = tk.Frame(body, bg=BG)
        right.pack(side="left", fill="both", expand=True)

        self.summary_frame = tk.Frame(right, bg=SURFACE, pady=10)
        self.summary_frame.pack(fill="x", pady=(0,10))
        for label, attr, color in [
            ("Total","lbl_total",TEXT),("OK","lbl_ok",SUCCESS),
            ("Erro","lbl_err",ERROR),("Aviso","lbl_warn",WARNING)]:
            f = tk.Frame(self.summary_frame, bg=SURFACE)
            f.pack(side="left", expand=True)
            lv = tk.Label(f, text="0", font=("Segoe UI",22,"bold"), bg=SURFACE, fg=color)
            lv.pack()
            tk.Label(f, text=label, font=("Segoe UI",8), bg=SURFACE, fg=MUTED).pack()
            setattr(self, attr, lv)

        cols = ("arquivo","campo","esperado","encontrado","status")
        self.tree = ttk.Treeview(right, columns=cols, show="headings", height=20)
        for c,h,w in zip(cols,("Arquivo XML","Campo","Esperado (Excel)","Encontrado (XML)","Status"),(180,160,180,180,90)):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=w, minwidth=60, anchor="w")

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background=SURFACE, foreground=TEXT,
                        fieldbackground=SURFACE, rowheight=26, font=("Segoe UI",9))
        style.configure("Treeview.Heading", background=BORDER, foreground=TEXT, font=("Segoe UI",9,"bold"))
        style.map("Treeview", background=[("selected",ACCENT)])
        self.tree.tag_configure("ok",      foreground=SUCCESS)
        self.tree.tag_configure("erro",    foreground=ERROR)
        self.tree.tag_configure("aviso",   foreground=WARNING)
        self.tree.tag_configure("missing", foreground=WARNING)

        sb = ttk.Scrollbar(right, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")

        self.status_var = tk.StringVar(value="Pronto. Selecione o Excel e os XMLs para comecar.")
        tk.Label(self, textvariable=self.status_var, bg=SURFACE, fg=MUTED,
                 font=("Segoe UI",8), anchor="w", padx=12, pady=6).pack(fill="x", side="bottom")

    def _section(self, parent, text):
        tk.Label(parent, text=text, bg=BG, fg=ACCENT, font=("Segoe UI",9,"bold")).pack(anchor="w", pady=(10,2))

    def _sep(self, parent):
        tk.Frame(parent, bg=BORDER, height=1).pack(fill="x", pady=8)

    def _btn(self, parent, text, cmd, color=ACCENT):
        return tk.Button(parent, text=text, command=cmd, bg=color, fg="white",
                         activebackground=ACCENT2, activeforeground="white",
                         relief="flat", bd=0, font=("Segoe UI",9), cursor="hand2", pady=8)

    def _load_excel(self):
        path = filedialog.askopenfilename(title="Selecionar gabarito Excel",
            filetypes=[("Excel","*.xlsx *.xls"),("Todos","*.*")])
        if path:
            self.excel_path.set(os.path.basename(path))
            self._excel_full_path = path
            self.status_var.set(f"Excel carregado: {os.path.basename(path)}")

    def _load_xmls(self):
        paths = filedialog.askopenfilenames(title="Selecionar arquivos XML",
            filetypes=[("XML","*.xml"),("Todos","*.*")])
        for p in paths:
            if p not in self.xml_paths:
                self.xml_paths.append(p)
                self.xml_listbox.insert("end", os.path.basename(p))
        self.status_var.set(f"{len(self.xml_paths)} arquivo(s) XML carregado(s).")

    def _remove_xml(self):
        sel = self.xml_listbox.curselection()
        if sel:
            idx = sel[0]
            self.xml_listbox.delete(idx)
            self.xml_paths.pop(idx)

    def _clear(self):
        self.excel_path.set("Nenhum arquivo selecionado")
        self._excel_full_path = None
        self.xml_paths.clear()
        self.xml_listbox.delete(0, "end")
        self.tree.delete(*self.tree.get_children())
        self.results.clear()
        self._update_summary()
        self.status_var.set("Limpo. Pronto para nova validacao.")

    def _run(self):
        if not hasattr(self, "_excel_full_path") or not self._excel_full_path:
            messagebox.showwarning("Atencao", "Selecione o arquivo Excel (gabarito) primeiro.")
            return
        if not self.xml_paths:
            messagebox.showwarning("Atencao", "Adicione pelo menos um arquivo XML.")
            return
        self.tree.delete(*self.tree.get_children())
        self.results.clear()
        try:
            gabarito = self._read_excel(self._excel_full_path)
        except Exception as e:
            messagebox.showerror("Erro ao ler Excel", str(e))
            return
        if not gabarito:
            messagebox.showerror("Erro", "Nenhum dado encontrado no Excel.")
            return
        ok = err = warn = 0
        for xml_path in self.xml_paths:
            try:
                xml_data = self._read_xml(xml_path)
                fname = os.path.basename(xml_path)
                for campo, esperado in gabarito.items():
                    encontrado = xml_data.get(campo)
                    if encontrado is None:
                        status, tag = "Ausente", "missing"
                        warn += 1
                    elif str(encontrado).strip() == str(esperado).strip():
                        status, tag = "OK", "ok"
                        ok += 1
                    else:
                        status, tag = "Divergente", "erro"
                        err += 1
                    row = dict(arquivo=fname, campo=campo, esperado=str(esperado),
                               encontrado=str(encontrado) if encontrado else "-", status=status)
                    self.results.append(row)
                    self.tree.insert("", "end",
                                     values=(fname, campo, esperado,
                                             encontrado if encontrado else "-", status),
                                     tags=(tag,))
            except Exception as e:
                fname = os.path.basename(xml_path)
                self.tree.insert("", "end", values=(fname,"-","-",f"Erro: {e}","Falha"), tags=("erro",))
                err += 1
        self._update_summary(ok+err+warn, ok, err, warn)
        self.status_var.set(f"Validacao concluida - {ok} OK, {err} Erros, {warn} Ausentes")

    def _read_excel(self, path):
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return {}
        result = {}
        first = rows[0]
        non_empty = [c for c in first if c is not None]
        if len(non_empty) <= 2:
            for row in rows:
                if row[0] is not None:
                    result[str(row[0]).strip()] = row[1] if len(row) > 1 else ""
        else:
            headers = [str(h).strip() for h in first if h is not None]
            if len(rows) > 1:
                values = rows[1]
                for i, h in enumerate(headers):
                    result[h] = values[i] if i < len(values) else ""
        wb.close()
        return result

    def _read_xml(self, path):
        tree = ET.parse(path)
        root = tree.getroot()
        data = {}
        self._flatten_xml(root, "", data)
        return data

    def _flatten_xml(self, element, prefix, data):
        tag = element.tag.split("}")[-1] if "}" in element.tag else element.tag
        key = f"{prefix}/{tag}" if prefix else tag
        for attr_name, attr_val in element.attrib.items():
            data[f"{key}@{attr_name}"] = attr_val
        text = (element.text or "").strip()
        if text:
            data[key] = text
        for child in element:
            self._flatten_xml(child, key, data)

    def _update_summary(self, total=0, ok=0, err=0, warn=0):
        self.lbl_total.config(text=str(total))
        self.lbl_ok.config(text=str(ok))
        self.lbl_err.config(text=str(err))
        self.lbl_warn.config(text=str(warn))

    def _export(self):
        if not self.results:
            messagebox.showwarning("Atencao", "Execute a validacao antes de exportar.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
            filetypes=[("Excel","*.xlsx"),("CSV","*.csv")], title="Salvar relatorio")
        if not path:
            return
        try:
            if path.endswith(".csv"):
                self._export_csv(path)
            else:
                self._export_excel(path)
            messagebox.showinfo("Exportado", f"Relatorio salvo em:\n{path}")
            self.status_var.set(f"Relatorio exportado: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Erro ao exportar", str(e))

    def _export_excel(self, path):
        from openpyxl.styles import PatternFill, Font, Alignment
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Resultado"
        ws.append(["Arquivo XML","Campo","Esperado (Excel)","Encontrado (XML)","Status"])
        hdr_fill = PatternFill("solid", fgColor="7C6AF7")
        for cell in ws[1]:
            cell.fill = hdr_fill
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center")
        fill_ok   = PatternFill("solid", fgColor="D6F5E3")
        fill_err  = PatternFill("solid", fgColor="FAD7D7")
        fill_warn = PatternFill("solid", fgColor="FFF0CC")
        for r in self.results:
            ws.append([r["arquivo"],r["campo"],r["esperado"],r["encontrado"],r["status"]])
            fill = (fill_ok if "OK" in r["status"] else
                    fill_err if "Divergente" in r["status"] or "Falha" in r["status"] else fill_warn)
            for cell in ws[ws.max_row]:
                cell.fill = fill
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 28
        wb.save(path)

    def _export_csv(self, path):
        import csv
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=["arquivo","campo","esperado","encontrado","status"])
            writer.writeheader()
            writer.writerows(self.results)

if __name__ == "__main__":
    app = XMLValidatorApp()
    app.mainloop()
