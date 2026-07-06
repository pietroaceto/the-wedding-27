import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from urllib.parse import urlencode

BASE_INVITE_URL = "https://the-wedding-27.web.app/invito.html"

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Lista Invitati"

# Columns
headers = [
    "Gruppo", "Nome", "Ruolo",
    "Adulto", "Bambino 8+", "Bambino 3-8", "Bambino 0-3",
    "Allergie o Intolleranze", "Probabilità di presenza", "Note"
]

# ─── Build rows ────────────────────────────────────────────────────────────────
# Each tuple: (gruppo, nome, ruolo, note)
rows = []
group_meta = {}  # gruppo -> URL params dict for the digital invitation

def coppia(g, n1, n2, note=""):
    rows.append((g, n1, "Sposo/a", note))
    rows.append((g, n2, "Sposa/o", note))
    group_meta[g] = {'tipo': 'coppia', 'nome1': n1, 'nome2': n2}

def coppia_figli(g, n1, n2, n_figli, note=""):
    rows.append((g, n1, "Sposo/a", note))
    rows.append((g, n2, "Sposa/o", note))
    for i in range(1, n_figli + 1):
        rows.append((g, f"Figlio/a {i} di {n1} e {n2}", "Figlio/a", note))
    group_meta[g] = {'tipo': 'coppia', 'nome1': n1, 'nome2': n2}

def singolo_partner(g, nome, sesso="", note=""):
    rows.append((g, nome, "Ospite", note))
    rows.append((g, f"Ragazzo/a di {nome}", "Partner", note))
    meta = {'tipo': 'singolo', 'nome': nome}
    if sesso: meta['sesso'] = sesso
    group_meta[g] = meta

def singolo_famiglia(g, nome, n_extra, note=""):
    """nome + moglie/marito + (n_extra-1) figli"""
    rows.append((g, nome, "Ospite", note))
    rows.append((g, f"Moglie/Marito di {nome}", "Coniuge", note))
    for i in range(1, n_extra):
        rows.append((g, f"Figlio/a {i} di {nome}", "Figlio/a", note))
    group_meta[g] = {'tipo': 'famiglia', 'cognome': nome}

def singolo(g, nome, ruolo="Ospite", sesso="", note=""):
    rows.append((g, nome, ruolo, note))
    meta = {'tipo': 'singolo', 'nome': nome}
    if sesso: meta['sesso'] = sesso
    group_meta[g] = meta

def famiglia_n(g, capofamiglia, n_totale, note=""):
    rows.append((g, capofamiglia, "Capofamiglia", note))
    for i in range(1, n_totale):
        rows.append((g, f"Familiare {i} di {capofamiglia}", "Familiare", note))
    group_meta[g] = {'tipo': 'famiglia', 'cognome': capofamiglia}

# ─── Entries ───────────────────────────────────────────────────────────────────

# Famiglia Pietro 27
famiglia_n("Famiglia Pietro", "Pietro", 27)

# Famiglia Ludo 23
famiglia_n("Famiglia Ludo", "Ludo", 23)

singolo("Vigna", "Vigna")
singolo("Lorella", "Lorella")
singolo("Gustavo", "Gustavo")
singolo("Marta Baroni", "Marta Baroni")
singolo("Masso", "Masso")
coppia("Berga e Maria", "Berga", "Maria")
singolo("Giulia Tusino", "Giulia Tusino")
singolo("Marco Iannetelli", "Marco Iannetelli")
coppia("Luca e Fiammetta", "Luca", "Fiammetta")

# Enrico e Reb +4 → coppia + 4 figli
coppia_figli("Enrico e Reb", "Enrico", "Reb", 4)

# Andrea e Marta +3 → coppia + 2 figli  (la moglie = 1, figli = +3-1=2)
coppia_figli("Andrea e Marta", "Andrea", "Marta", 2)

# Malo e Vale +2 → coppia + 1 figlio
coppia_figli("Malo e Vale", "Malo", "Vale", 1)

# Luca e Betta +3 → coppia + 2 figli
coppia_figli("Luca e Betta", "Luca", "Betta", 2)

# Luca e Martina +3 → coppia + 2 figli
coppia_figli("Luca e Martina", "Luca", "Martina", 2)

# Matteo e Alessia +3 → coppia + 2 figli
coppia_figli("Matteo e Alessia", "Matteo", "Alessia", 2)

# Chiara e Colle +3 → coppia + 2 figli
coppia_figli("Chiara e Colle", "Chiara", "Colle", 2)

# Giovanni e Noemi +2 → coppia + 1 figlio
coppia_figli("Giovanni e Noemi", "Giovanni", "Noemi", 1)

# Elena Elisei +1 → Elena + ragazzo/a
singolo_partner("Elena Elisei", "Elena Elisei")

# Raffo e Chiara +1 → coppia + 1 ragazzo/a
rows.append(("Raffo e Chiara", "Raffo", "Ospite", ""))
rows.append(("Raffo e Chiara", "Chiara", "Ospite", ""))
rows.append(("Raffo e Chiara", "Ragazzo/a di Raffo e Chiara", "Partner/Familiare", ""))

# FraBru e Elisa +3 → coppia + 2 figli
coppia_figli("FraBru e Elisa", "FraBru", "Elisa", 2)

# AleBru e Chiara +1 → coppia + 1 ragazzo/a
rows.append(("AleBru e Chiara", "AleBru", "Ospite", ""))
rows.append(("AleBru e Chiara", "Chiara", "Ospite", ""))
rows.append(("AleBru e Chiara", "Ragazzo/a di AleBru e Chiara", "Partner/Familiare", ""))

# Berna e Carm +2 → coppia + 1 figlio
coppia_figli("Berna e Carm", "Berna", "Carm", 1)

# Luigi e Cate +3 → coppia + 2 figli
coppia_figli("Luigi e Cate", "Luigi", "Cate", 2)

coppia("Budo e Jenia", "Budo", "Jenia")

# Andrea e Jadi +1 → coppia + 1 ragazzo/a
rows.append(("Andrea e Jadi", "Andrea", "Ospite", ""))
rows.append(("Andrea e Jadi", "Jadi", "Ospite", ""))
rows.append(("Andrea e Jadi", "Ragazzo/a di Andrea e Jadi", "Partner/Familiare", ""))

# Anna Belia e Andrea +2 → coppia + 1 figlio
coppia_figli("Anna Belia e Andrea", "Anna Belia", "Andrea", 1)

coppia("Jack e Diana", "Jack", "Diana")
coppia("Anglano e Irene", "Anglano", "Irene")
singolo("Federica Belia", "Federica Belia")
coppia("Piero e Chiara", "Piero", "Chiara")
coppia("Diba e Claudia", "Diba", "Claudia")
singolo("Adele", "Adele")
coppia("Maddalena e Andri", "Maddalena", "Andri")
singolo("Elisabetta", "Elisabetta")
coppia("Maddi e Fra Bini", "Maddi", "Fra Bini")
singolo("Benedetta Bianchini", "Benedetta Bianchini")
singolo("Chiara De Socio", "Chiara De Socio")
singolo("Nicola De Socio", "Nicola De Socio")
coppia("Giuseppe e Gabriella", "Giuseppe", "Gabriella")

# Meco e Ceci +3 → coppia + 2 figli
coppia_figli("Meco e Ceci", "Meco", "Ceci", 2)

singolo("Almerina", "Almerina")
singolo("Vincenzo Sorella e Famiglia", "Vincenzo Sorella e Famiglia", note="Da confermare")
coppia("Mamo e Fiorella", "Mamo", "Fiorella")

# Ale Menna +1
singolo_partner("Ale Menna", "Ale Menna")

# Benji+Moglie → Benji + Moglie
rows.append(("Benji e Moglie", "Benji", "Ospite", ""))
rows.append(("Benji e Moglie", "Moglie di Benji", "Coniuge", ""))

coppia("Mario e Lina", "Mario", "Lina")
coppia("Marino e Roberta", "Marino", "Roberta")
singolo("Daniele Bontempi", "Daniele Bontempi")
singolo("Valerio Ferrini", "Valerio Ferrini")

# Aurora basilica +1
singolo_partner("Aurora Basilica", "Aurora Basilica")

singolo("Gigliola e Famiglia", "Gigliola e Famiglia", note="Numero da confermare")

# Giulia e Stefano +2 → coppia + 1 figlio
coppia_figli("Giulia e Stefano", "Giulia", "Stefano", 1)

# Marco e Giovanna +2 → coppia + 1 figlio
coppia_figli("Marco e Giovanna", "Marco", "Giovanna", 1)

# Luigi e Anna +1 → coppia + 1 ragazzo/a
rows.append(("Luigi e Anna", "Luigi", "Ospite", ""))
rows.append(("Luigi e Anna", "Anna", "Ospite", ""))
rows.append(("Luigi e Anna", "Ragazzo/a di Luigi e Anna", "Partner/Familiare", ""))

coppia("Francesco e Claudia", "Francesco", "Claudia")

# Simone e Francesca +2 → coppia + 1 figlio
coppia_figli("Simone e Francesca", "Simone", "Francesca", 1)

# Cesare e Daniela +1 → coppia + 1 ragazzo/a
rows.append(("Cesare e Daniela", "Cesare", "Ospite", ""))
rows.append(("Cesare e Daniela", "Daniela", "Ospite", ""))
rows.append(("Cesare e Daniela", "Ragazzo/a di Cesare e Daniela", "Partner/Familiare", ""))

# Tomei e Giulia +1 → coppia + 1 ragazzo/a
rows.append(("Tomei e Giulia", "Tomei", "Ospite", ""))
rows.append(("Tomei e Giulia", "Giulia", "Ospite", ""))
rows.append(("Tomei e Giulia", "Ragazzo/a di Tomei e Giulia", "Partner/Familiare", ""))

# Letizia e Nic +1 → coppia + 1 ragazzo/a
rows.append(("Letizia e Nic", "Letizia", "Ospite", ""))
rows.append(("Letizia e Nic", "Nic", "Ospite", ""))
rows.append(("Letizia e Nic", "Ragazzo/a di Letizia e Nic", "Partner/Familiare", ""))

# Carmen e Luca +1 → coppia + 1 ragazzo/a
rows.append(("Carmen e Luca", "Carmen", "Ospite", ""))
rows.append(("Carmen e Luca", "Luca", "Ospite", ""))
rows.append(("Carmen e Luca", "Ragazzo/a di Carmen e Luca", "Partner/Familiare", ""))

coppia("Andrea e Federica", "Andrea", "Federica")
coppia("Gloria e Francesco", "Gloria", "Francesco")
coppia("Rita e Gabriele", "Rita", "Gabriele")
coppia("Roberta e Alessio", "Roberta", "Alessio")
coppia("Beatrice e Luca", "Beatrice", "Luca")
singolo("Mario????", "Mario", note="Da confermare presenza")
coppia("Erika e Giulio", "Erika", "Giulio")

# Famiglia Rita 4? → Rita + 3 familiari (incerto)
rows.append(("Famiglia Rita", "Rita", "Capofamiglia", "Numero da confermare (4?)"))
for i in range(1, 4):
    rows.append(("Famiglia Rita", f"Familiare {i} di Rita", "Familiare", "Da confermare"))

# Adrian +1
singolo_partner("Adrian", "Adrian")

# Raffaele +1
singolo_partner("Raffaele", "Raffaele")

# Riccardo +1
singolo_partner("Riccardo", "Riccardo")

# Davide +1
singolo_partner("Davide", "Davide")

# Roberta e Famiglia (3) → Roberta + 2 familiari
rows.append(("Roberta e Famiglia", "Roberta", "Capofamiglia", ""))
for i in range(1, 3):
    rows.append(("Roberta e Famiglia", f"Familiare {i} di Roberta", "Familiare", ""))

# Stefano e Famiglia (4) → Stefano + 3 familiari
rows.append(("Stefano e Famiglia", "Stefano", "Capofamiglia", ""))
for i in range(1, 4):
    rows.append(("Stefano e Famiglia", f"Familiare {i} di Stefano", "Familiare", ""))

# Gloria e Francesco inf +2 → coppia + 2 bambini piccoli
rows.append(("Gloria e Francesco (inf)", "Gloria", "Ospite", ""))
rows.append(("Gloria e Francesco (inf)", "Francesco", "Ospite", ""))
for i in range(1, 3):
    rows.append(("Gloria e Francesco (inf)", f"Bambino/a {i} di Gloria e Francesco", "Figlio/a", "Bambino piccolo (inf)"))

singolo("Federica", "Federica")
singolo("Marta", "Marta")
singolo("Suor Giovanna", "Suor Giovanna", ruolo="Religiosa")

# Gianluca e Adriana +1 → coppia + ragazzo/a
rows.append(("Gianluca e Adriana", "Gianluca", "Ospite", ""))
rows.append(("Gianluca e Adriana", "Adriana", "Ospite", ""))
rows.append(("Gianluca e Adriana", "Ragazzo/a di Gianluca e Adriana", "Partner/Familiare", ""))

# Daniele Mamusa +fid → Daniele + fidanzato/a
rows.append(("Daniele Mamusa", "Daniele Mamusa", "Ospite", ""))
rows.append(("Daniele Mamusa", "Fidanzato/a di Daniele Mamusa", "Partner", ""))

singolo("Emanuele Salera", "Emanuele Salera")
singolo("Maria Fiorelli", "Maria Fiorelli")
singolo("Anna Brusciano", "Anna Brusciano")
singolo("Giulio", "Giulio")
singolo("Letizia", "Letizia")
singolo("Maria Rondelli", "Maria Rondelli")

# Nicole+Ale → coppia
coppia("Nicole e Ale", "Nicole", "Ale")

coppia("Diana e Simone", "Diana", "Simone")
singolo("Erika De Marco", "Erika De Marco")

# Zio Silvano + → Silvano + accompagnatore (numero non specificato)
rows.append(("Zio Silvano", "Zio Silvano", "Ospite", ""))
rows.append(("Zio Silvano", "Accompagnatore/trice di Zio Silvano", "Partner/Familiare", "Numero non specificato"))

singolo("Camilla", "Camilla")

# Annalisa +1
singolo_partner("Annalisa", "Annalisa")

singolo("Natalia", "Natalia")

# Andrea Musio +moglie+3 → Andrea + moglie + 3 figli
rows.append(("Andrea Musio e Famiglia", "Andrea Musio", "Ospite", ""))
rows.append(("Andrea Musio e Famiglia", "Moglie di Andrea Musio", "Coniuge", ""))
for i in range(1, 4):
    rows.append(("Andrea Musio e Famiglia", f"Figlio/a {i} di Andrea Musio", "Figlio/a", ""))

singolo("Tommy", "Tommy")
coppia("Christy e Federica", "Christy", "Federica")
singolo("Federica Belia (2)", "Federica Belia", note="Verificare duplicato")
coppia("Mati e Mary", "Mati", "Mary")
singolo("Daniela Mammarella", "Daniela Mammarella")
coppia("Lucia Mammarella e Stefano", "Lucia Mammarella", "Stefano")
coppia("Giacomo e Annalisa", "Giacomo", "Annalisa")

# Nicola sorella +1 → Nicola (la sorella) + ragazzo/a
rows.append(("Nicola Sorella", "Nicola (sorella)", "Ospite", ""))
rows.append(("Nicola Sorella", "Ragazzo/a di Nicola (sorella)", "Partner", ""))

# Angelo Anglano +moglie → Angelo + moglie
rows.append(("Angelo Anglano e Moglie", "Angelo Anglano", "Ospite", ""))
rows.append(("Angelo Anglano e Moglie", "Moglie di Angelo Anglano", "Coniuge", ""))

singolo("Prof Pambianco", "Prof. Pambianco")

# Dibez +3 → Dibez + moglie/marito + 2 figli
singolo_famiglia("Dibez", "Dibez", 3)

singolo("Comunità Neocat", "Comunità Neocat", note="Numero da confermare")
singolo("Caterina Fichera", "Caterina Fichera")

# Domenico Salvemini +3 → Domenico + moglie + 2 figli
singolo_famiglia("Domenico Salvemini", "Domenico Salvemini", 3)

coppia("Stefania e Riccardo", "Stefania", "Riccardo")
singolo("Marchesi", "Marchesi")

# Alberto Gambelli +1
singolo_partner("Alberto Gambelli", "Alberto Gambelli")

singolo("Benedetta Carlotti", "Benedetta Carlotti")
coppia("Luigi e Michele di Paolo", "Luigi di Paolo", "Michele di Paolo")
singolo("Simone Moretti", "Simone Moretti")

# ─── group_meta per i gruppi definiti con rows.append diretti ────────────────────
group_meta.setdefault("Raffo e Chiara",           {'tipo': 'coppia',   'nome1': 'Raffo',          'nome2': 'Chiara'})
group_meta.setdefault("AleBru e Chiara",          {'tipo': 'coppia',   'nome1': 'AleBru',         'nome2': 'Chiara'})
group_meta.setdefault("Andrea e Jadi",            {'tipo': 'coppia',   'nome1': 'Andrea',         'nome2': 'Jadi'})
group_meta.setdefault("Benji e Moglie",           {'tipo': 'singolo',  'nome':  'Benji'})
group_meta.setdefault("Luigi e Anna",             {'tipo': 'coppia',   'nome1': 'Luigi',          'nome2': 'Anna'})
group_meta.setdefault("Cesare e Daniela",         {'tipo': 'coppia',   'nome1': 'Cesare',         'nome2': 'Daniela'})
group_meta.setdefault("Tomei e Giulia",           {'tipo': 'coppia',   'nome1': 'Tomei',          'nome2': 'Giulia'})
group_meta.setdefault("Letizia e Nic",            {'tipo': 'coppia',   'nome1': 'Letizia',        'nome2': 'Nic'})
group_meta.setdefault("Carmen e Luca",            {'tipo': 'coppia',   'nome1': 'Carmen',         'nome2': 'Luca'})
group_meta.setdefault("Famiglia Rita",            {'tipo': 'famiglia', 'cognome': 'Rita'})
group_meta.setdefault("Gloria e Francesco (inf)", {'tipo': 'coppia',   'nome1': 'Gloria',         'nome2': 'Francesco'})
group_meta.setdefault("Roberta e Famiglia",       {'tipo': 'famiglia', 'cognome': 'Roberta'})
group_meta.setdefault("Stefano e Famiglia",       {'tipo': 'famiglia', 'cognome': 'Stefano'})
group_meta.setdefault("Gianluca e Adriana",       {'tipo': 'coppia',   'nome1': 'Gianluca',       'nome2': 'Adriana'})
group_meta.setdefault("Daniele Mamusa",           {'tipo': 'singolo',  'nome':  'Daniele Mamusa'})
group_meta.setdefault("Zio Silvano",              {'tipo': 'singolo',  'nome':  'Zio Silvano'})
group_meta.setdefault("Andrea Musio e Famiglia",  {'tipo': 'famiglia', 'cognome': 'Andrea Musio'})
group_meta.setdefault("Nicola Sorella",           {'tipo': 'singolo',  'nome':  'Nicola (sorella)'})
group_meta.setdefault("Angelo Anglano e Moglie",  {'tipo': 'singolo',  'nome':  'Angelo Anglano'})

# ─── Write to Excel ────────────────────────────────────────────────────────────

# Styles
header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
header_fill = PatternFill("solid", fgColor="1F4E79")
group_fill_even = PatternFill("solid", fgColor="D6E4F0")
group_fill_odd = PatternFill("solid", fgColor="EBF5FB")
alt_fill = PatternFill("solid", fgColor="F2F9FF")
thin = Side(style="thin", color="AAAAAA")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)

# Header row
ws.append(headers)
for col_idx, _ in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col_idx)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = border

ws.row_dimensions[1].height = 28

# Track groups for color banding
groups_seen = {}
group_color_idx = 0

for data in rows:
    gruppo, nome, ruolo, note = data
    if gruppo not in groups_seen:
        groups_seen[gruppo] = group_color_idx % 2
        group_color_idx += 1
    
    row_idx = ws.max_row + 1
    ws.cell(row=row_idx, column=1).value = gruppo
    ws.cell(row=row_idx, column=2).value = nome
    ws.cell(row=row_idx, column=3).value = ruolo
    # columns 4-7 (Adulto, Bambino 8+, Bambino 3-8, Bambino 0-3) left blank for user
    # columns 8-9 (Allergie, Probabilità) left blank
    ws.cell(row=row_idx, column=10).value = note

    fill = group_fill_even if groups_seen[gruppo] == 0 else group_fill_odd
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=row_idx, column=col_idx)
        cell.fill = fill
        cell.border = border
        if col_idx == 1:
            cell.alignment = left
        elif col_idx == 2:
            cell.alignment = left
            cell.font = Font(name="Calibri", size=10)
        elif col_idx in (4, 5, 6, 7):
            cell.alignment = center
        else:
            cell.alignment = left

# ─── Column widths ─────────────────────────────────────────────────────────────
col_widths = {
    1: 30,   # Gruppo
    2: 36,   # Nome
    3: 22,   # Ruolo
    4: 10,   # Adulto
    5: 14,   # Bambino 8+
    6: 14,   # Bambino 3-8
    7: 14,   # Bambino 0-3
    8: 30,   # Allergie
    9: 24,   # Probabilità
    10: 35,  # Note
}
for col, width in col_widths.items():
    ws.column_dimensions[get_column_letter(col)].width = width

# Freeze header row
ws.freeze_panes = "A2"

# Auto-filter
ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"

# ─── Sheet 2: Link Inviti ──────────────────────────────────────────────────────────
ws2 = wb.create_sheet("Link Inviti")

link_headers = ["Gruppo", "Tipo", "Link Invito"]
ws2.append(link_headers)
for col_idx in range(1, 4):
    cell = ws2.cell(row=1, column=col_idx)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = border
ws2.row_dimensions[1].height = 28

link_font_normal = Font(name="Calibri", size=10)
link_font_url    = Font(name="Calibri", size=10, color="0563C1", underline="single")

for idx, (g, meta) in enumerate(group_meta.items()):
    tipo = meta['tipo']
    params = {'tipo': tipo}
    if tipo == 'famiglia':
        params['cognome'] = meta.get('cognome', g)
    elif tipo == 'coppia':
        params['nome1'] = meta.get('nome1', '')
        params['nome2'] = meta.get('nome2', '')
    else:
        params['nome'] = meta.get('nome', g)

    url = BASE_INVITE_URL + "?" + urlencode(params)
    row_fill = group_fill_even if idx % 2 == 0 else group_fill_odd

    row_idx = ws2.max_row + 1
    c1 = ws2.cell(row=row_idx, column=1, value=g)
    c1.font = link_font_normal; c1.fill = row_fill; c1.border = border; c1.alignment = left

    c2 = ws2.cell(row=row_idx, column=2, value=tipo)
    c2.font = link_font_normal; c2.fill = row_fill; c2.border = border; c2.alignment = center

    c3 = ws2.cell(row=row_idx, column=3, value=url)
    c3.hyperlink = url
    c3.font = link_font_url; c3.fill = row_fill; c3.border = border; c3.alignment = left

ws2.column_dimensions["A"].width = 32
ws2.column_dimensions["B"].width = 12
ws2.column_dimensions["C"].width = 95
ws2.freeze_panes = "A2"
ws2.auto_filter.ref = f"A1:C{ws2.max_row}"

out_path = r"d:\TheWedding27\lista_invitati_matrimonio.xlsx"
wb.save(out_path)
print(f"File salvato: {out_path}")
print(f"Totale righe invitati: {ws.max_row - 1}")
print(f"Totale gruppi con link invito: {ws2.max_row - 1}")
