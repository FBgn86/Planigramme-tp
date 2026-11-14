#!/usr/bin/env python3
"""Planigramme TP - Générateur automatique de planigramme depuis fiches techniques."""
import io, re, zipfile
from xml.etree import ElementTree as ET
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from reportlab.lib.pagesizes import A3, landscape
from reportlab.pdfgen import canvas

st.set_page_config(page_title="Planigramme TP", layout="wide")

VERT, BLEU, ROUGE = "C6EFCE", "BDD7EE", "FFC7CE"
COUL_PDF = {"sale": (0.78, 0.93, 0.80), "propre": (0.74, 0.84, 0.93), "cuisson": (1.0, 0.78, 0.80)}

def extraire_texte(fichier) -> str:
    nom = fichier.name.lower()
    data = fichier.read()
    if nom.endswith((".odt", ".idt", ".docx")):
        membre = "content.xml" if nom.endswith((".odt", ".idt")) else "word/document.xml"
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            root = ET.fromstring(z.read(membre))
        return "\n".join(t.strip() for t in ("".join(e.itertext()) for e in root.iter()) if t.strip())
    if nom.endswith(".pdf"):
        from pypdf import PdfReader
        return "\n".join((p.extract_text() or "") for p in PdfReader(io.BytesIO(data)).pages)
    if nom.endswith(".xlsx"):
        wb = load_workbook(io.BytesIO(data), data_only=True)
        return "\n".join(str(c.value) for ws in wb for row in ws.iter_rows() for c in row if c.value)
    return data.decode("utf-8", errors="ignore")

def type_technique(libelle: str) -> str:
    l = libelle.lower()
    return "surveillance" if any(m in l for m in ("cuisson", "cuire", "four", "pousse", "infusion", "mijoter")) else "travail"

def categorie(libelle: str) -> str:
    l = libelle.lower()
    if any(m in l for m in ("cuisson", "cuire", "four", "pousse")): return "cuisson"
    if any(m in l for m in ("laver", "éplucher", "eplucher", "ébarber", "vider", "écailler")): return "sale"
    return "propre"

def planifier(taches):
    plan, fin_actif, en_cours = [], 0, []
    for t in taches:
        debut = fin_actif if t["type"] == "travail" else 0
        while any(abs(debut - a["debut"]) < 5 for a in en_cours): debut += 5
        t.update(debut=debut, fin=debut + t["duree"])
        if t["type"] == "travail": fin_actif = t["fin"]
        en_cours = [a for a in en_cours if a["fin"] > debut] + [t]
    return plan

def generer_xlsx(plan) -> bytes:
    wb = Workbook()
    ws = wb.active; ws.title = "Planigramme"
    ws.append(["Heure début", "Heure fin", "Technique", "Type", "Catégorie"])
    for t in plan:
        ws.append([f"{t['debut']//60:02d}:{t['debut']%60:02d}", f"{t['fin']//60:02d}:{t['fin']%60:02d}",
                   t["nom"], t["type"], categorie(t["nom"])])
    for row in ws.iter_rows(min_row=2):
        fill = PatternFill("solid", fgColor={"sale": VERT, "propre": BLEU, "cuisson": ROUGE}[categorie(row[2].value)])
        for c in row: c.fill = fill
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

def generer_pdf(plan) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=landscape(A3))
    largeur, hauteur = landscape(A3)
    c.setFont("Helvetica-Bold", 16); c.drawString(40, hauteur - 40, "Planigramme de TP")
    duree_max = max((t["fin"] for t in plan), default=60)
    echelle = (largeur - 450) / duree_max
    marge, y = 40, hauteur - 80
    c.setFont("Helvetica", 9)
    for m in range(0, int(duree_max) + 1, 15):
        x = marge + 380 + m * echelle
        c.line(x, marge, x, y + 10); c.drawCentredString(x, marge - 12, f"{m//60:02d}:{m%60:02d}")
    for t in plan:
        c.drawString(marge, y, f"{t['nom'][:60]}")
        x1, x2 = marge + 380 + t["debut"] * echelle, marge + 380 + t["fin"] * echelle
        c.setFillColorRGB(*COUL_PDF[categorie(t["nom"])])
        c.rect(x1, y - 4, x2 - x1, 14, fill=1, stroke=1)
        if t["type"] == "surveillance":
            c.setFillColorRGB(0, 0, 0)
            x = x1
            while x < x2:
                c.line(x, y - 4, min(x + 6, x2), y + 10); x += 8
        c.setFillColorRGB(0, 0, 0)
        c.drawString(x2 + 5, y, f"{t['debut']//60:02d}:{t['debut']%60:02d}")
        y -= 22
        if y < marge + 30: c.showPage(); y = hauteur - 80
    c.save(); return buf.getvalue()

# ---------------- Interface ----------------
st.title("📋 Planigramme de TP")
st.markdown("🟩 technique sale · 🟦 technique propre · 🟥 cuisson (surveillance)")

if st.button("🔄 Nouvelle analyse (réinitialiser)", type="secondary"):
    st.cache_data.clear()
    st.rerun()

fichiers = st.file_uploader("Fiches techniques (DOCX, PDF, ODT, IDT, XLSX)",
                            accept_multiple_files=True,
                            type=["docx", "pdf", "odt", "idt", "xlsx"])
duree = st.slider("Durée par défaut d'une technique (min)", 5, 60, 15)

if fichiers and st.button("🚀 Générer le planigramme", type="primary"):
    texte = "\n".join(extraire_texte(f) for f in fichiers)
    noms = [l.strip(" -•\t") for l in texte.splitlines()
            if l.strip() and not re.fullmatch(r"[\d\s.:-]+", l.strip()) and len(l.strip()) > 3]
    taches = [{"nom": n, "duree": duree, "type": type_technique(n)} for n in dict.fromkeys(noms)]
    plan = planifier(taches)
    st.success(f"✅ {len(plan)} techniques planifiées")
    for t in plan:
        icone = "👁️" if t["type"] == "surveillance" else "🔧"
        st.write(f"🕐 **{t['debut']//60:02d}:{t['debut']%60:02d}** — {icone} {t['nom']}")

    col1, col2 = st.columns(2)
    with col1:
        st.download_button("⬇️ Excel (.xlsx)", generer_xlsx(plan), "planigramme.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
    with col2:
        st.download_button("⬇️ PDF (Gantt)", generer_pdf(plan), "planigramme.pdf",
                           "application/pdf", use_container_width=True)
