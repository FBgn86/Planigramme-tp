import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
import re
from docx import Document
import PyPDF2
from odf.opendocument import load
from odf.text import P
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Configuration de la page
st.set_page_config(page_title="Générateur de Planigramme TP", layout="wide")
st.title("📋 Générateur Automatique de Planigramme TP")

# Listes de techniques par catégorie
TECHNIQUES_SALES = [
    "éplucher", "laver", "émincer", "hacher", "tailler", "couper", "râper",
    "monder", "épépiner", "dénoyauter", "écailler", "vider", "nettoyer poisson",
    "lever filets", "désosser", "dépouiller", "écosser", "égrainer", "battre viande"
]

TECHNIQUES_PROPRES = [
    "nettoyer plan de travail", "désinfecter", "ranger", "organiser poste",
    "préparer matériel", "vérifier température frigo", "contrôler DLC",
    "étiqueter", "filmer", "stocker", "trier déchets", "nettoyer ustensiles"
]

TECHNIQUES_CUISSON = [
    "cuire au four", "frire", "sauter", "griller", "rôtir", "braiser", "mijoter",
    "pocher", "cuire vapeur", "sous vide", "blanchir", "confire", "dorer",
    "réchauffer", "maintenir au chaud", "cuisson basse température"
]

def detecter_categorie(technique):
    """Détecte la catégorie d'une technique"""
    tech_lower = technique.lower()
    for t in TECHNIQUES_SALES:
        if t in tech_lower:
            return "sale", "🟢"
    for t in TECHNIQUES_PROPRES:
        if t in tech_lower:
            return "propre", "🔵"
    for t in TECHNIQUES_CUISSON:
        if t in tech_lower:
            return "cuisson", "🔴"
    return "autre", "⚪"

def extraire_texte(fichier):
    """Extrait le texte selon le format du fichier"""
    nom = fichier.name.lower()

    if nom.endswith('.docx'):
        doc = Document(fichier)
        return "\n".join([p.text for p in doc.paragraphs])

    elif nom.endswith('.pdf'):
        reader = PyPDF2.PdfReader(fichier)
        texte = ""
        for page in reader.pages:
            texte += page.extract_text() + "\n"
        return texte

    elif nom.endswith('.odt'):
        doc = load(fichier)
        texte = []
        for para in doc.getElementsByType(P):
            texte.append(str(para))
        return "\n".join(texte)

    elif nom.endswith('.xlsx'):
        df = pd.read_excel(fichier, header=None)
        return df.to_string()

    elif nom.endswith('.idt') or nom.endswith('.txt'):
        return fichier.read().decode('utf-8', errors='ignore')

    return ""

def parser_fiche_technique(texte, nom_fichier):
    """Parse une fiche technique pour extraire les techniques et durées"""
    lignes = texte.split('\n')
    techniques = []

    # Patterns pour détecter les techniques et durées
    pattern_duree = r'(\d+)\s*(min|minute|h|heure)'

    for ligne in lignes:
        ligne = ligne.strip()
        if not ligne or len(ligne) < 3:
            continue

        # Détection de la durée
        duree = 15  # durée par défaut
        match = re.search(pattern_duree, ligne.lower())
        if match:
            val = int(match.group(1))
            unite = match.group(2)
            duree = val * 60 if unite in ['h', 'heure'] else val

        # Détection si c'est une technique de surveillance (cuisson)
        est_surveillance = any(t in ligne.lower() for t in TECHNIQUES_CUISSON)

        # Nettoyage du nom de la technique
        nom_tech = re.sub(r'\d+\s*(min|minute|h|heure)', '', ligne, flags=re.IGNORECASE).strip()
        nom_tech = nom_tech.strip(':-•* ')

        if len(nom_tech) > 2:
            categorie, emoji = detecter_categorie(nom_tech)
            techniques.append({
                'technique': nom_tech,
                'duree': duree,
                'categorie': categorie,
                'emoji': emoji,
                'surveillance': est_surveillance,
                'source': nom_fichier
            })

    return techniques

def planifier_techniques(toutes_techniques):
    """Organise les techniques dans le temps en respectant les contraintes"""
    # Trier par source puis par ordre d'apparition
    toutes_techniques.sort(key=lambda x: (x['source'],))

    heure_debut = datetime(2026, 11, 14, 8, 0)  # 8h00
    lignes_plan = []
    heure_courante = heure_debut
    heure_fin_surveillance = heure_debut

    for tech in toutes_techniques:
        if tech['surveillance']:
            # Les cuissons peuvent se faire en parallèle
            heure_fin = heure_courante + timedelta(minutes=tech['duree'])
            if heure_fin > heure_fin_surveillance:
                heure_fin_surveillance = heure_fin
        else:
            # Travail effectif : séquentiel avec décalage de 5 min si conflit
            heure_debut_tech = heure_courante
            # Vérifier s'il y a chevauchement avec une autre tâche effective
            for ligne in lignes_plan:
                if not ligne['surveillance']:
                    fin_precedente = ligne['heure_fin']
                    if heure_debut_tech < fin_precedente:
                        heure_debut_tech = fin_precedente + timedelta(minutes=5)

            heure_fin = heure_debut_tech + timedelta(minutes=tech['duree'])
            heure_courante = heure_fin

            lignes_plan.append({
                'heure_debut': heure_debut_tech,
                'heure_fin': heure_fin,
                'technique': tech['technique'],
                'duree': tech['duree'],
                'categorie': tech['categorie'],
                'emoji': tech['emoji'],
                'surveillance': False,
                'source': tech['source']
            })
            continue

        lignes_plan.append({
            'heure_debut': heure_courante,
            'heure_fin': heure_courante + timedelta(minutes=tech['duree']),
            'technique': tech['technique'],
            'duree': tech['duree'],
            'categorie': tech['categorie'],
            'emoji': tech['emoji'],
            'surveillance': True,
            'source': tech['source']
        })

    return sorted(lignes_plan, key=lambda x: x['heure_debut'])

def creer_excel_colore(lignes_plan):
    """Crée le fichier Excel avec une seule feuille colorée"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Planigramme"

    # En-têtes
    headers = ["Heure début", "Heure fin", "Durée (min)", "Technique", "Catégorie", "Type", "Source"]
    ws.append(headers)

    # Styles
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    fill_sale = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # Vert
    fill_propre = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")  # Bleu
    fill_cuisson = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # Rouge
    fill_autre = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")  # Gris

    # Style en-tête
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    # Remplissage des données
    for i, ligne in enumerate(lignes_plan, start=2):
        ws.cell(row=i, column=1, value=ligne['heure_debut'].strftime("%H:%M"))
        ws.cell(row=i, column=2, value=ligne['heure_fin'].strftime("%H:%M"))
        ws.cell(row=i, column=3, value=ligne['duree'])
        ws.cell(row=i, column=4, value=f"{ligne['emoji']} {ligne['technique']}")
        ws.cell(row=i, column=5, value=ligne['categorie'].capitalize())
        ws.cell(row=i, column=6, value="Surveillance" if ligne['surveillance'] else "Travail effectif")
        ws.cell(row=i, column=7, value=ligne['source'])

        # Application de la couleur selon la catégorie
        if ligne['categorie'] == 'sale':
            fill = fill_sale
        elif ligne['categorie'] == 'propre':
            fill = fill_propre
        elif ligne['categorie'] == 'cuisson':
            fill = fill_cuisson
        else:
            fill = fill_autre

        for col in range(1, len(headers) + 1):
            ws.cell(row=i, column=col).fill = fill

    # Ajustement des colonnes
    for col in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 20

    # Sauvegarde en mémoire
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Interface utilisateur
st.header("📁 Upload des fiches techniques")
fichiers = st.file_uploader(
    "Sélectionnez vos fiches techniques (.docx, .pdf, .odt, .xlsx, .idt)",
    accept_multiple_files=True,
    type=['docx', 'pdf', 'odt', 'xlsx', 'idt', 'txt']
)

if fichiers:
    st.success(f"✅ {len(fichiers)} fichier(s) uploadé(s)")

    if st.button("🚀 Générer le planigramme", type="primary"):
        with st.spinner("Analyse des fiches en cours..."):
            toutes_techniques = []

            for fichier in fichiers:
                texte = extraire_texte(fichier)
                if texte:
                    techniques = parser_fiche_technique(texte, fichier.name)
                    toutes_techniques.extend(techniques)
                    st.info(f"📄 {fichier.name} : {len(techniques)} technique(s) détectée(s)")

            if toutes_techniques:
                lignes_plan = planifier_techniques(toutes_techniques)
                excel_file = creer_excel_colore(lignes_plan)

                st.success("✅ Planigramme généré avec succès !")

                # Aperçu
                st.subheader("Aperçu du planning")
                df_preview = pd.DataFrame([{
                    'Début': l['heure_debut'].strftime("%H:%M"),
                    'Fin': l['heure_fin'].strftime("%H:%M"),
                    'Technique': f"{l['emoji']} {l['technique']}",
                    'Catégorie': l['categorie'].capitalize(),
                    'Type': 'Surveillance' if l['surveillance'] else 'Effectif'
                } for l in lignes_plan])
                st.dataframe(df_preview, use_container_width=True)

                # Téléchargement
                st.download_button(
                    label="📥 Télécharger le planigramme Excel (feuille unique colorée)",
                    data=excel_file,
                    file_name="planigramme_tp.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("❌ Aucune technique détectée. Vérifiez le format de vos fichiers.")
else:
    st.info("👆 Veuillez uploader vos fiches techniques pour commencer")
