import streamlit as st
import os
import json
import tempfile
import base64
from pptx import Presentation

# Import sécurisé pour la conversion PDF (Windows uniquement)
try:
    import pythoncom
    import comtypes.client
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False

# ==========================================
# CONFIGURATION ET SAUVEGARDE
# ==========================================
st.set_page_config(page_title="Générateur de CV PowerPoint & PDF", layout="wide")
DATA_FILE = "cv_data.json"

def load_data():
    """Charge les données depuis le fichier JSON s'il existe."""
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            pass
    return None

def save_data(silent=True):
    """Sauvegarde l'état actuel dans le fichier JSON (Sauvegarde Automatique)."""
    data = {
        'profil': st.session_state.profil,
        'competences': st.session_state.competences,
        'diplomes': st.session_state.diplomes,
        'certifications': st.session_state.certifications,
        'experiences': st.session_state.experiences
    }
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
    if not silent:
        st.sidebar.success("✅ Données sauvegardées localement !")

# ==========================================
# INITIALISATION DE L'ÉTAT (SESSION STATE)
# ==========================================
if 'initialized' not in st.session_state:
    saved = load_data()
    if saved:
        st.session_state.profil = saved.get('profil', {})
        st.session_state.competences = saved.get('competences', [])
        st.session_state.diplomes = saved.get('diplomes', [])
        st.session_state.certifications = saved.get('certifications', [])
        st.session_state.experiences = saved.get('experiences', [])
    else:
        st.session_state.profil = {
            'prenom': '', 'nom': '', 'poste': '', 'experience_ans': '', 
            'resume': '', 'expertise': '', 'positionnement': '', 'photo_b64': None
        }
        st.session_state.competences = []
        st.session_state.diplomes = []
        st.session_state.certifications = []
        st.session_state.experiences = []
    
    st.session_state.initialized = True

# Indicateur de sauvegarde dans la barre latérale
with st.sidebar:
    st.title("💾 Stockage")
    st.success("🟢 **Sauvegarde automatique activée.** \nToutes vos modifications sont enregistrées en temps réel sur votre ordinateur.")

# ==========================================
# FONCTIONS DE MANIPULATION DU PPTX
# ==========================================
def replace_profile_picture(slide, new_image_path):
    """Trouve l'image de profil dans le slide et la remplace par la nouvelle."""
    target_shape = None
    for shape in slide.shapes:
        if getattr(shape, "shape_type", None) == 13: # 13 = Image
            target_shape = shape
            break
            
    if target_shape:
        left, top = target_shape.left, target_shape.top
        width, height = target_shape.width, target_shape.height
        
        sp = target_shape._element
        sp.getparent().remove(sp)
        
        slide.shapes.add_picture(new_image_path, left, top, width, height)

def generate_pptx_cv(template_path, data):
    prs = Presentation(template_path)
    p = data['profil']
    
    # 0. Remplacement de la photo (si fournie)
    if p.get('photo_b64'):
        img_data = base64.b64decode(p['photo_b64'])
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
            tmp_img.write(img_data)
            tmp_img_path = tmp_img.name
        
        replace_profile_picture(prs.slides[0], tmp_img_path)
        os.unlink(tmp_img_path)

    # Préparation des textes
    exp_str = f"+{p['experience_ans']} années d'expérience" if p['experience_ans'].isdigit() else p['experience_ans']
    
    standard_replacements = {
        "[PRÉNOM]": p.get('prenom', ''),
        "[NOM]": p.get('nom', '').upper(),
        "[POSTE / FONCTION]": p.get('poste', ''),
        "[X années d’expérience]": exp_str,
        "[X années d'expérience]": exp_str,
        "[RÉSUMÉ PROFESSIONNEL — 2 à 4 lignes]": p.get('resume', ''),
        "[EXPERTISE / VALEUR AJOUTÉE — 2 à 4 lignes]": p.get('expertise', ''),
        "[POSITIONNEMENT / DOMAINES D’INTERVENTION — 2 à 4 lignes]": p.get('positionnement', '')
    }

    # Compétences
    for i in range(4):
        val = data['competences'][i] if i < len(data['competences']) else ""
        standard_replacements[f"#[COMPÉTENCE {i+1}]"] = f"#{val}" if val else ""

    # Diplômes
    for i in range(2):
        val = f"{data['diplomes'][i]['ecole']}\n{data['diplomes'][i]['titre']}" if i < len(data['diplomes']) else ""
        standard_replacements[f"[DIPLÔME / FORMATION {i+1}]"] = val

    # Certifications
    for i in range(4):
        val = f"{data['certifications'][i]['nom']}" if i < len(data['certifications']) else ""
        standard_replacements[f"[CERTIFICATION {i+1}]"] = val
        if i == 2: standard_replacements["[CERTIFICATION / LANGUE]"] = val

    # Expériences
    for i in range(4):
        idx = i + 1
        if i < len(data['experiences']):
            exp = data['experiences'][i]
            standard_replacements[f"[ENTREPRISE {idx}]"] = exp['entreprise']
            standard_replacements[f"[POSTE {idx}]"] = exp['poste']
            standard_replacements[f"[CONTEXTE / MISSION {idx} — 1 à 2 lignes]"] = exp['contexte']
            standard_replacements[f"[CONTEXTE / MISSION {idx} — 1 à 3 lignes]"] = exp['contexte']
            for j in range(3):
                real_val = exp['realisations'][j] if j < len(exp['realisations']) else ""
                standard_replacements[f"[RÉALISATION {idx}.{j+1}]"] = real_val
        else:
            standard_replacements[f"[ENTREPRISE {idx}]"] = ""
            standard_replacements[f"[POSTE {idx}]"] = ""
            standard_replacements[f"[CONTEXTE / MISSION {idx} — 1 à 2 lignes]"] = ""
            standard_replacements[f"[CONTEXTE / MISSION {idx} — 1 à 3 lignes]"] = ""
            for j in range(3):
                standard_replacements[f"[RÉALISATION {idx}.{j+1}]"] = ""

    # 1. Extraction et tri de tous les paragraphes de haut en bas (axe Y)
    paragraphs_with_y = []
    for slide in prs.slides:
        for shape in slide.shapes:
            def extract_paras(shp, current_top):
                if getattr(shp, "shape_type", None) == 6:
                    for c in shp.shapes:
                        extract_paras(c, current_top + getattr(c, "top", 0))
                if getattr(shp, "has_text_frame", False):
                    for paragraph in shp.text_frame.paragraphs:
                        paragraphs_with_y.append((current_top, paragraph))
                if getattr(shp, "has_table", False):
                    for row in shp.table.rows:
                        for cell in row.cells:
                            if getattr(cell, "text_frame", None):
                                for paragraph in cell.text_frame.paragraphs:
                                    paragraphs_with_y.append((current_top, paragraph))
            extract_paras(shape, getattr(shape, "top", 0))
            
    paragraphs_with_y.sort(key=lambda x: x[0])
    
    # 2. Remplacement textuel
    meta_counter = 0
    paragraphs_to_delete = []

    for _, paragraph in paragraphs_with_y:
        p_text = "".join(run.text for run in paragraph.runs)
        
        # Remplacement de la ligne "META"
        meta_tag = "[TYPE / CHARGE / DURÉE / LOCALISATION]"
        if meta_tag in p_text:
            if meta_counter < len(data['experiences']):
                exp = data['experiences'][meta_counter]
                meta_val = f"{exp['type']} – {exp['charge']} ({exp['duree']}), {exp['localisation']}"
            else:
                meta_val = ""
                
            for run in paragraph.runs:
                if meta_tag in run.text:
                    run.text = run.text.replace(meta_tag, meta_val)
            meta_counter += 1

        # Remplacement des autres balises
        for key, val in standard_replacements.items():
            if key in p_text:
                for run in paragraph.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, str(val))
        
        # Nettoyage : Si la ligne est vide (ou ne contient que des tirets/espaces car l'expérience est vide)
        new_p_text = "".join(run.text for run in paragraph.runs)
        if not new_p_text.strip(" -–—,.|•\t\n\r"): 
            paragraphs_to_delete.append(paragraph)

    # Suppression effective des paragraphes vides (pour ne pas casser l'espacement)
    for p in paragraphs_to_delete:
        p_element = p._p
        parent = p_element.getparent()
        if parent is not None:
            parent.remove(p_element)
                
    tmp_pptx = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    prs.save(tmp_pptx.name)
    tmp_pptx.close()
    return tmp_pptx.name

def convert_pptx_to_pdf(pptx_path):
    if not COM_AVAILABLE:
        raise Exception("Librairie comtypes absente.")
    
    pdf_path = pptx_path.replace('.pptx', '.pdf')
    pythoncom.CoInitialize()
    powerpoint = None
    try:
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        deck = powerpoint.Presentations.Open(os.path.abspath(pptx_path), WithWindow=False)
        deck.SaveAs(os.path.abspath(pdf_path), 32)
        deck.Close()
    finally:
        pass
    return pdf_path

# ==========================================
# INTERFACE UTILISATEUR
# ==========================================
st.title("💼 Générateur de CV Avancé")

tab1, tab2, tab3 = st.tabs(["1. Mon Profil", "2. Parcours & Expériences", "3. Génération & Export"])

# --- ONGLET 1 : PROFIL ---
with tab1:
    st.header("Informations générales")
    
    col_img, col_info = st.columns([1, 3])
    with col_img:
        st.write("📷 **Photo de profil**")
        if st.session_state.profil.get('photo_b64'):
            st.image(base64.b64decode(st.session_state.profil['photo_b64']), width=150)
        
        uploaded_file = st.file_uploader("Modifier la photo", type=["jpg", "jpeg", "png"], label_visibility="collapsed")
        if uploaded_file is not None:
            st.session_state.profil['photo_b64'] = base64.b64encode(uploaded_file.read()).decode()

    with col_info:
        c1, c2 = st.columns(2)
        st.session_state.profil['prenom'] = c1.text_input("Prénom", value=st.session_state.profil.get('prenom', ''))
        st.session_state.profil['nom'] = c2.text_input("Nom", value=st.session_state.profil.get('nom', ''))
        st.session_state.profil['poste'] = c1.text_input("Poste / Fonction", value=st.session_state.profil.get('poste', ''), placeholder="Ex: Product Owner")
        st.session_state.profil['experience_ans'] = c2.text_input("Années d'expérience", value=st.session_state.profil.get('experience_ans', ''), placeholder="Ex: 3")
    
    st.subheader("Textes de présentation")
    st.info("💡 **Conseil** : Formulez des phrases percutantes à la 3ème personne, comme dans l'exemple de Dan.")
    
    st.session_state.profil['resume'] = st.text_area(
        "Résumé (2-4 lignes)", 
        value=st.session_state.profil.get('resume', ''), 
        help="Ex: Issu du Programme Grande École de l'ESSEC, Dan cumule trois années d’expérience dans le domaine du conseil.",
        height=80
    )
    st.session_state.profil['expertise'] = st.text_area(
        "Expertise / Valeur ajoutée", 
        value=st.session_state.profil.get('expertise', ''), 
        help="Ex: En pilotant des projets de transformation digitale, il a développé une appétence pour la résolution de problème complexes...",
        height=80
    )
    st.session_state.profil['positionnement'] = st.text_area(
        "Domaines d'intervention", 
        value=st.session_state.profil.get('positionnement', ''), 
        help="Ex: Dan intervient à la fois dans le cadrage de la feuille de route digitale et dans la coordination des développements...",
        height=80
    )

# --- ONGLET 2 : AJOUT DE DONNÉES ---
with tab2:
    st.header("Alimenter la base de données")
    st.write("Ajoutez ici toutes vos expériences. Vous pourrez choisir lesquelles afficher à l'étape suivante.")
    
    c1, c2 = st.columns(2)
    
    with c1:
        with st.form("form_comp", clear_on_submit=True):
            st.subheader("Tags / Compétences")
            new_comp = st.text_input("Nouveau Tag (sans le #)", placeholder="Ex: discovery, erp, data...")
            if st.form_submit_button("➕ Ajouter") and new_comp:
                if new_comp not in st.session_state.competences: 
                    st.session_state.competences.append(new_comp)
                    st.success(f"Tag '{new_comp}' ajouté !")
                    
        with st.form("form_dip", clear_on_submit=True):
            st.subheader("Diplômes")
            ecole = st.text_input("École", placeholder="Ex: ESSEC Business School")
            titre = st.text_input("Titre du diplôme", placeholder="Ex: Master in Management, Programme Grande École")
            if st.form_submit_button("➕ Ajouter") and ecole:
                st.session_state.diplomes.append({"ecole": ecole, "titre": titre})
                st.success("Diplôme ajouté !")
                
        with st.form("form_cert", clear_on_submit=True):
            st.subheader("Certifications & Langues")
            nom_cert = st.text_input("Nom de la certification", placeholder="Ex: Certification PSPO / TOEIC : 990/990")
            if st.form_submit_button("➕ Ajouter") and nom_cert:
                st.session_state.certifications.append({"nom": nom_cert})
                st.success("Certification ajoutée !")

    with c2:
        with st.form("form_exp", clear_on_submit=True):
            st.subheader("Expérience Professionnelle")
            ent = st.text_input("Entreprise", placeholder="Ex: EDMOND DE ROTHSCHILD")
            poste = st.text_input("Poste", placeholder="Ex: Chef de projet")
            ctx = st.text_area("Contexte / Mission", placeholder="Ex: Accompagnement du groupe, spécialisé dans le courtage en assurance, pour la définition et l’implémentation de sa feuille de route data et IA.")
            
            st.write("**Détails de la mission**")
            c_meta1, c_meta2 = st.columns(2)
            meta_type = c_meta1.text_input("Type de contrat", placeholder="Ex: Forfait, Régie, CDI...")
            meta_charge = c_meta2.text_input("Charge de travail", placeholder="Ex: temps plein, temps partiel")
            meta_duree = c_meta1.text_input("Durée", placeholder="Ex: 2 mois, 1 an")
            meta_loc = c_meta2.text_input("Localisation", placeholder="Ex: Paris, Remote")
            
            st.write("**Réalisations (Commencez par un verbe d'action)**")
            reals = st.text_area("Listez 3 points maximum (un par ligne)", placeholder="Ex: Cartographie des cas d'usage...\nPriorisation selon les gains...\nConstruction d'un plan de pilotage...")
            
            if st.form_submit_button("➕ Ajouter l'expérience") and ent:
                real_list = [r.strip() for r in reals.split('\n') if r.strip()][:3]
                st.session_state.experiences.append({
                    "entreprise": ent, "poste": poste, "contexte": ctx,
                    "type": meta_type, "charge": meta_charge, "duree": meta_duree, "localisation": meta_loc,
                    "realisations": real_list
                })
                st.success(f"Expérience chez '{ent}' ajoutée !")

# --- ONGLET 3 : GÉNÉRATION ---
with tab3:
    st.header("Génération du document")
    st.write("Sélectionnez les éléments à inclure pour **ce CV spécifiquement**.")
    
    col_sel1, col_sel2 = st.columns(2)
    with col_sel1:
        selected_comps = st.multiselect("Compétences (Max 4 affichées)", st.session_state.competences, default=st.session_state.competences[:4])
        dip_opts = [f"{d['ecole']} - {d['titre']}" for d in st.session_state.diplomes]
        sel_dip = st.multiselect("Diplômes (Max 2 affichés)", dip_opts, default=dip_opts[:2])
        selected_diplomes = [d for d in st.session_state.diplomes if f"{d['ecole']} - {d['titre']}" in sel_dip]
    
    with col_sel2:
        cert_opts = [c['nom'] for c in st.session_state.certifications]
        sel_cert = st.multiselect("Certifications (Max 4 affichées)", cert_opts, default=cert_opts[:4])
        selected_certifications = [c for c in st.session_state.certifications if c['nom'] in sel_cert]
        
        exp_opts = [f"{e['entreprise']} - {e['poste']}" for e in st.session_state.experiences]
        sel_exp = st.multiselect("Expériences (Max 4 affichées)", exp_opts, default=exp_opts[:4])
        selected_experiences = [e for e in st.session_state.experiences if f"{e['entreprise']} - {e['poste']}" in sel_exp]

    template_file = "CV_Template_reutilisable.pptx"
    
    st.write("---")
    if not os.path.exists(template_file):
        st.error(f"❌ Fichier introuvable : `{template_file}`. Veuillez le placer dans le dossier.")
    else:
        if st.button("🚀 GÉNÉRER MON CV", type="primary", use_container_width=True):
            data_to_render = {
                "profil": st.session_state.profil,
                "competences": selected_comps,
                "diplomes": selected_diplomes,
                "certifications": selected_certifications,
                "experiences": selected_experiences
            }
            
            with st.spinner("Création du fichier PowerPoint..."):
                try:
                    pptx_path = generate_pptx_cv(template_file, data_to_render)
                    
                    c_dl1, c_dl2 = st.columns(2)
                    with open(pptx_path, "rb") as file:
                        c_dl1.download_button("📥 Télécharger le CV (PPTX)", data=file, file_name=f"CV_{st.session_state.profil.get('nom', 'Genere')}.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", use_container_width=True)
                    
                    with st.spinner("Conversion en PDF en cours..."):
                        try:
                            pdf_path = convert_pptx_to_pdf(pptx_path)
                            with open(pdf_path, "rb") as file:
                                c_dl2.download_button("📥 Télécharger le CV (PDF)", data=file, file_name=f"CV_{st.session_state.profil.get('nom', 'Genere')}.pdf", mime="application/pdf", use_container_width=True)
                        except Exception as e_pdf:
                            st.warning(f"⚠️ La conversion automatique en PDF a échoué. Ouvrez le fichier PPTX téléchargé et faites 'Enregistrer sous > PDF'.")
                            
                except Exception as e:
                    st.error(f"❌ Erreur lors du remplacement : {e}")

# ==========================================
# SAUVEGARDE AUTOMATIQUE EN FIN DE SCRIPT
# ==========================================
# Le code arrive ici à la fin de chaque interaction utilisateur. 
# On déclenche la sauvegarde silencieuse pour tout enregistrer en temps réel.
save_data(silent=True)