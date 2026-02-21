import streamlit as st
import pandas as pd
import unicodedata
import re
import io

# ==========================================
# FONCTIONS OUTILS
# ==========================================
def normalize_string(text):
    if pd.isna(text) or str(text).strip() == "":
        return ""
    text = str(text)
    text = ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')
    text = text.lower()
    text = re.sub(r'[\s\-_]', '', text)
    return text

# ==========================================
# CONFIGURATION DE LA PAGE
# ==========================================
st.set_page_config(page_title="Vérificateur P1", page_icon="✨", layout="wide")

# ==========================================
# EN-TÊTE DE L'APPLICATION
# ==========================================
st.title("✨ Assistant de Vérification P1 & Fournisseurs")
st.markdown("""
Bienvenue dans votre outil de nettoyage de base de données. 
Cet utilitaire croise intelligemment vos données pour détecter les **équipements en doublon**, les **anomalies fournisseurs** et les **codes orphelins**.
""")
st.divider()

# ==========================================
# INTERFACE PRINCIPALE (Colonnes)
# ==========================================
col_gauche, col_droite = st.columns([1, 2], gap="large")

with col_gauche:
    st.header("📂 Étape 1 : Import")
    st.info("Déposez votre fichier Excel. L'outil trouvera automatiquement les bons onglets (Commun P1 / Fournisseurs).")
    file_excel = st.file_uploader("Glissez votre fichier ici (.xlsx)", type=['xlsx'], label_visibility="collapsed")

with col_droite:
    st.header("⚙️ Étape 2 : Analyse & Résultats")
    
    if not file_excel:
        st.write("👈 *Veuillez importer un fichier dans la zone de gauche pour commencer.*")
    
    if file_excel:
        if st.button("🚀 Lancer le diagnostic complet", type="primary", use_container_width=True):
            with st.spinner("Analyse des milliers de lignes en cours... ⏳"):
                try:
                    # --- LECTURE ET RECHERCHE DES FEUILLES ---
                    xl = pd.ExcelFile(file_excel)
                    feuilles_disponibles = xl.sheet_names
                    
                    nom_feuille_p1 = None
                    nom_feuille_fournisseurs = None
                    
                    for f in feuilles_disponibles:
                        f_norm = f.strip().lower()
                        if "commun" in f_norm and "p1" in f_norm:
                            nom_feuille_p1 = f
                        elif "fournisseurs" in f_norm:
                            nom_feuille_fournisseurs = f
                    
                    if not nom_feuille_p1 or not nom_feuille_fournisseurs:
                        st.error(f"❌ Onglets introuvables ! Feuilles détectées : {feuilles_disponibles}.")
                        st.stop() # Arrête l'exécution ici

                    # --- CHARGEMENT DES DONNÉES ---
                    df_p1 = pd.read_excel(xl, sheet_name=nom_feuille_p1, dtype=str)
                    df_fournisseurs = pd.read_excel(xl, sheet_name=nom_feuille_fournisseurs, dtype=str)

                    # --- ANALYSE FOURNISSEURS ---
                    df_fournisseurs['Nom_Norm'] = df_fournisseurs['Nom'].apply(normalize_string)
                    fournisseurs_dict = dict(zip(df_fournisseurs['Code'].dropna(), df_fournisseurs['Nom_Norm'].dropna()))

                    doublons_fournisseurs = []
                    for nom_norm, group in df_fournisseurs.groupby('Nom_Norm'):
                        if len(group['Code'].unique()) > 1 and nom_norm != "":
                            noms_originaux = " / ".join(group['Nom'].dropna().unique())
                            codes_lies = " ; ".join(group['Code'].dropna().unique())
                            doublons_fournisseurs.append({
                                'Fabricant (Nom unifié)': noms_originaux,
                                'Codes Fournisseurs multiples': codes_lies
                            })
                    df_anomalies_fournisseurs = pd.DataFrame(doublons_fournisseurs)

                    # --- ANALYSE P1 & ORPHELINS ---
                    df_p1['K_Norm'] = df_p1['Code barre référence'].apply(normalize_string)
                    df_p1['L_Original'] = df_p1['Code référence constructeur'].fillna("")

                    codes_p1_uniques = set(df_p1['L_Original'][df_p1['L_Original'] != ""])
                    codes_fournisseurs_existants = set(df_fournisseurs['Code'].dropna())
                    codes_orphelins = codes_p1_uniques - codes_fournisseurs_existants
                    df_orphelins = pd.DataFrame([{"Code constructeur utilisé dans P1 mais inconnu dans la base": c} for c in codes_orphelins])

                    def get_manufacturer_norm(code_l):
                        if pd.isna(code_l) or code_l == "": return ""
                        if code_l in fournisseurs_dict: return fournisseurs_dict[code_l]
                        return normalize_string(code_l)

                    df_p1['Fabricant_Compare'] = df_p1['L_Original'].apply(get_manufacturer_norm)

                    # --- RECHERCHE DOUBLONS P1 ---
                    duplicates = []
                    for (k_norm, fab_norm), group in df_p1.groupby(['K_Norm', 'Fabricant_Compare']):
                        if len(group) > 1 and k_norm != "":
                            codes_catalogue = group['Code référence catalogue'].tolist()
                            libelles = group['Libellé référence catalogue'].tolist() if 'Libellé référence catalogue' in group.columns else ["N/A"] * len(group)
                            codes_barre = group['Code barre référence'].tolist()
                            codes_constructeur = group['Code référence constructeur'].tolist()
                            
                            l_norms = set(normalize_string(l) for l in group['Code référence constructeur'].dropna())
                            raison = "Doublon exact (Tolérance casse/espaces/accents)" if len(l_norms) <= 1 else "Code barre identique, mais constructeurs différents rattachés au même Fabricant"
                                
                            duplicates.append({
                                'Libellés des équipements': " | ".join(map(str, set(libelles))),
                                'Codes référence catalogue': " ; ".join(map(str, codes_catalogue)),
                                'Codes barre saisis': " ; ".join(map(str, set(codes_barre))),
                                'Codes constructeurs saisis': " ; ".join(map(str, set(codes_constructeur))),
                                'Raison du doublon': raison
                            })
                    df_report = pd.DataFrame(duplicates)

                    # --- GÉNÉRATION EXCEL ---
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_report.to_excel(writer, sheet_name='1 - Doublons Equipements', index=False)
                        if not df_anomalies_fournisseurs.empty:
                            df_anomalies_fournisseurs.to_excel(writer, sheet_name='2 - Doublons Fournisseurs', index=False)
                        if not df_orphelins.empty:
                            df_orphelins.to_excel(writer, sheet_name='3 - Orphelins P1', index=False)
                    
                    # ==========================================
                    # AFFICHAGE DU TABLEAU DE BORD (UX Améliorée)
                    # ==========================================
                    st.success("✅ Traitement terminé avec succès !")
                    st.divider()
                    
                    st.subheader("📊 Résumé des anomalies détectées")
                    
                    # Création de jolies métriques alignées
                    m1, m2, m3 = st.columns(3)
                    m1.metric(label="Doublons Équipements", value=f"{len(df_report)} groupes")
                    m2.metric(label="Anomalies Fournisseurs", value=f"{len(df_anomalies_fournisseurs)} cas")
                    m3.metric(label="Codes Orphelins", value=f"{len(df_orphelins)} codes")

                    # Aperçu interactif caché dans des menus déroulants
                    st.write("") # Espace
                    if not df_report.empty:
                        with st.expander("👀 Voir un aperçu des équipements en doublon"):
                            st.dataframe(df_report.head(15), use_container_width=True)
                    
                    st.write("") # Espace
                    
                    # Gros bouton de téléchargement final
                    st.download_button(
                        label="📥 TÉLÉCHARGER LE RAPPORT COMPLET (.xlsx)",
                        data=output.getvalue(),
                        file_name="Rapport_Verification_Global.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="primary"
                    )

                except Exception as e:
                    st.error(f"❌ Une erreur inattendue s'est produite : {e}")
