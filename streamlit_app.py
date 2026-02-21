import streamlit as st
import pandas as pd
import unicodedata
import re
import io

# Fonction de nettoyage des caractères (majuscules, accents, espaces...)
def normalize_string(text):
    if pd.isna(text) or str(text).strip() == "":
        return ""
    text = str(text)
    text = ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')
    text = text.lower()
    text = re.sub(r'[\s\-_]', '', text)
    return text

# Configuration de la page web
st.set_page_config(page_title="Vérificateur de Doublons", layout="centered")
st.title("🛠️ Outil de Vérification P1 & Fournisseurs")
st.write("Déposez votre fichier Excel. L'outil détectera automatiquement les onglets, **peu importe s'ils sont écrits en majuscules ou minuscules**.")

# Zone de glisser-déposer unique pour n'importe quel fichier Excel
file_excel = st.file_uploader("📥 Déposez votre fichier Excel (.xlsx)", type=['xlsx'])

if file_excel:
    if st.button("🚀 Lancer l'analyse"):
        with st.spinner("Lecture du fichier et analyse en cours, veuillez patienter..."):
            try:
                # 1. On analyse la structure du fichier
                xl = pd.ExcelFile(file_excel)
                feuilles_disponibles = xl.sheet_names
                
                # Recherche ultra-flexible des noms des feuilles (ignore la casse et les espaces)
                nom_feuille_p1 = None
                nom_feuille_fournisseurs = None
                
                for f in feuilles_disponibles:
                    f_norm = f.strip().lower() # Tout en minuscules sans espaces aux extrémités
                    if "commun" in f_norm and "p1" in f_norm:
                        nom_feuille_p1 = f
                    elif "fournisseurs" in f_norm:
                        nom_feuille_fournisseurs = f
                
                # Vérification si les feuilles ont bien été trouvées
                if not nom_feuille_p1 or not nom_feuille_fournisseurs:
                    st.error(f"❌ Onglets introuvables ! Votre fichier Excel contient les onglets suivants : {feuilles_disponibles}.")
                    st.warning("Veuillez vérifier que l'un des onglets contient le mot 'Commun' et 'P1', et l'autre le mot 'Fournisseurs'.")
                else:
                    # Lecture des données depuis les feuilles trouvées dynamiquement
                    df_p1 = pd.read_excel(xl, sheet_name=nom_feuille_p1, dtype=str)
                    df_fournisseurs = pd.read_excel(xl, sheet_name=nom_feuille_fournisseurs, dtype=str)

                    # 2. Analyse Fournisseurs
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

                    # 3. Analyse P1 & Orphelins
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

                    # 4. Recherche des doublons
                    duplicates = []
                    for (k_norm, fab_norm), group in df_p1.groupby(['K_Norm', 'Fabricant_Compare']):
                        if len(group) > 1 and k_norm != "":
                            codes_catalogue = group['Code référence catalogue'].tolist()
                            
                            # Gestion de la colonne Libellé si elle existe
                            if 'Libellé référence catalogue' in group.columns:
                                libelles = group['Libellé référence catalogue'].tolist()
                            else:
                                libelles = ["Non disponible"] * len(group)
                                
                            codes_barre = group['Code barre référence'].tolist()
                            codes_constructeur = group['Code référence constructeur'].tolist()
                            
                            l_norms = set(normalize_string(l) for l in group['Code référence constructeur'].dropna())
                            raison = "Doublon exact (aux espaces/tirets/accents/casse près)" if len(l_norms) <= 1 else "Code barre identique, mais rattachés au même Fabricant via des codes différents"
                                
                            duplicates.append({
                                'Libellés des équipements': " | ".join(map(str, set(libelles))),
                                'Codes référence catalogue': " ; ".join(map(str, codes_catalogue)),
                                'Codes barre saisis': " ; ".join(map(str, set(codes_barre))),
                                'Codes constructeurs saisis': " ; ".join(map(str, set(codes_constructeur))),
                                'Raison du doublon': raison
                            })
                    df_report = pd.DataFrame(duplicates)

                    # 5. Création du fichier Excel en mémoire
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_report.to_excel(writer, sheet_name='1 - Doublons Equipements', index=False)
                        if not df_anomalies_fournisseurs.empty:
                            df_anomalies_fournisseurs.to_excel(writer, sheet_name='2 - Doublons Fournisseurs', index=False)
                        if not df_orphelins.empty:
                            df_orphelins.to_excel(writer, sheet_name='3 - Orphelins P1', index=False)
                    
                    # Affichage des résultats
                    st.success(f"✅ Analyse terminée ! {len(df_report)} groupes de doublons trouvés.")
                    
                    # 6. Bouton de téléchargement
                    st.download_button(
                        label="📥 Télécharger le Rapport Complet (.xlsx)",
                        data=output.getvalue(),
                        file_name="Rapport_Verification_Global.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            
            except Exception as e:
                st.error(f"❌ Une erreur inattendue s'est produite lors de la lecture du fichier : {e}")
