import streamlit as st
import pandas as pd
import unicodedata
import re
import io

# Fonction de nettoyage
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
st.write("Déposez vos fichiers CSV ci-dessous pour générer le rapport d'anomalies.")

# Zones de glisser-déposer
file_p1 = st.file_uploader("📥 Déposez le fichier 'Commun Travail P1' (Format CSV)", type=['csv'])
file_fournisseurs = st.file_uploader("📥 Déposez le fichier 'FOURNISSEURS' (Format CSV)", type=['csv'])

# Si les deux fichiers sont déposés, on affiche le bouton d'analyse
if file_p1 and file_fournisseurs:
    if st.button("🚀 Lancer l'analyse"):
        with st.spinner("Analyse des données en cours, veuillez patienter..."):
            
            # 1. Lecture des fichiers
            df_p1 = pd.read_csv(file_p1, dtype=str, sep=';', encoding='latin1')
            df_fournisseurs = pd.read_csv(file_fournisseurs, dtype=str, sep=';', encoding='latin1')

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
                    libelles = group['Libellé référence catalogue'].tolist()
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
