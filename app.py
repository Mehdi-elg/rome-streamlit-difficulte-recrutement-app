import streamlit as st
import requests
import pandas as pd
from io import BytesIO


CLIENT_ID    = st.secrets["CLIENT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]

TOKEN_URL = "https://entreprise.francetravail.fr/connexion/oauth2/access_token?realm=/partenaire"
STATS_BASE_URL = "https://api.francetravail.io/partenaire/stats-offres-demandes-emploi"


DIFFICULTE_LABELS = {
    1: "Très faible",
    2: "Faible",
    3: "Moyenne",
    4: "Élevée",
    5: "Très élevée",
}


def get_token() -> str:
    resp = requests.post(
        TOKEN_URL,
        params={"realm": "/partenaire"},
        data={
            "grant_type":    "client_credentials",
            "client_id":     CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "scope":         "offresetdemandesemploi api_stats-offres-demandes-emploiv1",
        },
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def get_difficulte_persp2_for_rome(headers, code_rome: str):
    url = f"{STATS_BASE_URL}/v1/indicateur/stat-perspective-employeur"

    body = {
        "codeTypeTerritoire": "REG",
        "codeTerritoire": "11",
        "codeTypeActivite": "ROME",
        "codeActivite": code_rome,
        "codeTypePeriode": "ANNEE",
        "codeTypeNomenclature": "TYPE_TENSION",
        "listeCodeNomenclature": ["PERSPECTIVE"],
        "dernierePeriode": True,
    }

    resp = requests.post(url, headers=headers, json=body)
    resp.raise_for_status()
    data = resp.json()

    periodes = data.get("listeValeursParPeriode", [])
    if not periodes:
        return None

    periode = periodes[0]
    code_periode = periode.get("codePeriode")
    lib_periode = periode.get("libPeriode")
    code_rome_retour = periode.get("codeActivite")
    lib_metier = periode.get("libActivite")

    score = periode.get("valeurPrincipaleNombre")
    difficulte_libelle = DIFFICULTE_LABELS.get(score)

    return {
        "code_rome": code_rome_retour or code_rome,
        "libelle_metier": lib_metier,
        "annee": code_periode,
        "lib_periode": lib_periode,
        "score_difficulte": score,
        "difficulte_recrutement": difficulte_libelle,
    }

def get_all_rome_metiers(headers):
    url = f"{STATS_BASE_URL}/v1/referentiel/activites/ROME"
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    data = resp.json()

    activites = data.get("activites", [])
    metiers = []
    for a in activites:
        metiers.append({
            "code_rome": a.get("codeActivite"),
            "libelle_metier": a.get("libelleActivite"),
        })
    return metiers

DATAEMPLOI_METHODO_URL = "https://dataemploi.francetravail.fr/emploi/sources-donnees#difficulte-de-recrutement"

def build_df_for_rome_list(headers, liste_codes_rome, progress_bar=None, progress_text=None, stop_flag=None):
    lignes = []
    codes_sans_tension = []
    
    total = len([c for c in liste_codes_rome if c])
    done = 0

    for code in liste_codes_rome:
        if not code:
            continue
        
        if stop_flag is not None and stop_flag():
            break
        
        try:
            infos = get_difficulte_persp2_for_rome(headers, code)
        except requests.HTTPError as e:
            codes_sans_tension.append((code, f"Erreur HTTP: {e}"))
            done += 1
            if progress_bar:
                progress_bar.progress(done / total)
            if progress_text:
                progress_text.text(f"{done} / {total} codes traités")
            continue
        except Exception as e:
            codes_sans_tension.append((code, f"Erreur: {e}"))
            done += 1
            if progress_bar:
                progress_bar.progress(done / total)
            if progress_text:
                progress_text.text(f"{done} / {total} codes traités")
            continue

        if infos is None:
            codes_sans_tension.append((code, None))
        else:
            lignes.append(infos)

        done += 1
        if progress_bar:
            progress_bar.progress(done / total)
        if progress_text:
            progress_text.text(f"{done} / {total} codes traités")

    if not lignes:
        df = pd.DataFrame()
    else:
        df = pd.DataFrame(lignes)
        df = df[
            [
                "code_rome",
                "libelle_metier",
                "annee",
                "lib_periode",
                "score_difficulte",
                "difficulte_recrutement",
            ]
        ]

    return df, codes_sans_tension


def build_df_all_idf(headers):
    metiers = get_all_rome_metiers(headers)
    codes = [m["code_rome"] for m in metiers if m.get("code_rome")]

    df, codes_sans_tension = build_df_for_rome_list(headers, codes)

    return df


def df_to_excel_bytes(df_raw: pd.DataFrame) -> BytesIO:
    df = format_display_df(df_raw)

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Données")
        ws = writer.sheets["Données"]

        for col_cells in ws.columns:
            max_len = 0
            col_letter = col_cells[0].column_letter
            for cell in col_cells:
                try:
                    val = str(cell.value) if cell.value is not None else ""
                except:
                    val = ""
                max_len = max(max_len, len(val))
            ws.column_dimensions[col_letter].width = max_len + 2

        for cell in ws["C"]:
            if cell.row == 1:
                continue

    buffer.seek(0)
    return buffer

def format_display_df(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.copy()
    df = df[["code_rome", "libelle_metier", "annee", "difficulte_recrutement"]]

    df = df.rename(columns={
        "code_rome": "Code ROME",
        "libelle_metier": "Libellé Métier",
        "annee": "Année",
        "difficulte_recrutement": "Difficulté de recrutement",
    })

    df["Année"] = pd.to_numeric(df["Année"], errors="coerce")

    return df

def main():
    st.title("Difficulté de recrutement (IDF) par métier ROME")

    try:
        with st.spinner("Récupération du token d'authentification..."):
            token = get_token()
    except Exception as e:
        st.error(f"Erreur lors de l'authentification : {e}")
        return

    headers = {
    	"Authorization": f"Bearer {token}",
    	"Content-Type": "application/json",
    	"Accept": "application/json",
    }

    tab_all, tab_list = st.tabs(
        ["Extraction complète IDF", "Recherche par liste de codes ROME"]
    )

    with tab_all:
        st.subheader("Extraction complète de tous les métiers ROME pour l'Île-de-France")

        if "stop_full" not in st.session_state:
            st.session_state.stop_full = False
        if "df_all" not in st.session_state:
            st.session_state.df_all = None

        col1, col2 = st.columns(2)
        with col1:
            start_full = st.button("Lancer l'extraction complète IDF")
        with col2:
            stop_full = st.button("STOP", type="primary")

        if stop_full:
            st.session_state.stop_full = True

        progress_text = st.empty()
        progress_bar = st.progress(0)

        if start_full:
            st.session_state.stop_full = False

            def stop_flag():
                return st.session_state.stop_full

            with st.spinner("Récupération des métiers ROME et des tensions pour l'IDF..."):
                metiers = get_all_rome_metiers(headers)
                codes = [m["code_rome"] for m in metiers if m.get("code_rome")]

                df_all, _ = build_df_for_rome_list(
                    headers,
                    codes,
                    progress_bar=progress_bar,
                    progress_text=progress_text,
                    stop_flag=stop_flag,
                )

            st.session_state.df_all = df_all
            progress_text.text(f"Extraction terminée ou stoppée. {len(df_all)} codes ROME traités.")

        df_all = st.session_state.df_all
        if df_all is not None and not df_all.empty:
            st.success(f"{len(df_all)} codes ROME avec indicateur de tension disponible (résultat actuel).")

            st.dataframe(
                format_display_df(df_all),
                use_container_width=True,
                height=600,
            )

            excel_bytes = df_to_excel_bytes(df_all)
            st.download_button(
                label="Télécharger les résultats actuels (.xlsx)",
                data=excel_bytes,
                file_name="difficulte_recrutement_IDF_complet_ou_partiel.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


    with tab_list:
        st.subheader("Recherche ciblée par liste de codes ROME")

        st.markdown(
            "Colle une liste de codes ROME, **un code par ligne** (ex : A1203, M1801...). "
            "L'outil retournera l'année la plus récente et le libellé de difficulté de recrutement."
        )

        if "df_list" not in st.session_state:
            st.session_state.df_list = None
        if "codes_list" not in st.session_state:
            st.session_state.codes_list = []

        input_codes = st.text_area(
            "Liste de codes ROME (un par ligne)",
            height=250,
            placeholder="A1203\nM1801\nK1303",
        )

        lancer_recherche = st.button("Lancer la recherche ciblée")

        if lancer_recherche:
            liste_codes_rome = [
                line.strip().upper()
                for line in input_codes.splitlines()
                if line.strip()
            ]
            st.session_state.codes_list = liste_codes_rome

            if not liste_codes_rome:
                st.warning("Merci de saisir au moins un code ROME.")
            else:
                progress_text = st.empty()
                progress_bar = st.progress(0)

                with st.spinner("Récupération des difficultés de recrutement pour les codes ROME saisis..."):
                    df_list, codes_sans_tension = build_df_for_rome_list(
                        headers,
                        liste_codes_rome,
                        progress_bar=progress_bar,
                        progress_text=progress_text,
                        stop_flag=None,
                    )

                st.session_state.df_list = df_list

                if codes_sans_tension:
                    justification = (
                        "Pour certains métiers, l'indicateur principal de tension "
                        "n'est pas calculé. Selon la méthodologie de Data Emploi, l'indicateur "
                        "n'est produit que si l'année compte au moins 30 offres déposées, "
                        "30 projets de recrutement et 30 demandeurs d'emploi en catégorie A. "
                        f"[Voir le glossaire complet de Data Emploi]({DATAEMPLOI_METHODO_URL})."
                    )
                    st.markdown("### Codes ROME sans indicateur de tension calculé")
                    st.markdown(justification)

                    codes_txt = ", ".join(code for code, _ in codes_sans_tension)
                    st.markdown(f"Codes sans tension disponible : **{codes_txt}**")

        df_list = st.session_state.df_list

        if df_list is not None:
            if df_list.empty:
                st.warning("Aucune donnée de tension disponible pour les codes ROME saisis.")
            else:
                st.success(f"{len(df_list)} codes ROME avec indicateur de tension disponible.")
                st.dataframe(
                    format_display_df(df_list),
                    use_container_width=True,
                    height=400,
                )

                excel_bytes = df_to_excel_bytes(df_list)
                st.download_button(
                    label="Télécharger les résultats (.xlsx)",
                    data=excel_bytes,
                    file_name="difficulte_recrutement_IDF_selection.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                
        with st.expander("Exemple de liste de codes ROME"):
            st.code("A1203\nM1801\nK1303\nH1102", language="text")

if __name__ == "__main__":
    main()
