import pandas as pd
import os
import shutil
import re

# ================= CONFIGURATION DOSSIERS =================

BASE_RESEAU = r"\\192.109.69.46\GenerationGBA"

DOSSIER_ENTREE = os.path.join(BASE_RESEAU, "brutes")
DOSSIER_ARCHIVE = os.path.join(BASE_RESEAU, "Archives")
DOSSIER_EXPORT = os.path.join(BASE_RESEAU, "Export")
DOSSIER_PARAM = os.path.join(BASE_RESEAU, "parametres")

FICHIER_CLIENTS = os.path.join(DOSSIER_PARAM, "clients_419811.xlsx")

os.makedirs(DOSSIER_ARCHIVE, exist_ok=True)
os.makedirs(DOSSIER_EXPORT, exist_ok=True)

# ================= PARAMETRES COMPTABLES =================

JOURNAL = "C52"
TYPE_PIECE = "OD"
COMPTE_CONTREPARTIE = "41981"
REFERENCE_PAR_DEFAUT = "BA0086"
REFERENCE_ACTIF = "BAC"


# ================= CHARGEMENT CLIENTS =================

if not os.path.exists(FICHIER_CLIENTS):
    print("❌ Fichier clients introuvable")
    exit()

df_clients = pd.read_excel(FICHIER_CLIENTS)

df_clients["Des_Cli"] = df_clients["Des_Cli"].astype(str).str.strip()
df_clients["Reference"] = df_clients["Reference"].astype(str).str.strip()

ref_map_419811 = dict(zip(df_clients["Des_Cli"], df_clients["Reference"]))

print(f"✅ {len(ref_map_419811)} clients chargés")

# ================= RECHERCHE FICHIER =================

fichiers = []

for f in os.listdir(DOSSIER_ENTREE):

    # Filtrer extensions
    if not f.lower().endswith((".xlsx", ".xls", ".xlsb")):
        continue

    nom_normalise = " ".join(f.upper().split())

    if "BASE INFORMATIQUE" in nom_normalise and "DU" in nom_normalise:
        fichiers.append(os.path.join(DOSSIER_ENTREE, f))

if not fichiers:
    print("❌ Aucun fichier trouvé")
    exit()

FICHIER_SOURCE = fichiers[0]
nom_fichier = os.path.basename(FICHIER_SOURCE)

print(f"📂 Fichier détecté : {nom_fichier}")

# ================= EXTRACTION DATE ROBUSTE =================

def extraire_date_intervalle(nom_fichier):
    nom = nom_fichier.upper()

    # Supprimer extension
    nom = os.path.splitext(nom)[0]

    # Nettoyer espaces
    nom = " ".join(nom.split())

    # Récupérer après DU
    match = re.search(r"DU\s+(.+)", nom)
    if not match:
        return None

    bloc = match.group(1)

    # Nettoyage
    bloc = bloc.replace(" ", "")
    bloc = bloc.replace("AU", "_")

    return bloc


date_intervalle = extraire_date_intervalle(nom_fichier)

if not date_intervalle:
    print("❌ Erreur extraction date")
    exit()

print(f"📅 Intervalle : {date_intervalle}")

FICHIER_EXPORT = os.path.join(
    DOSSIER_EXPORT,
    f"Nouvelles_ecritures_GBA_{date_intervalle}.txt"
)

# ================= LECTURE EXCEL =================

if FICHIER_SOURCE.lower().endswith(".xlsb"):
    df_temp = pd.read_excel(FICHIER_SOURCE, engine="pyxlsb", header=None)
else:
    df_temp = pd.read_excel(FICHIER_SOURCE, header=None)

header_row = None

for i in range(15):
    ligne = df_temp.iloc[i].astype(str).str.lower().str.strip()
    if "date" in ligne.values and "des_cli" in ligne.values:
        header_row = i
        break

if header_row is None:
    print("❌ Header introuvable")
    exit()

if FICHIER_SOURCE.lower().endswith(".xlsb"):
    df = pd.read_excel(FICHIER_SOURCE, engine="pyxlsb", header=header_row)
else:
    df = pd.read_excel(FICHIER_SOURCE, header=header_row)

df.columns = df.columns.str.strip()

print("✅ Colonnes :", list(df.columns))

# ================= PREPARATION =================

def convertir_date(val):
    try:
        if isinstance(val, (int, float)):
            return pd.to_datetime("1899-12-30") + pd.to_timedelta(val, unit="D")
        return pd.to_datetime(val, dayfirst=True, errors="coerce")
    except:
        return pd.NaT

df["date"] = df["date"].apply(convertir_date)
df["Des_Cli"] = df["Des_Cli"].astype(str).str.strip()
df["Compte"] = df["Compte"].apply(lambda x: str(int(float(x))) if pd.notna(x) else "")
df["Mon_Bon"] = pd.to_numeric(df["Mon_Bon"], errors="coerce").fillna(0)

# ================= REGROUPEMENT =================

df_grouped = (
    df
    .groupby(["date", "Compte", "Des_Cli", "Cod_Sit"], as_index=False)
    .agg({"Mon_Bon": "sum"})
)

# ================= GENERATION ECRITURES =================

ecritures = []

for _, row in df_grouped.iterrows():

    date_op = row["date"]
    compte_debit = row["Compte"]
    des_cli = row["Des_Cli"]
    montant = row["Mon_Bon"]
    magasin = int(row["Cod_Sit"])

    date_str = date_op.strftime("%d/%m/%Y")
    MOIS_FR = {
        "jan": "jan",
        "feb": "fév",
        "mar": "mar",
        "apr": "avr",
        "may": "mai",
        "jun": "jun",
        "jul": "jul",
        "aug": "aoû",
        "sep": "sep",
        "oct": "oct",
        "nov": "nov",
        "dec": "déc"
    }

    mois_en = date_op.strftime("%b").lower()
    mois_fr = MOIS_FR.get(mois_en, mois_en)

    date_ech_str = f"{date_op.strftime('%d')}-{mois_fr}-{date_op.strftime('%y')}"
    libelle = f"Bon d'achat {des_cli} site {magasin} {date_str}"

    # ================= LOGIQUE REFERENCE =================

    if compte_debit == "419811":

        if des_cli == "ATCDCI":
            ref_debit = REFERENCE_PAR_DEFAUT

        else:
            ref_debit = ref_map_419811.get(des_cli, "")

    elif compte_debit == "419812":
        ref_debit = REFERENCE_PAR_DEFAUT

    else:
        ref_debit = ""

    lettrage_debit = "X" if ref_debit else ""

    # -------- DEBIT --------
    ecritures.append({
        "Journal": JOURNAL,
        "Date": date_str,
        "Type": TYPE_PIECE,
        "Compte": compte_debit,
        "Sens": "D",
        "Lettrage": lettrage_debit,
        "Montant": montant,
        "Reference": ref_debit,
        "Date_ech": date_ech_str,
        "Libelle": libelle
    })

    # -------- CREDIT --------
    ref_credit = REFERENCE_ACTIF + str(magasin) if COMPTE_CONTREPARTIE == "41981" else ""
    lettrage_credit = "X" if ref_credit else ""

    ecritures.append({
        "Journal": JOURNAL,
        "Date": date_str,
        "Type": TYPE_PIECE,
        "Compte": COMPTE_CONTREPARTIE,
        "Sens": "C",
        "Lettrage": lettrage_credit,
        "Montant": montant,
        "Reference": ref_credit,
        "Date_ech": date_ech_str,
        "Libelle": libelle
    })

# ================= EXPORT =================

df_od = pd.DataFrame(ecritures)

df_od.to_csv(
    FICHIER_EXPORT,
    sep=";",
    index=False,
    header=False,
    encoding="utf-8-sig"
)

print("✅ Export terminé")

# ================= ARCHIVAGE =================

chemin_archive = os.path.join(DOSSIER_ARCHIVE, nom_fichier)
shutil.move(FICHIER_SOURCE, chemin_archive)

print(f"📦 Archivé : {chemin_archive}")
print("🚀 ===== TRAITEMENT TERMINE =====")
input("Appuyez sur Entrée pour fermer...")