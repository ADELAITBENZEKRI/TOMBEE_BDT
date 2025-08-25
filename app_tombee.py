import streamlit as st
import pandas as pd
import plotly.express as px
import calendar
from datetime import datetime
from dateutil.relativedelta import relativedelta
from io import BytesIO
import re

# Configuration de la page
st.set_page_config(page_title="Tableau de Bord d'Analyse des BDT", layout="wide")

# Titre
st.title("Tableau de Bord d'Analyse des Obligations")

# Initialisation de l'état de la session
if 'raw_data' not in st.session_state:
    st.session_state.raw_data = None
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
if 'results' not in st.session_state:
    st.session_state.results = None
if 'step' not in st.session_state:
    st.session_state.step = 0
if 'raw_results' not in st.session_state:
    st.session_state.raw_results = None
if 'instruments_details' not in st.session_state:
    st.session_state.instruments_details = None

# Fonction pour nettoyer les nombres avec espaces
def clean_numeric_value(value):
    """Nettoie les valeurs numériques en supprimant les espaces et en convertissant les virgules en points"""
    if pd.isna(value):
        return value
    
    if isinstance(value, (int, float)):
        return value
    
    # Convertir en chaîne et nettoyer
    value_str = str(value).strip()
    
    # Supprimer tous les espaces (y compris les espaces insécables)
    value_str = value_str.replace(' ', '').replace('\u202f', '').replace('\xa0', '')
    
    # Remplacer les virgules par des points pour les décimales
    value_str = value_str.replace(',', '.')
    
    try:
        # Essayer de convertir en float
        return float(value_str)
    except ValueError:
        # Si la conversion échoue, retourner la valeur originale
        return value

# Fonction de prétraitement
def preprocess_bond_data(df):
    """
    Prétraite les données d'obligations avec les transformations spécifiées
    
    Args:
        df (DataFrame): DataFrame contenant les données brutes
        
    Returns:
        DataFrame: DataFrame prétraité avec les nouvelles colonnes
    """
    
    # Copie du DataFrame pour éviter les modifications sur l'original
    df_processed = df.copy()
    
    # Renommage des colonnes
    rename_dict = {
        'Code ISIN': 'INSTRID',
        'Maturit&eacute;': 'Maturite',
        'Date d\'&eacute;mission': 'ISSUEDT',
        'Date d\'&eacute;ch&eacute;ance': 'MATURITYDT_L',
        'Valeur Nominale': 'PARVALUE',
        'Taux Nominal %': 'INTERESTRATE',
        'Encours': 'ENCOURS'
    }
    
    df_processed = df_processed.rename(columns=rename_dict)
    
    # Nettoyage des valeurs numériques (suppression des espaces)
    numeric_columns = ['PARVALUE', 'ENCOURS', 'INTERESTRATE', 'Prix', 'Coupon Couru Unitaire']
    for col in numeric_columns:
        if col in df_processed.columns:
            df_processed[col] = df_processed[col].apply(clean_numeric_value)
    

    # Ajout de la colonne INTERESTPERIODCTY basée sur la maturité
    def determine_interest_period(maturite):
        if pd.isna(maturite):
            return 'ANLY'
            
        maturite_str = str(maturite).lower()
        
        if 'semaine' in maturite_str:
            if '26' in maturite_str or '52' in maturite_str:
                return 'HFLY'  # Semi-annuel
            elif '13' in maturite_str:
                return 'QTLY'  # Trimestriel
        elif any(keyword in maturite_str for keyword in ['an', 'ans', 'année', 'années']):
            return 'ANLY'  # Annuel
        
        return 'ANLY'  # Par défaut
    
    if 'Maturite' in df_processed.columns:
        df_processed['INTERESTPERIODCTY'] = df_processed['Maturite'].apply(determine_interest_period)
    
    return df_processed

# Fonctions utilitaires
def number_to_text(value):
    value = abs(value)
    if value >= 1_000_000_000:
        return f"{value / 1_000_000_000:,.2f} milliards"
    elif value >= 1_000_000:
        return f"{value / 1_000_000:,.2f} millions"
    else:
        return f"{value:,.2f}"

def format_amount(value):
    return f"{value:,.2f} ({number_to_text(value)})"

def calculate_coupon_dates(row):
    try:
        issue_date = pd.to_datetime(row["ISSUEDT"], errors='coerce')
        maturity_date = pd.to_datetime(row["MATURITYDT_L"], errors='coerce')
        
        if pd.isna(issue_date) or pd.isna(maturity_date):
            return [maturity_date]
        
        frequency = row["INTERESTPERIODCTY"]
        coupon_dates = []

        if frequency == "ANLY":
            first_coupon = maturity_date.replace(year=issue_date.year + 1)
            
            if first_coupon < maturity_date:
                current_date = first_coupon
                previous_year = None
                
                while current_date <= maturity_date:
                    current_year = current_date.year
                    if current_year != previous_year:
                        coupon_dates.append(current_date)
                        previous_year = current_year
                    current_date += relativedelta(years=1)
            
            if not coupon_dates or coupon_dates[-1].year != maturity_date.year:
                coupon_dates.append(maturity_date)
        else:
            coupon_dates.append(maturity_date)

        return coupon_dates
    except Exception as e:
        st.error(f"Erreur pour l'instrument {row.get('INSTRID', 'inconnu')}: {str(e)}")
        return [maturity_date]

# Interface utilisateur
st.sidebar.header("Contrôles")
uploaded_file = st.sidebar.file_uploader("Télécharger un fichier Excel", type=["xlsx"])

if st.sidebar.button("1. Charger et prétraiter les données") and uploaded_file is not None:
    try:
        # Chargement des données
        st.session_state.raw_data = pd.read_excel(uploaded_file)
        
        # Afficher les colonnes originales pour debug
        st.sidebar.info(f"Colonnes originales: {list(st.session_state.raw_data.columns)}")
        
        # Prétraitement des données
        st.session_state.processed_data = preprocess_bond_data(st.session_state.raw_data)
        
        # Afficher les colonnes après prétraitement pour debug
        st.sidebar.info(f"Colonnes après prétraitement: {list(st.session_state.processed_data.columns)}")
        
        # Vérification des colonnes requises
        required_cols = ['INSTRID', 'ISSUEDT', 'MATURITYDT_L', 'INTERESTPERIODCTY', 'ISSUESIZE', 'INTERESTRATE']
        missing = [col for col in required_cols if col not in st.session_state.processed_data.columns]
        
        if missing:
            st.sidebar.error(f"Colonnes manquantes après prétraitement: {', '.join(missing)}")
            # Afficher les données pour debug
            st.subheader("Données brutes (pour debug)")
            st.dataframe(st.session_state.raw_data.head(), use_container_width=True)
            
            st.subheader("Données prétraitées (pour debug)")
            st.dataframe(st.session_state.processed_data.head(), use_container_width=True)
        else:
            # Suppression des doublons basée sur INSTRID (garder la première occurrence)
            st.session_state.processed_data = st.session_state.processed_data.drop_duplicates(subset=['INSTRID'], keep='first')
            
            # Conversion des types de données
            st.session_state.processed_data['MATURITYDT_L'] = pd.to_datetime(st.session_state.processed_data['MATURITYDT_L'], errors='coerce')
            st.session_state.processed_data['ISSUEDT'] = pd.to_datetime(st.session_state.processed_data['ISSUEDT'], errors='coerce')
            st.session_state.processed_data['ISSUESIZE'] = pd.to_numeric(st.session_state.processed_data['ISSUESIZE'], errors='coerce') * 100_000
            
            st.session_state.step = 2
            st.sidebar.success("Chargement et prétraitement réussis! Doublons supprimés.")
            
            # Aperçu des données prétraitées
            st.subheader("Aperçu des données prétraitées")
            st.dataframe(st.session_state.processed_data.head(), use_container_width=True)
            
    except Exception as e:
        st.sidebar.error(f"Erreur: {str(e)}")
        import traceback
        st.error(f"Traceback: {traceback.format_exc()}")

if st.sidebar.button("2. Calculer les coupons") and st.session_state.step >= 2:
    try:
        df = st.session_state.processed_data.copy()
        df["CouponPayDate"] = df.apply(calculate_coupon_dates, axis=1)
        df["AnnualCouponAmount"] = df["ISSUESIZE"] * df["INTERESTRATE"] / 100
        
        def calculate_coupon_amount(row, coupon_date):
            if pd.isna(coupon_date):
                return 0
            freq = row["INTERESTPERIODCTY"]
            if freq == "ANLY":
                return row["AnnualCouponAmount"]
            elif freq == "HFLY":
                return row["AnnualCouponAmount"] # / 2
            elif freq == "QTLY":
                return row["AnnualCouponAmount"] # / 4
            else:
                return row["AnnualCouponAmount"]
        
        max_coupons = max(df["CouponPayDate"].apply(len)) if not df["CouponPayDate"].empty else 0
        for i in range(max_coupons):
            df[f"CouponPayDate_{i+1}"] = df["CouponPayDate"].apply(lambda x: x[i] if i < len(x) else pd.NaT)
            df[f"CouponAmount_{i+1}"] = df.apply(lambda row: calculate_coupon_amount(row, row[f"CouponPayDate_{i+1}"]), axis=1)
        
        date_cols = [col for col in df.columns if "CouponPayDate" in col]
        for col in date_cols:
            df[col] = df[col].dt.strftime('%d-%m-%Y') if df[col].dtype == 'datetime64[ns]' else df[col]
        
        st.session_state.processed_data = df
        st.session_state.step = 3
        st.sidebar.success("Calcul des coupons terminé!")
    except Exception as e:
        st.sidebar.error(f"Erreur: {str(e)}")

if st.sidebar.button("3. Analyser les résultats") and st.session_state.step >= 3:
    try:
        df = st.session_state.processed_data.copy()
        
        def get_coupons_by_month_year(row):
            coupons = {}
            for i in range(1, 32):
                date_col = f'CouponPayDate_{i}'
                amount_col = f'CouponAmount_{i}'
                if date_col in df.columns and pd.notna(row[date_col]):
                    try:
                        date = pd.to_datetime(row[date_col], dayfirst=True)
                        month_year = (date.month, date.year)
                        amount = float(row[amount_col]) if pd.notna(row[amount_col]) else 0
                        coupons[month_year] = coupons.get(month_year, 0) + amount
                    except:
                        continue
            return coupons
        
        results = {}
        instruments_details = {}
        
        for _, row in df.iterrows():
            maturity_date = row['MATURITYDT_L']
            if pd.notna(maturity_date):
                month_year = (maturity_date.month, maturity_date.year)
                issue_size = row['ISSUESIZE'] if pd.notna(row['ISSUESIZE']) else 0
                
                if month_year not in results:
                    results[month_year] = {
                        'total_issuesize': 0,
                        'total_coupons': 0,
                        'instruments': set(),
                        'coupon_instruments': set()
                    }
                    instruments_details[month_year] = {
                        'maturity_instruments': [],
                        'coupon_instruments': []
                    }
                
                results[month_year]['total_issuesize'] += issue_size
                results[month_year]['instruments'].add(row['INSTRID'])
                instruments_details[month_year]['maturity_instruments'].append({
                    'INSTRID': row['INSTRID'],
                    'ISSUESIZE': issue_size,
                    'MATURITYDT': maturity_date.strftime('%d-%m-%Y')
                })
            
            coupons = get_coupons_by_month_year(row)
            for month_year, amount in coupons.items():
                if month_year not in results:
                    results[month_year] = {
                        'total_issuesize': 0,
                        'total_coupons': 0,
                        'instruments': set(),
                        'coupon_instruments': set()
                    }
                    instruments_details[month_year] = {
                        'maturity_instruments': [],
                        'coupon_instruments': []
                    }
                
                results[month_year]['total_coupons'] += amount
                results[month_year]['coupon_instruments'].add(row['INSTRID'])

                coupon_date = None
                for i in range(1, 32):
                    date_col = f'CouponPayDate_{i}'
                    if date_col in row and pd.notna(row[date_col]):
                        try:
                            date = pd.to_datetime(row[date_col], dayfirst=True)
                            if date.year == month_year[1] and date.month == month_year[0]:
                                coupon_date = row[date_col]
                                break
                        except:
                            continue

                instruments_details[month_year]['coupon_instruments'].append({
                    'INSTRID': row['INSTRID'],
                    'CouponAmount': amount,
                    'CouponDate': coupon_date
                })
        
        for month_year in results:
            results[month_year]['instruments'] = list(results[month_year]['instruments'])
            results[month_year]['coupon_instruments'] = list(results[month_year]['coupon_instruments'])
        
        sorted_results = sorted(results.items(), key=lambda x: (x[0][1], x[0][0]))
        filtered_results = [(m_y, data) for m_y, data in sorted_results 
                          if (m_y[1] > 2025) or (m_y[1] == 2025 and m_y[0] >= 1)]
        
        st.session_state.results = filtered_results
        st.session_state.raw_results = results
        st.session_state.instruments_details = instruments_details
        st.session_state.step = 4
        st.sidebar.success("Analyse terminée!")
    except Exception as e:
        st.sidebar.error(f"Erreur: {str(e)}")

# Affichage de l'état
st.sidebar.header("État du processus")
steps = [
    "🟠 En attente de données",
    "🔴 Prétraitement nécessaire",
    "✅ Données prétraitées | 🔴 Calcul des coupons nécessaire",
    "✅ Données prétraitées | ✅ Coupons calculés | 🔴 Analyse nécessaire",
    "✅ Données prétraitées | ✅ Coupons calculés | ✅ Analyse terminée"
]
st.sidebar.info(steps[st.session_state.step])

# Recherche d'instrument par INSTRID
if st.session_state.step >= 2:
    st.header("🔍 Recherche d'instrument")
    
    search_instr = st.text_input("Entrez l'INSTRID de l'instrument à rechercher:")
    
    if search_instr and st.session_state.processed_data is not None:
        instrument_data = st.session_state.processed_data[
            st.session_state.processed_data['INSTRID'].astype(str).str.contains(search_instr, case=False)
        ]
        
        if not instrument_data.empty:
            st.subheader(f"Résultats pour: {search_instr}")
            st.dataframe(instrument_data, use_container_width=True)
            
            # Afficher les dates de coupon spécifiques si disponibles
            if st.session_state.step >= 3:
                coupon_cols = [col for col in instrument_data.columns if "CouponPayDate" in col or "CouponAmount" in col]
                if coupon_cols:
                    coupon_data = instrument_data[coupon_cols].transpose().reset_index()
                    coupon_data.columns = ['Colonne', 'Valeur']
                    st.subheader("Détails des coupons")
                    st.dataframe(coupon_data, use_container_width=True)
        else:
            st.warning(f"Aucun instrument trouvé avec INSTRID contenant '{search_instr}'")

# Nouvelle section d'onglets
if st.session_state.step >= 4:
    tab1, tab2 = st.tabs(["📊 Vue d'ensemble", "📅 Détails par mois"])
    
    with tab1:
        # Création du DataFrame de résultats
        results_df = []
        total_issuesize = 0
        total_coupons = 0
        
        for (month, year), data in st.session_state.results:
            month_name = f"{calendar.month_name[month]} {year}"
            total_issuesize += data['total_issuesize']
            total_coupons += data['total_coupons']
            
            results_df.append({
                "Mois/Année": month_name,
                "Taille Émission": data['total_issuesize'],
                "Coupons": data['total_coupons'],
                "Nb Instruments": len(data['instruments']),
                "Total": data['total_issuesize'] + data['total_coupons']
            })
        
        results_df = pd.DataFrame(results_df)
        
        # Filtres
        st.sidebar.header("Filtres")
        years = sorted({y for (_, y), _ in st.session_state.results})
        selected_year = st.sidebar.selectbox("Année", years)
        
        months_in_year = sorted({m for (m, y), _ in st.session_state.results if y == selected_year})
        month_names = [calendar.month_name[m] for m in months_in_year]
        selected_month = st.sidebar.selectbox("Mois", month_names)
        
        # Visualisations
        st.header("Visualisations")
        selected_years = st.multiselect("Années à afficher", years, default=[selected_year])
        
        if selected_years:
            filtered_df = results_df[results_df['Mois/Année'].str.contains('|'.join(map(str, selected_years)))]
            
            if not filtered_df.empty:
                col1, col2 = st.columns(2)
                
                with col1:
                    fig = px.bar(
                        filtered_df, 
                        x="Mois/Année", 
                        y=["Taille Émission", "Coupons"],
                        title="Taille d'émission et coupons",
                        barmode="group"
                    )
                    st.plotly_chart(fig, use_container_width=True)
                
                with col2:
                    fig = px.line(
                        filtered_df,
                        x="Mois/Année",
                        y="Total",
                        title="Total (Taille + Coupons)"
                    )
                    st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("Aucune donnée à afficher pour les années sélectionnées.")
        
        # Tableau complet
        st.header("Tous les résultats")
        
        # Préparation des données pour l'affichage
        display_data = []
        
        for (month, year), data in st.session_state.results:
            month_name = f"{calendar.month_name[month]} {year}"
            
            # Récupérer les détails des instruments
            details = st.session_state.instruments_details.get((month, year), {})
            
            # Instruments à maturité
            maturity_instruments = details.get('maturity_instruments', [])
            maturity_ids = "\n".join(instr['INSTRID'] for instr in maturity_instruments) if maturity_instruments else ""
            
            # Instruments avec coupons
            coupon_instruments = details.get('coupon_instruments', [])
            coupon_ids = "\n".join([instr['INSTRID'] for instr in coupon_instruments]) if coupon_instruments else ""
            
            display_data.append({
                "Mois/Année": month_name,
                "Total Taille Émission": data['total_issuesize'],
                "Total Coupons": data['total_coupons'],
                "Nb Instruments Maturité": len(maturity_instruments),
                "Instruments Maturité": maturity_ids,
                "Nb Instruments Coupons": len(coupon_instruments),
                "Instruments Coupons": coupon_ids,
                "Total (Taille Émission + Coupons)": data['total_issuesize'] + data['total_coupons']
            })
        
        # Création du DataFrame
        display_df = pd.DataFrame(display_data)
        
        # Formatage des montants
        display_df["Total Taille Émission"] = display_df["Total Taille Émission"].apply(format_amount)
        display_df["Total Coupons"] = display_df["Total Coupons"].apply(format_amount)
        display_df["Total (Taille Émission + Coupons)"] = display_df["Total (Taille Émission + Coupons)"].apply(format_amount)
        
        # Affichage du tableau avec mise en forme
        st.dataframe(
            display_df,
            use_container_width=True,
            column_config={
                "Instruments Maturité": st.column_config.TextColumn(
                    "Instruments Maturité",
                    help="Liste des instruments arrivant à échéance ce mois",
                    width="medium"
                ),
                "Instruments Coupons": st.column_config.TextColumn(
                    "Instruments Coupons",
                    help="Liste des instruments payant des coupons ce mois",
                    width="medium"
                )
            },
            hide_index=True
        )
    with tab2:
        st.header("Détails des instruments par mois")
        
        if hasattr(st.session_state, 'instruments_details'):
            # Sélection du mois
            months = sorted(st.session_state.instruments_details.keys(), key=lambda x: (x[1], x[0]))
            month_options = [f"{calendar.month_name[m]} {y}" for m, y in months]
            
            default_index = next((i for i, m in enumerate(month_options) if m == "January 2025"), 0)
            
            selected_month_str = st.selectbox(
                "Sélectionnez un mois", 
                month_options,
                index=default_index
            )
            
            selected_index = month_options.index(selected_month_str)
            selected_month = months[selected_index]
            details = st.session_state.instruments_details.get(selected_month, {})
      
            # Calcul des sommes totales
            total_coupons = sum(instr['CouponAmount'] for instr in details.get('coupon_instruments', []))
            total_maturity = sum(instr['ISSUESIZE'] for instr in details.get('maturity_instruments', []))
            total_flux = total_coupons + total_maturity
            
            # Affichage des totaux avec style
            st.markdown("---")
            col_sum1, col_sum2, col_sum3 = st.columns(3)
            with col_sum1:
                st.metric("Total coupons versés", format_amount(total_coupons), help="Somme des coupons payés ce mois")
            with col_sum2:
                st.metric("Total capitaux à échéance", format_amount(total_maturity), help="Somme des capitaux arrivant à échéance ce mois")
            with col_sum3:
                st.metric("Tombees totales du mois", 
                          format_amount(total_flux), 
                          help="Somme des flux financiers (coupons + capitaux)",
                          delta_color="off")
            
            st.markdown("""
            <style>
                div[data-testid="stMetric"]:nth-child(3) {
                    border: 1px solid #ff4b4b;
                    border-radius: 5px;
                    background-color: #fff0f0;
                    padding: 5px;
                }
                div[data-testid="stMetric"]:nth-child(3) > div > label {
                    color: #ff4b4b !important;
                    font-weight: bold !important;
                }
                div[data-testid="stMetric"]:nth-child(3) > div > div {
                    color: #ff4b4b !important;
                    font-weight: bold !important;
                    font-size: 1.3rem !important;
                }
            </style>
            """, unsafe_allow_html=True)
            
            # Graphique combiné des flux
            if total_flux > 0:
                flux_data = pd.DataFrame({
                    'Type': ['Coupons', 'Capitaux', 'Total'],
                    'Montant': [total_coupons, total_maturity, total_flux],
                    'Couleur': ['#1f77b4', '#ff7f0e', '#ff4b4b']
                })
                
                fig_flux = px.bar(flux_data, 
                                x='Type', 
                                y='Montant',
                                color='Couleur',
                                title=f"Flux financiers - {selected_month_str}",
                                labels={'Montant': 'Montant (€)', 'Type': ''},
                                text=[format_amount(x) for x in flux_data['Montant']])
                
                fig_flux.update_traces(textposition='outside',
                                    marker_color=flux_data['Couleur'],
                                    showlegend=False)
                fig_flux.update_layout(yaxis={'visible': False, 'showticklabels': False})
                
                st.plotly_chart(fig_flux, use_container_width=True)
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader(f"Instruments à échéance - {selected_month_str}")
                if details.get('maturity_instruments'):
                    maturity_df = pd.DataFrame(details['maturity_instruments'])
                    maturity_df['ISSUESIZE'] = maturity_df['ISSUESIZE'].apply(format_amount)
                    st.dataframe(maturity_df, hide_index=True, use_container_width=True)
                    
                    fig1 = px.bar(maturity_df, 
                                  x='INSTRID', 
                                  y='ISSUESIZE',
                                  title=f"Capital à échéance - {selected_month_str}",
                                  labels={'ISSUESIZE': 'Montant', 'INSTRID': 'Instrument'})
                    st.plotly_chart(fig1, use_container_width=True)
                else:
                    st.info("Aucun instrument arrivant à échéance ce mois-ci")
            
            with col2:
                st.subheader(f"Instruments avec coupons - {selected_month_str}")
                if details.get('coupon_instruments'):
                    coupon_df = pd.DataFrame(details['coupon_instruments'])
                    coupon_df['CouponAmount'] = coupon_df['CouponAmount'].apply(format_amount)
                    st.dataframe(coupon_df, hide_index=True, use_container_width=True)
                    
                    fig2 = px.bar(coupon_df, 
                                x='INSTRID', 
                                y='CouponAmount',
                                title=f"Coupons versés - {selected_month_str}",
                                labels={'CouponAmount': 'Montant', 'INSTRID': 'Instrument'})
                    st.plotly_chart(fig2, use_container_width=True)
                else:
                    st.info("Aucun coupon versé ce mois-ci")
        else:
            st.warning("Aucun détail d'instrument disponible. Veuillez relancer l'analyse.")

# Téléchargement des données
if st.session_state.step >= 2:
    st.header("📥 Téléchargement des données")
    
    @st.cache_data
    def convert_to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        return output.getvalue()
    
    if st.session_state.step >= 4:
        # Téléchargement des données avec coupons
        excel_data = convert_to_excel(st.session_state.processed_data)
        st.download_button(
            label="Télécharger les données avec coupons",
            data=excel_data,
            file_name="donnees_avec_coupons.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Téléchargement des résultats par année
        st.subheader("Télécharger les résultats par année")
        
        available_years = sorted({y for (_, y) in st.session_state.raw_results.keys()})
        selected_dl_year = st.selectbox("Sélectionnez une année à télécharger:", available_years)
        
        if st.button(f"Générer le rapport pour {selected_dl_year}"):
            # Création d'un DataFrame pour l'année sélectionnée
            year_data = []
            
            for (month, year), data in st.session_state.results:
                if year == selected_dl_year:
                    month_name = calendar.month_name[month]
                    
                    # Instruments à échéance
                    maturity_instr = st.session_state.instruments_details.get((month, year), {}).get('maturity_instruments', [])
                    for instr in maturity_instr:
                        year_data.append({
                            'Date': f"01-{month:02d}-{year}",
                            'Type': 'Maturité',
                            'INSTRID': instr['INSTRID'],
                            'Montant': instr['ISSUESIZE'],
                            'Date Flux': instr['MATURITYDT']
                        })
                    
                    # Instruments avec coupons
                    coupon_instr = st.session_state.instruments_details.get((month, year), {}).get('coupon_instruments', [])
                    for instr in coupon_instr:
                        year_data.append({
                            'Date': f"01-{month:02d}-{year}",
                            'Type': 'Coupon',
                            'INSTRID': instr['INSTRID'],
                            'Montant': instr['CouponAmount'],
                            'Date Flux': instr['CouponDate']
                        })
            
            year_df = pd.DataFrame(year_data)
            
            # Création du fichier Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                year_df.to_excel(writer, sheet_name=f"Flux {selected_dl_year}", index=False)
                
                # Ajout des feuilles supplémentaires
                if selected_dl_year in [2025, 2026]:
                    summary_df = pd.DataFrame([
                        {'Année': selected_dl_year, 
                         'Total Maturités': sum(x['ISSUESIZE'] for x in st.session_state.instruments_details.get((month, selected_dl_year), {}).get('maturity_instruments', [])),
                         'Total Coupons': sum(x['CouponAmount'] for x in st.session_state.instruments_details.get((month, selected_dl_year), {}).get('coupon_instruments', []))
                        } for month in range(1, 13) if (month, selected_dl_year) in st.session_state.instruments_details
                    ])
                    summary_df.to_excel(writer, sheet_name="Résumé", index=False)
                
                st.session_state.processed_data.to_excel(writer, sheet_name="Données complètes", index=False)
            
            st.download_button(
                label=f"Télécharger le rapport {selected_dl_year}",
                data=output.getvalue(),
                file_name=f"rapport_flux_{selected_dl_year}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        # Bouton pour télécharger un rapport complet
        if st.button("Télécharger un rapport complet (2025-2026)"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Feuille 1: Résultats finaux 2025-2026
                final_results = []
                for (month, year), data in st.session_state.results:
                    if year in [2025, 2026]:
                        month_name = calendar.month_name[month]
                        final_results.append({
                            'Année': year,
                            'Mois': month_name,
                            'Total Maturités': data['total_issuesize'],
                            'Total Coupons': data['total_coupons'],
                            'Nb Instruments': len(data['instruments'])
                        })
                pd.DataFrame(final_results).to_excel(writer, sheet_name="Résultats 2025-2026", index=False)
                
                # Feuille 2: Données traitées
                st.session_state.processed_data.to_excel(writer, sheet_name="Données traitées", index=False)
                
                # Feuille 3: Résultats après 2026
                post_2026 = []
                for (month, year), data in st.session_state.results:
                    if year > 2026:
                        month_name = calendar.month_name[month]
                        post_2026.append({
                            'Année': year,
                            'Mois': month_name,
                            'Total Maturités': data['total_issuesize'],
                            'Total Coupons': data['total_coupons'],
                            'Nb Instruments': len(data['instruments'])
                        })
                pd.DataFrame(post_2026).to_excel(writer, sheet_name="Résultats après 2026", index=False)
            
            st.download_button(
                label="Télécharger le rapport complet",
                data=output.getvalue(),
                file_name="rapport_complet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# Message initial
if st.session_state.step == 0:
    st.info("Veuillez télécharger un fichier Excel et suivre les étapes du processus.")


