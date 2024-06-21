import pandas as pd # pip install pandas openpyxl
import plotly.express as px # pip install plotly-express
import streamlit as st # pip install streamlit
import folium 
from streamlit_folium import st_folium # pip install streamlit-folium


st.set_page_config(page_title="Sales Dashboard", 
                   page_icon=":bar_chart:",
                   layout="wide")

@st.cache_data # mettre la donnée en cache pour éviter d'aller la copier à chaque fois dans le fichier excel
def get_data_from_excel1():
    df = pd.read_excel(
        io= "DATA.xlsx",
        engine= "openpyxl",
        sheet_name= "Base_avec_formules",
        usecols="A:M",
        nrows=1000
    )
    return df

@st.cache_data
def get_data_from_excel2():
    df2 = pd.read_excel(
        io= "DATA.xlsx",
        engine= "openpyxl",
        sheet_name= "Objectifs",
        usecols="A:G",
        nrows=10
    )
    return df2

df = get_data_from_excel1()
df2 = get_data_from_excel2()
df_merged = df.merge(df2, on="Produits", how="left") # faire un join des 2 sheets

villes = sorted(df_merged["Villes"].unique()) 
produits = sorted(df_merged["Produits"].unique())

st.sidebar.header("Filtrer les données ici:") #sidebar
produit = st.sidebar.multiselect(
    "Choississez le produit:",
    options = produits,
    default = produits
)

ville = st.sidebar.multiselect(
    "Choississez la Ville:",
    options = villes,
    default = villes
)


df_selection = df_merged.query(
    "Villes == @ville & Produits == @produit"
)

# Verifier si le dataframe est vide
if df_selection.empty:
    st.warning("Aucune donnée disponible sur la base des paramètres de filtre actuels !")
    st.stop() 

st.title(":bar_chart: Sales Dashboard")
st.markdown("##")

#KPIs
chiffre_daffaire_total = int(df_selection["Chiffre_d'affaires"].sum())
cout_total_de_production = int(df_selection["Cout_total"].sum())
marge_brute = chiffre_daffaire_total - cout_total_de_production

left_column, middle_column, right_column = st.columns(3)
with left_column:
    st.subheader("Chiffre d'affaires:")
    st.subheader(f"EURO € {chiffre_daffaire_total:,}")

with middle_column:
    st.subheader("Coût production:")
    st.subheader(f"EURO € {cout_total_de_production}")

with right_column:
    st.subheader("Marge brute:")
    st.subheader(f"EURO € {marge_brute}")


st.markdown("""---""")

#CHARTS
Quantité_vendue_par_produit = df_selection.groupby(by=["Produits"])[["Quantité_vendue"]].sum().sort_values(by="Quantité_vendue")
fig_quantité_vendue_par_produit = px.bar(
    Quantité_vendue_par_produit,
    x="Quantité_vendue",
    y=Quantité_vendue_par_produit.index,
    orientation="h",
    title="<b>Quantité vendue par produit</b>",
    color_discrete_sequence=["#0083B8"] * len(Quantité_vendue_par_produit),
    template="plotly_white",
)
fig_quantité_vendue_par_produit.update_layout(
    plot_bgcolor="rgba(0,0,0,0)",
    xaxis=(dict(showgrid=False))
)

fig_comparaison_CA = px.bar(
    df_selection, 
    x="Produits", 
    y=["CA produits", "CA objectifs"], 
    barmode="group",
    labels={"value": "Chiffre d'Affaires", "variable": "Type"},
    title="<b>CA produits vs CA objectifs</b>",
)
fig_comparaison_CA .update_layout(
    xaxis=dict(tickmode="linear"),
    plot_bgcolor="rgba(0,0,0,0)",
    yaxis=(dict(showgrid=False)),
)

fig3_pourcentage_produit = px.bar(
    df_selection, 
    x="Produits", 
    y="% actuel", 
    title="<b>% Objectifs atteints par produit</b>",
    labels={"% actuel": "% Objectifs atteints"}
)
fig3_pourcentage_produit.update_layout(
    xaxis=dict(tickmode="linear"),
    plot_bgcolor="rgba(0,0,0,0)",
    yaxis=(dict(showgrid=False)),
)

revenu_brut_par_produit = df_selection.groupby(by=["Produits"])[["Revenu_brut_pa_produit"]].sum() #GroupBy
fig_col_revenu_brut = px.bar(
    revenu_brut_par_produit,
    x=revenu_brut_par_produit.index,
    y="Revenu_brut_pa_produit",
    title="<b>Revenu brut par produit</b>",
    color_discrete_sequence=["#0083B8"] * len(revenu_brut_par_produit),
    template="plotly_dark"
)
fig_col_revenu_brut.update_xaxes(title="Produits")
fig_col_revenu_brut.update_yaxes(title="Revenu brut")
fig_col_revenu_brut.update_layout(
    xaxis=dict(tickmode="linear"),
    plot_bgcolor="rgba(0,0,0,0)",
    yaxis=(dict(showgrid=False)),
)



left_column, right_column = st.columns(2)
left_column.plotly_chart(fig_comparaison_CA, use_container_width=True)
right_column.plotly_chart(fig3_pourcentage_produit, use_container_width=True)

left_column2, right_column2 = st.columns(2)
left_column.plotly_chart(fig_quantité_vendue_par_produit, use_container_width=True)
right_column.plotly_chart(fig_col_revenu_brut, use_container_width=True)


# Partie Maps
st.header('CA par ville')

# Calculer le chiffre d'affaires total par ville
CA_par_ville = df_selection.groupby(by=["Villes"])[["Chiffre_d'affaires"]].sum().reset_index()

# Recupérer latite,longitude,Villes,CA dans une liste 
list = list(zip(df_selection["Latitude_Ville"], df_selection["Longitude_Ville"], df_selection["Villes"]))

CENTER = (46, 1.8883335) #lat, long de la France
map = folium.Map(location=CENTER, zoom_start=6)

# Markers
for (lat, lng, ville), ca in zip(list, CA_par_ville["Chiffre_d'affaires"]):
    folium.Marker(
        [lat, lng],
        popup=f"{ville}: {ca}€",
        tooltip="CHIFFRE D'AFFAIRES"
    ).add_to(map)

st_folium(map, width=725) # Afficher la Carte

hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)


# Pour lancer l'application en local: streamlit run streamlit_app.py
