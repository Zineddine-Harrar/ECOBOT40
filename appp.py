# -*- coding: utf-8 -*-
"""
Created on Wed Mar  6 09:35:19 2024

@author: zined
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from dash import Dash, dcc, html, Input, Output
import plotly.graph_objects as go

# Chargement du fichier Excel
file_path = 'TaskReport_GS142-0120-G2N-0001_20240318102004.xlsx'  # Remplacez par le chemin réel du fichier
data = pd.read_excel(file_path)

# Nettoyage des données
data = data.drop(index=0)  # Suppression de la ligne d'en-tête redondante
data['Task start time'] = pd.to_datetime(data['Task start time'])
data['End time'] = pd.to_datetime(data['End time'])
data['Task completion (%)'] = pd.to_numeric(data['Task completion (%)'])
data['Cleaning plan area (㎡)'] = data['Cleaning plan area (㎡)'].str.replace(',', '').astype(float)
data['Actual cleaning area(㎡)'] = data['Actual cleaning area(㎡)'].str.replace(',', '').astype(float)
data['Total time (h)'] = pd.to_numeric(data['Total time (h)'])
data['week'] = data['Task start time'].dt.isocalendar().week
data['successful_completion'] = data['Task completion (%)'] >= 90
data['suivi'] = data['Task completion (%)'] >= 0

# Calcul des KPI
# Taux de suivi planning par semaine (pour tâches complétées à 90% ou plus)
taux_suivi_planning_weekly = (data.groupby('week')['suivi'].sum()/23)*100

taux_completion_planning_weekly = (data.groupby('week')['successful_completion'].sum()/data.groupby('week')['suivi'].sum())*100

# Heures d'usage par semaine
usage_per_week = data.groupby('week')['Total time (h)'].sum()

# Surfaces couvertes par semaine
surfaces_couvertes_weekly = data.groupby('week')['Actual cleaning area(㎡)'].sum()


# Coût horaire par semaine
cost_per_month = 900
weekly_hours = data.groupby('week')['Total time (h)'].sum()
cost_per_week = cost_per_month / 4
cout_horaire_weekly = cost_per_week / weekly_hours

semaines_options = [{'label': f'Semaine {week}', 'value': week} for week in sorted(data['week'].unique())]

# Calcul des dates de début et de fin de semaine
# Calculer la date de début de la semaine pour chaque tâche
data['Start of Week'] = data['Task start time'] - pd.to_timedelta(data['Task start time'].dt.weekday, unit='d')

# Calculer la date de fin de la semaine pour chaque tâche
data['End of Week'] = data['Start of Week'] + pd.to_timedelta(6, unit='d')

# Obtenir les dates de début et de fin pour chaque semaine
weekly_dates = data.groupby('week').agg({'Start of Week':'min', 'End of Week':'max'}).reset_index()

# Formater les dates pour qu'elles appparaissent comme 'Semaine 8 (01/01/2024 - 07/01/2024)'
semaines_options = [
    {
        'label': f"Semaine {row['week']} ({row['Start of Week'].strftime('%d/%m/%Y')} - {row['End of Week'].strftime('%d/%m/%Y')})",
        'value': row['week']
    }
    for _, row in weekly_dates.iterrows()
]
# Initialisation de l'appplication Dash
appp = Dash(__name__)

# Style de l'encadrement extérieur et des couleurs de la page
external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']  # URL de la feuille de style externe
appp = Dash(__name__, external_stylesheets=external_stylesheets)

# Création des graphiques avec Plotly
fig_suivi = go.Figure([go.Bar(x=taux_suivi_planning_weekly.index, y=taux_suivi_planning_weekly.values)])
fig_suivi.update_layout(title='Taux de suivi du planning par semaine', xaxis_title='Semaine', yaxis_title='Taux de suivi (%)')

fig_usage = go.Figure([go.Scatter(x=usage_per_week.index, y=usage_per_week.values, mode='lines+markers')])
fig_usage.update_layout(title='Heures d\'usage par semaine', xaxis_title='Semaine', yaxis_title='Heures')

fig_surface = go.Figure([go.Bar(x=surfaces_couvertes_weekly.index, y=surfaces_couvertes_weekly.values)])
fig_surface.update_layout(title='Surfaces couvertes par semaine (㎡)', xaxis_title='Semaine', yaxis_title='Surface (㎡)')

fig_cout = go.Figure([go.Scatter(x=cout_horaire_weekly.index, y=cout_horaire_weekly.values, mode='lines+markers')])
fig_cout.update_layout(title='Coût horaire par semaine', xaxis_title='Semaine', yaxis_title='Coût (euros/heure)')

# Création de la liste des options pour le dropdown, basée sur les semaines disponibles
dropdown_options = [{'label': f"Semaine {week}", 'value': week} for week in taux_suivi_planning_weekly.index]
# Ajout des graphiques à l'appplication Dash
appp.layout = html.Div(style={'backgroundColor': '#f9f9f9'}, children=[
    html.Div([
        html.Img(src='https://atalian.fr/wp-content/uploads/sites/4/2013/05/atalian-logo.png', style={'height': '130px', 'width': 'auto'})
    ], style={'textAlign': 'center', 'margin-bottom': '20px'}),
    
    html.H1('Tableau de bord des KPI ECOBOT 40', style={'textAlign': 'center', 'color': '#007BFF'}),
    
    html.Div([
        dcc.Dropdown(
            id='week-dropdown',
            options=semaines_options,  # Utilisation des nouvelles options avec les dates de début et de fin
            value=weekly_dates['week'].iloc[0]  # Valeur par défaut est la première semaine disponible
        ),
    ], style={'width': '100%', 'margin': '20px 0', 'textAlign': 'center'}),

    
    html.Div(id='gauges-container', children=[
        dcc.Graph(id='gauge-taux-suivi'),
        dcc.Graph(id='gauge-taux-completion')
    ], style={'display': 'flex', 'justify-content': 'space-around', 'margin-bottom': '20px'}),
    
    dcc.Graph(id='graph-suivi', figure=fig_suivi),
    dcc.Graph(id='graph-usage', figure=fig_usage),
    dcc.Graph(id='graph-surface', figure=fig_surface),
    dcc.Graph(id='graph-cout', figure=fig_cout)
])

# Assumons que vos callbacks sont défi
@appp.callback(
    [Output('gauge-taux-suivi', 'figure'),
     Output('gauge-taux-completion', 'figure')],
    [Input('week-dropdown', 'value')]
)
def update_gauges(selected_week):
    # Calcul pour le taux de suivi (Assurez-vous d'avoir un calcul pour le taux de suivi ici)
    taux_suivi = taux_suivi_planning_weekly.loc[selected_week]
    fig_gauge_suivi = go.Figure(go.Indicator(
        mode="gauge+number",
        value=taux_suivi,
        domain={'x': [0, 1], 'y': [0, 1]},
        title={'text': f"Taux de suivi - Semaine {selected_week}"},
        gauge={
            'axis': {'range': [0, 100], 'tickwidth': 1, 'tickcolor': "darkblue"},
            'bar': {'color': "lightblue"},
            'bgcolor': "white",
            'borderwidth': 2,
            'bordercolor': "gray",
            'steps': [
                {'range': [0, 50], 'color': 'red'},
                {'range': [50, 75], 'color': 'yellow'},
                {'range': [75, 100], 'color': 'green'}
            ],
        }
    ))
    
    # Calcul pour le taux de complétion
    taux_completion = taux_completion_planning_weekly.loc[selected_week]
    fig_gauge_completion = go.Figure(go.Indicator(
        mode="gauge+number",
        value=taux_completion,
        domain={'x': [0, 1], 'y': [0, 1]},
        title={'text': f"Taux de complétion - Semaine {selected_week}"},
        gauge={
            'axis': {'range': [0, 100], 'tickwidth': 1, 'tickcolor': "darkblue"},
            'bar': {'color': "darkblue"},
            'bgcolor': "white",
            'borderwidth': 2,
            'bordercolor': "gray",
            'steps': [
                {'range': [0, 50], 'color': 'red'},
                {'range': [50, 75], 'color': 'yellow'},
                {'range': [75, 100], 'color': 'green'}
            ],
        }
    ))
    
    return fig_gauge_suivi, fig_gauge_completion

if __name__ == '__main__':
    appp.run_server(debug=True, port=8050)
