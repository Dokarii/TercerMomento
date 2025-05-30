import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Configuración visual para los gráficos
sns.set_theme(style='whitegrid', palette='pastel')

# --- 1. Carga de datos desde Excel y parseo de fechas ---
df = pd.read_excel('CALIDAD.xlsx', decimal=',', parse_dates=['Fecha'])

# Crear columnas adicionales para facilitar análisis por día y mes
df['FechaSolo'] = df['Fecha'].dt.date
df['Mes'] = df['Fecha'].dt.to_period('M')

# --- 2. Exploración básica: estaciones únicas y conteo de registros ---
estaciones = df['Estacion'].unique()
conteo_estaciones = df['Estacion'].value_counts()

# --- 3. Estadísticas diarias: promedio de PM2.5 y PM10 por día ---
media_diaria = df.groupby('FechaSolo')[['PM2.5', 'PM10']].mean().reset_index()
alertas_pm25 = media_diaria[media_diaria['PM2.5'] > 35]  # Días con alerta de PM2.5 alta

# --- 4. Tendencias mensuales de NO2 y O3 por estación ---
mensual = df.groupby(['Mes', 'Estacion'])[['NO2', 'O3']].mean().reset_index()
mensual_sorted = mensual.sort_values(by='NO2', ascending=False)

# --- 5. Gráfico de barras: promedio mensual de PM2.5 por estación ---
prom_pm25_mensual = df.groupby(['Mes', 'Estacion'])['PM2.5'].mean().reset_index()
plt.figure(figsize=(12,6))
sns.barplot(data=prom_pm25_mensual, x='Mes', y='PM2.5', hue='Estacion')
plt.title('Promedio mensual de PM2.5 por estación')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig('bar_promedio_pm25_mensual.png')
plt.close()

# --- 6. Gráfico de línea: serie diaria de PM10 en estación más contaminada ---
estacion_critica = df.groupby('Estacion')['PM2.5'].mean().idxmax()
df_critica = df[df['Estacion'] == estacion_critica]
serie_pm10 = df_critica.groupby('FechaSolo')['PM10'].mean().reset_index()

plt.figure(figsize=(12,6))
sns.lineplot(data=serie_pm10, x='FechaSolo', y='PM10')
plt.title(f'Serie diaria de PM10 - Estación {estacion_critica}')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig('line_pm10_estacion_critica.png')
plt.close()

# --- 7. Gráfico circular: porcentaje de días con alerta PM2.5 > 35 por estación ---
alertas = df[df['PM2.5'] > 35]
dias_alerta_estacion = alertas['Estacion'].value_counts()
plt.figure(figsize=(8,8))
plt.pie(dias_alerta_estacion, labels=dias_alerta_estacion.index, autopct='%1.1f%%', startangle=140, colors=sns.color_palette('pastel'))
plt.title('Distribución de días de alerta PM2.5 > 35 por estación')
plt.tight_layout()
plt.savefig('pie_alertas_estacion.png')
plt.close()

# --- 8. Heatmap: matriz de correlación entre contaminantes ---
corr = df[['PM2.5', 'PM10', 'NO2', 'O3']].corr()
plt.figure(figsize=(8,6))
sns.heatmap(corr, annot=True, cmap='coolwarm', vmin=-1, vmax=1)
plt.title('Matriz de Correlación de Contaminantes')
plt.tight_layout()
plt.savefig('heatmap_correlacion.png')
plt.close()

# --- 9. Preparar tablas para el reporte ---
head10 = df.head(10)
conteo_estaciones_table = conteo_estaciones.to_frame().reset_index()
conteo_estaciones_table.columns = ['Estacion', 'Conteo']

promedios = df[['PM2.5', 'PM10', 'NO2', 'O3']].mean().reset_index()
promedios.columns = ['Contaminante', 'Promedio']

# Encontrar fila con mayor contaminación total (suma de contaminantes)
df['Total'] = df[['PM2.5', 'PM10', 'NO2', 'O3']].sum(axis=1)
max_row = df.loc[df['Total'].idxmax()]

# --- 10. Leer plantilla HTML y generar reporte final ---
with open('template.html', 'r', encoding='utf-8') as f:
    template = f.read()

html_filled = template.format(
    head10_table=head10.to_html(index=False, classes="table table-striped"),
    estaciones=', '.join(estaciones),
    conteo_table=conteo_estaciones_table.to_html(index=False, classes="table table-striped"),
    prom_table=promedios.to_html(index=False, classes="table table-striped"),
    max_fecha=max_row['Fecha'].strftime('%Y-%m-%d %H:%M:%S'),
    max_estacion=max_row['Estacion'],
    max_valor=round(max_row['Total'], 2)
)

with open('reporte_calidad_aire.html', 'w', encoding='utf-8') as f:
    f.write(html_filled)

print("✅ Reporte generado: reporte_calidad_aire.html")
