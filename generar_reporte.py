import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Cargar el archivo csv en un dataframe de pandas
df = pd.read_csv("notasV.csv")

# Mostrar las primeras 10 filas del dataframe
head10 = df.head(10)
print("Primeras 10 filas:")   
print(head10)

# Mostrar las columnas únicas para "Materia" y "Estudiantes"
Materia = df["Materia"].unique()
Estudiante = df["Estudiante"].unique()
print("\nMaterias únicas:")
print(Materia)
print("\nEstudiantes únicos:")
print(Estudiante)

# Calcular promedio de notas para cada fila
columnas_notas = [f"Nota{i}" for i in range(1, 10)]
df["NotaFinal"] = df[columnas_notas].mean(axis=1)

# Mostrar dataframe con las notas finales
print("\nDataFrame con NotaFinal calculada:")
print(df.head())

# Calcular el promedio de notas por materia
promedio_notas_materia = df.groupby("Materia")["NotaFinal"].mean()
print("\nPromedio de notas por materia:")
print(promedio_notas_materia)

# Calcular el promedio de notas por estudiante
promedio_notas_estudiante = df.groupby("Estudiante")["NotaFinal"].mean()
print("\nPromedio de notas por estudiante:")
print(promedio_notas_estudiante)

# Mostrar la mejor nota final
mejor_nota = df.loc[df["NotaFinal"].idxmax()]
print("\nMejor nota:")
print(mejor_nota)

# Configuración de estilos seaborn para todos los gráficos
sns.set_theme(style="whitegrid", palette="pastel")

# --- Gráfico 1: Pie Chart de promedios por materia ---
plt.figure(figsize=(8, 8))
plt.pie(
    promedio_notas_materia,
    labels=promedio_notas_materia.index,
    autopct='%1.1f%%',
    startangle=90,
    colors=sns.color_palette("pastel"),
    explode=[0.05] * len(promedio_notas_materia)
    )  # Separación uniforme
plt.title("Distribución del Promedio de Notas por Materia", fontsize=14)
plt.savefig("PieChart_NotasPorMateria.png")
plt.show()

# --- Gráfico 2: Barras de promedio de notas por materia ---
plt.figure(figsize=(8, 5))
sns.barplot(x=promedio_notas_materia.index, y=promedio_notas_materia.values)
plt.title("Promedio de Notas por Materia")
plt.xlabel("Materia")
plt.ylabel("Promedio de Notas")
plt.xticks(rotation=45)
plt.savefig("PromedioNotasPorMateria.png")
plt.show()

# --- Gráfico 3: Barras de promedio de notas por estudiante ---
plt.figure(figsize=(8, 5))
sns.barplot(x=promedio_notas_estudiante.index, y=promedio_notas_estudiante.values)
plt.title("Promedio de Notas por Estudiante")
plt.xlabel("Estudiante")
plt.ylabel("Promedio de Notas")
plt.xticks(rotation=45)
plt.savefig("PromedioNotasPorEstudiante.png")
plt.show()

# --- Gráfico 4: Histograma de distribución de notas finales ---
plt.figure(figsize=(8, 5))
sns.histplot(df["NotaFinal"], bins=10, kde=True)
plt.title("Distribución de Notas Finales")
plt.xlabel("Nota Final")
plt.ylabel("Frecuencia")
plt.savefig("DistribucionNotasFinales.png")
plt.show()

# --- Gráfico 5: Boxplot de notas por materia (adicional) ---
plt.figure(figsize=(8, 5))
sns.boxplot(data=df, x="Materia", y="NotaFinal")
plt.title("Distribución de Notas por Materia")
plt.xticks(rotation=45)
plt.savefig("BoxplotNotasPorMateria.png")
plt.show()