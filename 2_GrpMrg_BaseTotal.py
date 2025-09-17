#%%
#pandas openpyxl
import pandas as pd

#%% Cargamos todas las bases agrupadas y las juntamos en un solo DF

ruta_trabajo = "C:/Users/cgonzalezb/Downloads/SeminarioTit/ProyectoEquipo4/BasesProyecto/"
nombres_bases_Grp = ["Grp_CENTRO Base Ene-Dic 2022.xlsx", 
                     "Grp_CENTRO Base Ene-Dic 2023.xlsx", 
                     "Grp_CENTRO Base Ene-Dic 2024.xlsx", 
                     "Grp_CENTRO Base Ene-Jul 2025.xlsx"]
pestanias = ["ProvLev","ProvLev","ProvLev","ProvLev"]

df_total = pd.DataFrame()
for nombre_base, nombre_pestania in zip(nombres_bases_Grp, pestanias):
    print(f"leyendo base: {nombre_base}, pestania: {nombre_pestania}")
    
    file_base = pd.ExcelFile(ruta_trabajo + nombre_base)  
    df_base = file_base.parse(sheet_name= nombre_pestania)
    
    # Concatenando verticalmente
    df_total = pd.concat([df_total, df_base], axis=0)

#%%
#Filtramos solo DICO CENTRO
df_total = df_total[df_total["Region"]=="DICO CENTRO"]

#%% Guardamos la base totaal agrupada en un nuevo archivo Excel
df_total.to_excel(ruta_trabajo + "Grp_CENTRO BaseTotal 2022-2025.xlsx", index=False, sheet_name= "ProvLev")

# %%
