#%%
#pandas openpyxl
import pandas as pd
print("Hola, mundo!")

#%%
def QuitarMultiplicados(PD_DataFrame, columnaIDs):
    tabla_frecuencia = PD_DataFrame[columnaIDs].value_counts()
    lista_IDs = tabla_frecuencia[ tabla_frecuencia > 1 ].index.tolist()
    print(f"Cantidad de IDs con duplicados: {len(lista_IDs)}")
    for ID in lista_IDs:
        #df_block = pd.DataFrame(columns= PD_DataFrame.columns)
        #df_block = pd.concat([df_block, PD_DataFrame[:][PD_DataFrame[columnaIDs] == ID]], ignore_index=True) 
        df_block = PD_DataFrame[:][PD_DataFrame[columnaIDs] == ID]
        PD_DataFrame = PD_DataFrame[:][PD_DataFrame[columnaIDs] != ID]
        PD_DataFrame = pd.concat([PD_DataFrame, df_block.iloc[[0],:]], ignore_index=True)
    
    return PD_DataFrame

#%% Definimos la ruta de trabajo y los nombres de los archivos
ruta_trabajo = "C:/Users/cgonzalezb/Downloads/SeminarioTit/ProyectoEquipo4/BasesProyecto/"
nombre_base = "CENTRO Base Ene-Jul 2025.xlsx"
pestania_base = "ProvLev"
nombre_cat_tiendas = "CATALOGO DE TIENDAS_R.xlsx"
pestania_cat_tiendas = "CatalogoTiendas"
nombre_cat_familias = "REGIONES_TOTALES Familia II.xlsx"
pestania_cat_familias = "CatalogoFamilias"

nombre_base_Ckp = "CkP_" + nombre_base
nombre_base_Grp = "Grp_" + nombre_base
Ckp = False


#%% Leemos las pestañas de los archivos Excel
file_base = pd.ExcelFile(ruta_trabajo + nombre_base)  
print(file_base.sheet_names)

file_cat_tiendas = pd.ExcelFile(ruta_trabajo + nombre_cat_tiendas)  
print(file_cat_tiendas.sheet_names)

file_cat_familias = pd.ExcelFile(ruta_trabajo + nombre_cat_familias)  
print(file_cat_familias.sheet_names)


#%% Leemos la base y los catálogos de tiendas y familias
df_base = file_base.parse(sheet_name= pestania_base)
df_catalogo_tiendas = file_cat_tiendas.parse(sheet_name= pestania_cat_tiendas)
df_catalogo_familias = file_cat_familias.parse(sheet_name= pestania_cat_familias)


#%% Quitamos los duplicados en los catálogos de tiendas y familias
df_catalogo_familias = QuitarMultiplicados(df_catalogo_familias, 'Aux')
df_catalogo_tiendas = QuitarMultiplicados(df_catalogo_tiendas, 'ID_Sucursal')


#%%NOS QUEDAMOS SOLO CON LAS COLUMNAS QUE NOS INTERESAN
columnas_selc = ['Fecha',
 'Categoria',
 'Familia',
 'Articulo',
 'Descripcion',
 'Cantidad',
 'Venta',
 'Costo',
 'Utilidad',
 'cliente',
 'Sucursal',
 'ID_Suc']
df_base = df_base[columnas_selc]
df_base = df_base.rename(columns={'ID_Suc': 'ID_Sucursal', 'Familia': 'FamiCH'})
del columnas_selc


#%%#CAMBIAMOS EL FORMATO DE LA FECHA A YYYY-MM-DD
df_base['Fecha'] = pd.to_datetime(df_base['Fecha']).dt.strftime('%Y/%m/%d')


#%%
#left join con df_catalogo_tiendas
df_base= pd.merge(df_base, df_catalogo_tiendas[['ID_Sucursal', 'Nom_Sucursal', 'Clasificacion',
       'Region', 'Zona']], how='left', on='ID_Sucursal')


#%%
#left join con df_catalogo_familias
#CAMBIAOS A TIPO STRING Y CREAMOS UNA NUEVA COLUMNA AUX CON LA CONCATENACION DE CATEGORIA Y FamiCH
df_base['Categoria'] = df_base['Categoria'].astype(str)
df_base['FamiCH'] = df_base['FamiCH'].astype(str)
df_base['Aux'] = df_base['Categoria'] + df_base['FamiCH']
df_base = pd.merge(df_base, df_catalogo_familias[['Aux', 'FamiliaII']], how='left', left_on='Aux', right_on='Aux')


# %%
# Reorganizamos las columnas
column_order = df_base.columns.tolist()
print(column_order)
index_famiCH = column_order.index('FamiCH')  # Encontramos el índice de FamiCH
new_order = column_order[:index_famiCH + 1] + ['Aux', 'FamiliaII'] + column_order[index_famiCH + 1:-2]
df_base = df_base[new_order]
print(df_base.columns.tolist())
df_base = df_base.rename(columns={'FamiliaII': 'Familia'})


#%%
# Guardamos el resultado en un nuevo archivo Excel (Checkpoint)
df_base.to_excel(ruta_trabajo + nombre_base_Ckp, index=False, sheet_name= pestania_base)




#%% Solo en caso de iniciar desde el Checkpoin
if Ckp:
    file = pd.ExcelFile(ruta_trabajo + nombre_base_Ckp)  
    print(file.sheet_names)

    #df_base_datos = pd.read_excel(ruta_base_datos, sheet_name='ProvLev')
    df_base = file.parse(sheet_name= pestania_base)




#%%
# Aseguramos que la columna 'Fecha' esté en formato datetime para extraer año y mes
df_base['Fecha'] = pd.to_datetime(df_base['Fecha'])
df_base['Año'] = df_base['Fecha'].dt.year
df_base['Mes'] = df_base['Fecha'].dt.month


#%%
df_base = df_base.rename(columns={'Cantidad': 'Piezas'})
# Agrupamos y aplicamos la función de agregación
df_base = df_base.groupby(['Año', 'Mes', 
                                            'Region', 'Zona', 'Clasificacion', 'Nom_Sucursal', 
                                            'Categoria', 'Familia']
                                            ).agg({
                                                'Piezas': 'sum',
                                                'Venta': 'sum',
                                                'Costo': 'sum',
                                                'Utilidad': 'sum'
                                            }).reset_index()

# Ordenamos el resultado en orden descendente por año, mes y Venta
df_base = df_base.sort_values(
    by=['Año', 'Mes', 'Region', 'Nom_Sucursal', 'Venta', 'Piezas', 'Utilidad'], 
    ascending=[False, False, True, True, False, False, False]
)


#%% Guardamos la base agrupada en un nuevo archivo Excel
df_base.to_excel(ruta_trabajo + nombre_base_Grp, index=False, sheet_name= pestania_base)
