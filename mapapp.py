import sys
import os
from datetime import datetime
import pandas as pd
import folium
from folium.plugins import MarkerCluster
from folium.plugins import BeautifyIcon

from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel
from PyQt5.QtCore import Qt
import webbrowser

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Arrastra y suelta un archivo Excel")
        self.setGeometry(100, 100, 600, 400)

        self.label = QLabel("Arrastra y suelta un archivo Excel aquí", self)
        self.label.setGeometry(50, 50, 500, 300)
        self.label.setAlignment(Qt.AlignCenter)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
                self.process_excel(file_path)

    def process_excel(self, file_path):
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))

        # Cargar el archivo Excel y el archivo GeoJSON desde la ruta empaquetada
        cp_file_path = os.path.join(base_path, 'data', 'Listado-de-CP.xlsx')
        geojson_file_path = os.path.join(base_path, 'data', '0.geojson')

        # Leer los archivos Excel
        cp_df = pd.read_excel(cp_file_path)

        # Intentar cargar la primera hoja
        try:
            direcciones_df = pd.read_excel(file_path)
            direcciones_df.columns = direcciones_df.columns.str.replace('\ufeff', '', regex=True).str.strip()
        except Exception as e:
            return

        # Mapear nombres alternativos de las columnas
        columnas_mapeo = {
            'Evt_Label': ['ID_ORDEN', 'Evt_Label'],
            'Evt_PROVINCIA': ['COD_POSTAL', 'Evt_PROVINCIA'],
            'Res_Label': ['NOM_INGENIEROS', 'Res_Label'],
            'Dat_StartDate': ['FEC_PLANIF', 'Dat_StartDate']
        }

        # Función para encontrar la columna adecuada en el DataFrame
        def encontrar_columna(df, posibles_nombres):
            for nombre in posibles_nombres:
                if nombre in df.columns:
                    return nombre
            return None  # Devuelve None si no se encuentra la columna

        # Intentar encontrar las columnas en la primera hoja
        col_evt_label = encontrar_columna(direcciones_df, columnas_mapeo['Evt_Label'])
        col_evt_provincia = encontrar_columna(direcciones_df, columnas_mapeo['Evt_PROVINCIA'])
        col_res_label = encontrar_columna(direcciones_df, columnas_mapeo['Res_Label'])
        col_dat_startdate = encontrar_columna(direcciones_df, columnas_mapeo['Dat_StartDate'])

        # Si no se encuentran todas las columnas, intentar cargar la hoja 'DATA'
        if not all([col_evt_label, col_evt_provincia, col_res_label, col_dat_startdate]):
            try:
                direcciones_df = pd.read_excel(file_path, sheet_name='DATA')
                direcciones_df.columns = direcciones_df.columns.str.replace('\ufeff', '', regex=True).str.strip()

                # Volver a buscar las columnas en la hoja 'DATA'
                col_evt_label = encontrar_columna(direcciones_df, columnas_mapeo['Evt_Label'])
                col_evt_provincia = encontrar_columna(direcciones_df, columnas_mapeo['Evt_PROVINCIA'])
                col_res_label = encontrar_columna(direcciones_df, columnas_mapeo['Res_Label'])
                col_dat_startdate = encontrar_columna(direcciones_df, columnas_mapeo['Dat_StartDate'])

                if not all([col_evt_label, col_evt_provincia, col_res_label, col_dat_startdate]):
                    raise ValueError("No se encontraron todas las columnas requeridas en la hoja 'DATA'.")

            except Exception as e:
                return

        # Procesar los DataFrames
        cp_df['codigo_postal'] = cp_df['codigo_postal'].astype(str).apply(lambda x: x.zfill(5))
        direcciones_df.dropna(subset=[col_evt_provincia], inplace=True)
        cp_df.dropna(subset=['codigo_postal', 'Latitud', 'Longitud'], inplace=True)

        # Realizar las transformaciones necesarias
        direcciones_df['codigo_postal'] = direcciones_df[col_evt_provincia].str.split('-').str[0].astype(str)
        cp_df['codigo_postal'] = cp_df['codigo_postal'].astype(str)

        direcciones_df[col_res_label] = direcciones_df[col_res_label].str.split('_').str[0]

        direcciones_df[col_dat_startdate] = pd.to_datetime(direcciones_df[col_dat_startdate], dayfirst=True)
        direcciones_df[col_dat_startdate] = direcciones_df[col_dat_startdate].dt.strftime('%d/%m/%Y')
        direcciones_df['DiaSemana'] = pd.to_datetime(direcciones_df[col_dat_startdate], format='%d/%m/%Y').dt.day_name()

        # Unir los DataFrames por 'codigo_postal'
        direcciones_geo = pd.merge(direcciones_df, cp_df, left_on='codigo_postal', right_on='codigo_postal', how='left')


        # Crear el mapa
        mapa = folium.Map(location=[40.0, -3.7], zoom_start=6)
        marker_cluster = MarkerCluster().add_to(mapa)

        for i, row in direcciones_geo.iterrows():
            lat = row['Latitud']
            lon = row['Longitud']
            if pd.notnull(lat) and pd.notnull(lon):
                # Asignar un icono diferente según el día de la semana con forma personalizada
                if row['DiaSemana'] == 'Monday':
                    icon = BeautifyIcon(
                        icon_shape='marker',
                        border_color='blue',
                        border_width=3,
                        text_color='white',
                        icon='briefcase',
                        prefix='fa',
                        background_color='blue'
                    )
                elif row['DiaSemana'] == 'Tuesday':
                    icon = BeautifyIcon(
                        icon_shape='circle-dot',  # Círculo con punto en el centro
                        border_color='green',
                        border_width=11,  # Borde más grueso para destacar
                        text_color='white',
                        icon='cloud',
                        prefix='fa',
                        background_color='green'
                    )
                elif row['DiaSemana'] == 'Wednesday':
                    icon = BeautifyIcon(
                        icon_shape='marker',
                        border_color='orange',
                        border_width=3,
                        text_color='orange',
                        icon='star',
                        prefix='fa'
                    )
                elif row['DiaSemana'] == 'Thursday':
                    icon = BeautifyIcon(
                        icon_shape='circle',
                        border_color='purple',
                        border_width=3,
                        text_color='purple',
                        icon='calendar',
                        prefix='fa'
                    )
                elif row['DiaSemana'] == 'Friday':
                    icon = BeautifyIcon(
                        icon_shape='circle',
                        border_color='red',
                        border_width=3,
                        text_color='red',
                        icon='flag',
                        prefix='fa'
                    )
                else:
                    # Cambiar el punto amarillo a una forma más visible y un color diferente
                    icon = BeautifyIcon(
                        icon_shape='rectangle',
                        border_color='gray',
                        border_width=3,
                        text_color='black',
                        icon='circle',
                        prefix='fa'
                    )

                # Crear el contenido del popup con el formato requerido
                popup_text = f"""
                <strong>{row.get(col_evt_provincia, 'Sin provincia')}</strong><br>
                - Nombre del técnico: {row.get(col_res_label, 'No disponible')}<br>
                - Número de orden: {row.get(col_evt_label, 'No disponible')}<br>
                - Fecha de la cita: {row.get(col_dat_startdate, 'No disponible')}<br>
                """
                # Definir el tooltip con 'Res_Label'
                tooltip_text = f"{row[col_res_label]} - {row[col_evt_label]}"

                folium.Marker(
                    location=[lat, lon],
                    popup=folium.Popup(popup_text, max_width=300),
                    tooltip=tooltip_text,
                    icon=icon
                ).add_to(marker_cluster)

        with open(geojson_file_path, encoding='utf-8') as f:
            geojson_data = f.read()

        # Añadir el contorno de las provincias con GeoJSON (sin relleno)
        folium.GeoJson(
            geojson_data,  # Cambia esto por la ruta de tu archivo GeoJSON
            name="Límites de Provincias",
            style_function=lambda feature: {
                'color': 'black',     
                'weight': 1.5,        
                'fillOpacity': 0      
            }
        ).add_to(mapa)


                # Guardar el mapa en un archivo HTML con un nombre único
        timestamp = datetime.now().strftime("%m%d_%H%M")
        map_filename = f'mapa_visitas_cluster_{timestamp}.html'
        mapa.save(map_filename)
        webbrowser.open(map_filename)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
