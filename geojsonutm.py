import streamlit as st
import zipfile
import os
import geojson
import xml.etree.ElementTree as ET
import pandas as pd
from io import BytesIO
from pyproj import Proj

# Fungsi untuk mendekompres file KMZ dan mengambil file KML
def extract_kml_from_kmz(kmz_file, output_folder):
    with zipfile.ZipFile(kmz_file, 'r') as kmz:
        kmz.extractall(output_folder)
        for file in os.listdir(output_folder):
            if file.endswith(".kml"):
                return os.path.join(output_folder, file)
    return None

# Fungsi untuk mengonversi KML menjadi GeoJSON
def kml_to_geojson(kml_file, kmz_filename):
    tree = ET.parse(kml_file)
    root = tree.getroot()

    # Namespace yang digunakan di file KML
    namespace = {'ns': 'http://www.opengis.net/kml/2.2'}

    geojson_data = {
        "type": "FeatureCollection",
        "features": []
    }

    # List untuk menyimpan nama file GeoJSON yang akan dihasilkan
    geojson_files = []

    for placemark in root.findall('.//ns:Placemark', namespace):
        name = placemark.find('./ns:name', namespace).text if placemark.find('./ns:name', namespace) is not None else "Unnamed"
        
        # Mengambil data koordinat
        coordinates = placemark.find('.//ns:coordinates', namespace).text if placemark.find('.//ns:coordinates', namespace) is not None else None
        
        if coordinates:
            coords = coordinates.strip().split(' ')
            coords = [coord.split(',') for coord in coords]  # Pisahkan setiap koordinat ke [longitude, latitude]
            coords = [[float(coord[0]), float(coord[1])] for coord in coords]  # Konversi koordinat ke tipe float

            utm_coords = [decimal_degrees_to_utm(coord[1], coord[0])[:2] for coord in coords]

            # Menangani geometri Poligon
            if len(coords) > 2:
                feature = {
                    "type": "Feature",
                    "geometry": {
                        "type": "Polygon",
                        "coordinates": [coords]  # Membungkus koordinat dalam list untuk Poligon
                    },
                    "properties": {
                        "name": name,
                        "kmz_filename": kmz_filename,  # Menambahkan nama file KMZ
                        "coordinates_dd": coords,      # Tambahkan koordinat DD
                        "coordinates_utm": utm_coords  # Tambahkan koordinat UTM
                    }
                }
                geojson_data["features"].append(feature)

                # Simpan setiap feature sebagai file GeoJSON terpisah dengan nama sesuai dengan name dari placemark
                geojson_filename = f"{name.replace(' ', '_')}.geojson"
                geojson_files.append((geojson_filename, geojson.dumps({
                    "type": "FeatureCollection",
                    "features": [feature]
                })))

    return geojson_files


# Fungsi untuk mengonversi Decimal Degrees ke UTM
def decimal_degrees_to_utm(lat, lon):
    zone = int((lon + 180) / 6) + 1  # Menghitung zona UTM
    proj_utm = Proj(proj="utm", zone=zone, ellps="WGS84")
    easting, northing = proj_utm(lon, lat)
    
    # Menangani belahan bumi selatan
    if lat < 0:
        northing += 10000000  # UTM menggunakan offset 10 juta untuk belahan bumi selatan

    # Membulatkan Easting dan Northing ke dua angka di belakang koma
    easting = round(easting, 2)
    northing = round(northing, 2)
    
    return easting, northing, zone



# Fungsi untuk membuat GeoJSON dan Excel konsisten
def geojson_to_excel(geojson_files):
    data = []

    for geojson_filename, geojson_data in geojson_files:
        # Memuat data GeoJSON
        features = geojson.loads(geojson_data).get("features", [])

        for feature in features:
            name = feature['properties'].get('name', 'Unnamed')
            kmz_filename = feature['properties'].get('kmz_filename', 'Unknown')
            coordinates_dd = feature['properties'].get('coordinates_dd', [])
            coordinates_utm = feature['properties'].get('coordinates_utm', [])

            # Sinkronisasi data untuk Excel
            for dd, utm in zip(coordinates_dd, coordinates_utm):
                lon, lat = dd
                easting, northing = utm
                zone = decimal_degrees_to_utm(lat, lon)[2]  # Ambil zona dari koordinat DD
                data.append({
                    "GeoJSON Filename": geojson_filename,
                    "Name": name,
                    "Longitude (DD)": lon,
                    "Latitude (DD)": lat,
                    "Easting (UTM)": easting,
                    "Northing (UTM)": northing,
                    "UTM Zone": zone,
                })

    # Membuat DataFrame dari data
    df = pd.DataFrame(data)
    return df

# Fungsi utama tetap menggunakan geojson_files untuk membuat Excel dan GeoJSON


# Fungsi utama aplikasi Streamlit
def main():
    st.title("KMZ to GeoJSON and Excel Converter")
    st.write("Upload a KMZ file, and this app will convert it into GeoJSON files. You can also download it as an Excel file.")

    # Upload file KMZ
    uploaded_file = st.file_uploader("Upload KMZ file", type=["kmz"])

    if uploaded_file:
        with st.spinner("Processing KMZ file..."):
            # Temp folder untuk ekstraksi file
            temp_folder = "temp_kmz"
            os.makedirs(temp_folder, exist_ok=True)

            # Simpan file KMZ sementara
            kmz_path = os.path.join(temp_folder, "uploaded.kmz")
            with open(kmz_path, "wb") as f:
                f.write(uploaded_file.read())

            # Ekstrak file KML
            kml_file = extract_kml_from_kmz(kmz_path, temp_folder)
            if not kml_file:
                st.error("Failed to extract KML from KMZ file.")
                return

            # Ambil nama file KMZ (untuk digunakan dalam GeoJSON dan Excel)
            kmz_filename = uploaded_file.name

            # Konversi KML ke GeoJSON
            geojson_files = kml_to_geojson(kml_file, kmz_filename)

            # Menyiapkan file ZIP untuk download
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for geojson_filename, geojson_data in geojson_files:
                    zipf.writestr(geojson_filename, geojson_data)

            zip_buffer.seek(0)

            # Tombol untuk mendownload file ZIP yang berisi banyak file GeoJSON
            st.download_button(
                label="Download GeoJSON Files (ZIP)",
                data=zip_buffer,
                file_name="geojson_files.zip",
                mime="application/zip"
            )

            # Tambahkan tombol untuk mengonversi GeoJSON menjadi Excel
            st.write("Download GeoJSON as Excel file")
            excel_df = geojson_to_excel(geojson_files)
            
            # Menyimpan DataFrame sebagai file Excel dalam memori
            excel_file = BytesIO()
            with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
                excel_df.to_excel(writer, index=False)
            excel_file.seek(0)

            st.download_button(
                label="Download Excel",
                data=excel_file,
                file_name=kmz_filename + ".xlsx",  # Nama file Excel sesuai dengan file KMZ yang diunggah
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # Hapus folder sementara
        if os.path.exists(temp_folder):
            for file in os.listdir(temp_folder):
                os.remove(os.path.join(temp_folder, file))
            os.rmdir(temp_folder)

if __name__ == "__main__":
    main()
