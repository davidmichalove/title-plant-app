import os
import requests
import json
import re
import urllib.parse
from bs4 import BeautifulSoup

try:
    import geopandas as gpd
    from shapely.geometry import Point, Polygon, shape
except ImportError:
    gpd = None

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PARCEL_SHP = os.path.join(BASE_DIR, "Polygon_Belmont_County_Web_Parcels_20260501085529 (1).zip")
OG_SHP_DIR = os.path.join(BASE_DIR, "o_g_shape_files")

LOCAL_SHAPES = {
    "Unitizations": "Unitizations_belmont.shp",
    "Horizontal Drilling Units": "HorizontalDrillingUnits (1)_belmont.shp",
    "Well Pads": "OG_WellPads_belmont.shp",
    "Urban Areas": "UrbanAreas_belmont.shp"
}

API_ENDPOINTS = {
    "Bottom Holes": "https://services5.arcgis.com/ajRlmtxbNBjZggOT/arcgis/rest/services/Oil_and_Gas_Wells_Bottom_Holes/FeatureServer/0/query",
    "Well Bores": "https://services5.arcgis.com/ajRlmtxbNBjZggOT/arcgis/rest/services/Oil_and_Gas_Well_Bores/FeatureServer/0/query"
}

def _generate_report_for_buffer(parcel_num, geom, buffer_size, base_out_dir, loaded_shapes, log_callback):
    log_callback(f"Generating report for {buffer_size}ft buffer...")
    
    out_dir = os.path.join(base_out_dir, f"Buffer_{buffer_size}ft")
    os.makedirs(out_dir, exist_ok=True)
    pdf_dir = os.path.join(out_dir, "PDFs")
    os.makedirs(pdf_dir, exist_ok=True)
    
    buffered = geom.buffer(buffer_size)
    bounds = buffered.bounds
    env_str = f"{bounds[0]},{bounds[1]},{bounds[2]},{bounds[3]}"
    
    report_lines = [
        f"Oil & Gas Activity Report for Parcel: {parcel_num}",
        f"Buffer: {buffer_size} feet",
        "="*50,
        ""
    ]
    
    # 2. Check Local Shapefiles
    for name, gdf_og in loaded_shapes.items():
        try:
            intersecting = gdf_og[gdf_og.intersects(buffered)]
            
            report_lines.append(f"--- {name} ---")
            if intersecting.empty:
                report_lines.append("No intersections found.")
            else:
                report_lines.append(f"Found {len(intersecting)} intersecting feature(s):")
                for _, row in intersecting.iterrows():
                    attrs = row.drop(['geometry'], errors='ignore').to_dict()
                    cleaned_attrs = {k: v for k, v in attrs.items() if str(v) != 'nan' and v is not None}
                    report_lines.append(f"  - {cleaned_attrs}")
            report_lines.append("")
        except Exception as e:
            report_lines.append(f"--- {name} ---")
            report_lines.append(f"Error checking shapefile: {e}")
            report_lines.append("")
            
    # 3. Check Live APIs for Wells
    api_numbers = set()
    for name, url in API_ENDPOINTS.items():
        try:
            params = {
                "where": "1=1",
                "geometry": env_str,
                "geometryType": "esriGeometryEnvelope",
                "spatialRel": "esriSpatialRelIntersects",
                "inSR": 3735,
                "outFields": "*",
                "returnGeometry": "true",
                "outSR": 3735,
                "f": "json"
            }
            resp = requests.post(url, data=params, timeout=30)
            data = resp.json()
            
            report_lines.append(f"--- {name} (Live API) ---")
            features = data.get('features', [])
            
            exact_features = []
            for feat in features:
                attrs = feat.get('attributes', {})
                geom_data = feat.get('geometry', {})
                if 'x' in geom_data and 'y' in geom_data:
                    pt = Point(geom_data['x'], geom_data['y'])
                    if buffered.intersects(pt):
                        exact_features.append(attrs)
                        if 'API_WELLNO' in attrs:
                            api_numbers.add(attrs['API_WELLNO'])
            
            if not exact_features:
                report_lines.append("No intersections found.")
            else:
                report_lines.append(f"Found {len(exact_features)} intersecting well(s):")
                for attrs in exact_features:
                    cleaned_attrs = {k: v for k, v in attrs.items() if v is not None}
                    report_lines.append(f"  - {cleaned_attrs}")
            report_lines.append("")
            
        except Exception as e:
            report_lines.append(f"--- {name} (Live API) ---")
            report_lines.append(f"Error accessing API: {e}")
            report_lines.append("")

    # 4. Fetch PDF Reports
    if api_numbers:
        log_callback(f"({buffer_size}ft buffer) Found API numbers: {api_numbers}. Fetching reports...")
        report_lines.append("--- PDF Downloads ---")
        for api in api_numbers:
            try:
                card_url = f"https://gis.ohiodnr.gov/MapViewer/WellSummaryCard.asp?api={api}"
                html_resp = requests.get(card_url, timeout=30)
                html_resp.raise_for_status()
                
                soup = BeautifulSoup(html_resp.text, 'html.parser')
                links = soup.find_all('a', href=True)
                download_links = [a for a in links if 'download.ashx?' in a['href']]
                
                if download_links:
                    report_lines.append(f"API {api}: Found {len(download_links)} document(s).")
                    for link in download_links:
                        doc_name = link.text.strip().replace(" ", "_")
                        href = link['href']
                        if href.startswith('/'):
                            href = "https://gis.ohiodnr.gov" + href
                            
                        pdf_resp = requests.get(href, timeout=60)
                        
                        disp = pdf_resp.headers.get('content-disposition', '')
                        filename = f"{api}_{doc_name}.pdf"
                        if 'filename=' in disp:
                            filename = disp.split('filename=')[-1].strip('"')
                            
                        pdf_path = os.path.join(pdf_dir, filename)
                        with open(pdf_path, 'wb') as f:
                            f.write(pdf_resp.content)
                        report_lines.append(f"  - Downloaded: {filename}")
                else:
                    report_lines.append(f"API {api}: No PDF documents found on summary card.")
            except Exception as e:
                report_lines.append(f"API {api}: Failed to download documents ({e})")
    
    # Write report
    report_path = os.path.join(out_dir, "O_AND_G_REPORT.txt")
    with open(report_path, "w") as f:
        f.write("\n".join(report_lines))


def check_parcel_og_activity(parcel_num, log_callback=print, out_dir=None):
    if gpd is None:
        log_callback("GeoPandas not found. O&G check failed.")
        return
        
    log_callback(f"Starting O&G check for parcel: {parcel_num}")
    
    # 1. Load parcel geometry
    try:
        gdf_parcel = gpd.read_file(f"zip://{PARCEL_SHP}")
        parcel = gdf_parcel[gdf_parcel['parcel_no'] == parcel_num]
        if parcel.empty:
            log_callback(f"Parcel {parcel_num} not found in shapefile.")
            return
            
        # Convert to Ohio State Plane South (ft) EPSG:3735
        parcel = parcel.to_crs(epsg=3735)
        geom = parcel.geometry.iloc[0]
        
    except Exception as e:
        log_callback(f"Error loading parcel geometry: {e}")
        return
        
    # Setup base output dir
    if out_dir is None:
        out_dir = os.path.join(BASE_DIR, f"PID {parcel_num}", "WELL INFO")
    os.makedirs(out_dir, exist_ok=True)
    
    # Preload local shapefiles so we don't reload them for every buffer!
    log_callback("Pre-loading local shapefiles...")
    loaded_shapes = {}
    for name, filename in LOCAL_SHAPES.items():
        try:
            path = os.path.join(OG_SHP_DIR, filename)
            if filename.endswith(".zip"):
                path = f"zip://{path}"
            gdf_og = gpd.read_file(path)
            if gdf_og.crs is None:
                gdf_og.set_crs(epsg=3735, inplace=True)
            elif gdf_og.crs.to_string() != "EPSG:3735":
                gdf_og = gdf_og.to_crs(epsg=3735)
            loaded_shapes[name] = gdf_og
        except Exception as e:
            log_callback(f"Failed to load {name}: {e}")
    
    # Generate reports for different buffer sizes
    buffer_sizes = [500, 1000, 2000, 5000]
    for size in buffer_sizes:
        try:
            _generate_report_for_buffer(parcel_num, geom, size, out_dir, loaded_shapes, log_callback)
        except Exception as e:
            log_callback(f"Error generating {size}ft report: {e}")
            
    log_callback(f"O&G Checks complete. Reports saved to: {out_dir}")

if __name__ == "__main__":
    check_parcel_og_activity("53-01031.000")
