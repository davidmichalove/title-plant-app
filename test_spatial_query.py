import requests
import json
import geopandas as gpd

gdf = gpd.read_file("zip:///Users/davidmichalove/Desktop/automate/Polygon_Belmont_County_Web_Parcels_20260501085529 (1).zip")
gdf = gdf.to_crs(epsg=3735)
geom_3735 = gdf.geometry.iloc[0]
buffered = geom_3735.buffer(500)

bounds = buffered.bounds # minx, miny, maxx, maxy
env = f"{bounds[0]},{bounds[1]},{bounds[2]},{bounds[3]}"

url = "https://services5.arcgis.com/ajRlmtxbNBjZggOT/arcgis/rest/services/Oil_and_Gas_Wells_Bottom_Holes/FeatureServer/0/query"
params = {
    "where": "1=1",
    "geometry": env,
    "geometryType": "esriGeometryEnvelope",
    "spatialRel": "esriSpatialRelIntersects",
    "inSR": 3735,
    "outFields": "API_WELLNO",
    "returnGeometry": "true",
    "outSR": 3735,
    "f": "json"
}

resp = requests.post(url, data=params)
print(resp.status_code)
data = resp.json()
if 'features' in data:
    print(f"Found {len(data['features'])} features")
else:
    print(data)
