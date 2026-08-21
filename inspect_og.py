import geopandas as gpd
import warnings
warnings.filterwarnings('ignore')

zips = [
    "HorizontalDrillingUnits (1).zip",
    "OGWells_war.zip",
    "OG_WellPads.zip",
    "Unitizations.zip",
    "UrbanAreas.zip"
]

for z in zips:
    path = f"zip:///Users/davidmichalove/Desktop/automate/o_g_shape_files/{z}"
    try:
        gdf = gpd.read_file(path)
        print(f"--- {z} ---")
        print(f"Columns: {list(gdf.columns)}")
        print(f"Rows: {len(gdf)}")
        print(f"CRS: {gdf.crs}")
        if not gdf.empty:
            print(gdf.drop(columns=['geometry'], errors='ignore').head(1).to_dict('records'))
    except Exception as e:
        print(f"Failed to read {z}: {e}")
