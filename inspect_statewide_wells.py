import geopandas as gpd
path = "zip:///Users/davidmichalove/Desktop/automate/o_g_shape_files/OGWells_statewide.zip"
gdf = gpd.read_file(path, layer='OGWells_statewide')
print(f"Columns: {list(gdf.columns)}")
if not gdf.empty:
    print(gdf.drop(columns=['geometry'], errors='ignore').head(1).to_dict('records'))
