import geopandas as gpd
path = "zip:///Users/davidmichalove/Desktop/automate/Polygon_Belmont_County_Web_Parcels_20260501085529 (1).zip"
gdf = gpd.read_file(path)
print(f"Belmont CRS: {gdf.crs}")
