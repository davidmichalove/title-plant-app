import geopandas as gpd

gdf = gpd.read_file('/Users/davidmichalove/Desktop/automate/app/shape_files/Belmont_County_Parcels.shp')
print("Columns:")
print(gdf.columns.tolist())
print("First 5 rows of related columns:")
for col in gdf.columns:
    if 'vol' in col.lower() or 'page' in col.lower() or 'book' in col.lower() or 'deed' in col.lower():
        print(col, gdf[col].head().tolist())
