import geopandas as gpd

gdf = gpd.read_file('/Users/davidmichalove/Desktop/automate/app/shape_files/Belmont_County_Parcels.shp')
print(gdf.iloc[0])
print(list(gdf.columns))
