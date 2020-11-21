import pandas as pd
import xarray as xr
from datetime import datetime, timedelta
#!/usr/bin/python
import os
import glob

# Folder where the files are
os.chdir(r"C:\Users\USUARIO\Desktop\Rafaela\Down")

# File extension
extension = 'nc'

# Get al files in folder and save in a list
all_filenames = [i for i in glob.glob('*.{}'.format(extension))]

# Loop to get all data from the files and save in a xlsx file
index = 0
while(index < len(all_filenames)):

    # Get all files name
    nc_file = all_filenames[index]
    
    # With Xarray he open ao file and takes your data
    NC = xr.open_dataset(str(all_filenames[index]),decode_times=True)

    # Open csv file with lat and lon (can get many points)
    csv = r"C:\Users\USUARIO\Desktop\Rafaela\Down\Cities.csv"
    df = pd.read_csv(csv,delimiter=',')
    Newdf = pd.DataFrame([])
   
    # grid point lists
    lat = df["Latitude"]
    lon = df["Longitude"]

    point_list = zip(lat,lon)

    for i, j in point_list:

        # Another way to get the data with a lat and lon error
        # dsloc = NC.isel(lat=int(lat),lon=int(lon))

        # Takes the mesh data closest to lat and lon
        dsloc = NC.sel(lat=i,lon=j,method='nearest')

        # Transform to DataFrame to facilitate data manipulation
        DT=dsloc.to_dataframe()
        df = DT.copy(deep=True)
        
        # Drop columns if you not use (the name and number of columns can variable of the type of file netcfd)
        df = df.drop(columns=['DSR','DQF','time_bounds','goes_lat_lon_projection','lat_image','lat_image_bounds','lon_image','lon_image_bounds','nominal_satellite_subpoint_lat','nominal_satellite_subpoint_lon',
                               'nominal_satellite_height','geospatial_lat_lon_extent','retrieval_local_zenith_angle','quantitative_local_zenith_angle','retrieval_local_zenith_angle_bounds','quantitative_local_zenith_angle_bounds','retrieval_solar_zenith_angle','quantitative_solar_zenith_angle','retrieval_solar_zenith_angle_bounds','quantitative_solar_zenith_angle_bounds','dsr_product_wavelength','dsr_product_wavelength_bounds','retrieval_pixel_count','lza_pixel_count','outlier_pixel_count','percent_uncorrectable_GRB_errors','percent_uncorrectable_L0_errors','algorithm_dynamic_input_data_container','processing_parm_version_container','algorithm_product_version_container'])

        # Remove the duplicates rows based on specific column
        df = df.drop_duplicates(subset=['t'], keep='first')

        # Save data
        Newdf=Newdf.append(df,sort=False)
    # Convert DataFrame to Excel file
    Newdf.to_excel(r'Dados'+' '+str(all_filenames[index])+r'.xlsx',index=False)
    index+=1
