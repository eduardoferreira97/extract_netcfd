import pandas as pd
import xarray as xr
from datetime import datetime, timedelta
#!/usr/bin/python
import os
import glob

# Pasta de onde vai pegar os arquivos
os.chdir(r"C:\Users\USUARIO\Desktop\Rafaela\Down")

# Extensão/tipo do arquivo que vai ser usado
extension = 'nc'

# Percorre o diretório e pega todos os nomes dos arquivos e salva em uma Lista
all_filenames = [i for i in glob.glob('*.{}'.format(extension))]

# Laço para parar apenas depois de converter todos os arquivos em xlsx
index = 0
while(index < len(all_filenames)):

    # Pegacada nome de cada arquivo e salva numa String
    nc_file = all_filenames[index]
    
    # Com Xarray ele abre cada arquivo .nc para pegar os dados
    NC = xr.open_dataset(str(all_filenames[index]),decode_times=True)

    # Abre o csv com as latitudes e longitudes, com isso é possivel pegar dados de vários pontos ao mesmo tempo
    csv = r"C:\Users\USUARIO\Desktop\Rafaela\Down\Cities.csv"
    df = pd.read_csv(csv,delimiter=',')
    Newdf = pd.DataFrame([])
   
    # grid point lists
    lat = df["Latitude"]
    lon = df["Longitude"]

    point_list = zip(lat,lon)

    for i, j in point_list:

        # Outra forma de pegar os dados a partir de latitude e longitude, mas com um erro de distancia não muito grande
        # dsloc = NC.isel(lat=int(lat),lon=int(lon))

        #Pega os valores de Latitude e Longitude mais proximos da latitude e longitude dada
        dsloc = NC.sel(lat=i,lon=j,method='nearest')

        # Transforma em DataFrame para facilicar a manipulação dos dados
        DT=dsloc.to_dataframe()
        df = DT.copy(deep=True)
        
        # Colunas que vão sair antes de salvar os dados no excel
        df = df.drop(columns=['DSR','DQF','time_bounds','goes_lat_lon_projection','lat_image','lat_image_bounds','lon_image','lon_image_bounds','nominal_satellite_subpoint_lat','nominal_satellite_subpoint_lon',
                               'nominal_satellite_height','geospatial_lat_lon_extent','retrieval_local_zenith_angle','quantitative_local_zenith_angle','retrieval_local_zenith_angle_bounds','quantitative_local_zenith_angle_bounds','retrieval_solar_zenith_angle','quantitative_solar_zenith_angle','retrieval_solar_zenith_angle_bounds','quantitative_solar_zenith_angle_bounds','dsr_product_wavelength','dsr_product_wavelength_bounds','retrieval_pixel_count','lza_pixel_count','outlier_pixel_count','percent_uncorrectable_GRB_errors','percent_uncorrectable_L0_errors','algorithm_dynamic_input_data_container','processing_parm_version_container','algorithm_product_version_container'])

        # Remove as linhas duplicadas a partir dos valores da coluna dada no subset
        df = df.drop_duplicates(subset=['t'], keep='first')

        # Salva os dados numa lista
        Newdf=Newdf.append(df,sort=False)
    
    Newdf.to_excel(r'Dados'+' '+str(all_filenames[index])+r'.xlsx',index=False)
    index+=1

os.system('python juntar_csv.py')
