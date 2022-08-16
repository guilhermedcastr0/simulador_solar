from tools import simulation, Locations
import warnings
warnings.filterwarnings("ignore")


path = './Rio do Vento/SRA/!Energia/20220713 Safira/2. SRDV/solar/'

locations = Locations(path = path)


for location in locations.SITE_NAME: 
      
      simulation(path = path, site_name = location, pvsyst_validation = True)

      