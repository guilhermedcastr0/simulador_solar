import re
import pandas as pd
import datetime
import pvlib
from pvlib.tools import cosd, sind
import numpy as np
import math
import logging

version = 'CVER 1.0.0'

class Locations:
    
    def __init__(self, path: str) -> list:
        
        """
                Com esta função é possível realizar a leitura dos site_names disponíveis no arquivo 
                que contém os parâmetros de base da simulação. 
                O arquivo de base deve ser do formato .xlsx e possuir a coluna: 'site_name'  
                
                A função possui somente um argumento.
                
                -------------------
                path : str - Recebe o endereço do data base file.
         """
           
        locations = pd.read_excel(path + 'cver/DataBase.xlsx')
        self.SITE_NAME = locations.site_name
              

class DataLocations:

    def __init__(self, path: str, site_name: str or int) -> None:
     
        """
                Com esta função é possível realizar a leitura do arquivo que contém os parâmetros 
                de base da simulação. 
               O arquivo de base deve ser do formato .xlsx e possuir as seguintes colunas:
                   
                   ['site_name', 'solar_series', 'pan_file', 'ond_file',
                   'LAT', 'LON', 'L', 'D', 'altitude', 'MODULES_IN_SERIES',
                   'MODULES_IN_PARALLEL', 'INVERTERS', 'LID_LOSS', 'QUALITY_LOSS',
                   'STC_OHM_LOSS', 'SOILING_LOSS', 'MISMATCH_LOSS', 'ALBEDO',
                   'GHI_MIN_THRESHOLD', 'MAX_ANGLE']   
                
                A função possui somente dois argumentos.
                
                -------------------
                path : str - Recebe o endereço da pasta raíz onde está armazenado os arquivos do solar do projeto.
                site_name : int or srt - Recebe um valor numérico, correspondente a linha do arquivo. 
                Ou o nome do site_name.  
         """
           
        locations = pd.read_excel(path + 'cver/DataBase.xlsx')
        locations.index = locations.site_name
        try:
            location = locations.loc[site_name]
        except:
            location = locations.iloc[site_name]
        
        self.SITE_NAME = location['site_name']
        self.SOLAR_SERIES_FILE = path + 'ts/' + location['solar_series_file']
        self.LAT = location['LAT']
        self.LON = location['LON']
        self.L = location['L']
        self.D = location['D']
        self.ALTITUDE = location['ALTITUDE']
        self.MODULES_IN_SERIES = location['MODULES_IN_SERIES']
        self.MODULES_IN_PARALLEL = location['MODULES_IN_PARALLEL']
        self.INVERTERS = location['INVERTERS']
        self.SOILING_LOSS = location['SOILING_LOSS'] 
        self.STC_OHM_LOSS = location['STC_OHM_LOSS']
        self.LID_LOSS = location['LID_LOSS'] 
        self.QUALITY_LOSS = location['QUALITY_LOSS']
        self.ALBEDO = location['ALBEDO']
        self.MAX_ANGLE = location['MAX_ANGLE']
        self.GHI_MIN_THRESHOLD = location['GHI_MIN_THRESHOLD']
        self.MISMATCH_LOSS = location['MISMATCH_LOSS']
        self.GCR = self.L / self.D
        self.PAN_FILE = path + 'cver/module/' + location['pan_file']
        self.OND_FILE = path + 'cver/inverter/' +location['ond_file']
        self.U_c = location['U_c']
        self.U_v = location['U_v']
        self.STC_OHM_LOSS_AC = location['STC_OHM_LOSS_AC']
        self.FUSO = location['FUSO']
        self.IRON_LOSS = location['MV_IRON_LOSS']
        self.COPPER_LOSS = location['MV_COPPER_LOSS'] 
        self.PMAX_OUT = location['PMAX_OUT']
        self.MV_LOSS_STC = location['MV_LOSS_STC'] 
        
        try:
            self.PVSYST_FILE = path + 'cver/simulation_PVSyst/' + location['pvsyst_file']
        except:
            self.PVSYST_FILE = None
     

class PVModulo:

     def __init__(self, path: str) -> None: 
         
         """
                A função realiza a leitura dos dados do arquivo .PAN. 
               
                
                A função possui somente um argumento.
                
                -------------------
                path : str - Recebe o endereço do aquivo .pan.
        """
         
        
         def _parse_tree(lines) -> None:
            
            """
            Parse an indented outline into (level, name, parent) tuples.  Each level
            of indentation is 2 spaces.
            """
            
            regex = re.compile(r'^(?P<indent>(?: {2})*)(?P<name>\S.*)')
            stack = []
            
            for line in lines:
                match = regex.match(line)
                level=0
                if not match:
                    continue #skip last line or empty lines
                stack[level:] = [match.group('name')]
                yield level, match.group('name'), (stack[level - 1] if level else None)
                
         def e_gap(pan_data: object, technol: str) -> float:
               """
               Com esta função é obter o valor de energia do gap do semicondutor
           
               -------------------
               technol - recebe a tecnologia do módulo extraída do arquivo .PAN. 
           
               """  
               technol_module = pan_data.loc[technol, 'value']
               e_gap = {'mtSiPoly':1.12,
                     'mtCdTe': 1.5,
                     'mtSiMono':1.12}
            
               return e_gap[technol_module]
           
         def value(pan_data: object, parameter: str) -> float:
                
            return float(pan_data.loc[parameter, 'value'])
             

            
         logging.addLevelName(5,"VERBOSE")
         logging.basicConfig(level=logging.INFO)
    
         with open(path, mode='r', encoding='utf-8-sig') as file:
                    raw = file.read()
                    if raw[:3] == "ï»¿": # this is utf-8-BOM
                        raw = raw[3:] #remove BOM
             
         lines = (raw.split('\n'))
    
         pan_data_tuple=[]
    
         for index, name, parent in _parse_tree(lines):
            try:
                if len(tuple(re.split('=',name))) == 3:
                    pan_data_tuple.append(tuple(re.split('=',name))[0]+' '+
                                          tuple(re.split('=',name))[1], tuple(re.split('=',name))[2])
                else:
                    pan_data_tuple.append(tuple(re.split('=',name)))
            except:
                continue
    
         pan_data = pd.DataFrame(pan_data_tuple, columns=['key','value'])
         pan_data.index = pan_data.key
         pan_data.index = pan_data['key']
         pan_data = pan_data[['value']]
        
         
         self.parameters = {'alpha_sc': value(pan_data,'muISC') / 1000, #muISC - Coeficiente de temperatura da corrente de curto circuito
                            'gamma_ref': value(pan_data,'Gamma'), #Gamma - Fator de idealidade do Diodo
                            'mu_gamma': value(pan_data,'muGamma'), #muGamma - Coeficiente de temperatura para o fator de idealidade do diodo
                            'R_sh_ref': value(pan_data,'RShunt'), #RShunt - Resistência shunt em condições de referência
                            'R_sh_0': value(pan_data,'Rp_0'), #Rp_0 - Resistência shunt em condições de irradiância zero
                            'R_s': value(pan_data,'RSerie'), #RSerie - Resistência em série nas condições de referência
                            'cells_in_series': value(pan_data,'NCelS'), #NCelS - Número de células em série
                            'cells_in_parallel': value(pan_data,'NCelP'), #NCelP - Número de células em paralelo
                            'R_sh_exp':value(pan_data,'Rp_Exp'), #Rp_Exp - Expoente da equação para resistência de derivação (Padrão 5.5)
                            'EgRef': e_gap(pan_data,'Technol'), #Bandgap de energia na temperatura de referência
                            'irrad_ref': value(pan_data,'GRef'), #GRef - Irradiância de referência
                            'temp_ref': value(pan_data,'TRef') + 273.15 , #TRef - Temperatura de referência [K]
                            'cell_area': value(pan_data,'CellArea') / 10000,# CellArea
                            'nominal_power': value(pan_data,'Imp') * value(pan_data,'Vmp'), #PNom = Imp * Vmp
                            'Imp': value(pan_data,'Imp'),
                            'Vmp': value(pan_data,'Vmp'),
                            'Isc_ref': value(pan_data,'Isc'), #Isc
                            'Voc_ref': value(pan_data,'Voc'), #Voc
                            'Width': value(pan_data,'Width'),
                            'Height':  value(pan_data,'Height'),
                            'surface': value(pan_data,'Width')*value(pan_data,'Height')} 

        
         curva = ['Point_1', 'Point_2', 'Point_3', 'Point_4', 'Point_5', 'Point_6', 
                  'Point_7', 'Point_8', 'Point_9']
         self.iam_curve = pan_data[pan_data.index.isin(curva)].reset_index(drop=True)[:9]
         self.iam_curve['Angle'] = [float(re.split(',',poits)[0] )for poits in self.iam_curve['value']]
         self.iam_curve['FIAM'] = [float(re.split(',',poits)[1]) for poits in self.iam_curve['value']]
         self.iam_curve.drop(columns='value', inplace=True)
         

class Inverter:
    
    
    def __init__(self, path: str) -> None:
        
        """
                A função realiza a leitura dos dados do arquivo .OND. 
               
                
                A função possui somente um argumento.
                
                -------------------
                path : str - Recebe o endereço do aquivo .ond.
        """
        
        def value(ond_data: object, parameter:str) -> float:
            
            if parameter == 'MonoTri':
                if ond_data.loc[parameter, 'value'] == 'Tri':
                    return 3
                elif ond_data.loc[parameter, 'value'] == 'Mono':
                    return 1
            else:   
                return float(ond_data.loc[parameter, 'value'])
           
            #Read Ond File
        
        def _parse_tree(lines: list) -> None:
            """
            Parse an indented outline into (level, name, parent) tuples.  Each level
            of indentation is 2 spaces.
            """
            regex = re.compile(r'^(?P<indent>(?: {2})*)(?P<name>\S.*)')
            stack = []
            for line in lines:
                match = regex.match(line)
                level=0
                if not match:
                    continue #skip last line or empty lines
                stack[level:] = [match.group('name')]
                yield level, match.group('name'), (stack[level - 1] if level else None)
                
        """
        Com esta função é possível extrair um dataframe contendo as curvas P_dc/P_ac 
        extraídas do arquivo .OND 
    
        -------------------
        path - recebe o diretório onde se encontra o arquivo .ond.
    
        """    
        logging.addLevelName(5,"VERBOSE")
        logging.basicConfig(level=logging.INFO)
    
        try:
            with open(path, mode='r', encoding='utf-8-sig') as file:
                    raw = file.read()
                    if raw[:3] == "ï»¿": # this is utf-8-BOM
                        raw = raw[3:] #remove BOM
        except:   
            
            try: 
                with open(path, mode='r') as file:
                        raw = file.read()
                        if raw[:3] == "ï»¿": # this is utf-8-BOM
                            raw = raw[3:] #remove BOM
            except:
                return 'Erro - O arquivo está cofidicado de forma inadequada'
    
        lines = (raw.split('\n'))
    
        ond_data_tuple=[]
    
        for index, name, parent in _parse_tree(lines):
            try:
                if len(tuple(re.split('=',name))) == 3:
                    ond_data_tuple.append(tuple(re.split('=',name))[0]+' '+
                                          tuple(re.split('=',name))[1], tuple(re.split('=',name))[2])
                else:
                    ond_data_tuple.append(tuple(re.split('=',name)))
            except:
                continue
    
        ond_data = pd.DataFrame(ond_data_tuple, columns=['key','value'])
        ond_data.index = ond_data.key
        
        def ond_read_curves(ond_data: object) -> object:
            """
            Com esta função é possível extrair um dataframe contendo as curvas P_dc/P_ac 
            extraídas do dataframe com os dados do .OND. 
        
            -------------------
            ond_data - recebe um dataframe contendo os dados do arquivo .ond.
        
            """
            curva = ['Point_1', 'Point_2', 'Point_3', 'Point_4', 'Point_5', 'Point_6', 
                     'Point_7', 'Point_8', 'Point_9', 'Point_10', 'Point_11']
            ond_data['key'] = ond_data.index
            ond_curves = ond_data[ond_data['key'].isin(curva)].reset_index(drop=True)
            ond_curves['P_dc'] = [float(re.split(',',poits)[0] )for poits in ond_curves['value']]
            ond_curves['P_ac'] = [float(re.split(',',poits)[1]) for poits in ond_curves['value']]
            ond_curves['eff']=ond_curves['P_ac']/ond_curves['P_dc']
        
            
            if ond_data[ond_data['key']=='Point_1'].shape[0] > 3:
                ond_curves['input_voltage']=['' if int(index)<11
                                    else 'Vmin' if ((int(index)>=11)&(int(index)<22))
                                    else 'Vnom' if ((int(index)>=22)&(int(index)<33))
                                    else 'Vmax' 
                                    for index, _ in ond_curves.iterrows()]
            else:
                ond_curves['input_voltage']=['Vmin' if int(index)<11
                                    else 'Vnom' if ((int(index)>=11)&(int(index)<22))
                                    else 'Vmax' 
                                    for index, _ in ond_curves.iterrows()]
            
            ond_curves['VNomEff']=[0 if int(index)<11
                                else float(re.split(',',ond_data.loc['VNomEff']['value'])[0]) 
                                if ((int(index)>=11)&(int(index)<22))
                                else float(re.split(',',ond_data.loc['VNomEff']['value'])[1]) 
                                if ((int(index)>=22)&(int(index)<33))
                                else float(re.split(',',ond_data.loc['VNomEff']['value'])[2]) 
                                for index, _ in ond_curves.iterrows()]
        
            ond_curves.dropna(inplace=True)
            ond_curves.drop(columns=['key', 'value'], inplace=True)
        
            return ond_curves 
        
        
        self.parameters =  {'PNomConv': value(ond_data,'PNomConv') , #PNomConv - Potência do inversor dada em W
                            'EfficMax': value(ond_data,'EfficMax'), #EfficMax - Eficiência máxima do inversor
                            'PMaxOUT': value(ond_data,'PMaxOUT'), #PMaxOUT - Potência máxima de entrada AC
                            'TPNom': value(ond_data,'TPNom'), #Temperatura de operação para Pnon
                            'TPMax': value(ond_data,'TPMax'), #Temperatura de operação para Pmax
                           'VOutConv': value(ond_data,'VOutConv'), #Tensão de saída do conversor
                           'MonoTri': value(ond_data,'MonoTri')} #Número de fases do inversor
        
        self.curve = ond_read_curves(ond_data)
                            

class Simulation:  
    
    def __init__(self, location:object, modulo:object, inversor:object):
            
        solar_series = pd.read_csv(location.SOLAR_SERIES_FILE)
        solar_series['date'] = pd.to_datetime(solar_series['time'], dayfirst=True)
        solar_series['date'] = \
                    pd.to_datetime(solar_series['date']).dt.tz_localize('UTC').dt.tz_convert('Etc/GMT+'+str(location.FUSO))
        solar_series['date'] = solar_series['date']  + datetime.timedelta(hours=int(location.FUSO))
        t_shift = solar_series['date'] + datetime.timedelta(minutes=30)
        solar_series.index =   t_shift
        self.simulation_output = pd.DataFrame(index=t_shift)
        self.simulation_output['date'] = solar_series['date']

        ################# Angles #################
        def pvlib_elevation_correction(apparent_elevation_pvlib:float):
            """
            Com esta função é possível realizar o ajuste dos valores de elevação aparente
            do sol obtidos pelo PVLib com o SPA do NREL, de modo que após o ajuste a curva 
            de correlação com os dados do PVSyst possua valor de R² igual a 1.
            
            A função possui somente um argumento [apparent_elevation_pvlib]
            
            -------------------
            apparent_elevation_pvlib - Recebe os valores de elevação solar aparente calculados pelo PVLib.
            
            """
            
            if apparent_elevation_pvlib >= 7:
                return apparent_elevation_pvlib
            elif -7 < apparent_elevation_pvlib and apparent_elevation_pvlib < 7:
                return 0.5*apparent_elevation_pvlib + 3.5
            else:
                return 0
     
        #apparent_elevation 
        pvlib_solar_position = pvlib.solarposition.get_solarposition(time=t_shift, 
                                                                 latitude=location.LAT, 
                                                                 longitude=location.LON)
     
        self.simulation_output['HSol'] =  (pvlib_solar_position['apparent_elevation'].
                                           apply(pvlib_elevation_correction))
        
        #azimuth
        self.simulation_output['AzSol'] =  pvlib_solar_position['azimuth']
    
        #zenith
        ZSol =  90 - self.simulation_output['HSol']
    
    
        #Tracker position
        tracker = pvlib.tracking.singleaxis(apparent_zenith=ZSol,  
                                        apparent_azimuth=self.simulation_output['AzSol'],
                                        axis_tilt=0,
                                        axis_azimuth=0,
                                        max_angle=location.MAX_ANGLE,
                                        backtrack=True,
                                        gcr=location.GCR)
    
        self.simulation_output['AngInc'] =  tracker['aoi']
        self.simulation_output['PhiAng'] =  tracker['tracker_theta']
    
    
        ################# MetData #################

        self.simulation_output['GlobHor'] =  solar_series['GHI']
        self.simulation_output['DiffHor'] =  solar_series['DIF']
          
        """Note que o PVsyst não permite fornecer as três componentes ao mesmo tempo, apenas duas. O time de projetos solares da
            Casa dos Ventos optou por usar a irradiação global horizontal (GHI) e a componente difusa (DIF), por entender que são
            dados mais robustos. No início e no final do dia, o cálculo da irradiação direta envolve a divisão por um número pequeno,
            o que pode levar (e leva) a imprecisões de origem numérica."""
            
        self.simulation_output['BeamHor'] =  (self.simulation_output['GlobHor']- self.simulation_output['DiffHor']) /\
                                                  cosd(ZSol)
                                                  
        
        self.simulation_output['T_Amb'] =  solar_series['TEMP']
        self.simulation_output['WindVel'] =  solar_series['WS']
    
        ################# Transpo #################

            
        # Irradiação direta no plano inclinado
        self.simulation_output['BeamInc'] = pvlib.irradiance.beam_component(surface_tilt=tracker['surface_tilt'],
                                                surface_azimuth=tracker['surface_azimuth'],
                                                solar_zenith=ZSol,
                                                solar_azimuth=self.simulation_output['AzSol'],
                                                dni=self.simulation_output['BeamHor'])
            
        # Irradiação extraterrestre (ou topo da atmosfera, ToA) numa superfície normal ao sol
        dni_extra=pvlib.irradiance.get_extra_radiation(datetime_or_doy=self.simulation_output.index)
        
        
        # Massa de ar absoluta (ajustada à pressão) 
        airmass = pvlib.atmosphere.get_absolute_airmass(pvlib.atmosphere.get_relative_airmass(zenith=ZSol))
        
        # Irradiação difusa no plano inclinado, modelo de Perez-Ineichen.
        self.simulation_output['DifSInc'] = pvlib.irradiance.perez(surface_tilt=tracker['surface_tilt'],
                                                                  surface_azimuth=tracker['surface_azimuth'],
                                                                  dhi=self.simulation_output['DiffHor'],
                                                                  dni=self.simulation_output['BeamHor'],
                                                                  dni_extra= dni_extra,
                                                                  solar_zenith=ZSol,
                                                                  solar_azimuth=self.simulation_output['AzSol'],
                                                                  airmass=airmass)
        
        self.simulation_output['DifSInc'].fillna(0, inplace=True)
        
        # Irradiação refletida do solo no plano inclinado
        self.simulation_output['Alb_Inc'] = pvlib.irradiance.get_ground_diffuse(tracker['surface_tilt'],
                                           self.simulation_output['GlobHor'],
                                           albedo=location.ALBEDO)
        
        # Irradiação global no plano inclinado
        self.simulation_output['GlobInc'] = (self.simulation_output['BeamInc'] + 
                                             self.simulation_output['DifSInc'] + 
                                             self.simulation_output['Alb_Inc'])

             
        ################# IncColl #################
    
        
        from scipy import interpolate
        
        # Vetor da direção do Sol (x,y,z)
        sun_x = cosd(self.simulation_output['HSol']) * \
            sind(self.simulation_output['AzSol'])
        
        sun_z = sind(self.simulation_output['HSol'])
        
        # Os cálculos subsequentes serão feitos em radianos
        theta = np.deg2rad(self.simulation_output['PhiAng'])
        
        # É obtido o ângulo do sol relativo ao plano xz
        psi = np.arctan2(sun_x, sun_z)
        
       #Near Shading Loss
        
       #Near shading bean loss
        # beam_loss_factor = perda / total 
        beam_loss_factor = (np.tan(psi) * np.tan(theta) + 1 - 1/(location.GCR * np.cos(theta))) /\
            (np.tan(psi) * np.tan(theta) + 1)
            
        # Aplica 0 nos valores negativos
        beam_loss_factor = beam_loss_factor.apply(lambda x: max(x, 0))

        #Obtém a perdas na componente direta fazendo a ponderação da componente direta pelo fator de perda
        self.simulation_output['ShdBLss'] = self.simulation_output['BeamInc'] * beam_loss_factor
        
        #Near shading diffuse loss
        
        masking_angle = pvlib.shading.masking_angle(surface_tilt = tracker['surface_tilt'], 
                                          gcr = location.GCR, 
                                          slant_height = 0)
        
        
        shading_loss_factor = pvlib.shading.sky_diffuse_passias(masking_angle)
        
        self.simulation_output['ShdDLss'] = self.simulation_output['DiffHor']*shading_loss_factor
        

        
        #Near shadings albedo loss
        C_albedo = 1.15

        self.simulation_output['ShdALss'] = (C_albedo*location.GCR*self.simulation_output['Alb_Inc']*
                                             (math.pi - abs(theta)))
        
        
        #Near shadings loss
        self.simulation_output['ShdLoss'] = (self.simulation_output['ShdBLss']  + 
                                             self.simulation_output['ShdDLss']  + 
                                             self.simulation_output['ShdALss']) 
        
        beam_after_shading = self.simulation_output['BeamInc']  - self.simulation_output['ShdBLss']

        diff_after_shading = self.simulation_output['DifSInc'] -  self.simulation_output['ShdDLss']
        
        albedo_after_shading = self.simulation_output['Alb_Inc'] - self.simulation_output['ShdALss']
        
        #Global corrected for shadings
        self.simulation_output['GlobShd'] = beam_after_shading + diff_after_shading + albedo_after_shading
        
        
        #IAM Loss
        
        tck = interpolate.interp1d(modulo.iam_curve['Angle'], modulo.iam_curve['FIAM'],
                           fill_value='extrapolate')

        iam_loss_factor_beam = tck(self.simulation_output['AngInc'])
        
        # Incidence beam loss
        
        beam_after_iam = beam_after_shading * iam_loss_factor_beam

        diff_after_iam = diff_after_shading

        albedo_after_iam = albedo_after_shading
        
        #Global corrected for IAM
        self.simulation_output['GlobIAM'] = beam_after_iam + diff_after_iam + albedo_after_iam
        
        # Soiling Loss
        
        self.simulation_output['SlgLoss'] = self.simulation_output['GlobIAM'] * location.SOILING_LOSS
        
        # Global corrected for Soiling
        self.simulation_output['GlobSlg'] = (self.simulation_output['GlobIAM'] - 
                                             self.simulation_output['SlgLoss'])
        
        # Global effective
        
        self.simulation_output['GlobEff'] = self.simulation_output['GlobSlg'] 
    
        ################# Array e Inverter #################
        
        from scipy import constants
        
        # A eficiência na condição de operação padrão é calculada como sendo a razão da 
        #potência nominal pela potência disponível na área do painel. 
        stc_efficiency = \
            modulo.parameters['nominal_power'] /\
            (modulo.parameters['surface'] * modulo.parameters['irrad_ref'])
            

        # Tensão térmica
        #Vt = k * T / q
        Vt = constants.Boltzmann * modulo.parameters['temp_ref'] / constants.e 
        
        # Cálculo da corrente de saturação do diodo a partir das condições (0, V_oc) e (I_sc, 0)
        #I_0 = (Isc  - (Voc - Isc * Rs) / Rsh) * math.exp(-Voc / (ns * ni * Vt))
        I_0 = ((modulo.parameters['Isc_ref']  - (modulo.parameters['Voc_ref'] - 
                modulo.parameters['Isc_ref'] * modulo.parameters['R_s']) / 
                modulo.parameters['R_sh_ref']) * math.exp(-modulo.parameters['Voc_ref'] / 
                (modulo.parameters['cells_in_series'] * modulo.parameters['gamma_ref'] * Vt))) 
                                                         
        #I_ph = I_0 * math.exp(Voc / (ni * ns * Vt)) + Voc / Rsh
        #Cálculo da fotocorrente pela equação obtida em (0, V_oc)
        I_ph = (I_0 * math.exp(modulo.parameters['Voc_ref'] / (modulo.parameters['gamma_ref'] * 
            modulo.parameters['cells_in_series'] * Vt)) + modulo.parameters['Voc_ref'] 
                / modulo.parameters['R_sh_ref']) 
        
        # Potência Nominal
        self.simulation_output['EArrNom'] = self.simulation_output['GlobEff'] *stc_efficiency*\
                    modulo.parameters['surface'] *\
                    location.MODULES_IN_SERIES *\
                    location.MODULES_IN_PARALLEL

        self.simulation_output['TArray'] = pvlib.temperature.pvsyst_cell( poa_global=self.simulation_output['GlobEff'], 
                                            temp_air=self.simulation_output['T_Amb'], 
                                            wind_speed=self.simulation_output['WindVel'] , 
                                            u_c = location.U_c, 
                                            u_v = location.U_v, 
                                            eta_m = stc_efficiency, 
                                            alpha_absorption = 0.9)
      
        
                            
        pvsyst_params = pvlib.pvsystem.calcparams_pvsyst(effective_irradiance = self.simulation_output['GlobEff']*(1-location.LID_LOSS-location.QUALITY_LOSS), #Irradiância que é convertida em fotocorrente
                                                 temp_cell = self.simulation_output['TArray'], #temperatura média das células
                                                 alpha_sc = modulo.parameters['alpha_sc'], #Coeficiente de temperatura da corrente de curto circuito
                                                 gamma_ref = modulo.parameters['gamma_ref'], #Fator de idealidade do diodo
                                                 mu_gamma = modulo.parameters['mu_gamma'], #Coeficiente de temperatura para o fator de idealidade do diodo 
                                                 I_L_ref = I_ph, #Fotocorrente nas condições de referência
                                                 I_o_ref = I_0, #Corrente de saturação reversa nas condições de referência
                                                 R_sh_ref = modulo.parameters['R_sh_ref'], #Resistência shunt em condições de referência
                                                 R_sh_0 = modulo.parameters['R_sh_0'], #Resistência shunt em condições de irradiância zero
                                                 R_s = modulo.parameters['R_s'], #Resistência série
                                                 cells_in_series = modulo.parameters['cells_in_series'], #Número de células em série
                                                 R_sh_exp=5.5, #The exponent in the equation for shunt resistance
                                                 EgRef = modulo.parameters['EgRef'], #Bandgap de engergia na temperatura de referência
                                                 irrad_ref = modulo.parameters['irrad_ref'], #Irradiância de referência
                                                 temp_ref = modulo.parameters['temp_ref']- 273.15) #Temperatura da célula de referência em Célsius

        photocurrent = pvsyst_params[0]
        
        saturation_current = pvsyst_params[1]
        
        resistance_series = pd.Series(np.empty(len(photocurrent)))
        
        resistance_series = resistance_series.apply(lambda x: pvsyst_params[2])
        
        resistance_series.index = photocurrent.index
        
        resistance_shunt = pvsyst_params[3]
        
        nNsVth = pvsyst_params[4]
        
        single_diode_params = pd.DataFrame({'photocurrent': photocurrent, 
                                            'saturation_current': saturation_current,
                                            'resistance_series': resistance_series,
                                            'resistance_shunt': resistance_shunt,
                                            'nNsVth': nNsVth})
        
        single_diode = pvlib.pvsystem.singlediode(single_diode_params.photocurrent,
                                                  single_diode_params.
                                                  saturation_current,
                                                  single_diode_params.
                                                  resistance_series,
                                                  single_diode_params.resistance_shunt,
                                                  single_diode_params.nNsVth,
                                                  ivcurve_pnts=None, 
                                                  method='lambertw')
        
        scaled_value = pvlib.pvsystem.scale_voltage_current_power(single_diode,
                                                                  location.MODULES_IN_SERIES,
                                                                  location.MODULES_IN_PARALLEL)
                            
        
        
        # Ohmic wiring loss
        
        # Stc parameter
        single_diode_stc = pvlib.pvsystem.singlediode(I_ph, I_0, modulo.parameters['R_s'], modulo.parameters['R_sh_ref'],
                                              modulo.parameters['gamma_ref'] * modulo.parameters['cells_in_series'] * Vt)


        single_diode_stc = pd.DataFrame({'v_mp': single_diode_stc['v_mp'],
                                         'v_oc': single_diode_stc['v_oc'],
                                         'i_mp': single_diode_stc['i_mp'],
                                         'i_x': single_diode_stc['i_x'],
                                         'i_xx': single_diode_stc['i_xx'],
                                         'i_sc': single_diode_stc['i_sc'],
                                         'p_mp': single_diode_stc['p_mp']},
                                        index=[0])
        
        #Os parâmetros de operação são escalados para que sejam obtidos os parâmetros de todo o sistema
        scaled_value_stc = pvlib.pvsystem.scale_voltage_current_power(single_diode_stc,
                                                              location.MODULES_IN_SERIES,
                                                              location.MODULES_IN_PARALLEL)

        #Cálculo do R_dc equivalente para as condições de operação padrão
        R_equiv_dc = location.STC_OHM_LOSS * scaled_value_stc['p_mp'] /\
            (scaled_value_stc['i_mp'] ** 2)
        
        self.R_equiv_dc = R_equiv_dc[0]
        
        
        
        #Cálculo da perda ôhmica aplicando a resistência equivalente às condições normais de operação do sistema
        self.simulation_output['OhmLoss']  = self.R_equiv_dc * (scaled_value['i_mp'] ** 2)
        self.simulation_output['OhmLoss'] = [0 if row['GlobEff']==0 else row['OhmLoss'] for index, row in self.simulation_output.iterrows()]
       
        #EM AVALIAÇÃO
        self.simulation_output['FOhmLoss'] = self.simulation_output['OhmLoss']/scaled_value['p_mp'] 
        
        
        # MisLoss
        
        self.simulation_output['MisLoss'] = (scaled_value['p_mp']) * location.MISMATCH_LOSS
        self.simulation_output['MisLoss'] = [0 if row['GlobEff']==0 else row['MisLoss'] for index, row in self.simulation_output.iterrows()]
        
        
        # Array virtual energy at MPP                    
        
        self.simulation_output['EArrMPP']  = scaled_value['p_mp'] - self.simulation_output['OhmLoss'] - self.simulation_output['MisLoss']
        self.simulation_output['EArrMPP'] = [0 if row['GlobEff']==0 else row['EArrMPP'] for index, row in self.simulation_output.iterrows()]
        
        
        
        
        # Modelo do inversor
        
        #Curva de eficiência do inversor 
        
        efficiency_curve = inversor.curve[inversor.curve['input_voltage']=='Vnom']
        
        eff = interpolate.interp1d(efficiency_curve['P_dc'].astype(float), 
                                   efficiency_curve['eff'].astype(float),
                                   fill_value='extrapolate')
        
        self.simulation_output['eff_inverter'] = eff(self.simulation_output['EArrMPP']/location.INVERTERS)
        
        self.simulation_output['eff_inverter'] = self.simulation_output['eff_inverter'].apply(lambda x: max(0, x))
        
        #Potência AC máxima de entrada dos Inversores
        if location.PMAX_OUT == 0:
            PMaxOUT = inversor.parameters['PMaxOUT']*1000*location.INVERTERS
        else:
            PMaxOUT = location.PMAX_OUT*1000*location.INVERTERS
        
        
        #Potência DC máxima de entrada dos Inversores com base na curva de eficiência 
        self.simulation_output['PMaxIN']=(PMaxOUT/self.simulation_output['eff_inverter'])
        
        #Corficientes da curva de correção térmica do Inversor
        a = (inversor.parameters['PMaxOUT'] - inversor.parameters['PNomConv'])/(inversor.parameters['TPMax'] - inversor.parameters['TPNom'])
        b = (inversor.parameters['PMaxOUT'] - inversor.parameters['TPMax']*a)
        
       
        
        #Corrige o valor da potência DC máxima de entrada em função da curva potência x temperatura
        self.simulation_output['PMaxIN'] = [row['PMaxIN'] 
                                            if (row['T_Amb']<=inversor.parameters['TPMax'])|(row['GlobEff']<location.GHI_MIN_THRESHOLD)  
                                            else (((((a)*row['T_Amb'])+b)/row['eff_inverter'])*1000*location.INVERTERS) 
                                            for index, row in self.simulation_output.iterrows()]
        
        
        #Calcula a potência AC de saída do inversor
        self.simulation_output['EOutInv'] = pvlib.inverter.pvwatts(pdc=self.simulation_output['EArrMPP'],
                                                                    pdc0= self.simulation_output['PMaxIN'],
                                                                    eta_inv_nom=self.simulation_output['eff_inverter'])
        
        
        self.simulation_output['EOutInv'].fillna(0, inplace=True)
        
        # Aplicação do Clipping na Pmpp
        self.simulation_output['EArray'] = [row['EArrMPP'] if row['PMaxIN']>row['EArrMPP'] else row['PMaxIN'] for index, row in self.simulation_output.iterrows()]
        
        self.simulation_output['UArray'] = scaled_value['v_mp']
        
        self.simulation_output['IArray'] = scaled_value['i_mp']
        
        self.simulation_output['VocArray'] = scaled_value['v_oc']
        
        self.simulation_output['resistance_shunt'] = resistance_shunt
        
        self.simulation_output['saturation_current'] = saturation_current
        
        self.simulation_output['photocurrent'] = photocurrent
        
        self.simulation_output['resistance_series'] = resistance_series
        
        self.simulation_output['nNsVth'] = nNsVth

        v_clippado=[]
        i_clippado=[]
        
        for index, row in self.simulation_output.iterrows():
            if row['EArrMPP']>row['PMaxIN']:
                v_calculated = 0
                i_calculated = 0
                diff_min = row['EArray']
                v = row['UArray']
                cont=0 
                
                #incrementa a tensão em passos de 10 V
                for v in np.arange(row['UArray'], row['VocArray'], 10):   
        
                    #Calcula a corrente para a tensão especificada
                    i = pvlib.pvsystem.i_from_v(row.resistance_shunt, 
                                                  row.resistance_series, 
                                                  row.nNsVth, 
                                                  v/location.MODULES_IN_SERIES, 
                                                  row.saturation_current,
                                                  row.photocurrent, 
                                                  method='lambertw')*location.MODULES_IN_PARALLEL
                    
                    
                    #Aplica a perda ôhmica e de mismatch no valor de corrente
                    i = i *(1 - location.MISMATCH_LOSS - row['FOhmLoss'])
                    
                    #Calcula a diferença absoluta entre a 
                    diff = abs(i*v - row['EArray'])
                    
                    #Verifica se a diferença é inferior à diferença mínima 
                    if diff < diff_min:
                        diff_min = diff
                        v_calculated = v 
                        i_calculated = i
                        
                    else:
                        cont+=1
                        if cont>5:
                            break  
                        
                v_clippado.append(v_calculated) 
                i_clippado.append(i_calculated) 
                
            else:
                v_clippado.append(row['UArray']) 
                i_clippado.append(row['IArray'])
        
        
        self.simulation_output['UArray'] = v_clippado
        self.simulation_output['UArray'] = [0 if row['GlobEff']==0 else row['UArray']
                                            *(1 - location.MISMATCH_LOSS - row['FOhmLoss']) 
                                            for index, row in self.simulation_output.iterrows()]
        
        self.simulation_output['IArray'] = self.simulation_output['EArray']/self.simulation_output['UArray']
        self.simulation_output['IArray'].fillna(0, inplace=True)
        self.simulation_output['IArray'] = [0 if row['GlobEff']==0 else row['IArray'] for index, row in self.simulation_output.iterrows()]
        
        #Perdas ôhmicas AC
        P_AC_STC = (modulo.parameters['nominal_power']*location.MODULES_IN_SERIES*location.MODULES_IN_PARALLEL)*\
            (inversor.parameters['EfficMax']/100)

        Phase = inversor.parameters['MonoTri']

        VOutConv = inversor.parameters['VOutConv']*(Phase**0.5)
        
        Rac = (location.STC_OHM_LOSS_AC*VOutConv)/\
            ((P_AC_STC)/(VOutConv))

        IoutConv = self.simulation_output['EOutInv']/(VOutConv)
        
        self.simulation_output['EACOhmL'] = ((Phase**0.5) * ((IoutConv)**2)*Rac)/3

        # Perdas no Medium Voltage Transformer

        Res_EMVTrfL = (P_AC_STC*location.COPPER_LOSS/((P_AC_STC/VOutConv)**2))

        self.simulation_output['EMVTrfL'] = (IoutConv**(2))*(Res_EMVTrfL) + P_AC_STC*location.IRON_LOSS

        Res_EMVOhmL = (P_AC_STC*location.MV_LOSS_STC/((P_AC_STC/VOutConv)**2))

        self.simulation_output['EMVOhmL'] = (IoutConv**(2))*(Res_EMVOhmL)

        
        self.simulation_output['E_Grid'] =  self.simulation_output['EOutInv'] - self.simulation_output['EACOhmL'] \
            - self.simulation_output['EMVTrfL'] - self.simulation_output['EMVOhmL']
        
        self.simulation_output.drop(columns=['PMaxIN', 'nNsVth', 'resistance_series', 'photocurrent', 'saturation_current', 
                                                   'resistance_shunt', 'VocArray', 'eff_inverter', 'FOhmLoss'], inplace=True)
        
                       
class MetricsComplete:
    
    def __init__(self, location: object, output_simulation: object):
        
        """
                A função realiza o cálculo das métricas de desempenho do simulador.  
               
                
                A função possui dois argumentos.
                
                -------------------
                location : object - Recebe objeto onde estão armazenados os parâmetros base da simulação.
                output_simulation: object - Recebe o dataframe de saída da simulação.
        """
        
        from sklearn.linear_model import LinearRegression
        from sklearn.metrics import mean_squared_error
        
        def r2_calculation(df_x: object, df_y: object, parameter: str, ghi_min_threshold: int = location.GHI_MIN_THRESHOLD):
    
            """
            Com esta função é possível realizar o cálculo do valor do R² de um modelo linear obtido a partir da variável alvo y
            e da variável preditora x.
            A função possui três argumentos [df_x, df_y, parameter]
            
            -------------------
            df_x : dataframe - Recebe o dataframe com os dados do PVLib
            df_y : dataframe - Recebe o dataframe com os dados do PVSyst
            x    : str       - Recebe o parâmetro que será avaliado
            
            """

            df = pd.DataFrame({'pvlib': df_x[parameter],
                               'pvsyst': df_y[parameter],
                               'ghi': df_y['GlobEff']})
            df = df[df['ghi']>ghi_min_threshold]
            df.dropna(inplace=True)
            modelo = LinearRegression()
            modelo.fit(df[['pvlib']],df[['pvsyst']])
            return round(modelo.score(df[['pvlib']],df[['pvsyst']]), 4)

        def rsme_calculation(df_x: object, df_y: object, parameter: str, ghi_min_threshold: int = location.GHI_MIN_THRESHOLD):
    
            """
            Com esta função é possível realizar o cálculo do valor do RMSE percentual de um modelo linear obtido a partir da variável alvo y
            e da variável preditora x.
            A função possui três argumentos [df_x, df_y, parameter]
            
            -------------------
            df_x : dataframe - Recebe o dataframe com os dados do PVLib
            df_y : dataframe - Recebe o dataframe com os dados do PVSyst
            x    : str       - Recebe o parâmetro que será avaliado
            
            """
            df = pd.DataFrame({'pvlib': df_x[parameter],
                               'pvsyst': df_y[parameter],
                               'ghi': df_y['GlobEff']})
            df = df[df['ghi']>ghi_min_threshold]
            df.dropna(inplace=True)
            mse_value = mean_squared_error(df[['pvlib']],df[['pvsyst']])
            rmse = np.sqrt(mse_value)
            rmse_percentual = rmse/np.mean(df['pvlib'])*100
            return round(rmse_percentual, 4)
        
        def diff_ratio(df_x: object, df_y: object, parameter: str, ghi_min_threshold: int = location.GHI_MIN_THRESHOLD):

            """
            Com esta função é possível realizar o cálculo da diferença percentual de um modelo linear obtido a partir da variável alvo y
            e da variável preditora x.
            A função possui três argumentos [df_x, df_y, parameter]
        
            -------------------
            df_x : dataframe - Recebe o dataframe com os dados do PVLib
            df_y : dataframe - Recebe o dataframe com os dados do PVSyst
            x    : str       - Recebe o parâmetro que será avaliado
        
            """
            df = pd.DataFrame({'pvlib': df_x[parameter],
                               'pvsyst': df_y[parameter],
                               'ghi': df_y['GlobEff']})
            df = df[df['ghi']>ghi_min_threshold]
            df.dropna(inplace=True)
                    
            diff_per_cent = ((df['pvlib'].sum()-df['pvsyst'].sum())/df['pvsyst'].sum())*100
            return round(diff_per_cent, 4)        

        # Os azimutes do PVsyst seguem uma convenção diferente do pvlib
        def convert_pvsyst_azimuth_to_pvlib(x:float):
        
            if x == 0:
        
                return x
        
            elif x > 0:
        
                return 360.0 - x
        
            else:
        
                return -x        
        
        if location.PVSYST_FILE is None:
            
            return ('Arquivo pvsyst inválido. Por favor, verifique o diretório indicado na base de dados.')
        
        else: 
            pvsyst_data = pd.read_csv(location.PVSYST_FILE,
                              sep=';', skiprows=10, encoding='ISO-8859-1')
            
            # Removendo linha com unidades
            pvsyst_data = pvsyst_data.loc[1:, ]
            
            # Conversão dos dados. dayfirst indica que no formato da data o dia vem primeiro
            pvsyst_data['date'] = pd.to_datetime(pvsyst_data['date'], dayfirst=True)
            # Aplica a informação de fuso horário a data
            pvsyst_data.date = \
              pvsyst_data.date.dt.tz_localize('UTC').dt.tz_convert('Etc/GMT+'+str(location.FUSO)) 
            #Soma as três horas perdidas no processo de aplicação do d
            pvsyst_data.date = pvsyst_data.date + datetime.timedelta(hours=int(location.FUSO)) 
            
            # Lista as colunas do dataframe
            columns = list(pvsyst_data.columns) 
            #Remove a coluna de data
            columns = columns[1:] 
            
            
            for column in columns:
                #Converte todos os dados das demais colunas em dados numéricos
                pvsyst_data[column] = pd.to_numeric(pvsyst_data[column]) 
            
            # Os valores de cada timestamp correspondem aos centros dos intervalos
            # horários.
            t = pvsyst_data['date']
            
            # Para evitar problemas mais à frente
            pvsyst_data.index = list(t + datetime.timedelta(minutes=30))

            self.metrics_output = pd.DataFrame(index = [location.SITE_NAME],columns = ['Site'])
            self.metrics_output['Site'][location.SITE_NAME] = location.SITE_NAME
                    
            pvsyst_data['AzSol']  =  pvsyst_data['AzSol'].apply(convert_pvsyst_azimuth_to_pvlib)
            pvsyst_data['PhiAng'] = -pvsyst_data['PhiAng']
            pvsyst_data['DifSInc'] += pvsyst_data['CircTrp']
            
            output_simulation.drop(columns = 'date', inplace=True)

            for parameter in output_simulation.columns:
                
                if parameter in pvsyst_data.columns:
                    self.metrics_output[parameter+'_r2'] = r2_calculation(output_simulation, pvsyst_data, parameter)
                    self.metrics_output[parameter+'_RMSE'] = rsme_calculation(output_simulation, pvsyst_data, parameter)  
                    self.metrics_output[parameter+'_diff_per_cent_signal'] = diff_ratio(output_simulation, pvsyst_data, parameter)
                    self.metrics_output[parameter+'_diff_per_cent'] = abs(self.metrics_output[parameter+'_diff_per_cent_signal'])
            
                        
def create_csv (location:object, output:object, path: str):
    
    import csv

    header = [version+'\n'+
      ';' + 'File' ';' +'\n' +
      'Site_name' + ';' + str(location.SITE_NAME) +'\n' +
     'Pvsyst_file' + ';' + str(location.PVSYST_FILE) +'\n' +
     'Solar_series' + ';' + str(location.SOLAR_SERIES_FILE) +'\n' +
     'Pan_file' + ';' + str(location.PAN_FILE) +'\n' +
     'Ond_file' + ';' + str(location.OND_FILE) +'\n' + 

     'LAT' +';' +'LON' +';' +'ALTITUDE'';' +'ALBEDO' +';' +'MAX_ANGLE' +';' +'D' +';' +'L' +';' +'INVERTERS' +';' +
      'MODULES_IN_SERIES' +';' +'MODULES_IN_PARALLEL' +';' +'U_c' +';' +'U_v' +';' +'STC_OHM_LOSS' +
      ';' +'STC_OHM_LOSS_AC' +';' +'QUALITY_LOSS' +';' +'LID_LOSS' +';' +'MISMATCH_LOSS' +';' +'SOILING_LOSS' + ';' 
      +'GHI_MIN_THRESHOLD' +';' +'FUSO' +';' +'MV_IRON_LOSS'+';' +'MV_COPPER_LOSS'+';' +'PMAX_OUT'+'\n' +
     str(location.LAT) +';' +str(location.LON) +';' +str(location.ALTITUDE) +';' +str(location.ALBEDO) +';' 
     +str(location.MAX_ANGLE) +';' +str(location.D) +';' +str(location.L) +';' +str(location.INVERTERS) +
      ';' +str(location.MODULES_IN_SERIES) +';' +str(location.MODULES_IN_PARALLEL) +';' +str(location.U_c) +';' 
      +str(location.U_v) +';' +str(location.STC_OHM_LOSS) +';' +str(location.STC_OHM_LOSS_AC) +';' 
      +str(location.QUALITY_LOSS) +';' +str(location.LID_LOSS) +';' +str(location.MISMATCH_LOSS) +';' 
      +str(location.SOILING_LOSS) +';' +str(location.GHI_MIN_THRESHOLD) +';' +'-'+str(location.FUSO)
      +';' +str(location.IRON_LOSS)+';' +str(location.COPPER_LOSS)+';' +str(location.PMAX_OUT)+'\n']
    
    output['date'] = output.date.dt.strftime("%Y-%m-%d %H:%M:%S %Z") 

    with open(path + 'cver/simulation_CVER/' + location.SITE_NAME + '.csv', mode = "w", encoding = 'utf-8-sig') as f:
            writer = csv.writer(f, quotechar = ' ')
            writer.writerow(header)
            writer.writerow(['date; GlobHor; E_Grid \n ; W/m²; W \n'])

    output[['date', 'GlobHor', 'E_Grid']].to_csv(path +'cver/simulation_CVER/' + location.SITE_NAME + '.csv', 
    index = False, header=False, sep = ';', mode = 'a')

        
def simulation(path: str, site_name: str or int, pvsyst_validation: bool):

    """
                Com esta função é possível realizar a simulação do cenário para uma determinada localidade.   
                
                A função possui somente três argumentos.
                
                -------------------
                path : str - Recebe o endereço do data base file.
                site_name : int or srt - Recebe um valor numérico, correspondente a linha do arquivo. 
                Ou o nome do site_name. 
                pvsyst_validation: bool - Recebe o valor booleano que ativa ou desativa a comparação 
                dos dados do simulador com os dados obtidos pelo PVSyst. 
    """
    
    location = DataLocations(path = path, site_name = site_name)
    module = PVModulo(path = location.PAN_FILE)
    inverter = Inverter(path = location.OND_FILE)
    datasimulation = Simulation(location = location, modulo = module, inversor = inverter)
    create_csv(location = location, output = datasimulation.simulation_output, path = path)

    if (pvsyst_validation == True) & (location.PVSYST_FILE != None):

        metrics = MetricsComplete(location = location, output_simulation = datasimulation.simulation_output)
        resultados_simulacao = metrics.metrics_output.append(pd.read_csv(path + 'cver/simulation_metrics.csv', sep = ';'))
        resultados_simulacao.to_csv(path + 'cver/simulation_metrics.csv', sep = ';', index = False)

    return  print('Arquivo', site_name, 'criado em', datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S'))




        
        
        
        
        
        
        
        
        
        
        
        
    

            
        
            
        
        
    
