<p>
<img src="https://github.com/Casa-dos-Ventos/simulador_solar/blob/main/pictures/logo.png" alt="Scver_logo" style="float:right;width:149px;height:38px;">


# Documentation of Solar Simulator CVER

</p> 

## List of Contents
  - [Purpose](#purpose)
  - [Overall](#overall)
  - [External data](#external-data)
  - [Instalation](#instalation)
  - [How to use it](#how-to-use-it)
  - [Output](#output)
  - [To do list](#to-do-list)


---

> ## Purpose

The objective of this repository is to create a routine that receives as input meteorological data from a location and the configuration of a photovoltaic plant and returns the energy data of the plant.

---

> ## Overall

The code gets information from two files one .csv and one .xlsx: the first contains local weather data generated from hybridsim; the other containing the data for the plant simulation, such as the directories of the weather data file, .PAN file, .OND file, module configuration, tracker configuration, inverter configuration and system losses.
 
 
 ---

> ## External data

The external information used are listed below:

<p id = external-data-table >

- **./solar/ts/TMY.csv**
  - time
  - GHI: float [W/m²]
  - DIF: float [W/m²]
  - TEMP: float [°C]
  - WS: float [m/s] 


- **./solar/cver/DataBase.xlsx**
    - site_name: str
    - pvsyst_file: str (optional)
    - solar_series: str
    - pan_file: str
    - ond_file: str
    - LAT: float [°]
    - LON: float [°]
    - ALTITUDE:int [m]
    - ALBEDO: float [0 to 1]
    - MAX_ANGLE: int [°]
    - D: float [m]
    - L: float [m]
    - INVERTERS: int
    - MODULES_IN_SERIES: int
    - MODULES_IN_PARALLEL: int
    - PNOM_RATIO:float (not used in simulation)
    - U_c: int [10 to 50]
    - U_v: float [0 to 12]
    - STC_OHM_LOSS: float [0 to 1]
    - STC_OHM_LOSS_AC: float [0 to 1] 
    - MV_IRON_LOSS: float [0 to 1] 
    - MV_COPPER_LOSS: float [0 to 1] 
    - MV_LOSS_STC: float [0 to 1] 
    - QUALITY_LOSS: float [0 to 1] 
    - LID_LOSS: float [0 to 1] 
    - MISMATCH_LOSS: float [0 to 1] 
    - SOILING_LOSS: float [0 to 1] 
    - GHI_MIN_THRESHOLD: int [W/m²]
    - FUSO: int (use positive integer "3" for "GMT -3")
    - PMAX_OUT: int [kW]

</p>

---

> ## Instalation
1. First thing you need to do is clone this repo into your folder with:

````
git clone https://github.com/Casa-dos-Ventos/simulador_solar.git
````

2. After that, *in the `G:\Drives compartilhados\[GITHUB] Casos de uso\simulador_solar`** folder, make a copy of the `Rio do Vento` template folder or follow the same nomenclature standardization for use in an already existing project. existing.

3. Than you need to install all the requirements.

````
pip install -r requirements.txt
````

> ## How to use it
Suppose you want to run code for some simulation scenarios with varying pitch and Pnom ratio, all you'll need to do is:

1. Use https://github.com/Casa-dos-Ventos/hybridsim to create the TMY of the location from satellite data and weather station data.

2. Indicate the file directories and the system configuration, following the same pattern used in the DataBase.xlsx file in the 'Casos de uso' folder.

3. If you want to obtain a comparison between the data from the cver simulator and the PVSyst in the output, it is necessary to add the files referring to the scenarios simulated in PVSyst to the './solar/cver/simulation_CVER/' folder.

``` python
######## INPUT ######## 

# ./solar/cver/DataBase.xlsx

> ## Output

The possible results are listed here:

- simulation_data in csv
- simulation_metrics in csv


You can decide whether to save the simulation_metrics file by changing this variable of the "simulation" function:

```python
# Set whether to save simulation_metrics as output:

# saves simulation_metrics
pvsyst_validation = True 
```

If the variable is "True", then the file will be created and will be inside the ".solar/cver/" folder. 

> ## To do list

This section list some future improvements that coluld be done.

1. https://docs.google.com/document/d/10aEmv8ojR_2bPooWyMcNjelHx-04bce0xrK2qvDIjJQ/edit?pli=1#heading=h.cdtbb1jejxf3
