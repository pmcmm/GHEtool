# import all the relevant functions
import GHEtool as ghe
import xlwings as xw
import numpy as np

def main():
    try:
        sheet = xw.sheets[0]
    except:
        print ("Please open the excel to run this function")
        return
    
    # relevant borefield data for the calculations
    data = {"H": sheet['bh_depth'].value,           # depth (m)
            "B": sheet['bh_spacing'].value,           # borehole spacing (m)
            "N_1":int(sheet['bh_width'].value),           # width of rectangular field (#)
            "N_2":int(sheet['bh_length'].value),           # length of rectangular field (#)}
            "Rb":sheet['bh_resistance'].value,           # equivalent borehole resistance (K/W) 0.08-0.18
            "Tg":sheet['soil_temperature'].value,            # Ground temperature at infinity (degrees C)
            "k_s": sheet['soil_conductivity'].value,           # conductivity of the soil (W/mK)
            }
    HeatingDemand = np.arange(12)
    HeatingPeak = np.arange(12)
    
    CoolingDemand = np.arange(12)
    CoolingPeak = np.arange(12)
    
    for i in range(0, 12):
        HeatingDemand[i] = sheet['heating_demand'].offset(i + 1,0).value
        CoolingDemand[i] = sheet['cooling_demand'].offset(i + 1,0).value
        HeatingPeak[i] = sheet['heating_peak'].offset(i + 1,0).value
        CoolingPeak[i] = sheet['cooling_peak'].offset(i + 1,0).value

    # Montly loading values
    #peakCooling = [0., 0, 34., 69., 133., 187., 213., 240., 160., 37., 0., 0.]              # Peak cooling in kW
    #peakHeating = [160., 142, 102., 55., 0., 0., 0., 0., 40.4, 85., 119., 136.]             # Peak heating in kW

    # annual heating and cooling load
    print ("Heating demand: %i kWh"%HeatingDemand.sum())
    print ("Cooling demand: %i kWh"%CoolingDemand.sum())
    annualHeatingLoad = HeatingDemand.sum()
    annualCoolingLoad = CoolingDemand.sum()
    
    
    
    # percentage of annual load per month (15.5% for January ...)
    #print ("Monthly heating demand:")
    #print (HeatingDemand / HeatingDemand.sum())
    monthlyLoadHeatingPercentage = HeatingDemand / HeatingDemand.sum()
    monthlyLoadCoolingPercentage = CoolingDemand / CoolingDemand.sum()

    # resulting load per month
    monthlyLoadHeating = list(map(lambda x: x * annualHeatingLoad, monthlyLoadHeatingPercentage))   # kWh
    monthlyLoadCooling = list(map(lambda x: x * annualCoolingLoad, monthlyLoadCoolingPercentage))   # kWh

    # create the borefield object

    borefield = ghe.Borefield(simulationPeriod=20,
                          peakHeating=HeatingPeak,
                          peakCooling=CoolingPeak,
                          baseloadHeating=monthlyLoadHeating,
                          baseloadCooling=monthlyLoadCooling)

    borefield.setGroundParameters(data)

    # set temperature boundaries
    borefield.setMaxGroundTemperature(16)   # maximum temperature
    borefield.setMinGroundTemperature(0)    # minimum temperature

    print ("Imbalance: %i kWh"%borefield.imbalance) # print imbalance

    # size borefield
    depth= borefield.size(100)
    print ("Optimized depth: %.1f m"%depth)
    sheet['results_depth'].value = depth

    # plot temperature profile for the calculated depth
    #borefield.printTemperatureProfile(legend=True)

    #plot temperature profile for a fixed depth
    borefield.printTemperatureProfileFixedDepth(depth=data['H'],legend=False)

    # print gives the array of montly tempartures for peak cooling without showing the plot
    borefield.calculateTemperatures(depth=90)
    #print (borefield.resultsPeakCooling)
    
if __name__ == "__main__":
    main()