import matplotlib.pyplot as plt
import pandas as pd
import math
import scipy.stats as stat


#class to organize data
class run:
    def __init__(self,df,name,conc,col1,col2,ipa,ipc,corrected_col2, epa, epc):
        self.df = df
        self.name = name
        self.conc = conc
        self.col1 = col1
        self.col2 = col2
        self.ipa = ipa
        self.ipc = ipc
        self.corrected = corrected_col2
        self.epa = epa
        self.epc = epc

def get_data():

    #parameters for excel read
    intlist = [x for x in range(1,18)]
    io=r"C:\Users\elija\Documents\exp8.xlsx"
    sheet_name= intlist
    usecols=0,1
    skiprows=0
    nrows=1202

    #to get sheet names
    

    #read excel file into dataframe dict
    df = pd.read_excel(io,sheet_name=sheet_name,usecols=usecols,skiprows=skiprows,nrows=nrows)
    #convert dict into list, ignoring sheet1
    data = [df[x] for x in range(1,18)]
    return data

def clean_data(data):
    io=r"C:\Users\elija\Documents\exp8.xlsx"
    xl = pd.ExcelFile(io)
#convert data to numeric form; ignore NaN
    for d in data:
        d["Column1"] = pd.to_numeric(d["Column1"],errors="coerce")
        d["Column2"] = pd.to_numeric(d["Column2"],errors="coerce")
        d = d.dropna()

    #create list for run objects; get list of sheet names
    runs = []
    sheets = xl.sheet_names


    #set all sheets to run objects
    for i in range(17):
        sheet = data[i]
        name = sheets[i+1]
        col1 = [x for x in sheet["Column1"]]
        col1 = [item for item in col1 if not(math.isnan(item))]
        col2 = [x for x in sheet["Column2"]]
        col2 = [item for item in col2 if not(math.isnan(item))]
        ipa = 0
        ipc = 0
        epa = 0
        epc = 0

        #set conc of unknown and background to 0
        if i < 3:
            conc = 0
        #get conc 10 sheets out of the way since they have two digits
        elif i < 5 and i > 2:
            conc = 10
        #get conc from name
        else:
            conc = int(name[0])
        run_i = run(sheet, name, conc, col1, col2, ipa, ipc, col2, epa, epc)
        runs.append(run_i)

    #correct current measurements by subtracting background
    for item in runs:
        corrected_col2 = []
        for i in range(len(item.col2)):
            corrected = item.col2[i] - runs[2].col2[i]
            corrected_col2.append(corrected)
        item.corrected_col2 = corrected_col2
        item.ipa = min(corrected_col2)
        item.ipc = max(corrected_col2)
        for j in range(len(item.col2)):
            if item.col2[j] == item.ipa:
                item.epa = item.col1[j]
            elif item.col2[j] == item.ipc:
                item.epc = item.col1[j]
            else:
                pass
    return runs

def organize_data(runs, choice):
    #get all 0.1 scan rate runs where the conc is known excluding background
    trunc_runs = [x for x in runs if x.name != "unknownmM_2" and x.name != "unknownmM_1" and x.name != "background" and x.name != "2mM_500mV" and x.name != "2mM_200mV" and x.name != "2mM_50mV" and x.name != "2mM_20mV"]

    #get all the first runs
    first_runs = [trunc_runs[x] for x in range(1,10,2)]
    #get all the second runs
    sec_runs = [trunc_runs[x] for x in range(0,9,2)]
    #get all the 2 mM runs with different scan rates
    mM_2_runs = [x for x in runs if x.name[0] == '2']

    if choice == 0:
        return first_runs
    elif choice == 1:
        return sec_runs
    else:
        return mM_2_runs

def regress(chosen_run, runs):
#get conc, ipc, and ipa values from specified runs
    concentrations = [x.conc for x in chosen_run]
    ipcs = [x.ipc for x in chosen_run]
    ipas = [x.ipa for x in chosen_run]



    #create linear regressions of cathodic and anodic current vs. conc
    clin = stat.linregress(concentrations, ipcs)
    alin = stat.linregress(concentrations, ipas)
    cslope = clin[0]
    cintercept = clin[1]
    aslope = alin[0]
    aintercept = alin[1]

    print("Ipc vs. conc")
    print(clin)
    print('\n')
    
    print("Ipa vs. conc")
    print(alin)
    print('\n')

    #estimate unknown conc from regression
    unknown1 = runs[0].ipc
    cunknown1 = (unknown1 + cintercept)/cslope
    unknown1 = runs[0].ipa
    aunknown1 = (unknown1 + aintercept)/aslope

    unknown2 = runs[1].ipc
    cunknown2 = (unknown2 + cintercept)/cslope
    unknown2 = runs[1].ipa
    aunknown2 = (unknown2 + aintercept)/aslope
    output = []
    output.append(concentrations)
    output.append(ipcs)
    output.append(ipas)
    output.append(cunknown1)
    output.append(aunknown1)
    output.append(cunknown2)
    output.append(aunknown2)
    output.append(clin)
    output.append(alin)
    return output

def graph(output):

    #unpack output
    concentrations = output[0]
    ipcs = output[1]
    ipas = output[2]
    cunknown1 = output[3]
    aunknown1 = output[4]
    cunknown2 = output[5]
    aunknown2 = output[6]
    clin = output[7]
    alin = output[8]
    cslope = clin[0]
    cintercept = clin[1]
    aslope = alin[0]
    aintercept = alin[1]

    #create trendlines
    ctrendx = [x for x in range(11)]
    ctrendy = [(x*cslope + cintercept) for x in ctrendx]
    atrendx = [x for x in range(11)]
    atrendy = [(x*aslope + aintercept) for x in atrendx]

    #plot points
    fig, ax = plt.subplots()
    ax.scatter(concentrations, ipcs)
    ax.scatter(concentrations, ipas)

    #plot trendlines
    ax.plot(ctrendx, ctrendy, c = 'black')
    ax.plot(atrendx, atrendy, c = 'black')

    #decorate graph
    ax.set_title("Ip vs. Concentration")
    ax.set_xlabel("[FeCN]")
    ax.set_ylabel("Ip")
    ax.tick_params(axis="both", labelsize=14)
    ax.minorticks_on()
    ax.grid()

    #print estimated conc of unknowns
    print("Unknown1 concentration in mM (determined by ipc): " + str(cunknown1))
    print("Unknown1 concentration in mM(determined by ipa): " + str(aunknown1))
    print('\n')
    print("Unknown2 concentration in mM(determined by ipc): " + str(cunknown2))
    print("Unknown2 concentration in mM(determined by ipa): " + str(aunknown2))
    print('\n')
    #show graph
    plt.show()

def regress_scans(chosen_run, runs):
#get conc, ipc, and ipa values from specified runs
    rates = [500,200,50,20,100,100]
    rates = [math.sqrt(x) for x in rates]
    ipcs = [x.ipc for x in chosen_run]
    ipas = [x.ipa for x in chosen_run]



    #create linear regressions of cathodic and anodic current vs. conc
    clin = stat.linregress(rates, ipcs)
    alin = stat.linregress(rates, ipas)
    cslope = clin[0]
    cintercept = clin[1]
    aslope = alin[0]
    aintercept = alin[1]

    print("Ipc vs. sqrt(v)")
    print(clin)
    print('\n')
    
    print("Ipa vs. sqrt(v)")
    print(alin)
    print('\n')

    #estimate unknown conc from regression
    unknown1 = runs[0].ipc
    cunknown1 = (unknown1 + cintercept)/cslope
    unknown1 = runs[0].ipa
    aunknown1 = (unknown1 + aintercept)/aslope

    unknown2 = runs[1].ipc
    cunknown2 = (unknown2 + cintercept)/cslope
    unknown2 = runs[1].ipa
    aunknown2 = (unknown2 + aintercept)/aslope
    output = []
    output.append(rates)
    output.append(ipcs)
    output.append(ipas)
    output.append(cunknown1)
    output.append(aunknown1)
    output.append(cunknown2)
    output.append(aunknown2)
    output.append(clin)
    output.append(alin)
    return output

def graph_scans(output):

    #unpack output
    rates = output[0]
    ipcs = output[1]
    ipas = output[2]

    clin = output[7]
    alin = output[8]
    cslope = clin[0]
    cintercept = clin[1]
    aslope = alin[0]
    aintercept = alin[1]

    #create trendlines
    ctrendx = [x for x in range(math.floor(math.sqrt(500)+1))]
    ctrendy = [(x*cslope + cintercept) for x in ctrendx]
    atrendx = [x for x in range(math.floor(math.sqrt(500)+1))]
    atrendy = [(x*aslope + aintercept) for x in atrendx]

    #plot points
    fig, ax = plt.subplots()
    ax.scatter(rates, ipcs)
    ax.scatter(rates, ipas)

    #plot trendlines
    ax.plot(ctrendx, ctrendy, c = 'black')
    ax.plot(atrendx, atrendy, c = 'black')

    #decorate graph
    ax.set_title("Ip vs. Sqrt(Rates)")
    ax.set_xlabel("v1/2")
    ax.set_ylabel("Ip")
    ax.tick_params(axis="both", labelsize=14)
    ax.minorticks_on()
    ax.grid()



    #show graph
    plt.show()

def diffusion_coefficent(slope):
    k = 2.69e5
    c = 0.002
    pi = math.pi
    r = 1.5e-6
    r *=100000
    a = pi*(r**2)

    d = slope
   
    d /= c
   
    d /= a
   
    d **= 2
    print("Diffusion coefficient")
    print(d)
    d = 1/d
    print('\n')
    print("1/d")
    print(d)
    print('\n')



#run program
data = get_data()
runs = clean_data(data)
chosen = organize_data(runs, 0)
output = regress(chosen, runs)
graph(output)

chosen = organize_data(runs, 1)
output = regress(chosen, runs)
graph(output)

scans = organize_data(runs, 2)

output = regress_scans(scans, runs)
graph_scans(output)
slope = output[7][0]
diffusion_coefficent(slope)
slope = output[8][0]
diffusion_coefficent(slope)

ratios = []
good = []
bad = []
e0 = []
for item in runs:
    try:
        ratio = item.ipa/item.ipc
        ratio = abs(ratio)
        ratios.append(ratio)
    except ZeroDivisionError:
        ratio = 0
    
    if ratio > 1.5 or ratio < 0.5:
        bad.append(ratio)
    else:
        good.append(ratio)
    
    e01 = item.ipa + item.ipc
    e01 /= 2
    e0.append(e01)


print("Ratios")
print(ratios)
print("Good")
print(good)
print("Bad")
print(bad)
print('\n')
print("Formal potentials")
print(e0)

#create duck plots for each run
# i = 0
# io=r"C:\Users\elija\Documents\exp8.xlsx"
# xl = pd.ExcelFile(io)
# names = xl.sheet_names
# for d in data:
#     i += 1
#     fig, ax = plt.subplots()
#     ax.plot(d["Column1"], d["Column2"], linestyle = "solid", color = "purple")
#     ax.set_title(names[i])
#     ax.set_xlabel("Voltage")
#     ax.set_ylabel("Current")
#     ax.tick_params(axis="both", labelsize=14)
#     ax.minorticks_on()
#     ax.grid()

# plt.show()


