import pandas as pd
import matplotlib.pyplot as plt
plt.rcParams["font.family"] = "Times New Roman"
def graphs(main_excel_file, main_sheet_name, x_column, y_column, start_row, output_file, axis_label):
    
    df_main = pd.read_excel(main_excel_file, sheet_name=main_sheet_name, header=start_row)
    x_data = df_main[x_column][0:200]
    y_data = df_main[y_column][0:200]
    
    plt.plot(x_data, y_data, marker='', linestyle='-', label='Data')
    plt.xlabel(x_column)
    plt.ylabel(axis_label)
    plt.legend()
    plt.savefig(output_file)
    plt.close()
    print(f'{output_file}')


columns = ["k/q", "T", "q_t", "q_t^tilde", "(1-Bjergtop)", "D_t", "SCC", "g_q", "g_k"]
names = ["Kapital-outputforhold", "Temperaturstigning", "Nettooutput per arbejder", " Nettooutput per effektiv arbejder", "Reduktionsomkostninger", "Skadesomkostninger", "Social cost of carbon", "Vækst i nettooutput per arbejder", "Vækst i kapital per arbejder"]
def many():
    for x in ["Baseline","2050", "2100", "1,5°", "2,0°", "3,0°"]:
        for y,z in zip(columns, names):
            graphs(
                main_excel_file= "Følsomhedsanalyse.xlsx",
                main_sheet_name= x,
                x_column='Tid',
                y_column= y,
                start_row=1,
                output_file="Udvikling i " +  z + " " + x + ".png",
                axis_label = z
            )

def netoutput():
    columns2 = ["q_t", "q_t^tilde", "SCC", "S_t", "T", "D_t", "Abatement", "k/q", "1-D", "1-(1-Bjergtop)*D", "Akkumulerede E"]
    names2 = ["Nettooutput per arbejder", "Nettooutput per effektiv arbejder", "Optimal skattesats", "Akkumuleret CO2 i atmosfæren", "Global gennemsnitstemperatur i grader celcius", "Skadesomkostninger", "Reduktionsomkostninger", "Kapital-outputforhold", "Klimaskadernes effekt på output", "Samlet skade", "Akkumulerede udledninger i atmosfæren"]
    legend_names = ["Basisscenarie", "2050", "2100", "1,5°", "2,0°", "3,0°"]
    for a, b in zip(columns2, names2):
        for c in ["Baseline", "2050", "2100", "1,5°", "2,0°", "3,0°"]:
            df_main = pd.read_excel("Følsomhedsanalyse.xlsx", sheet_name=c, header=1)
            x_data = df_main["Tid"][0:200]
            y_data = df_main[a][0:200]
            if  a == "D_t" or a == "Abatement" or a == "1-D" or b == "Samlet skade":
                y_data = [100 * x for x in y_data]
            plt.plot(x_data, y_data, marker='', linestyle='-', label='Data')
        plt.xlabel("Tid")
        if a == "S_t":
            plt.ylim(4500, 12500)
        if b == "Optimal skattesats":
            plt.ylabel("Amerikanske dollars per ton CO2")
        elif b == "Akkumuleret CO2 i atmosfæren" or b == "Akkumulerede udledninger i atmosfæren":
            if b == "Akkumulerede udledninger i atmosfæren":
                plt.ylim(0, 6000)
            plt.ylabel("Gigaton CO2 i atmosfæren")
        elif b == "Global gennemsnitstemperatur i grader celcius":
            plt.ylim(1, 3.5)
            plt.ylabel("Temperaturforskel i grader celcius")
        elif b == "Reduktionsomkostninger":
            plt.ylabel("Procent af bruttooutput fratrukket klimaskade")
        elif b == "Skadesomkostninger":
            plt.ylim(90, 100)
            plt.ylabel("Procent af bruttooutput")
        elif b == "Samlet skade" or b == "Klimaskadernes effekt på output":
            plt.ylim(0, 10)
            plt.ylabel("Procent af bruttooutput")
        elif b == "Kapital-outputforhold":
            plt.ylim(2, 4)
            plt.ylabel("Kapital-outputforhold")
        elif b == "Nettooutput per arbejder" or b == "Nettooutput per effektiv arbejder":
             if b == "Nettooutput per effektiv arbejder":
                 plt.ylim(1, 2)
             plt.ylabel("Tusinder dollars, PPP-justeret")  
        else:
            plt.ylabel("Billioner dollars, PPP-justeret")
        
        plt.legend(legend_names)
        
        plt.savefig(b + ".png")
        plt.close()
        print(b + ".png")



def temperature():
    groups = [["2050", "2100"], ["1,5°", "2,0°", "3,0°"]]
    developments = ["Temperaturudvikling ved forskellige tidspunkter for nuludledning", "Temperaturudvikling ved forskellige temperaturmål"]
    colours = ["blue", "red", "green"]
    for e, h in zip(groups, developments):
        for f, g in zip(e, colours):
            df_main = pd.read_excel("Følsomhedsanalyse.xlsx", sheet_name=f, header=1)
            x_data = df_main["Tid"][0:200]
            y_data = df_main["T"][0:200]
            plt.plot(x_data, y_data, marker='', linestyle='-', label='Data')
            df_lower = pd.read_excel("Følsomhedsanalyse.xlsx", sheet_name= f + (f != "Standard Solow") *" - TCRE 0,32", header=1)
            df_upper = pd.read_excel("Følsomhedsanalyse.xlsx", sheet_name= f + (f != "Standard Solow") * " - TCRE 0,62", header=1)
            lower_bound = df_lower["T"][0:200]
            upper_bound = df_upper["T"][0:200]
            plt.fill_between(x_data, upper_bound, lower_bound, color=g, alpha=0.2, label='Konfidensintervaller' + " for " + f)
        plt.xlabel("Tid")
        plt.ylabel("Temperaturforskel i grader celcius")
        if e[0] == "2050":
           plt.legend([e[0], "Konfidensintervaller for " + e[0], e[1], "Konfidensintervaller for " + e[1]]) 
        else:
            plt.legend([e[0], "Konfidensintervaller for " + e[0], e[1], "Konfidensintervaller for " + e[1], e[2], "Konfidensintervaller for " + e[2]])
        
        plt.savefig(h + ".png")
        plt.close()
        print(h + ".png")


def tipping_points():
    temperatures = ["1,5°", "2,0°", "3,0°"]
    names = ["Albedo", "Uden tipping point", "Permafrost"]
    for i, j in zip(temperatures, names):
        df_main = pd.read_excel("Følsomhedsanalyse.xlsx", sheet_name=i, header=1)
        x_data = df_main["Tid"][0:200]
        y_data = df_main["T"][0:200]
        plt.plot(x_data, y_data, marker='', linestyle='-', label='Data')
        if i == "1,5°":
            lower = max(y_data)
            limit = [1.5] * 200
            plt.fill_between(x_data, limit, lower, color="blue", alpha=0.3, label='Andre drivhusgasser')
            albedo = [1.97] *  200
            plt.fill_between(x_data, albedo, lower, color="blue", alpha=0.2, label='Albedoeffekten')
            plt.legend([i, "Andre drivhusgasser", "Ismasser smelter"])
        elif i == "3,0°":
            lower = max(y_data)
            limit = [3] *  200
            plt.fill_between(x_data, limit, lower, color="blue", alpha=0.3, label='Andre drivhusgasser')
            permafrost_1 = [3.36] * 200
            plt.fill_between(x_data, permafrost_1, lower, color="blue", alpha=0.2, label='Permafrosteffekten')
            permafrost_2 = [4.02] *  200
            plt.fill_between(x_data, permafrost_2, lower, color="blue", alpha=0.1, label='Permafrosteffekten_2')
            plt.legend([i, "Andre drivhusgasser", "30% af permafrosten smelter", "85% af permafrosten smelter"])

        else:
            lower = 1.9
            limit = [2] *  200
            plt.fill_between(x_data, limit, lower, color="blue", alpha=0.2, label='Andre drivhusgasser')
            plt.legend([i, "Andre drivhusgasser"])

        plt.xlabel("Tid")
        plt.ylabel("Temperaturforskel i grader celcius")
        plt.xlim(0, max(x_data) + 0,1)
        
        
        plt.savefig(j+ ".png")
        plt.close()
        print(j + ".png")

def production():
    new_columns = ["y_t", "y_skade", "q_t"]
    new_names = ["Bruttooutput", "Bruttooutput fratrukket skade", "Nettooutput"]
    for k in new_columns:
        df_main = pd.read_excel("Følsomhedsanalyse.xlsx", sheet_name="1,5°", header=1)
        x_data = df_main["Tid"][0:200]
        y_data = df_main[k][0:200]
        plt.plot(x_data, y_data, marker='', linestyle='-', label='Data')
    plt.xlabel("Tid")
    plt.ylabel("Tusinder dollars, PPP-justeret")
    plt.xlim(0, 80) 
    plt.ylim(30, 175) 
    plt.legend(new_names)
        
    plt.savefig(k + ".png")
    plt.close()
    print("Produktion.png")


def other_damage():
       
    analyser = ["Følsomhedsanalyse.xlsx", "Følsomhedsanalyse2.xlsx", "Følsomhedsanalyse3.xlsx"]
    grafer = ["q_t", "q_t^tilde", "1-(1-Bjergtop)*D", "1-D"]
    names = ["Skade - Nettooutput per arbejder", "Skade - Nettooutput per effektiv arbejder", "Skade - Total skade", "Klimaskade på bruttooutput"]
    for x, y in zip(grafer, names):
        for m in analyser:
            df_main = pd.read_excel(m, sheet_name="3,0°", header=1)
            x_data = df_main["Tid"][0:200]
            y_data = df_main[x][0:200]
            if x == "1-(1-Bjergtop)*D" or x == "1-D":
                y_data = [i * 100 for i in y_data]
            plt.plot(x_data, y_data, marker='', linestyle='-', label='Data')
        plt.xlabel("Tid")

        if y == "Skade - Total skade" or y == "Klimaskade på bruttooutput":
            plt.ylabel("Procent af bruttooutput")
        elif y == "Reduktionsomkostninger":
            plt.ylabel("Andel af nettoooutput") 
        elif y == "Skade - Nettooutput per arbejder" or y == "Skade - Nettooutput per effektiv arbejder":
            plt.ylabel("Tusinder dollars, PPP-justeret")    
        else:
            plt.ylabel("Billioner dollars, PPP-justeret")
        
        plt.legend(["Udgangspunkt", "Værst tænkelige situation", "DICE-2013"])
        
        plt.savefig(y + ".png")
        plt.close()
        print(y + ".png")

def accumulation():
    for x in ["Akkumuleret skade", "g_q"]:
        for y in ["1,5°", "3,0°"]:
            df_main = pd.read_excel("Følsomhedsanalyse.xlsx", sheet_name=y, header=1)
            x_data = df_main["Tid"][0:200]
            y_data = df_main[x][0:200]

            plt.plot(x_data, y_data, marker='', linestyle='-', label='Data')
        plt.xlabel("Tid")
        if x == "Akkumuleret skade":
            plt.ylabel("Tusinder dollars, PPP-justeret")
        else:
            plt.ylabel("Vækstrate for nettooutput per arbejder")
        plt.legend(["1,5°", "3,0°"])
        plt.savefig("{}.png".format(x))

        plt.close()
    print("Akkumuleret skade.png")
accumulation()
other_damage()
tipping_points()
many()
netoutput()
temperature()
production()