import pulp as p 
import numpy as np
import xlrd, xlsxwriter

def exceptionR(x):
    return (1+np.log(x))/x if x else 0

def exceptionD(y):
    return y if y else 0

def solveLP(alpha):
    # Create a LP Minimization problem 
    Lp_prob = p.LpProblem('Vaccine_Allocation', p.LpMinimize) 
    
    # Create problem Variables 
    ind = range(len(N)) #total number of provinces and cities in ncr
 #   lind = range(len(N)*(1/3))
    m = p.LpVariable.dicts("m", ind, cat='Integer') 
    
    # Objective Function 
    # Lp_prob += p.lpSum([(N[i]-C[i]-m[i])*(1-exceptionR(R[i]))*exceptionD(D[i]) for i in ind])
    # Lp_prob += p.lpSum([((N[i]-C[i]-m[i])*(1-exceptionR(R[i])))*(1-np.exp(-Delta[i]/194581))*exceptionD(D[i]) for i in ind])
    Lp_prob += p.lpSum([(N[i]-C[i]-m[i])*(1-exceptionR(R[i]))*(1-np.exp(-Delta[i]/maxDelta))*exceptionD(D[i]) for i in ind])
    # Lp_prob += p.lpSum([((N[i]-C[i]-m[i])*(1-exceptionR(R[i])))/(N[i]-C[i]) for i in ind])
    
    # Constraints:
    Lp_prob += p.lpSum([m[j] for j in ind])==int(V*alpha)
    for t in ind:
        Lp_prob += m[t] <= N[t]-C[t]
        Lp_prob += float(alpha)*P[t] <= m[t]
    
    #Display the problem 
    # print(Lp_prob) 
    
    status = Lp_prob.solve() # Solver  
    
    # print(p.LpStatus[status])
    
    solution = np.array([m[i].varValue for i in ind])
    # print(p.value(Lp_prob.objective))
    
    return solution, p.value(Lp_prob.objective)


# To open Workbook
wb = xlrd.open_workbook("new-data-covid-ph2.xlsx")
sheet = wb.sheet_by_index(1) #1 since sa 2nd na sheet
# workbook = xlsxwriter.Workbook('dontOpenThis_forDebuggingLangTo_VacAllo_LP_initial_result.xlsx') 
workbook = xlsxwriter.Workbook('dec17_VacAllo_LP_initial_result.xlsx') 
worksheet = workbook.add_worksheet() 

N=[] #population per city/province
C=[] #recovered cases
D=[] #case fatality rate
R=[] #Rt
P=[] #preferred minimum vaccinated
Delta=[] #population density

cityNames = 1
n_col = 2
p_col = 5
delta_col = 3
c_col = 17
d_col = 23
r_col = 30
V = sheet.cell_value(2,n_col)
wp = sheet.cell_value(2, 5)
# print(wp)
city = []
index = np.array([3,21,29,34,40,50,57,64,71,80,88,96,102,110,117,124,131,138,156
#                  138,156,135+29,135+34,135+40,135+50,135+57,135+64,135+71,135+80,135+88,135+96,135+102,135+110,135+117,135+124,135+131
                  ])
k = np.array([.1, .2, .3, .4, .5, .6, .7, .8, .9, 1]) #percent of V (total population of Philippines)

#get size of m here: m=len(N)=len(C)=len(D)=len(R)=len(city)
for j in range(17):
    regionTotalPop = index[j]
    city_row = index[j+1]
    for i in range(regionTotalPop+1, city_row):
        N.append(sheet.cell_value(i,n_col))
        C.append(sheet.cell_value(i,c_col))
        D.append(sheet.cell_value(i,d_col))
        R.append(sheet.cell_value(i,r_col))
        P.append(0.4*sheet.cell_value(i,p_col)/wp*V)
        Delta.append(sheet.cell_value(i,delta_col))
        city.append(sheet.cell_value(i,cityNames))

maxDelta=max(Delta)
        
#write city/province in excel, column 1
for h in range(len(N)):
    worksheet.write(h+2,0, city[h])

#write headers in excel
worksheet.write(0,0, "City/Province")
for l in range(len(k)):
    worksheet.write(0,l+1, "%s*V" %k[l])

#loop for different values of alpha
# for z in range(1):
# for z in range(1):
for z in range(len(k)):
    [sol, objFunction] = solveLP(k[z])
    # [sol, objFunction] = solveLP(0.5)
    worksheet.write(1,z+1, objFunction)
    for t in range(len(sol)):
        worksheet.write(t+2,z+1, sol[t])
    
workbook.close()

# print(P)