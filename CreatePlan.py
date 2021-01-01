# Optimization Project
# Fall 2020
# Author: Dhanya Lakshmi (dl998)
# imports
import xlrd
from ortools.linear_solver import pywraplp
import seaborn as sns; sns.set_theme()
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import Ellipse


# Some common variables
# path to dosematrix.xlsx and variables.xlsx - file created to label all the variables 
DM = 'DoseMatrix.xlsx'
VARS = 'Variables.xlsx'
# structure: {voxel:[intensity1, intensity2, ..., intensity60]}
dosage_matrix = {}
# create lists to hold the respective voxels
ctv_voxels = []
bld_voxels = []
uns_voxels = []
rec_voxels = []
lfh_voxels = []
rfh_voxels = []
# number of voxels and beamlets
V = 400
B = 60
M = 300

# UNCOMMENT FOR DIFFERENT PLANS
# ===============================
# If using Plan 1 uncomment the following lines
z_coeffs = [1, 1, 5, 5, 1, 150, 150, 1, 1]
IND = 0
LB = 80.73
UB = 84.78

# If using Plan 2 uncomment the following lines
# z_coeffs = [1, 1, 1, 1, 200, 5000, 1000, 1000, 600]
# IND = 1
# LB = 79.5
# UB = 85

# If using Plan 3 uncomment the following lines
# z_coeffs = [1000, 1, 1000, 1000, 200, 200, 200, 100, 100]
# IND = 1
# LB = 80.73
# UB = 84.78

# creating the dosage_matrix dictionary
workbook = xlrd.open_workbook(DM)
sheet = workbook.sheet_by_index(0)
for i in range(1, V+1):
    vox = int(sheet.cell_value(i, 0))
    dosage_matrix[vox] = []
    for j in range(1, B+1):
        dosage_matrix[vox].append(float(sheet.cell_value(i, j)))

# creating the lists
var_wb = xlrd.open_workbook(VARS)
var_sheet = var_wb.sheet_by_index(IND)
labels = {'Unspecified region':uns_voxels, 'Bladder':bld_voxels, 'Right Femur Head':rfh_voxels, 'Left Femur Head':lfh_voxels, 'CTV':ctv_voxels, 'Rectal Solid':rec_voxels}
for i in range(0, V):
    ind = int(var_sheet.cell_value(i, 0))
    lst = labels[str(var_sheet.cell_value(i, 1))]    
    lst.append(ind)

print(len(ctv_voxels), len(bld_voxels), len(rec_voxels))
# ============================================================================================
def InitializeLP():
    # initiate the solver
    solver = pywraplp.Solver('PROJECT', pywraplp.Solver.CBC_MIXED_INTEGER_PROGRAMMING)
    inf = solver.infinity()
    # set the objective to minimization
    objective=solver.Objective()
    # objective.SetMaximization()
    objective.SetMinimization()

    # create the variables for the optimization problem
    # beamlets - 60 of them
    beamlets = [0 for i in range(60)]
    z_neg = [0 for i in range(9)]
    bladder_int = [0 for i in range(len(bld_voxels))]
    lfh_int = [0 for i in range(len(lfh_voxels))]
    rfh_int = [0 for i in range(len(rfh_voxels))]

    # add to the variable lists and also to the objective function 
    for b in range(B):
        beamlets[b] = solver.NumVar(0, inf, 'b'+str(b))

    # Create z_neg only - errors for each constraint
    # Objective: Minimize z_neg 
    for z in range(len(z_neg)):
        z_neg[z] = solver.NumVar(0, inf, 'zneg'+str(z))
        objective.SetCoefficient(z_neg[z], z_coeffs[z])

    # creating int variables for bladder
    for i in range(len(bld_voxels)):
        bladder_int[i] = solver.IntVar(0, 1, 'bi'+str(i))

    # creating int variables for lfh
    for i in range(len(lfh_voxels)):
        lfh_int[i] = solver.IntVar(0, 1, 'lfhi'+str(i))

    # creating int variables for bladder
    for i in range(len(rfh_voxels)):
        rfh_int[i] = solver.IntVar(0, 1, 'rfhi'+str(i))

    # Now, the constraints
    # 1st constraint: ctv: each ctv dosage must be between 80.4 and 84.8
    ctv_consts = [0]*len(ctv_voxels)

    for i in range(len(ctv_voxels)):
        ctv_consts[i] = solver.Constraint(LB, UB)
        for b in range(len(beamlets)):
            ctv_consts[i].SetCoefficient(beamlets[b], dosage_matrix[ctv_voxels[i]][b])

    # 2nd set of constraints: BLADDER
    # 1: max b_dose should be leq 81
    bladder_constraints_bound = [0]*len(bld_voxels)
    for i in range(len(bld_voxels)):
        bladder_constraints_bound[i] = solver.Constraint(0, 81)
        for b in range(len(beamlets)):
            bladder_constraints_bound[i].SetCoefficient(beamlets[b], dosage_matrix[bld_voxels[i]][b])


    # 2: mean should be leq 50
    # i.e. sum of all bladder doses should be leq 50*num_bladder_constraints
    bladder_constraints_mean = solver.Constraint(0, len(bld_voxels)*50)
    bladder_constraints_mean.SetCoefficient(z_neg[0], -1)
    for b in range(len(beamlets)):
        coeff = 0
        for v in range(len(bld_voxels)):
            coeff += dosage_matrix[bld_voxels[v]][b]
        bladder_constraints_mean.SetCoefficient(beamlets[b], coeff)

    # # 3: at most 10% of voxels should receive dose over 65%
    bladder_constraints_atmost = [0]*len(bld_voxels)
    for i in range(len(bld_voxels)):
        bladder_constraints_atmost[i] = solver.Constraint(0, 65)
        bladder_constraints_atmost[i].SetCoefficient(bladder_int[i], -1*M)
        bladder_constraints_atmost[i].SetCoefficient(z_neg[1], -1)
        for b in range(len(beamlets)):
            bladder_constraints_atmost[i].SetCoefficient(beamlets[b], dosage_matrix[bld_voxels[i]][b])

    bladder_constraints_atmost.append(solver.Constraint(0, 0.10*len(bld_voxels)))
    for bi in bladder_int:
        bladder_constraints_atmost[-1].SetCoefficient(bi, 1)

    # 3rd set of constraints: RECTUM
    # 1: max dose to rectum should be leq 79.2
    rectum_constraints_bound = [0]*len(rec_voxels)
    for i in range(len(rec_voxels)):
        rectum_constraints_bound[i] = solver.Constraint(0, 79.2)
        rectum_constraints_bound[i].SetCoefficient(z_neg[2], -1)
        for b in range(len(beamlets)):
            rectum_constraints_bound[i].SetCoefficient(beamlets[b], dosage_matrix[rec_voxels[i]][b])

    # 2: mean should be leq 40
    # i.e. sum of all rectum doses should be leq 40*num_rectum_constraints
    rectum_constraints_mean = solver.Constraint(0, len(rec_voxels)*40)
    rectum_constraints_mean.SetCoefficient(z_neg[3], -1)
    for b in range(len(beamlets)):
        coeff = 0
        for v in range(len(rec_voxels)):
            coeff += dosage_matrix[rec_voxels[v]][b]
        rectum_constraints_mean.SetCoefficient(beamlets[b], coeff)

   # 4th set of constraints: UNSPECIFIED
    # 1: max dose to uns should be leq 72.0
    uns_constraints_bound = [0]*len(uns_voxels)
    for i in range(len(uns_voxels)):
        uns_constraints_bound[i] = solver.Constraint(0, 72.0)
        uns_constraints_bound[i].SetCoefficient(z_neg[4], -1)
        for b in range(len(beamlets)):
            uns_constraints_bound[i].SetCoefficient(beamlets[b], dosage_matrix[uns_voxels[i]][b])

   # 5th set of constraints: LFH
    # 1: max dose to lfh should be leq 50
    lfh_constraints_bound = [0]*len(lfh_voxels)
    for i in range(len(lfh_voxels)):
        lfh_constraints_bound[i] = solver.Constraint(0, 50)
        lfh_constraints_bound[i].SetCoefficient(z_neg[5], -1)
        for b in range(len(beamlets)):
            lfh_constraints_bound[i].SetCoefficient(beamlets[b], dosage_matrix[lfh_voxels[i]][b])

    # 2: at most 15% > 40
    lfh_constraints_atmost = [0]*len(lfh_voxels)
    for i in range(len(lfh_voxels)):
        lfh_constraints_atmost[i] = solver.Constraint(0, 40)
        lfh_constraints_atmost[i].SetCoefficient(lfh_int[i], -1*M)
        lfh_constraints_atmost[i].SetCoefficient(z_neg[6], -1)
        for b in range(len(beamlets)):
            lfh_constraints_atmost[i].SetCoefficient(beamlets[b], dosage_matrix[lfh_voxels[i]][b])

    lfh_constraints_atmost.append(solver.Constraint(0, 0.15*len(lfh_voxels)))
    for lfhi in lfh_int:
        lfh_constraints_atmost[-1].SetCoefficient(lfhi, 1)


   # 6th set of constraints: RFH
    # 1: max dose to rfh should be leq 50
    rfh_constraints_bound = [0]*len(rfh_voxels)
    for i in range(len(rfh_voxels)):
        rfh_constraints_bound[i] = solver.Constraint(0, 50)
        rfh_constraints_bound[i].SetCoefficient(z_neg[7], -1)
        for b in range(len(beamlets)):
            rfh_constraints_bound[i].SetCoefficient(beamlets[b], dosage_matrix[rfh_voxels[i]][b])


    # 2: at most 15% > 40
    rfh_constraints_atmost = [0]*len(rfh_voxels)
    for i in range(len(rfh_voxels)):
        rfh_constraints_atmost[i] = solver.Constraint(0, 40)
        rfh_constraints_atmost[i].SetCoefficient(rfh_int[i], -1*M)
        rfh_constraints_atmost[i].SetCoefficient(z_neg[8], -1)
        for b in range(len(beamlets)):
            rfh_constraints_atmost[i].SetCoefficient(beamlets[b], dosage_matrix[rfh_voxels[i]][b])

    rfh_constraints_atmost.append(solver.Constraint(0, 0.15*len(rfh_voxels)))
    for rfhi in rfh_int:
        rfh_constraints_atmost[-1].SetCoefficient(rfhi, 1)
    
    return solver, objective, beamlets, z_neg

# ============================================================================================
if __name__ == "__main__":

    solver, objective, beamlets, z_neg = InitializeLP()
    status = solver.Solve() # optimize the model
    if status == solver.OPTIMAL:
        print('Problem solved in %f milliseconds' %solver.wall_time())
    elif status == solver.FEASIBLE:
        print('Solver claims feasibility but not optimality')
        exit(1)
    else:
        print('Solver ran to completion but did not find an optimal solution')
        exit(1)
    print('Objective value = %f' %(objective.Value()))

    avg_ctv_dose = 0
    ctv_doses = []
    avg_bld_dose = 0
    bld_doses = []
    avg_rec_dose = 0
    rec_doses = []
    avg_uns_dose = 0
    uns_doses = []
    avg_lfh_dose = 0
    lfh_doses = []
    avg_rfh_dose = 0
    rfh_doses = []
    voxel_dosages = []
    for i in range(1, 401):
        voxel_dosages.append(0)
        beamlet_sum = 0
        for b in range(len(beamlets)):
            voxel_dosages[i-1] += beamlets[b].solution_value()*dosage_matrix[i][b]
            beamlet_sum += beamlets[b].solution_value()*dosage_matrix[i][b]
        if(i in ctv_voxels):
            ctv_doses.append(beamlet_sum)
            avg_ctv_dose += beamlet_sum
        elif(i in bld_voxels):
            bld_doses.append(beamlet_sum)
            avg_bld_dose += beamlet_sum
        elif(i in rec_voxels):
            rec_doses.append(beamlet_sum)
            avg_rec_dose += beamlet_sum
        elif(i in uns_voxels):
            uns_doses.append(beamlet_sum)
            avg_uns_dose += beamlet_sum
        elif(i in lfh_voxels):
            lfh_doses.append(beamlet_sum)
            avg_lfh_dose += beamlet_sum
        elif(i in rfh_voxels):
            rfh_doses.append(beamlet_sum)
            avg_rfh_dose += beamlet_sum
    print("Average CTV dosage = ", avg_ctv_dose/len(ctv_voxels))
    print("Maximum CTV dosage = ", max(ctv_doses))
    print("Average Bladder dosage = ", avg_bld_dose/len(bld_voxels))
    print("Maximum Bladder dosage = ", max(bld_doses))
    print("Average Rectal Solid dosage = ", avg_rec_dose/len(rec_voxels))
    print("Maximum Rectal Solid dosage = ", max(rec_doses))
    print("Average Unspecified region dosage = ", avg_uns_dose/len(uns_voxels))
    print("Maximum Unspecified region dosage = ", max(rec_doses))
    print("Average RFH dosage = ", avg_rfh_dose/len(rfh_voxels))
    print("Maximum RFH dosage = ", max(rfh_doses))
    print("Average LFH dosage = ", avg_lfh_dose/len(lfh_voxels))
    print("Maximum LFH dosage = ", max(lfh_doses))
    print("====================================================")
    for b in range(len(beamlets)):
        print("Intensity of beamlet ", str(b+1), "is: ", beamlets[b].solution_value())
    print("====================================================")
    for i in range(len(z_neg)):
        print("Value of "+ str(z_neg[i]) + " is: ", z_neg[i].solution_value())

    # Plotting the heatmap, drawing the boundaries
    voxel_dosages = np.reshape(np.array(voxel_dosages), (-1, 20))
        
    ax = sns.heatmap(voxel_dosages, cmap="mako", annot=True, fmt='.2f')
    plt.xticks(np.arange(1,20,1))
    plt.yticks(np.arange(1,20,1))     
    # Ellipses: Bladder, CTV, Rectal Solid, RFH, LFH
    ells = [Ellipse(xy=(10.5, 14), width=7.5, height=4, lw=2, facecolor='none'), Ellipse(xy=(10.5, 10), width=8, height=4, lw=2, facecolor='none'), Ellipse(xy=(10.5, 6), width=5.6, height=4, lw=2, facecolor='none'), Ellipse(xy=(3, 11), width=2.6, height=8, lw=2, facecolor='none'), Ellipse(xy=(17, 11), width=2.6, height=8, lw=2, facecolor='none')]
    for e in ells:
        ax.add_artist(e)
    ax.invert_yaxis()
    # Finally, show plot
    plt.show()