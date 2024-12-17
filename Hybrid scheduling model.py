#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Mar  7 20:36:01 2024

@author: yangjiatian
"""

from pyomo.environ import * 
import pandas as pd 
import numpy  
import math 

#data preparation 
#read dataset 
file_name='Appendix.1 MSBA-Term---Workshop-and-Seminar-Allocations-group28.xlsx' 
df_class = pd.read_excel(file_name,'Class', index_col=0) 
df_time = pd.read_excel(file_name,'Time', index_col=0) 

# different classes in the schedule - k = 23  
class_name = df_class.columns[1:]  

# number of students i.e. 229 MSBA 
stu_id = df_class.index[:-1] 

# Ak - set of students in k for 23 classes
enrolled_stu = df_class.iloc[0:-1,1:] 

# Input time t for each class - t = 23
time_table = df_time.index[:] 
time_TB = df_time.iloc[:,:] 

# class room capacity ck
Rk_class = df_class.loc['capacity'] 
Rk_class = Rk_class.iloc[1:] 

# Group number M 
groups = ['A','B','C'] 

# Social distance ratio
D = 0.3 

# Weights for the three metrics
weight_TE = 1 
weight_TD = 0.25 
weight_SSE = 0

# excess room capacity 
E = 25 

# Modelling
model = ConcreteModel() 
model.pi = Var(stu_id, groups, domain=Binary)  
model.ejk = Var(groups, class_name, domain=NonNegativeIntegers)  
model.djk = Var(groups, class_name, domain=NonNegativeReals)  
model.stj = Var(time_table, groups, domain=NonNegativeIntegers)  
model.s = Var(domain=NonNegativeIntegers) 
  
# Objective Function - Minimise TE, TD, SSE
model.obj = Objective(expr= weight_TE * sum(model.ejk[j,k] for j in groups for k in class_name) + weight_TD * sum(model.djk[j,k] for j in groups for k in class_name) + weight_SSE * model.s, sense=minimize) 
 
# Constraints

# one student assigned to only 1 group 
def rule_1(model,i): #for each student in each group  
    return sum(model.pi[i,j] for j in groups) == 1  
model.constr_1 = Constraint(stu_id, rule=rule_1) 

# constraint for total excess
def rule_te(model,j,k): #for each class for each group  
    return (sum(model.pi[i,j] * enrolled_stu.loc[i,k] for i in stu_id ) - math.floor(Rk_class.loc[k]*D)) <= model.ejk[j,k]  
model.constr_te = Constraint(groups, class_name, rule=rule_te) 

# constraint for total deviation
def rule_td1(model,j,k): #for each class for each group  
    return sum(model.pi[i,j] * enrolled_stu.loc[i,k] for i in stu_id) - (sum(enrolled_stu.loc[i,k] for i in stu_id) / len(groups)) <= model.djk[j,k] 
model.constr_td1 = Constraint(groups, class_name, rule=rule_td1) 

def rule_td2(model,j,k): #for each class for each group  
    return sum(model.pi[i,j] * enrolled_stu.loc[i,k] for i in stu_id) - (sum(enrolled_stu.loc[i,k] for i in stu_id) / len(groups)) >= -model.djk[j,k] 
model.constr_td2 = Constraint(groups, class_name, rule=rule_td2) 

# constraint for SSE
def rule_s(model,t,j): #for each time of class for each group 
    return model.s >= model.stj[t,j] - E 
model.constr_s = Constraint(time_table, groups, rule=rule_s) 

# constraint for Stj
def rule_stj(model,t,j): #for each time of class for each group  
    return sum(model.ejk[j,k] * time_TB.loc[t,k] for k in class_name) <= model.stj[t,j] 
model.constr_stj = Constraint(time_table, groups, rule=rule_stj) 

# Call the Solver and solve the problem
solver = SolverFactory('gurobi') 
results = solver.solve(model,load_solutions=False,tee=True) 
if (results.solver.status == SolverStatus.ok) and (results.solver.termination_condition == TerminationCondition.optimal): 
    model.solutions.load_from(results) 
elif (results.solver.termination_condition == TerminationCondition.maxTimeLimit): 
    print("Solve terminated due to time limit. No solution loaded.") 
else: 
    print("Solve failed. No solution loaded.")     
   
# Print student allocation     
for i in stu_id: 
    for j in groups: 
        if value(model.pi[i,j]) == 1.0: 
          print("Student ID", i, "is assigned to group:", j) 
oval = value(model.obj) 
print("Weighted excess value =", oval)  

#print TE 
total_e = 0.0 
for j in groups: 
    for k in class_name: 
        if  value(model.ejk[j,k]) > 0.0: 
            total_e += value(model.ejk[j,k]) 
print("TE is", total_e)      

#print TD   
total_delta = 0.0 
for j in groups: 
    for k in class_name:  
        if  value(model.djk[j,k]) > 0.0: 
            total_delta += value(model.djk[j,k])             
print("TD is", total_delta) 

#print SE 
max_stj_value = 0.0 
for t in time_table: 
    for j in groups: 
        if value(model.stj[t,j]) > max_stj_value: 
            max_stj_value = value(model.stj[t,j]) 
if max_stj_value > 0.0: 
    print(f"SE is {max_stj_value} in group",j, "of time", t) 
else: 
    print("SE is 0") 
    
#print SSE
max_stj_value = 0.0 
sse_value = 0.0
for t in time_table:  
    for j in groups:  
        if value(model.stj[t,j]) > max_stj_value:  
            max_stj_value = value(model.stj[t,j])
            sse_value = max_stj_value - math.floor(E*D)
if sse_value > 0.0:  
    print("SSE is", sse_value)  

else:  

    print("SSE is 0")
