import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import win32com.client 
import NkTi2Ax 
import time 
from pycromanager import Core
import pycromanager
import utilities as ut

print('initializing...')

# Establish connection with microscope
print('connecting microscope')
microscope = ut.connect()
print('connected')

# Get origin 
# print('getting origin')
x0,y0,z0 = ut.get_coord_var(microscope)
print(f'Origin: x0 = {x0}, y0 = {y0}, z0 = {z0}')
# print(dir(x0))

# Translate stage
print('translating stage...')
ut.grid_test(microscope,3,0.2)
print('translation complete')

# Get shutters
print('testing shutter')
shutter = ut.get_turret_shutter(microscope,2)
shutter.Value = 0
time.sleep(0.5)
shutter.Value = 1
print('passed')

# Get turrets
print('testing turrets...')
turret1,turret2 = ut.get_turret(microscope)

# test turrets
for i in range(1,7):
    turret1.Value = i
    time.sleep(0.5)

for i in range(1,7):
    turret2.Value = i
    time.sleep(0.5)
print('Passed...')

# test LightPath
for i in range(2):
    ut.flip_mirror(microscope,i)
    time.sleep(1)

print('passed')

print('Initializing completed')

