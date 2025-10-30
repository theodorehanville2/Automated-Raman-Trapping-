import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import win32com.client 
import NkTi2Ax 
import time 
from pycromanager import Core
import pycromanager
import utilities as ut
import inspect
# from pyvcam import pvc
# from pyvcam.camera import Camera

print('initializing...')

# Establish connection with microscope
print('connecting microscope')
microscope = ut.connect()
print('connected')

# Point 
x0,y0,z0 = ut.get_coord_var(microscope)
print(f'P2: x0 = {x0}, y0 = {y0}, z0 = {z0}')

# Coordinate trnasform (Alignment)
p1_mic = (59249, -263044)      # measured microscope coordinates of P1 (µm or stage units)
p2_mic = (60780, -263011)      # measured microscope coordinates of P2
p1_tgt = (0.0, 0.0)           # P1 in target frame (pick origin)
p2_tgt = (1733, 0.0)         # P2 in target frame

# Compute rotation angle
theta = ut.compute_rotation_angle(p1_mic, p2_mic, p1_tgt, p2_tgt)
print("Rotation (rad):", theta, "deg:", np.degrees(theta))

time_ = 1
n = 5
origin = (0,0)
step = 17.33

ut.grid_test_trnasform(microscope,origin,n,time_, p1_tgt,
                         p1_mic, theta,step)

# # Translate stage
# print('translating stage...')
# ut.grid_test(microscope,5,0.5)
# print('translation complete')

# grid_test_trnasform(microscope,(0,0),5,1, p1_tgt,
#                          p1_mic, theta,17.33)

# # Get shutters
# print('testing shutter')
# shutter = ut.get_turret_shutter(microscope,2)
# shutter.Value = 0
# time.sleep(0.5)
# shutter.Value = 1
# print('passed')

# # Get turrets
# print('testing turrets...')
# turret1,turret2 = ut.get_turret(microscope)

# # test turrets
# for i in range(1,7):
#     turret1.Value = i
#     time.sleep(0.5)

# for i in range(1,7):
#     turret2.Value = i
#     time.sleep(0.5)
# print('Passed...')

# # test LightPath
# for i in range(2):
#     ut.flip_mirror(microscope,i)
#     time.sleep(1)

# print('passed')

# print('Initializing completed')

