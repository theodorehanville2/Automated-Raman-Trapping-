import numpy as np
import time
import pandas as pd
import matplotlib.pyplot as plt
from pyAndorSpectrograph.spectrograph import ATSpectrograph
from pyAndorSDK2 import atmcd, atmcd_codes, atmcd_errors
import cv2
import utilities as ut
import win32com.client 
import NkTi2Ax 

print('initializing Nikon...')

# Establish connection with microscope
print('connecting microscope')
microscope = ut.connect()
print('connected')

# Get origin 
# print('getting origin')
x0,y0,z0 = ut.get_coord_var(microscope)
print(f'Origin: x0 = {x0}, y0 = {y0}, z0 = {z0}')

# Get turrets
print('testing turrets...')
turret2,turret1 = ut.get_turret(microscope)

ut.flip_mirror(microscope,0)
time.sleep(0.5)

print('Nikon initialization completed')
#----------------------------------------------------------------------------------
# User parameters
#----------------------------------------------------------------------------------
exposure_time = 3
slit_width = 100
acquisition_type = 'fvb'
wavelength = 890
temperature = -80

spc,sdk,codes,xpixels, ypixels = ut.RamanSetup(slit_width,wavelength,temperature)

if acquisition_type=='image':
     ut.prepare_acquisition_(2000,0,'image',0.1,spc,sdk,codes)
elif acquisition_type=='fvb':
     ut.prepare_acquisition_(slit_width,wavelength,'fvb',exposure_time,spc,sdk,codes)

plt.ion()
fig, ax = plt.subplots()
ax.set_xlim(600, 2200)
ax.set_xlabel("Raman shift")
ax.set_ylabel("Intensity")
ax.set_title("Raman Spectrum")
line, = ax.plot([], [], lw=1, color='blue')
plt.grid(True)

# Force initial draw
fig.canvas.draw()
plt.show(block=False)

# initialize offset positions
z0.Value = z0.Value + (-30*10) 

ut.flip_mirror(microscope,1)
turret2.Value = 5
time.sleep(1)

for i in range(1,61):

    # Move in 1 micron steps
    z0.Value += 1*10  
    time.sleep(0.2)

    # acquire Raman spectrum 
    raman_shift, data = ut.acquire_image('fvb', spc, sdk, xpixels, ypixels)

    line.set_data(raman_shift, data)
    ax.relim()
    ax.autoscale_view()

    raman_df = pd.DataFrame({'Raman Shift': raman_shift, f'Intensity': data})
    raman_df.to_csv(f'Raman_Zsweep_Position_{i}_Z_microns.csv', index=False)

    fig.canvas.draw()
    fig.canvas.flush_events()
    print(f'Collecting Raman at {i}, Z = {z0.Value} microns')

ut.flip_mirror(microscope,0)
turret2.Value = 3
plt.ioff()
plt.show()
