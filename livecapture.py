import numpy as np
import time
import pandas as pd
import matplotlib.pyplot as plt
from pyAndorSpectrograph.spectrograph import ATSpectrograph
from pyAndorSDK2 import atmcd, atmcd_codes, atmcd_errors
import cv2
import utilities as ut
from matplotlib.animation import FuncAnimation

#----------------------------------------------------------------------------------
# User parameters
#----------------------------------------------------------------------------------
exposure_time = 0.1
slit_width = 100
acquisition_type = 'fvb'
wavelength = 0
temperature = -80

spc,sdk,codes,xpixels, ypixels = ut.RamanSetup(slit_width,wavelength,temperature)

#-----------------------
#Start Acquisition
#-----------------------
acquisition_type = 'fvb'

if acquisition_type=='image':
     ut.prepare_acquisition_(2000,0,'image',0.1,spc,sdk,codes)
elif acquisition_type=='fvb':
     ut.prepare_acquisition_(100,890,'fvb',3,spc,sdk,codes)

plt.ion()
fig, ax = plt.subplots()
ax.set_xlim(600, 2200)
ax.set_xlabel("Raman shift")
ax.set_ylabel("Intensity")
ax.set_title("Raman Spectrum")
line, = ax.plot([], [], lw=1, color='blue')

# Force initial draw
fig.canvas.draw()
plt.show(block=False)

while True:

    if acquisition_type == 'image':
        img = ut.acquire_image('image', spc, sdk, xpixels, ypixels)
        cv2.imshow("Image", img)
        key = cv2.waitKey(1)
        if key & 0xFF == ord('q'):
            break

    elif acquisition_type == 'fvb':
        raman_shift, data = ut.acquire_image('fvb', spc, sdk, xpixels, ypixels)

        line.set_data(raman_shift, data)
        ax.relim()
        ax.autoscale_view()

        fig.canvas.draw()
        fig.canvas.flush_events()

    else:
        print('Invalid acquisition type')
        break

cv2.destroyAllWindows()
plt.ioff()
plt.show()