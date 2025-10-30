import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import win32com.client 
import NkTi2Ax 
import time 
from pycromanager import Core
import pycromanager
from math import atan2, cos, sin
from pyAndorSpectrograph.spectrograph import ATSpectrograph
from pyAndorSDK2 import atmcd, atmcd_codes, atmcd_errors
import cv2
from scipy.signal import find_peaks

# -------------------------------------------------------
# Nikon helper functions
#--------------------------------------------------------

def coord(x0,x,state='relative'):
    """
    This does a coordinate mapping from x -> x0 where 
    x0 is the absolute cordinate 
    """
    um = 10 # Scale factor
    if state=='relative':
        return x0 + (x*um)
    elif state=='absolute':
        return x0*um
    else:
        print("Invalid state: enter 'relative' or 'absolute'")

def connect():
    """
    This estabilshes a connection with the Nikon Ti2 eclipse microscope
    """
    microscope: NkTi2Ax.NikonTi2AxAutoConnectMicroscope = win32com.client.Dispatch(NkTi2Ax.NikonTi2AxAutoConnectMicroscope.CLSID)
    microscope.DedicatedCommand(r"SHOW_SIMULATION_WINDOW",r"0,1")  
    return microscope

def get_coord(microscope):
    """
    This gets the coordinates of the manually determined origin 
    """
    xpos: NkTi2Ax.NikonTi2AxData = microscope.XPosition
    ypos: NkTi2Ax.NikonTi2AxData = microscope.YPosition
    zpos: NkTi2Ax.NikonTi2AxData = microscope.ZPosition

    return xpos.Value,ypos.Value,zpos.Value

def get_coord_var(microscope):
    """
    This gets the coordinates of the manually determined origin 
    """
    xpos: NkTi2Ax.NikonTi2AxData = microscope.XPosition
    ypos: NkTi2Ax.NikonTi2AxData = microscope.YPosition
    zpos: NkTi2Ax.NikonTi2AxData = microscope.ZPosition

    return xpos,ypos,zpos

def get_turret(microscope):
    """
    This gets the turret
    """
    turret1: NkTi2Ax.NikonTi2AxSetting = microscope.Turret1Pos
    turret2: NkTi2Ax.NikonTi2AxSetting = microscope.Turret2Pos

    return turret1,turret2

def get_turret_shutter(microscope,i):
    """
    Get's the microscope shuttet so we can turn the laser on/off on demand
    """
    shutter1: NkTi2Ax.NikonTi2AxSetting= microscope.Turret1Shutter
    shutter2: NkTi2Ax.NikonTi2AxSetting= microscope.Turret2Shutter

    shutters = {1:shutter1,
                2:shutter2}
    if i==1:
        return shutters[1]
    elif i==2:
        return shutters[2]
    else:
        print("Invalid shutter index")

def flip_mirror(microscope,i):
    """
    flip the output form the objective between the camera and spectrometer
    i = 0: Camera
    i = 1: Spectrometer
    """
    if i == 0:
        index = 4
    elif i == 1:
        index = 2
    else:
        print("Invalid index")

    LightPath: NkTi2Ax.NikonTi2AxMicroscope= microscope.LightPath
    LightPath.Value = index

def grid_test(microscope,n,time_):
    """
    Drives the microscope stage in a XY grid
    """
    x0,y0,z0 = get_coord_var(microscope)
    
    for i in range(n+1):
        if i==0:
            pass
        else:
            x,y,z = get_coord(microscope)
            x0.Value = coord(x,5)
            time.sleep(time_)

        for j in range(n):
            x,y,z = get_coord(microscope)
            if i%2==0:
                y0.Value = coord(y,5)
                time.sleep(time_)
            else:
                y0.Value = coord(y,-5)
                time.sleep(time_)

def compute_rotation_angle(p1_mic, p2_mic, p1_tgt, p2_tgt):
    # vectors
    v_m = np.array(p2_mic) - np.array(p1_mic)
    v_t = np.array(p2_tgt) - np.array(p1_tgt)
    theta_m = atan2(v_m[1], v_m[0])
    theta_t = atan2(v_t[1], v_t[0])
    # rotation from target frame to microscope frame
    theta = theta_m - theta_t
    return theta

def rotate_point(p, theta):
    c = cos(theta)
    s = sin(theta)
    x, y = p
    return np.array([c*x - s*y, s*x + c*y])

def map_target_to_microscope(Q_tgt, p1_tgt, p1_mic, theta, scale=1.0):
    # Q_tgt: (x,y) in target frame
    # p1_tgt: origin in target frame (point +1)
    # p1_mic: corresponding origin in microscope frame
    q = np.array(Q_tgt) - np.array(p1_tgt)       # translate
    q = q * scale                                # optional scale (units)
    q_rot = rotate_point(q, theta)               # rotate into microscope orientation
    Q_mic = q_rot + np.array(p1_mic)             # translate into microscope origin
    return Q_mic
 
def grid_test_trnasform(microscope,origin,n,time_, p1_tgt,
                         p1_mic, theta,step):
    """
    Drives the microscope stage in a XY grid
    """
    x0,y0,z0 = get_coord_var(microscope)
    Q_mic = map_target_to_microscope(origin, p1_tgt, p1_mic, theta)
    x0.Value = Q_mic[0]
    y0.Value = Q_mic[1]

    for i in range(n+1):
        if i==0:
            xpos = origin[0]
            ypos = origin[1]
            pass
        else:
            xpos = xpos+(step*10)
            Q_tgt = (xpos, ypos)
            Q_mic = map_target_to_microscope(Q_tgt, p1_tgt, p1_mic, theta)

            x0.Value = Q_mic[0]
            # Raman collection code goes here..
            time.sleep(time_)

        for j in range(n+1):
            if i%2==0:
                ypos = ypos+(step*10)
            else:
                ypos = ypos+(step*(-10))
            Q_tgt = (xpos, ypos)
            Q_mic = map_target_to_microscope(Q_tgt, p1_tgt, p1_mic, theta)

            y0.Value = Q_mic[1]
            # Raman collection code goes here..
            time.sleep(time_)

# -------------------------------------------------------
# Andoe helper functions
#--------------------------------------------------------

def wavelength_to_raman(wavelength_nm, excitation_nm=785.0):
    """
    Convert wavelength(s) (nm) to Raman shift (cm^-1) relative to excitation_nm (nm).
    Accepts scalar or array-like. Returns numpy array (or scalar if input scalar).
    Formula: shift = 1e7 * (1/excitation_nm - 1/wavelength_nm)
    """
    arr = np.asanyarray(wavelength_nm, dtype=float)
    shift = 1e7 * (1.0 / float(excitation_nm) - 1.0 / arr)
    # return scalar if input scalar
    if np.isscalar(wavelength_nm):
        return float(shift)
    return shift

def raman_to_wavelength(raman_cm1, excitation_nm=785.0):
    """
    Inverse: convert Raman shift (cm^-1) back to wavelength (nm).
    Formula: 1/lambda = 1/excitation - shift/1e7  -> lambda = 1 / (1/excitation - shift/1e7)
    """
    s = np.asanyarray(raman_cm1, dtype=float)
    inv = 1.0/float(excitation_nm) - s/1e7
    return (1.0 / inv)  # array or scalar as appropriate

def perform_fvb(image_2d: np.ndarray, method: str = "sum"):
    """
    Perform Full Vertical Binning (FVB) on a 2D image.
    
    Parameters
    ----------
    image_2d : np.ndarray
        2D array representing the CCD image (rows × columns).
    method : str, optional
        Binning method: "sum" (default) or "mean".
    
    Returns
    -------
    fvb_image : np.ndarray
        2D image where each row is identical and represents the FVB spectrum.
    spectrum_1d : np.ndarray
        1D array representing the FVB spectrum (intensity vs pixel).
    """
    if image_2d.ndim != 2:
        raise ValueError("Input must be a 2D NumPy array.")
    
    # Compute FVB spectrum by summing or averaging along the vertical axis (rows)
    if method == "sum":
        spectrum_1d = np.sum(image_2d, axis=0)
    elif method == "mean":
        spectrum_1d = np.mean(image_2d, axis=0)
    else:
        raise ValueError("Invalid method. Use 'sum' or 'mean'.")
    
    # Expand the 1D spectrum back to a 2D image for visualization
    fvb_image = np.tile(spectrum_1d, (image_2d.shape[0], 1))
    
    return fvb_image, spectrum_1d

def detect_peaks(x,y, height=100, distance=50,xlims=(750,2200),max_peaks=True):
    """
    Detect peaks in a 1D array.
    
    Parameters
    ----------
    y : array-like
        Input data (1D array).
    height : float or tuple, optional
        Required height of peaks. Default is None.
    distance : int, optional
        Required minimum horizontal distance (in samples) between neighboring peaks. Default is None.
        
    Returns
    -------
    peaks : ndarray
        Indices of detected peaks in the input array.
    properties : dict
        Properties of the detected peaks.
    """
    # Detect peaks
    peaks, properties = find_peaks(y, height=100, distance=50)  # adjust parameters

    peaks_df = pd.DataFrame({'Wavenumber': x[peaks],
                            'Intensity': y[peaks],
                                "index": peaks})
    peaks_df = peaks_df[(peaks_df['Wavenumber']>xlims[0]) & (peaks_df['Wavenumber']<xlims[1])]
    peaks_df.sort_values(by='Intensity', inplace=True,ascending=False)

    row = peaks_df.iloc[0]
    print(f"Highest peak at Wavenumber: {row['Wavenumber']} \nIntensity: {row['Intensity']}, \nIndex: {row['index']}")
    return peaks_df, row if max_peaks else peaks_df
    
def prepare_acquisition_(slit_width,wavelength,mode,exposure_time,spc,sdk,codes):
    print("preparing to start acquisition...")
    #Start Acquisition
    shm = spc.SetSlitWidth(0, 1,slit_width)
    shm = spc.SetWavelength(0, wavelength)
    if mode=='image':
        ret = sdk.SetReadMode(codes.Read_Mode.IMAGE)
    elif mode=='fvb':
        ret = sdk.SetReadMode(codes.Read_Mode.FULL_VERTICAL_BINNING)
    ret = sdk.SetExposureTime(exposure_time)

    print(f"Function SetExposureTime returned {ret}")

def acquire_image(mode,spc,sdk,xpixels,ypixels):
    ret = sdk.StartAcquisition()
    print("Function StartAcquisition returned {}".format(ret))
    ret = sdk.WaitForAcquisition()
    print("Function WaitForAcquisition returned {}".format(ret))

    if mode=='image':
        imageSize = xpixels * ypixels
        ret, first, last = sdk.GetNumberNewImages()
        ret, data,validfirst,validlast = sdk.GetImages16(first, last, imageSize)
        print("Function GetImages16 returned {} first pixel = {} size = {}".format(ret, data[0], imageSize))   
        img = np.reshape(data, (ypixels, xpixels))
        return img
    
    elif mode=='fvb':
        imageSize = xpixels
        ret, first, last = sdk.GetNumberNewImages()
        ret, data,validfirst,validlast = sdk.GetImages16(first, last, imageSize)

        ret, xsize, ysize = sdk.GetPixelSize()
        print("Function GetPixelSize returned {} xsize = {} ysize = {}".format(
            ret, xsize, ysize))


        shm = spc.SetNumberPixels(0, xpixels)
        print("Function SetNumberPixels returned: {}".format(
            spc.GetFunctionReturnDescription(shm, 64)[1]))

        shm = spc.SetPixelWidth(0, xsize)
        print("Function SetPixelWidth returned: {}".format(
            spc.GetFunctionReturnDescription(shm, 64)[1]))

        shm, calibrationValues = spc.GetCalibration(0, xpixels)
        print(f"Function GetCalibration returned: {spc.GetFunctionReturnDescription(shm, 64)[0]}, min: {np.min(calibrationValues)}, max: {np.max(calibrationValues)}, shape: {np.shape(calibrationValues)}")
        print(f"Data min: {np.min(data)}, Data max: {np.max(data)}, Data shape: {np.shape(data)}")

        raman_shift = wavelength_to_raman(calibrationValues, excitation_nm=786.0) 
        wavelength = calibrationValues 
        return raman_shift, data
   
def RamanSetup(slit_width,wavelength,temperature):
     
    sdk = atmcd()
    spc = ATSpectrograph()
    codes = atmcd_codes
    print("Spectrometer and Camera libraries have been loaded successfully")

    #Initialize libraries
    shm = spc.Initialize("")
    print(f"Spectrometer Initialize returned {spc.GetFunctionReturnDescription(shm, 64)[1]}")

    ret = sdk.Initialize("")
    print(f"Camera Initialize returned {ret}")

    # ----------------------------
    # Configure Camera
    # ----------------------------
    ret = sdk.SetTemperature(temperature)
    print(f"Function SetTemperature returned {ret} (target = -80°C)")

    ret = sdk.CoolerON()
    print(f"Function CoolerON returned {ret}")

    # Stabilize temperature
    print("Stabilizing temperature...")
    while ret != atmcd_errors.Error_Codes.DRV_TEMP_STABILIZED:
            time.sleep(5)
            ret, temperature = sdk.GetTemperature()
            print(f"Function GetTemperature returned {ret}, current temperature = {temperature}", end="\r")

    print("\nTemperature stabilized")

    print("Configuring acquisition parameters...")
    # Setting acquisition parameters
    ret = sdk.SetReadMode(codes.Read_Mode.IMAGE)
    # ret = sdk.SetReadMode(codes.Read_Mode.FULL_VERTICAL_BINNING)
    print(f"Function SetReadMode = IMAGE returned {ret}")

    ret = sdk.SetTriggerMode(codes.Trigger_Mode.INTERNAL)
    print(f"Function SetTriggerMod=INTERNAL returned {ret}")

    ret = sdk.SetAcquisitionMode(codes.Acquisition_Mode.SINGLE_SCAN)
    print(f"Function SetAcquisitionMode=SINGLE_SCAN returned {ret}")

    ret, xpixels, ypixels = sdk.GetDetector()
    print(f"Function GetDetector returned {ret}: x = {xpixels}, y = {ypixels}")

    ret = sdk.SetImage(1, 1, 1, xpixels, 1, ypixels)
    print(f"Function SetImage returned {ret}")

    # ----------------------------
    # Configure Spectrograph
    # ----------------------------
    print("Configuring spectrograph parameters..." )
    shm = spc.SetGrating(0, 1)
    print(f"Function SetGrating=1 returned {spc.GetFunctionReturnDescription(shm, 64)[1]}")

    shm, grat = spc.GetGrating(0)
    print(f"Function GetGrating returned {grat}")

    shm = spc.SetWavelength(0, wavelength)
    print(f"Function SetWavelength=0 returned {spc.GetFunctionReturnDescription(shm, 64)[1]}")

    shm, wave = spc.GetWavelength(0)
    print(f"Function GetWavelength returned: {spc.GetFunctionReturnDescription(shm, 64)[1]}, wavelength = {wave}")

    shm, wl_min, wl_max = spc.GetWavelengthLimits(0, grat)
    print(
        f"Function GetWavelengthLimits returned {spc.GetFunctionReturnDescription(shm, 64)[1]}, "
        f"min = {wl_min}, max = {wl_max}"
    )

    # Set slit width for spectrograph 0, slit 1, width 100
    shm, present = spc.IsSlitPresent(0, 1)
    print(f"IsSlitPresent: {spc.GetFunctionReturnDescription(shm,64)[1]}, present = {present}")

    if not present:
        print("Slit not present or not motorised. Cannot set slit width.")
    else:
        # request move and check status
        shm = spc.SetSlitWidth(0, 1, slit_width)
        print(f"SetSlitWidth returned: {spc.GetFunctionReturnDescription(shm,64)[1]}")

        ret, width = spc.GetSlitWidth(0, 1)
        print(f"GetSlitWidth returned {ret}, width = {width}")

    # ensure spectrograph knows the number of pixels
    shm = spc.SetNumberPixels(0, xpixels)
    print("SetNumberPixels:", spc.GetFunctionReturnDescription(shm, 64)[1])

    # get camera pixel physical size and set pixel width in spectrograph
    ret, xsize, ysize = sdk.GetPixelSize()
    print("GetPixelSize:", ret, "xsize=", xsize, "ysize=", ysize)
    shm = spc.SetPixelWidth(0, xsize)
    print("SetPixelWidth:", spc.GetFunctionReturnDescription(shm, 64)[1])

    # request calibration and inspect full return
    shm, calibrationValues = spc.GetCalibration(0, xpixels)
    print("GetCalibration status:", spc.GetFunctionReturnDescription(shm, 64)[1])
    print("calibrationValues type:", type(calibrationValues), "len:", len(calibrationValues))

    print('spectrograph configured successfully')
    
    return spc,sdk,codes,xpixels, ypixels

def OptimizeRamanIntensity():
    import time
    import numpy as np

    # ---- USER SETTINGS ----
    step_size = 0.2          # µm — how much to move when intensity drops
    exposure_time = 0.5      # seconds between measurements
    z_pos = 0.0              # current Z position
    last_signal = 0.0
    direction = 1             # +1 = move up, -1 = move down

    # ---- REPLACE THESE WITH YOUR HARDWARE COMMANDS ----
    def move_stage(z):
        """Move stage to position z (µm)."""
        # e.g., stage.move_to(z)
        print(f"Moving to z = {z:.2f} µm")

    def read_signal():
        """Measure Raman signal intensity."""
        # Replace this with real acquisition code
        true_focus = 10.0
        noise = np.random.normal(0, 0.02)
        return np.exp(-0.5 * ((z_pos - true_focus)/1.2)**2) + noise
    # ----------------------------------------------------

    print("Starting continuous Raman focus optimization... (Ctrl+C to stop)")

    try:
        while True:
            # Measure signal
            signal = read_signal()
            print(f"z = {z_pos:.2f} µm → signal = {signal:.3f}")

            # Compare with last measurement
            if signal < last_signal:  # got worse → reverse direction
                direction *= -1
                step_size *= 0.9      # optionally shrink step slightly for stability

            # Update position
            z_pos += direction * step_size
            move_stage(z_pos)

            # Prepare for next loop
            last_signal = signal
            time.sleep(exposure_time)

    except KeyboardInterrupt:
        print("\nOptimization stopped.")



































# # Example usage:
# p1_mic = (1000.0, 200.0)      # measured microscope coordinates of +1 (µm or stage units)
# p2_mic = (1500.0, 210.0)      # measured microscope coordinates of +2
# p1_tgt = (0.0, 0.0)           # +1 in target frame (pick origin)
# p2_tgt = (100.0, 0.0)         # +2 is supposed to be +100 units along target x-axis

# theta = compute_rotation_angle(p1_mic, p2_mic, p1_tgt, p2_tgt)
# print("Rotation (rad):", theta, "deg:", np.degrees(theta))

# # Suppose we want to send Q_tgt=(100, 0) (i.e., +2) to the microscope:
# Q_tgt = (100.0, 0.0)
# Q_mic = map_target_to_microscope(Q_tgt, p1_tgt, p1_mic, theta)
# print("Microscope coords to move to:", Q_mic)