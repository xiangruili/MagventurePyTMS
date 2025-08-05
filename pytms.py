#!/usr/bin/env python
# 250714 "translated" Matlab TMS.m (Xiangrui.Li at gmail.com)
# 250725 add TMS_GUI part
# 250727 add components for motor threshold (need matplotlib)

import serial # pySerial is the only non-builtin module for TMS and GUI
from serial.tools.list_ports import comports
import json
import threading

import tkinter as tk
from tkinter import filedialog,ttk # ttk for Combobox only
from idlelib.tooltip import Hovertip

try: # allow TMS() and TMS_GUI() to work without matplotlib
    import matplotlib as mpl
    mpl.rcParams['toolbar'] = 'None'
    import matplotlib.pyplot as plt
    from matplotlib.widgets import Button
    import numpy as np
except:
    print("matplotlib module needed for motor threshold test")

path = serial.os.path
sleep = serial.time.sleep

__version__ = "2025.08.05"
__all__ = ['TMS', 'TMS_GUI', 'rMT', 'EMGCheck']

class TMS:
    """TMS controls Magventure TMS system (tested under X100).
    
    Usage syntax:
     from pytms import TMS
     T = TMS() # Connect to Magventure stimulator and return handle for later use
     T.disp() # Display most parameters
     T.load('myParams.mat') # load pre-saved stimulation parameters
     T.enabled = True # Enable it like pressing the button at stimulator
     T.amplitude = 60 # set amplitude to 60%
     T.firePulse() # send a single pulse/burst stimulation
    
     T.waveform = "Biphasic Burst" # set wave form
     T.burstPulses = 3 # set number of pulses in a burst, 2 to 5
     T.IPI = 20 # set inter pulse interval in ms
     T.train.RepRate= 5 # etc., set train parameter
      Here are important train parameters:
        RepRate: number of pulses per second in each train
        PulsesInTrain: number of pulses or bursts in each train
        NumberOfTrains: number of trains in the sequence
        ITI: Seconds between last pulse and first pulse in next train
     T.fireTrain() # Start train of pulse/burst stimulation
    """
    
    # first 8 in __slots__ are used in save() & load()
    __slots__ = ('_mode', '_currentDirection', '_waveform', '_burstPulses', '_IPI', 
        '_BARatio', '_delays', '_CoilTypeDisplay', '_ExRate', '_MEP', '_Model', 
        '_amplitude', '_didt', '_enabled', '_filename', '_info', '_page', '_port', 
        '_temperature', '_train', '_trainRunning', '_MODELs', '_TCs', '_PAGEs')
    _INS = None # store the single instance

    def __new__(cls):
        if TMS._INS: return TMS._INS # reuse existing instance
        return super().__new__(cls) # create new instance

    def __del__(self):
        TMS._INS = None
        try: self._port.close()
        except: pass

    def __init__(self):
        if TMS._INS: return
        TMS._INS = self
        self._MODELs = ("R30", "X100", "R30+Option", "X100+Option", "R30+Option+Mono", "MST")
        self._TCs = ("Sequence", "External Trig", "Ext Seq. Start", "Ext Seq. Cont")
        self._PAGEs = {1:"Main", 2:"Timing", 3:"Trig", 4:"Config", 6:"Download", \
                       7:"Protocol", 8:"MEP", 13:"Service", 15:"Treatment", \
                       16:"Treat Select", 17:"Service2", 19:"Calculator"}
        self._ExRate = False # ExtendedRepRate with 0.1 step
        self._enabled = False
        self._amplitude = (0,0)
        self._didt = (0,0)
        self._temperature = 21
        self._MEP = {}
        self._info = {"CoilType":""}
        self._filename = ""
        self._Model = "X100" # default to make Scales happy
        self._mode = "Standard"
        self._currentDirection = "Normal"
        self._waveform = "Biphasic"    
        self._BARatio = 1.0
        self._burstPulses = 2
        self._IPI = 10.0
        self._delays = [0.0, 0.0, 0]
        self._page = "Main"
        self._CoilTypeDisplay = True
        self._trainRunning = False
        self._train = trainParam()

        for p in comports():
            if "bluetooth" in p[0].lower(): continue
            try: ser = serial.Serial(p[0], 38400, timeout=0.3)
            except: continue
            ser.write(bytes((254,1,0,0,255))); sleep(0.1)
            if ser.in_waiting>12: break
            ser.close()

        if 'ser' not in locals() or not ser.is_open:
            print("Failed to connect to stimulator.")
            return # warn rather than error so code can run with no connection

        self._port = ser
        def worker(self):
            while True:
                try: n0 = self._port.in_waiting
                except: return
                sleep(0.05)
                try: n1 = self._port.in_waiting
                except: return
                if n1 == n0 > 7: self._read()
                sleep(0.1)

        thread = threading.Thread(target=worker, args=(self,))
        thread.daemon = True  # Allows the main program to exit
        thread.start()
        self.resync()

    def firePulse(self):
        '''Start single pulse or single burst stimulation'''
        assert self._enabled, "Need to enable"
        self._write((3,))

    def fireTrain(self):
        '''Start train stimulation'''
        assert self._enabled, "Need to enable"
        self.page = "Timing"
        self._write((4,))
        
    def fireProtocol(self):
        '''Start stimulation in Protocol'''
        assert self._enabled, "Need to enable"
        self.page = "Protocol"
        self._write((4,))

    def disp(self):
        '''Display major status / parameters'''
        print(f'        enabled: {self._enabled}')
        print(f'      amplitude: {self._amplitude}')
        print(f'           didt: {self._didt}')
        print(f'    temperature: {self._temperature}')
        print(f'           info: {self._info}')
        print(f'           mode: {self._mode}')
        print(f'       waveform: {self._waveform}')
        print(f'    burstPulses: {self._burstPulses}')
        print(f'            IPI: {self._IPI:g}')
        print(f' Train:')
        print(f'        RepRate: {self.train._RepRate:g}')
        print(f'  PulsesInTrain: {self.train._PulsesInTrain}')
        print(f' NumberOfTrains: {self.train._NumberOfTrains}')
        print(f'            ITI: {self.train._ITI:g}')
        print(f'      trainTime: {self.trainTime}')
        print(f'   trainRunning: {self._trainRunning}')

    @property
    def amplitude(self):
        '''Show and control stimulation amplitude in percent'''
        return self._amplitude

    @amplitude.setter
    def amplitude(self, amp):
        if type(amp) not in (list,tuple): amp = (amp, self._amplitude[1])
        amp = tuple(amp)
        assert 0<=amp[0]<=100 and isinstance(amp[0], int), "amplidute must be integer from 0 to 100."
        self._amplitude = amp
        self._write((1,)+amp)
        
    @property
    def enabled(self):
        '''Show and control stimulator is enabled or disabled'''
        return self._enabled

    @enabled.setter
    def enabled(self, tf):
        self._enabled = bool(tf)
        self._write((2,self._enabled));

    @property
    def didt(self):
        '''Realized di/dt in A/µs, as shown on the stimulator'''
        return self._didt

    @property
    def temperature(self):
        '''Coil temperature as shown on the stimulator'''
        return self._temperature

    @property
    def info(self):
        '''Some info/status: SerialNo, Connected coil, etc.'''
        return self._info

    @property
    def MEP(self):
        '''Info for MEP if available'''
        return self._MEP

    @property
    def filename(self):
        '''File name from which the parameters are loaded'''
        return self._filename

    @property
    def Model(self):
        '''Stimulator model'''
        return self._Model

    @property
    def mode(self):
        '''Stimulator mode: "Standard" "Power" "Twin" or "Dual", depending on model'''
        return self._mode

    @mode.setter
    def mode(self, arg):
        if self._mode==arg: return
        k = key(self._MODEs, arg)
        assert k>=0, "Valid mode input: "+lists(self._MODEs)
        self._mode = arg
        self._setParam9()

    @property
    def currentDirection(self):
        '''Current direction: "Normal" or "Reverse" if available'''
        return self._currentDirection

    @currentDirection.setter
    def currentDirection(self, arg):
        if self._curDirs==arg: return
        k = key(self._curDirs, arg)
        assert k>=0, "Valid CurrentDirection input: "+lists(self._curDirs)
        self._currentDirection = arg
        self._setParam9()

    @property
    def waveform(self):
        '''Wave form: "Monophasic" "Biphasic" "HalfSine" or "Biphasic Burst"'''
        return self._waveform

    @waveform.setter
    def waveform(self, arg):
        if self._waveform==arg: return
        k = key(self._wvForms, arg)
        assert k>=0, "Valid waveform input: "+lists(self._wvForms)
        self._waveform = arg
        self._setParam9()

    @property
    def burstPulses(self):
        '''Number of pulses in a burst: 2 to 5'''
        return self._burstPulses

    @burstPulses.setter
    def burstPulses(self, arg):
        if self._burstPulses==arg: return
        assert arg in (2,3,4,5), "burstPulses must be 2 to 5"
        self._burstPulses = arg;
        self._setParam9()

    @property
    def IPI(self):
        '''Inter pulse interval in ms. Will be adjusted to supported value'''
        return self._IPI

    @IPI.setter
    def IPI(self, arg):
        if self._IPI==arg: return
        self._IPI,dev = closestVal(arg, self._IPIs)
        if dev>0.001: print(f"IPI adjusted to {self._IPI}")
        self._setParam9()

    @property
    def BARatio(self):
        '''Pulse B/A Ratio. Will be adjusted to supported value'''
        return self._BARatio

    @BARatio.setter
    def BARatio(self, arg):
        if self._BARatio==arg: return
        self._BARatio,dev = closestVal(arg, frange(0.2, 5.01, 0.05))
        if dev>0.001: print(f"BARatio adjusted to {self._BARatio}")
        self._setParam9()

    @property
    def page(self):
        '''Current page on the stimulator'''
        return self._page

    @page.setter
    def page(self, arg):
        k = key(self._PAGEs, arg);
        assert k>=0, "Valid page input: "+lists(self._PAGEs)
        self._page = arg
        self._write((7, k))
        self._write((12, 0))
        if self._page != arg: print("Failed to switch to page: "+arg)

    @property
    def delays(self):
        '''[DelayInputTrig, DelayOutputTrig, ChargeDelay] in ms'''
        return self._delays

    @delays.setter
    def delays(self, arg):
        if self._delays==arg: return
        di,_ = closestVal(arg[0], (*frange(0, 10.01, 0.1), *range(11,101), *range(110,6501,10)))
        do,_ = closestVal(arg[1], (*range(-100,-9), *frange(-9.9, 10.01, 0.1), *range(11, 101)))
        dc,_ = closestVal(arg[2], (*range(0, 101, 10), *range(125, 4001, 25), *range(4050, 12001, 50)))
        self._delays = [di,do,dc]
        self._write((10,1)+int2byte((di*10, round(do*10)&0xffff, dc)))
        self._write((10,0))

    @property
    def CoilTypeDisplay(self):
        '''Display coil type if true. For now must update manually if changed on Stimulator'''
        return self._CoilTypeDisplay

    @CoilTypeDisplay.setter
    def CoilTypeDisplay(self, arg): # only set by user for now
        self._CoilTypeDisplay = bool(arg);
        self._write((0,)) # update info.CoilType

    @property
    def train(self):
        '''Train parameters:
         RepRate: number of pulses per second in each train
         PulsesInTrain: number of pulses or bursts in each train
         NumberOfTrains: number of trains in the sequence
         ITI: Seconds between last pulse and first pulse in next train
         PriorWarningSound: beeps to remind the start of each train'''
        return self._train

    @property
    def _IPIs(self):
        if self._mode not in ["Twin", "Dual"]:
            return (*frange(0.5, 10.01, 0.1), *frange(10.5, 20.1, 0.5), *range(21, 101))[::-1]
        i0 = 2 if self._waveform == "Monophasic" else 1
        return (*frange(i0, 10.01, 0.1), *frange(10.5, 20.1 ,0.5), *range(21, 101), \
                *range(110, 501, 10), *range(550, 1001, 50), *range(1100, 3001, 100))[::-1]

    @property
    def _RATEs(self):
        L,H = (1,100)
        if self._Model == "R30":
            L,H = (20,30) if self._ExRate else (1,30)
        elif self._Model == "R30+Option":
            if self._mode in ["Twin", "Dual"]: L,H = (5,0) if self._ExRate else (1,5)
            else: L,H = (1, 30)
        elif self._Model == "X100":
            if self._waveform == "Biphasic Burst": L,H = (20,0) if self._ExRate else (1,20)
            elif self._ExRate: L,H = (20,100) 
        elif self._Model == "X100+Option":
            if self._waveform == "Biphasic Burst": L,H = (20,0) if self._ExRate else (1,20)
            elif self._mode in ["Twin", "Dual"]:
                if self._waveform == "Monophasic": L,H = (5,0) if self._ExRate else (1,5)
                else: L,H = (20,50) if self._ExRate else (1,50)
            elif self._ExRate: L,H = (20,100)
        else:
            print(f"Scales for {self._Model} is not supported for now.")
        return (*frange(0.1, L+0.01, 0.1), *range(L+1, H+1))
 
    @property
    def _MODEs(self):
        if self._Model in ("R30", "X100"): return {0:"Standard"}
        elif self._Model == "R30+Option": return {0:"Standard", 2:"Twin", 3:"Dual"}
        else: return {0:"Standard", 1:"Power", 2:"Twin", 3:"Dual"}
    
    @property
    def _curDirs(self):
        if self._Model in ("R30", "R30+Option"): return {0:"Normal"}
        else: return {0:"Normal", 1:"Reverse"}
    
    @property
    def _wvForms(self):
        if self._Model == "R30": return {1:"Biphasic"}
        elif self._Model == "R30+Option": return {0:"Monophasic", 1:"Biphasic"}
        elif self._Model == "X100": return {0:"Monophasic", 1:"Biphasic", 3:"Biphasic Burst"}
        else: return {0:"Monophasic", 1:"Biphasic", 2:"Halfsine", 3:"Biphasic Burst"}

    @property
    def trainRunning(self):
        '''Indicate if train sequence is running'''
        return self._trainRunning

    @property
    def trainTime(self):
        '''Total time to run the sequence, based on train parameters'''
        S = self.train
        ss = ((S._PulsesInTrain-1)/S._RepRate + S._ITI) * S._NumberOfTrains - S._ITI
        mm,ss = divmod(round(ss), 60)
        return f"{mm:02d}:{ss:02d}"

    def resync(self):
        '''Update the parameters from stimulator, in case changes at stimulator'''
        self._write((5,)) # basic info
        self._write((9, 0)) # burst parameters etc
        self._write((10, 0)) # delays
        self._write((11, 0)) # train
        self._write((12, 0)) # page, stimCount

    def disconnect(self):
        '''Release the associated serial port, so other app can connect'''
        self.__del__()
    
    def save(self, fName=None):
        '''Save parameters to a file for future sessions to load'''
        if fName is None:
            root = tk.Tk()
            root.withdraw() # Hide the main window
            fName = filedialog.asksaveasfilename(title="Specify a file to save", \
                filetypes=[("JSON files", "*.json"),])                
            if not fName: return
        
        D = dict()
        for f in self.__slots__[:8]: D[f[1:]] = getattr(self, f)
        D["train"] = dict()
        for f in self.train.__slots__: D["train"][f[1:]] = getattr(self.train, f)
        with open(fName, "w") as file: json.dump(D, file, indent=4)
        
    def load(self, fName=None):
        '''Load parameters in .json/.CG3 file and set to stimuluator'''
        if fName is None:
            root = tk.Tk()
            root.withdraw() # Hide the main window
            fName = filedialog.askopenfilename(title="Select a file to load", \
                filetypes=((".CG3, .json files", "*.json *.CG3"), ))
            if not fName: return

        if fName.lower().endswith(".json"):
            with open(fName, "r") as file: D = json.load(file)
            assert "IPI" in D, "Invalid parameter file"
            for f in self.__slots__[:6]: setattr(self, f, D[f[1:]])
            self.delays = D["delays"]
            self.CoilTypeDisplay = D["CoilTypeDisplay"]
            for f in self.train.__slots__: setattr(self.train, f, D["train"][f[1:]])
        elif fName.upper().endswith(".CG3"):
            with open(fName, "r") as file: ch = file.read()
            ch = tk.re.split(r"\[protocol Line \d", ch)
            def getval(k): return int(tk.re.search(k+r"=(\d+)",ch[0]).groups()[0])
            self._mode = self._MODEs[getval("Mode")]
            self._currentDirection = self._curDirs[getval("Current Direction")]
            self._waveform = self._wvForms[getval("Wave Form")]
            a = getval("Burst Pulses")
            if 1<a<6: self._burstPulses = a
            self._IPI = getval("Inter Pulse Interval")/10
            self._BARatio = getval("Pulse BA Ratio")/100
            d3 = [getval(k) for k in ("Delay Input Trig", "Delay Output Trig", "Charge Delay")]
            self.delays = [d3[0]/10, d3[1]/10, d3[2]]
            self.CoilTypeDisplay = getval("Coil Type Display")
            self.train._TimingControl = self._TCs[getval("Timing Control")]
            self.train._RepRate = getval("Rep Rate")/10
            self.train._PulsesInTrain = getval("Pulses in train")
            self.train._NumberOfTrains = getval("Number of Trains")
            self.train._ITI = getval("Inter Train Interval")/10
            self.train._PriorWarningSound = getval("Prior Warning Sound")>0
            try: # .CG3 may not have these 2
                self.train._RampUp = getval("RampUp")/100
                self.train._RampUpTrains = getval("RampUpTrains")
            except: pass
        else: assert [], "Unsupported file type"

        self._filename = fName
        self._setParam9()
        self._setTrain()
        self.resync()
        if TMS_GUI._INS: TMS_GUI._INS.update() # update GUI if exists

    def _write(self, b):
        # print(f"Sent {b}") # for debug
        try: self._port.write(bytes((254,len(b))+b+(CRC8(b),255)))
        except: pass # quiet if fail

    def _setParam9(self): # Shortcut to set parameters via commandID=9
        b9 = [0]*9
        b9[0] = self._MODELs.index(self._Model)
        b9[1] = key(self._MODEs, self._mode)
        b9[2] = key(self._curDirs, self._currentDirection)
        b9[3] = key(self._wvForms, self._waveform)
        b9[4] = 5 - self._burstPulses
        b9[5:7] = int2byte(self._IPI*10)
        b9[7:9] = int2byte(self._BARatio*100)
        self._write((9, 1)+tuple(b9))
        self._write((9, 0)) # sync

    def _setTrain(self): # Shortcut to set train parameters
        self.page = "Timing" # save time for fireTrain()
        S = self.train
        tc = self._TCs.index(S._TimingControl) # +self._ExRate*16 # not settable
        b16 = (S._RepRate*10, S._PulsesInTrain, S._NumberOfTrains, S._ITI*10)
        # control for RampUp/RampUpTrains not working
        self._write((11, 1, tc)+int2byte(b16)+(int(S._PriorWarningSound),))
        self._write((11, 0)) # sync

    def _setCoilType(self, k):
        if not self._CoilTypeDisplay: coil = "Hidden"
        elif k==60: coil = "Cool-B65"
        elif k==72: coil = "C-B60"
        # elif k==72: coil = "" # add more pairs
        else: coil = str(k)
        self._info["CoilType"] = coil

    def _read(self): # Update parameters from stimulator
        byts = self._port.read_all()
        i0 = [i for i,val in enumerate(byts) if val==254]
        WvForms = ("Monophasic", "Biphasic", "Halfsine", "Biphasic Burst")
        def by2int(b, signed=False): return int.from_bytes(b, 'big', signed=signed)
        for i in range(len(i0)):
            try: b = byts[i0[i]:i0[i]+byts[i0[i]+1]+4]
            except: continue
            # valid packet: [254 len(b)-4 b[2:-2] CRC8(b[2:2]) 255]
            if b[-1]!=255 or b[-2]!=CRC8(b[2:-2]): continue
            # print(f" {list(b[2:-2])}") # for debug
            if b[2] in (0,5):
                def bit2int(i,n): return (b[3]>>i)&((1<<n)-1)
                self._mode = self._MODEs[bit2int(0,2)]
                self._waveform = WvForms[bit2int(2,2)]
                self._enabled = bit2int(4,1)>0
                self._Model = self._MODELs[bit2int(5,3)]
                self._info["SerialNo"] = by2int(b[4:7])
                self._temperature = b[7]
                self._setCoilType(b[8])
                self._amplitude = tuple(b[9:11])
                if b[2]==5: # Localite sends 5 twice a second
                    # amplitudeOriginal = tuple(b[11:13])
                    # self.protocol(1).AmplitudeAGain = b[13]/100
                    # self._page = self._PAGEs[b[15]] # not reliable for some
                    self._trainRunning = b[16]>0
            elif b[2] in (1,2,3,6,7,8):
                if b[2]==1: # amplitude & enable/disable
                    self._amplitude = tuple(b[3:5])
                elif b[2]==2: # fire pulse/train/protocol
                    self._didt = tuple(b[3:5])
                elif b[2]==3: # enable/disable
                    self._temperature = b[3]
                    self._setCoilType(b[4])
                elif b[2]==6: # only at Protocol page
                    print(f"amplitudeOriginal = {b[3]} {b[4]}")
                elif b[2]==7: # only at Protocol page
                    print(f"protocol.AmplitudeAGain = {b[3]/100}")
                elif b[2]==8: # page change & train stim
                    self._page = self._PAGEs[b[3]]
                    self._trainRunning = b[4]>0
                def bit2int(i,n): return (b[5]>>i)&((1<<n)-1)
                self._mode = self._MODEs[bit2int(0,2)]
                self._waveform = WvForms[bit2int(2,2)]
                self._enabled = bit2int(4,1)>0
                # self._Model = self._MODELs(bit2int(5,3))
            elif b[2]==4:
                self._MEP["maxAmplitude"] = by2int(b[3:7]) # in µV
                self._MEP["minAmplitude"] = by2int(b[7:11])
                self._MEP["maxTime"] = by2int(b[11:15]) # in µs
            elif b[2]==9: # basic param
                # self._Model = self._MODELs[b[4]]
                self._mode = self._MODEs[b[5]]
                self._currentDirection = self._curDirs[b[6]]
                self._waveform = WvForms[b[7]]
                self._burstPulses = 5 - b[8]
                if self._waveform=="Biphasic Burst": self._IPI = self._IPIs[b[10]*256+b[9]]
                if self._mode=="Twin": self._BARatio = 5 - b[11]*0.05
            elif b[2]==10: # delay
                self._delays = [by2int(b[4:6])/10, by2int(b[6:8],True)/10, by2int(b[8:10])]
            elif b[2]==11: # Timing menu params
                self.train._TimingControl = self._TCs[b[4]&15]
                self.train._RepRate = by2int(b[5:7])/10
                self.train._PulsesInTrain = by2int(b[7:9])
                self.train._NumberOfTrains = by2int(b[9:11])
                self.train._ITI = by2int(b[11:13])/10
                self.train._PriorWarningSound = b[13]>0
                self.train._RampUp = b[14]/100
                self.train._RampUpTrains = b[15]
                self._ExRate = b[4]>>4&1>0
            elif b[2]==12:
                # self._NumberOfTrains = by2int(b[4:6])
                # by2int(b[7:10]): PulsesInTrain * NumberOfTrains
                self._info["stimCount"] = by2int(b[11:13]) # b[9:11] too?
                self._page = self._PAGEs[b[15]]
            else:
                print(f"Unknown b[2]={b[2]}\n {b[3:-3]}") # 240,241

        if TMS_GUI._INS: TMS_GUI._INS.update() # update GUI if exists

def CRC8(u8):
    '''Compute CRC8 (Dallas/Maxim) checksum for polynomial x^8 + x^5 + x^4 + 1'''
    if not hasattr(CRC8, "C8"): # cache CRC8 for 0:256
        CRC8.C8 = bytearray(256)
        for c in range(1,256):
            a = c
            for i in range(8):
                if a&1: a ^= 0b100011001 # poly8 LE
                a >>= 1
            CRC8.C8[c] = a
    rst = CRC8.C8[0]
    for b in u8: rst = CRC8.C8[rst^b]
    return rst

def int2byte(u16):
    if not isinstance(u16,tuple): u16 = (u16,)
    return sum([divmod(int(i),256) for i in u16],())

def closestVal(val, vals):
    '''Return the closest value for val inside vals, and the deviation'''
    item = min(enumerate(vals), key=lambda x:abs(x[1]-val))
    dev = abs(vals[item[0]] - val)
    if val!=0: dev /= abs(val)
    return item[1],dev

def key(D, val):
    '''Return the key for val in dict (unique keys and values)'''
    for k,v in D.items():
        if v == val: return k
    return -1

def lists(D):
    '''Return joined dict or tuple values for error message'''
    if isinstance(D, dict): D = D.values()
    return '"'+'", "'.join(D)+'"'

def frange(start, stop, step):
    '''float version of range(). Rounded to 2 decimal digits (for TMS())'''
    f = (stop-start) / step
    n = int(f)
    if n<f: n += 1
    return [round(x*step+start, 2) for x in range(n)]

## Next Section for motor threshold ##
def rMT(startAmp=None):
    ''' Start to measure motor threshold.
    The optional input is the start amplitude for the threshold estimation. If
    not provided, the current amplitude on the stimulator will be used if it is
    greater than 30, otherwise 65 will be the starting amplitude.
  
    When asked if you see motor response, click "Yes" or "No" (or press key Y or 
    N), and the amplitude will be adjusted accordingly for the next trial. In case 
    you are unsure for a response, or the stimulation target needs to be adjusted, 
    click "Retry" to keep the amplitude unchanged for the next trial. The code will 
    set the default choice based on the trace, and you can press Spacebar if the 
    default choice is correct.
  
    When the estimate converges, the protocol will stop, and threshold will be
    shown on the title of the figure.
  
    Closing the figure will stop the test.'''
    ADC = RTBoxADC()
    ADC.flush()
    
    fig = plt.figure(figsize=[12,4], num=77)
    fig.canvas.manager.set_window_title('Motor Threshold')
    plt.get_current_fig_manager().window.geometry("+40+40")
    win = fig.canvas.manager.window
    try: win.resizable(False, False)
    except: win.setFixedSize(manager.window.size())

    ax = plt.axes([0.05, 0.12, 0.84, 0.7])
    ax.get_yaxis().set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.plot([-5,-5], [3,4],  color='black', linewidth=6)
    ax.text(-4.3, 3.4, "1mV")
    xdata = np.linspace(0, 80, 288) # in ms with fs=3600
    line, = ax.plot(xdata, bandpass(ADC.read()))
    ax.set_ylim(-4, 4)
    ax.set_xlim(-5, 80)
    plt.xlabel("ms")
    
    def setButton(event, x): rMT.pressed = x
    bs = [None]*3
    bs[0] = Button(plt.axes([0.94, 0.6, 0.04, 0.06]), "Yes")
    bs[0].on_clicked(lambda event:setButton(event, "Yes"))
    bs[1] = Button(plt.axes([0.94, 0.45, 0.04, 0.06]), "No")
    bs[1].on_clicked(lambda event:setButton(event, "No"))
    bs[2] = Button(plt.axes([0.94, 0.3, 0.04, 0.06]), "Retry")
    bs[2].on_clicked(lambda event:setButton(event, "Retry"))
    aT = plt.axes([0.895, 0.75, 0.1, 0.2])
    aT.set_axis_off()
    aT.text(0, 0, "Motor Response?")

    def on_key_press(event):
        if   event.key == 'y': rMT.pressed = 'Yes'
        elif event.key == 'n': rMT.pressed = 'No'
        elif event.key == 'r': rMT.pressed = 'Retry'
        elif event.key == ' ': rMT.pressed = dft

    def buttonVisible(tf, focus="Retry"):
        for b in bs:
            b.set_active(tf)
            b.ax.set_visible(tf)
            if tf: # mimick foucus to user: key event can set rMT.pressed
                if b.label.get_text()==focus:
                    b.ax.patch.set_linewidth(2)
                    b.ax.patch.set_edgecolor('blue')
                else:
                    b.ax.patch.set_linewidth(1)
                    b.ax.patch.set_edgecolor('black')
        aT.set_visible(tf)
        fig.canvas.draw_idle()

    buttonVisible(False)
    plt.show(block=False)

    T = TMS()
    amp = startAmp if startAmp else T.amplitude[0]
    if amp<30: amp = 65
    # T.waveform = "Biphasic"
    T.enabled = True
    step = 4
    btn0 = ''
    i = 1
        
    while True:
        T.amplitude = amp # change amplitude for next trial
        plt.pause(2+np.random.random()*2) # sleep() won't update button visible
        T.firePulse() # fire a pulse
        y = ADC.read() # acquire 288 points of EMG
        y[32:] = bandpass(y[32:]) # leave trigger artifact
        line.set_ydata(y)

        ratio = np.std(y[71:180]) / np.std(y[180:])
        if ratio>3: dft = "Yes"
        elif i<2: dft = "Retry"
        elif ratio<1.2: dft = "No"
        else: dft = "Retry"

        rMT.pressed = None
        buttonVisible(True, dft)
        cid = fig.canvas.mpl_connect('key_press_event', on_key_press)

        while rMT.pressed is None:
            plt.pause(0.02)
            if not plt.fignum_exists(77): # figure closed
                print('Motor threshold test stopped')
                return None

        buttonVisible(False)
        fig.canvas.mpl_disconnect(cid)
        if rMT.pressed=="Retry": continue
        if rMT.pressed=="Yes":
            if step<1:
                thre = amp
                break
            else: amp -= step
        elif rMT.pressed=="No":
            if step<1:
                thre = amp+1
                break
            else: amp += step
        if btn0 and rMT.pressed!=btn0: step -= 1
        print(f" Trial {i:2d}: amp={T.amplitude[0]:2d}, response={rMT.pressed}")
        btn0 = rMT.pressed
        i += 1
    ax.set_title(f"Motor threshold is {thre}")
    plt.pause(0.02)
    return thre

class EMGCheck:
    '''Work like an oscilliscope to show trace from RTBoxADC()'''
    def __init__(self, roll=False):
        self.roll = roll
        self.i = 0
        self.count = 0
        self.N = 7200
        xdata = np.linspace(0, 2, self.N)
        self.y = np.zeros(self.N)

        self.fig, ax = plt.subplots(figsize=[12,4], num=3)
        plt.get_current_fig_manager().window.geometry("+40+40")
        self.line, = ax.plot(xdata, self.y, linewidth=0.5)
        ax.set_ylim(-5, 5)
        ax.set_xlim(0, xdata[-1])
        # ax.grid(True, which='both', linestyle='--', linewidth=0.5)
        self.ADC = RTBoxADC()
        self.ADC.flush()
        self.ADC.start()

        def worker(self):
            while plt.fignum_exists(3):
                if self.ADC._port.in_waiting<5: sleep(0.02)
                try: self.update()
                except: return

        thread = threading.Thread(target=worker, args=(self,))
        thread.daemon = True
        thread.start()

        plt.xlabel("Seconds")
        plt.ylabel("mV")
        plt.show() # must block for thread

    def update(self):
        N = int(self.ADC._port.in_waiting/5)
        if N<1: return
        b = self.ADC._port.read(N*5)
        y = byte2vol(b)
        self.count += len(y)
        if self.count>=288:
            self.ADC.start() # start conversion again
            self.count = 0

        if self.roll:
            self.y = np.concatenate((self.y[len(y):], y))
        else:
            N = self.i + len(y)
            if N<self.N: 
                self.y[self.i:N] = y
            else:
                N -= self.N
                j = self.N - self.i
                self.y[self.i:] = y[:j]
                self.y[:N] = y[j:]
            self.i = N

        self.line.set_ydata(self.y)
        self.fig.canvas.draw_idle()

class RTBoxADC:
    '''Return handle for RTBoxADC.
    The differential signal input are DB25 pin 1 and 2 for RTBox v5 and higher.
    The gain is fixed at 200 for now.'''

    _INS = None # store for single instance

    def __new__(cls):
        if RTBoxADC._INS: return RTBoxADC._INS # reuse existing instance
        return super().__new__(cls) # create new instance

    def __del__(self):
        RTBoxADC._INS = None
        try:
            self._port.write(b'R') # return to RTBox main
            self._port.close()
        except: pass

    def __init__(self):
        if RTBoxADC._INS: return
        for p in comports():
            if "bluetooth" in p[0].lower(): continue # too slow to close for MAC
            try: ser = serial.Serial(p[0], 115200, timeout=0.3)
            except: continue
            ser.write(b'R'); sleep(0.1) # in case in ADC
            ser.write(b'X')
            b = ser.read(21)
            if len(b)>20: break
            ser.close()

        assert 'ser' in locals() and ser.is_open, "Failed to find RTBox."
        RTBoxADC._INS = self
        assert b[17:19]==b'v5' or b[17:19]==b'v6', "RTBox version 5 or later needed."
    
        ser.write(b'G') # jump into ADC function
        ser.write(bytes((67,75))); sleep(0.1) # differential ADC1-ADC0 gain=200, vref=5
        # ser.write(bytes((70,2)); sleep(0.1) # rate=3600, default
        ser.write(bytes((110,1,32))) # 3600*0.08 = 288 samples
        ser.flushInput()
        self._port = ser

    def start(self):
        '''Start conversion for 288 samples, and return immediately'''
        self._port.write(b'\x02')

    def flush(self):
        '''Flush the input buffer before start()'''
        n = self._port.in_waiting
        while True:
            n1 = self._port.in_waiting
            if n1 == n: break
            n = n1
            sleep(0.02)
        self._port.flushInput()

    def read(self):
        '''Start conversion and wait to return 288 samples'''
        self._port.write(b'\x02')
        sleep(0.08)
        for i in range(4):
            if self._port.in_waiting<360: sleep(0.02)
            else: break

        N = int(self._port.in_waiting/5) * 5
        b = self._port.read(N)
        return byte2vol(b)

class trainParam:
    '''Train parameters object'''
    __slots__ = ('_TimingControl', '_RepRate', '_PulsesInTrain', '_NumberOfTrains',\
                 '_ITI', '_PriorWarningSound', '_RampUp', '_RampUpTrains')
    
    def __init__(obj):
        obj._TimingControl = "Sequence"
        obj._RepRate = 1.0
        obj._PulsesInTrain = 5
        obj._NumberOfTrains = 3
        obj._ITI = 1.0
        obj._PriorWarningSound = True
        obj._RampUp = 1.0
        obj._RampUpTrains = 10

    @property
    def TimingControl(obj):
        '''Timing control options'''
        return obj._TimingControl

    @TimingControl.setter
    def TimingControl(obj, arg):
        if obj._TimingControl==arg: return
        try: k = TMS._INS._TCs.index(arg)
        except: assert [], "Valid TimingControl input: "+lists(TMS._INS._TCs)
        obj._TimingControl = arg
        TMS._INS._setTrain()

    @property
    def RepRate(obj):
        '''Number of pulses per second in each train'''
        return obj._RepRate

    @RepRate.setter
    def RepRate(obj, arg):
        if obj._RepRate==arg: return
        obj._RepRate,dev = closestVal(arg, TMS._INS._RATEs) # pps
        if dev>0.001: print(f"RepRate adjusted to {obj._RepRate}")
        TMS._INS._setTrain()
        
    @property
    def PulsesInTrain(obj):
        '''Number of pulses per second in each train'''
        return obj._PulsesInTrain

    @PulsesInTrain.setter
    def PulsesInTrain(obj, arg):
        if obj._PulsesInTrain==arg: return
        rg = (*range(1,1001), *range(1100,2001,100))
        obj._PulsesInTrain,dev = closestVal(arg, rg)
        if dev>0.001: print(f"PulsesInTrain adjusted to {obj._PulsesInTrain}")
        TMS._INS._setTrain()
        
    @property
    def NumberOfTrains(obj):
        '''Number of Trains in the sequence'''
        return obj._NumberOfTrains

    @NumberOfTrains.setter
    def NumberOfTrains(obj, arg):
        if obj._NumberOfTrains==arg: return
        obj._NumberOfTrains,dev = closestVal(arg, range(1,501))
        if dev>0.001: print(f"NumberOfTrains adjusted to {obj._NumberOfTrains}")
        TMS._INS._setTrain()

    @property
    def ITI(obj):
        '''Inter Train Interval in seconds'''
        return obj._ITI

    @ITI.setter
    def ITI(obj, arg):
        if obj._ITI==arg: return
        obj._ITI,dev = closestVal(arg, frange(0.1, 300.01, 0.1))
        if dev>0.001: print(f"ITI adjusted to {obj._ITI}")
        TMS._INS._setTrain()
        
    @property
    def PriorWarningSound(obj):
        # Sound warning before each train if True
        return obj._PriorWarningSound

    @PriorWarningSound.setter
    def PriorWarningSound(obj, arg):
        if obj._PriorWarningSound==bool(arg): return
        obj._PriorWarningSound = bool(arg)
        TMS._INS._setTrain()
        
    @property
    def RampUp(obj):
        '''A factor of 0.7~1.0 setting the level for the first Train'''
        return obj._RampUp

    @RampUp.setter
    def RampUp(obj, arg):
        if obj._RampUp==arg: return
        obj._RampUp,dev = closestVal(arg, frange(0.7, 1.01, 0.05))
        if dev>0.001: print(f"RampUp adjusted to {obj._RampUp}")
        TMS._INS._setTrain()
        
    @property
    def RampUpTrains(obj):
        # Number of trains during which the Ramp up function is active
        return obj._RampUpTrains

    @RampUpTrains.setter
    def RampUpTrains(obj, arg):
        if obj._RampUpTrains==arg: return
        obj._RampUpTrains,dev = closestVal(arg, range(1,11))
        if dev>0.001: print(f"RampUpTrains adjusted to {obj._RampUpTrains}")
        TMS._INS._setTrain()

def byte2vol(b):
    # 5-bytes packet is for 4 samples (10 bits): 5th for higher 2 bit for other four
    b = np.frombuffer(b, dtype=np.uint8)
    b = np.int16(np.reshape(b, (-1,5)))
    for i in range(4): b[:,i] |= (b[:,4]<<8-2*i)&0x300
    b = b[:,0:4].flatten()
    b[b>511] -= 1024
    return b/20.48 # mV: 1000/200*2/1024*5

def bandpass(x, band=[5,500], fs=3600):
    ''' Apply bandpass filter to signal x.
    The band input is [hp, lp] in Hz, and fs is sampling rate in Hz. '''
    n = len(x)
    x = np.fft.fft(x)
    i0 = round(band[0]/fs*n + 1)
    x[:i0] = 0
    x[n+2-i0:] = 0 # always remove mean: cufoff too steep?
    if not np.isinf(band[1]) and band[1] is not None:
        i1 = round(band[1]/fs*n + 1)
        if i1>1 and i1<n: x[i1:n+2-i1] = 0
    return np.real(np.fft.ifft(x))


## Next section for GUI ##
class addLabel:
    '''Place label at left of widget. widget and label share tooltip/enable/disable.'''
    def __init__(self, widget, label_text, tooltip_text=""):
        self.widget = widget
        self.label = tk.Label(widget.master, text=label_text, anchor="se")
        p = widget.place_info()
        self.label.place(x=1, y=p['y'], width=int(p['x'])-4, height=p['height'])
        if tooltip_text:
            Hovertip(self.label, tooltip_text)
            Hovertip(widget, tooltip_text)

    def disable(self):
        '''Disable widget and its label'''
        self.label.config(foreground="gray") # looks disabled
        self.widget.config(state="disabled")

    def enable(self):
        '''Enable widget and its label'''
        self.label.config(foreground="black")
        try: self.widget.config(state="normal")
        except: self.widget.state(["!disable"])

class TMS_GUI:
    """
    TMS_GUI() shows and controls the Magventure TMS parameters. 
    
    While TMS() works independently in script or command line, TMS_GUI() is only
    an interface to show and control the stimulator through TMS().
    
    The control panel contains help information for each parameter and item. By
    hovering mouse onto an item for a second, the help will show up.
    
    The Coil Status panel is information only, and the coil temperature and di/dt
    will update after stimulation is applied.
    
    The Basic Control panel contains common basic parameters. The Burst Parameters
    are active only for Waveform of Biphasic Burst. The Train Control panel 
    contains the parameters for train stimulation.
    
    All parameter change will be effective immediately at the stimulator.
    
    From File menu, the parameters can be saved to a JSON file, so they can be 
    loaded for the future sessions. The "Load" function will send all parameters 
    to the stimulator. Then once the desired amplitude is set, it is ready to 
    "Trig" a pulse/burst or "Start Train". This is the easy and safe way to set 
    up all parameters.
    
    From Serial menu, one can Resync status from the stimulator, in case any
    parameter is changed at the stimulator (strongly discouraged). One can also
    Disconnect the serial connection to the stimulator, as this is necessary if
    another program, e.g. E-Prime, will trigger the stimulation.
    """
    _INS = None
    
    def on_closing(self):
        self.root.destroy()
        TMS_GUI._INS = None
        try: TMS._INS.disconnect()
        except: pass
        try: RTBoxADC._INS.__del__()
        except: pass

    def __init__(self):
        TMS_GUI._INS = self
        self.root = tk.Tk()
        self.amplitude = tk.IntVar(self.root, 0)
        self.BARatio = tk.StringVar(self.root, "1")
        self.burstPulses = tk.IntVar(self.root, 3)
        self.IPI = tk.StringVar(self.root, "10")        
        self.RepRate = tk.StringVar(self.root, "1")
        self.PulsesInTrain = tk.StringVar(self.root, "5")
        self.NumberOfTrains = tk.StringVar(self.root, "3")
        self.ITI = tk.StringVar(self.root, "1")
        self.PriorWarningSound = tk.BooleanVar(self.root, True)
        self.create_GUI()

    def create_GUI(self):
        T = TMS()
        def enable(): T.enabled = not T.enabled
        def amplitude_cb(_=0): T.amplitude = self.amplitude.get()
        def burstPulses_cb(_=0): T.burstPulses = self.burstPulses.get()
        def mode_cb(_): T.mode = self.mode.get()
        def currentDirection_cb(_): T.currentDirection = self.currentDirection.get()
        def waveform_cb(_): T.waveform = self.waveform.get()
        def BARatio_cb(_): T.BARatio = float(self.BARatio.get())
        def IPI_cb(_): T.IPI = float(self.IPI.get())
        def RepRate_cb(_): T.train.RepRate = float(self.RepRate.get())
        def PulsesInTrain_cb(_): T.train.PulsesInTrain = int(self.PulsesInTrain.get())
        def NumberOfTrains_cb(_): T.train.NumberOfTrains = int(self.NumberOfTrains.get())
        def ITI_cb(_): T.train.ITI = float(self.ITI.get())
        def PriorWarningSound_cb(): T.train.PriorWarningSound = self.PriorWarningSound.get()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        dFontSz = 3 if "darwin" in tk.sys.platform else 0 # arbituary adjustment for MAC
        fnt = ('Helvetica', 9+dFontSz, 'bold')
        self.root.option_add("*Font", f"Helvetica {9+dFontSz} normal")
        fNam = path.join(path.dirname(path.abspath(__file__)),'TIcon.png')
        try: self.root.iconphoto(False, tk.PhotoImage(file=fNam))
        except: pass
        self.root.geometry("490x406+100+500")
        self.root.resizable(False, False)

        self.enable_txt = tk.Label(self.root, text="Disabled", font=fnt)
        self.enable_txt.place(x=31, y=10, width=62, height=32)
        enabled = tk.Button(self.root, text=chr(11096), command=enable, font=("Helvetica", 12+dFontSz))
        enabled.place(x=94, y=10, width=32, height=32)
        Hovertip(enabled, 'Push to enable/disable stimulation')
    
        self.firePulse = tk.Button(self.root, text="Trig", command=T.firePulse)
        self.firePulse.place(x=155, y=10, width=82, height=32)
        Hovertip(self.firePulse, 'Trigger a pulse or burst')

        self.EMGcheck = tk.Button(self.root, text="EMG Check", command=EMGCheck)
        self.EMGcheck.place(x=31, y=53, width=82, height=24)
        Hovertip(self.EMGcheck, 'Show continuous EMG trace')

        self.motorThr = tk.Button(self.root, text="Motor Threshold", command=rMT)
        self.motorThr.place(x=131, y=53, width=106, height=24)
        Hovertip(self.motorThr, 'Start motor threshold estimate')

        # Basic Control
        basicFrame = tk.LabelFrame(self.root, text="Basic Control", relief=tk.RIDGE, font=fnt)
        basicFrame.place(x=31, y=90, width=206, height=178)
    
        amplitude = tk.Spinbox(basicFrame, from_=0, to=100, justify=tk.RIGHT, command=amplitude_cb, textvariable=self.amplitude)
        amplitude.place(x=156, y=10, width=40, height=22)
        amplitude.bind("<Return>", amplitude_cb)
        amplitude.bind("<FocusOut>", amplitude_cb)
        addLabel(amplitude, "Amplitude (%)", 'Stimulation amplitude in percent')
        
        self.mode = ttk.Combobox(basicFrame, values=("Standard",), state="readonly")
        self.mode.place(x=112, y=40, width=84, height=22)
        self.mode.current(0)
        self.mode.bind("<<ComboboxSelected>>", mode_cb)
        addLabel(self.mode, "Mode", 'Stimulator mode other than "Standard" only available for MagOption')

        self.currentDirection = ttk.Combobox(basicFrame, values=("Normal","Reverse"), state="readonly")
        self.currentDirection.place(x=116, y=70, width=80, height=22)
        self.currentDirection.current(0)
        self.currentDirection.bind("<<ComboboxSelected>>", currentDirection_cb)
        addLabel(self.currentDirection, "Current Direction")

        self.waveform = ttk.Combobox(basicFrame, values=("Biphasic",), state="readonly")
        self.waveform.place(x=86, y=100, width=110, height=22)
        self.waveform.current(0)
        self.waveform.bind("<<ComboboxSelected>>", waveform_cb)
        addLabel(self.waveform, "Waveform")

        def isFloat(P):
            try:
                float(P)
                return True
            except:
                return False

        vcmd = (self.root.register(isFloat), "%P")
        entryArg = {"justify":tk.RIGHT, "validate":"key", "validatecommand":vcmd}
        BARatio = tk.Entry(basicFrame, textvariable=self.BARatio, **entryArg)
        BARatio.place(x=160, y=130, width=36, height=22)
        BARatio.bind("<Return>", BARatio_cb)
        BARatio.bind("<FocusOut>", BARatio_cb)
        self.h_BARatio = addLabel(BARatio, "Pulse B/A Ratio", 'Amplitude ratio of Pulse B over Pulse A for Twin mode')
    
        # Burst Parameters
        burstFrame = tk.LabelFrame(self.root, text="Burst Parameters", relief=tk.RIDGE, font=fnt)
        burstFrame.place(x=31, y=288, width=206, height=94)

        burstPulses = tk.Spinbox(burstFrame, from_=2, to=5, justify=tk.RIGHT, command=burstPulses_cb, textvariable=self.burstPulses)
        burstPulses.place(x=163, y=10, width=34, height=22)
        self.burstPulses_h = addLabel(burstPulses, "Burst Pulses", 
            'Number of pulses in a burst (2 to 5).\nClick dial or press up/down arrow key to change')
        
        IPI = tk.Entry(burstFrame, textvariable=self.IPI, **entryArg)
        IPI.place(x=163, y=40, width=32, height=22)
        IPI.bind("<Return>", IPI_cb)
        IPI.bind("<FocusOut>", IPI_cb)
        self.IPI_h = addLabel(IPI, "Inter Pulse Interval (ms)", 
            'Duration between the beginning of the first pulse to the beginning of the second pulse')

        # Coil Status
        coilFrame = tk.LabelFrame(self.root, text="Coil Status", relief=tk.RIDGE, font=fnt)
        coilFrame.place(x=256, y=10, width=206, height=122)

        self.CoilType = tk.Label(coilFrame, anchor="e", relief='groove')
        self.CoilType.place(x=96, y=10, width=100, height=22)
        addLabel(self.CoilType, "Type/Number", 'Connected coil type or number')
   
        self.temperature = tk.Label(coilFrame, anchor="e", relief='groove')
        self.temperature.place(x=165, y=40, width=31, height=22)
        addLabel(self.temperature, "Temperature (°C)", 'Coil temperature in Celsius')

        self.didt = tk.Label(coilFrame, anchor="e", relief='groove')
        self.didt.place(x=140, y=70, width=56, height=22)
        addLabel(self.didt, "Realized di/dt (A/µs)", 'Coil current gradient')
   
        # Train Control
        trainFrame = tk.LabelFrame(self.root, text="Train Control", relief=tk.RIDGE, font=fnt)
        trainFrame.place(x=256, y=144, width=206, height=238)

        RepRate = tk.Entry(trainFrame, textvariable=self.RepRate, **entryArg)
        RepRate.place(x=160, y=10, width=36, height=22)
        RepRate.bind("<Return>", RepRate_cb)
        RepRate.bind("<FocusOut>", RepRate_cb)
        addLabel(RepRate, "Rep Rate (pps)", 'Number of pulses per second')
    
        ITI = tk.Entry(trainFrame, textvariable=self.ITI, **entryArg)
        ITI.place(x=160, y=100, width=36, height=22)
        ITI.bind("<Return>", ITI_cb)
        ITI.bind("<FocusOut>", ITI_cb)
        addLabel(ITI, "Inter Train Interval (s)", 
            'The time interval between two trains described as\n'+
            'the time period between the last pulse in the first\n'+
            'train to the first pulse in the next train')
    
        def isInt(P):
            try:
                int(P)
                return True
            except:
                return False

        entryArg["validatecommand"] = (self.root.register(isInt), "%P")
        PulsesInTrain = tk.Entry(trainFrame, textvariable=self.PulsesInTrain, **entryArg)
        PulsesInTrain.place(x=160, y=40, width=36, height=22)
        PulsesInTrain.bind("<Return>", PulsesInTrain_cb)
        PulsesInTrain.bind("<FocusOut>", PulsesInTrain_cb)
        addLabel(PulsesInTrain, "Pulses in Train", 'Number of pulses or bursts in each train')
    
        NumberOfTrains = tk.Entry(trainFrame, textvariable=self.NumberOfTrains, **entryArg)
        NumberOfTrains.place(x=160, y=70, width=36, height=22)
        NumberOfTrains.bind("<Return>", NumberOfTrains_cb)
        NumberOfTrains.bind("<FocusOut>", NumberOfTrains_cb)
        addLabel(NumberOfTrains, "Number of Trains", 'Total amount of trains arriving in one sequence')
    
        PriorWarningSound = tk.Checkbutton(trainFrame, text="", 
            variable=self.PriorWarningSound, command=PriorWarningSound_cb)
        PriorWarningSound.place(x=180, y=130, width=22, height=22)
        addLabel(PriorWarningSound, "Prior Warning Sound", 'When on, a beep will sound 2 seconds before each train')

        self.trainTime = tk.Label(trainFrame, anchor="e", relief='groove')
        self.trainTime.place(x=140, y=160, width=56, height=22)
        addLabel(self.trainTime, "Total Time", 'Total time to run the sequence, based on above parameters')
   
        self.fireTrain = tk.Button(trainFrame, text="Start Train", command=T.fireTrain)
        self.fireTrain.place(x=52, y=190, width=100, height=24)
        Hovertip(self.fireTrain, 'Start / Stop train sequence')
    
        # Menus
        menubar = tk.Menu(self.root); self.root.update()
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Load", command=T.load, underline=0)
        filemenu.add_command(label="Save", command=T.save, underline=0)
        menubar.add_cascade(label="File", menu=filemenu)
        serialmenu = tk.Menu(menubar, tearoff=0)
        serialmenu.add_command(label="Resync", command=T.resync, underline=0)
        serialmenu.add_command(label="Disconnect", command=self.on_closing, underline=0)
        menubar.add_cascade(label="Serial", menu=serialmenu)
        helpmenu = tk.Menu(menubar, tearoff=0)
        helpmenu.add_command(label="Help about TMS()", command=lambda:help(TMS))
        helpmenu.add_command(label="Help about TMS_GUI()", command=lambda:help(TMS_GUI))
        menubar.add_cascade(label="Help", menu=helpmenu)
        self.root.config(menu=menubar)
        self.update()
        self.root.mainloop()

    def update(self):
        '''Update UI, mainly called by TMS._read()'''
        T = TMS()
        fName = path.splitext(path.basename(T.filename))[0]
        if hasattr(T, "_port"): self.root.title("Magventure "+T.Model+" "+fName)
        else: self.root.title("NotConnected "+fName)
        self.amplitude.set(T.amplitude[0])
        self.mode['values'] = tuple(T._MODEs.values())
        self.mode.set(T.mode)
        self.currentDirection['values'] = tuple(T._curDirs.values())
        self.currentDirection.set(T.currentDirection)
        self.waveform['values'] = tuple(T._wvForms.values())
        self.waveform.set(T.waveform)
        self.BARatio.set(f'{T.BARatio:g}')
        self.burstPulses.set(T.burstPulses)
        self.IPI.set(f'{T.IPI:g}')
        self.RepRate.set(f'{T.train.RepRate:g}')
        self.PulsesInTrain.set(f'{T.train.PulsesInTrain}')
        self.NumberOfTrains.set(f'{T.train.NumberOfTrains}')
        self.ITI.set(f'{T.train.ITI:g}')
        self.PriorWarningSound.set(T.train.PriorWarningSound)
        self.CoilType.config(text=T.info["CoilType"])
        self.temperature.config(text=f'{T.temperature}')
        self.didt.config(text=f'{T.didt[0]}  {T.didt[1]}')
        self.trainTime.config(text=T.trainTime)

        st = tk.NORMAL if T.enabled and T.amplitude[0]>0 else tk.DISABLED
        self.firePulse.config(state=st)
        self.fireTrain.config(state=st)
        self.fireTrain.config(text='Stop Train' if T.trainRunning else 'Start Train')
        if T.mode=="Twin": self.h_BARatio.enable()
        else: self.h_BARatio.disable()
        if T.waveform=="Biphasic Burst":
            self.burstPulses_h.enable()
            self.IPI_h.enable()
        else:
            self.burstPulses_h.disable()
            self.IPI_h.disable()

        self.temperature.config(background="red" if T.temperature>40 else "#f0f0f0")
        if T.enabled: self.enable_txt.config(background="#0f0", text="Enabled")
        else: self.enable_txt.config(background="#f00", text="Disabled")
        
        try:
            tb = mpl.rcParams['toolbar']
            ADC = RTBoxADC()
            self.EMGcheck.config(state=tk.NORMAL)
            self.motorThr.config(state=tk.NORMAL)
        except:
            self.EMGcheck.config(state=tk.DISABLED)
            self.motorThr.config(state=tk.DISABLED)

if __name__ == "__main__": TMS_GUI()
    
