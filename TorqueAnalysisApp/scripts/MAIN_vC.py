#!/usr/bin/env python
# coding: utf-8

########################################################################################################
# main file                                                                                            
# created by:      RHM
# References:	   This tool is developed from the previous Torque Analysis Jupyter scripts created by RHM & KZE.
# purpose:         Calculates the results for the torque measurement and provides visual results in plots/graphs
#                  Provide UI window --> user chooses wheelbase configuration
#                  Calculate relevant values (e.g. slewrate) based on wheelbase config
# Created on:      24.05.2023
# version:         2.0
########################################################################################################

########################################################################################################
# Change log:                                                                                    
# Version 2.1 :
#        30.10.2023:
#        Multiple measured files can be processed at once.
#        Updated the file names to be saved.
########################################################################################################
import subprocess
import sys
import os

# Install required packages from requirements.txt file
#subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "-r", "requirements.txt"])

# import libraries
import pandas as pd
import numpy as np
import time
import matplotlib.pyplot as plt
import openpyxl
import xlsxwriter
import tkinter as tk
from tkinter import *
from tkinter import ttk
from PEAK import *
from LIN import *
from config import input_path, result_path,config_path
import logging
import warnings
warnings.filterwarnings("ignore", message="Ignoring invalid distribution.*")

class TorqueAnalysis():
    def __init__(self):
        self.window = None

    def read_file(self):
        ######### STEP 1: Read excel config
        # read excel config, index is configuration name
        
        self.config_file = os.path.join(self.config_path, "WheelbaseConfigurationFile.xlsx")
        self.dataExcelSmall = pd.read_excel(self.config_file, index_col="WheelbaseConfiguration")
        # read excel config (without changing index)
        self.dataExcelFull = pd.read_excel(self.config_file)
        # make a list of all wheelbase configurations
        self.configurationList = self.dataExcelFull["WheelbaseConfiguration"].values.tolist()
        ######### STEP 1: Read excel config

    # function: get choosen wheelbase configuration
    def check_cbox(self, event):
        ######## Default configuration setting
        self.processedEnteredConfig = 'CSLDD_GTDD_180W100FFB'
        self.processedEnteredConfig = self.cmb.get()
        #print(self.processedEnteredConfig)

    def createwindow(self):
        # create UI window
        if self.window is None:
            self.window = Tk()
            self.window.title('Wheelbase Selection')
            self.window.geometry('350x250')
            self.window.eval('tk::PlaceWindow . center')

            ttk.Label(self.window, text="Wheelbase configuration list",
                    background='black', foreground="white",
                    ).grid(row=0, column=2, padx=70, pady=10)
            # list containing all wheelbase configurations
            self.cmb = ttk.Combobox(self.window, values=self.configurationList, height=20, width=30, state='readonly')
            self.cmb.grid(column=2, row=2, padx=70, pady=10)
            self.cmb.current(0)
            self.cmb.bind("<<ComboboxSelected>>", self.check_cbox)
            # close window if button is clicked
            btn = Button(self.window, text="Select and close", command=self.closeWindow)
            btn.grid(column=2, row=4, padx=70, pady=80)
            self.window.mainloop()

    ######### STEP 2: get wheelbase configuration from user
    # open window with configuration list
    # user: choose config, press "select and close"
    # function: close UI window
    def closeWindow(self):
        self.window.destroy()

    def exl_cfg(self):
        # get the choosen wheelbase setting from all excel configurations
        self.getCorrectSetting = self.dataExcelSmall.loc[self.processedEnteredConfig]
        # assign all the calculation variables from config excel
        maxPeakTRQRef = self.getCorrectSetting['MaxPeakTRQRef']
        holdingTRQRef = self.getCorrectSetting['HoldingTRQRef']
        maxPeakTRQRefLLFactor = self.getCorrectSetting['MaxPeakTRQRefLLFactor']
        maxPeakTRQRefULFactor = self.getCorrectSetting['MaxPeakTRQRefULFactor']
        self.preciseMaxPeakTRQLLFactor = self.getCorrectSetting['PreciseMaxPeakTRQLLFactor']
        self.preciseMaxPeakTRQULFactor = self.getCorrectSetting['PreciseMaxPeakTRQULFactor']
        self.TRQNinetyLLFactor = self.getCorrectSetting['TRQNinetyLLFactor']
        self.TRQNinetyULFactor = self.getCorrectSetting['TRQNinetyULFactor']
        self.TRQTenLLFactor = self.getCorrectSetting['TRQTenLLFactor']
        self.TRQTenULFactor = self.getCorrectSetting['TRQTenULFactor']
        self.avgHoldingTRQLLFactor = self.getCorrectSetting['AvgHoldingTRQLLFactor']
        self.avgHoldingTRQULFactor = self.getCorrectSetting['AvgHoldingTRQULFactor']
        self.preciseAverageHoldingTRQLLFactor = self.getCorrectSetting['PreciseAverageHoldingTRQLLFactor']
        self.preciseAverageHoldingTRQULFactor = self.getCorrectSetting['PreciseAverageHoldingTRQULFactor']
        ######### STEP 2: get wheelbase configuration from user

        ######### STEP 3: get measurement data
        # skip the first lines of the excel sheet; these lines contain information that are not relevant to the calculations
        #inputfilepath = "../Input/" + inputfilename + ".xlsx"
        dataInput = pd.read_excel(self.fullname, skiprows=50, header=None)
        # name the input columns
        dataInput.columns =['Time', 'Torque', 'WeDontKnowYet']
        # remove 3rd column
        dataInput = dataInput.drop('WeDontKnowYet', axis=1)
        ######### STEP 3: get measurement data

        ######### STEP 4: Check: clockwise or anticlockwise
        dataCheckClockwise = dataInput.copy()
        # count all data where torque is lower than -0.5 Nm
        counterNegativeValues = len(dataCheckClockwise[(dataCheckClockwise['Torque'] <= -0.5)])
        # initialize flag variable
        self.antiClockwise = False
        # if there are negative values (more than 50000) --> data is from anticlockwise measurement
        if counterNegativeValues > 50000:
            # multiply all torque values with (-1)
            dataCheckClockwise['Torque'] = dataCheckClockwise['Torque'].multiply(-1)
            # set flag variable to true
            self.antiClockwise = True
        self.measurementData = dataCheckClockwise.copy()
        ######### STEP 4: Check: clockwise or anticlockwise

        ######### STEP 4.5: Initial plot
        fig = plt.figure(figsize=(20,10))
        ax = fig.add_subplot(111)
        x = list(self.measurementData['Time'])
        y = list(self.measurementData['Torque'])
        line, = ax.plot(x, y)
        plt.title("Raw measurement data")
        plt.xlabel("Time[Seconds]")
        plt.ylabel("Torque[Nm]")
        #plt.show()
        ######### STEP 4.5: Initial plot

        ######### STEP 5a) Calculate max overshoot torque
        # get max overshoot torque
        self.maxOvershootTorqueValue = self.measurementData['Torque'].max()
        timeMaxOvershoot = self.measurementData[self.measurementData['Torque'] == self.maxOvershootTorqueValue].iloc[0]
        valueTimeMaxOvershoot = timeMaxOvershoot['Time']
        # get index value of max overshoot torque
        self.indexMaxOvershoot = self.measurementData['Torque'].idxmax()

        # split data
        # dataset goes until intIndexMaxOvershoot
        self.dataUntilMaxOvershoot = self.measurementData.iloc[:self.indexMaxOvershoot+1,:]
        # dataset starts after intIndexMaxOvershoot
        dataAfterMaxOvershoot = self.measurementData.iloc[self.indexMaxOvershoot+1:,:]
        ######### STEP 5a) Calculate max overshoot torque

        ######### STEP 5b) Calculate max peak torque (initial)
        datasetMaxPeakTRQ = self.measurementData.iloc[self.indexMaxOvershoot:self.indexMaxOvershoot+4000,:]
        self.maxPeakTRQ = datasetMaxPeakTRQ['Torque'].mean()
        ######### STEP 5b) Calculate max peak torque (initial)

        ######### STEP 5c) Calculate holding torque (initial)
        self.dataForHoldingTorque = self.measurementData.iloc[self.indexMaxOvershoot+10000:,:]
        # remove all data where torque <= holding TRQ reference (to remove the 0 Nm torque data)
        self.dataHoldingTorque = self.dataForHoldingTorque.loc[(self.dataForHoldingTorque['Torque'] >= holdingTRQRef)]
        # calc. the average of the last 10.000 data to get the holding torque
        self.tenThousandData = self.dataHoldingTorque.tail(10000)
        # get average holding torque
        self.averageHoldingTorque = self.tenThousandData['Torque'].mean()
        ######### STEP 5c) Calculate holding torque (initial)

        ######### STEP 5d) Calculate measurement error
        # get data of the first 0.5 seconds
        dataErrorCalculation = self.measurementData.loc[(self.measurementData['Time'] <= 0.5)]
        # get absolute values
        dataErrorCalculation = dataErrorCalculation.abs()
        # get max error
        self.maxZeroError = dataErrorCalculation['Torque'].max()
        # get min error
        self.minZeroError = dataErrorCalculation['Torque'].min()
        # get average error
        self.averageZeroError = dataErrorCalculation['Torque'].mean()
        ######### STEP 5d) Calculate measurement error

        ######### STEP 6: Check: PEAK or LIN?
        # calculate percentage --> holding torque and max peak
        percentageCheckValue = self.averageHoldingTorque / self.maxPeakTRQ
        # calculate difference --> holding torque and max peak
        differenceCheckValue = self.maxPeakTRQ - self.averageHoldingTorque
        # Check which code has to be executed
        if percentageCheckValue >= 0.9 and differenceCheckValue <= 0.5:
            print("Execute LIN (or PEAK with low ffb)")
            # Code LIN
            #get_ipython().run_line_magic('run', './LIN.ipynb')
            calculateMorePreciseMaxPeakTRQLIN(self)
            calculateSlewrateLIN(self)
            calculateRisetimeLIN(self)
            calculateSettlingTimeLIN(self)
            getMeasurementTimeLIN(self)
            calculatePreciseHoldingTRQ(self)
            printResultsLIN(self)
            getPlotsLIN(self)
            writeExcelResultsLIN(self)
            #plt.show()
        else:
            print("Execute PEAK")
            # Code PEAK
            #get_ipython().run_line_magic('run', './PEAK.ipynb')
            getMorePreciseMaxPeakTRQ(self)
            calculateSlewratePEAK(self)
            calculateRisetimePEAK(self)
            getMorePreciseHoldingTRQ(self)
            calculateSettlingTimes(self)
            getMeasurementTime(self)
            printResultsPEAK(self)
            getPlotsPEAK(self)
            writeExcelResultsPEAK(self)
            #plt.show()
        ######### STEP 6: Check: PEAK or LIN?

    def create_result_folder(self, result_folder):
        # Check if the result folder exists, and create it if not
        if not os.path.exists(result_folder):
            os.makedirs(result_folder)
            print(f"Result folder '{result_folder}' created.")

    def main(self):
        # Determine the base directory dynamically
        base_directory = os.path.dirname(os.path.abspath(__file__))

        # Build the full paths
        self.input_path = os.path.join(os.path.dirname(base_directory), input_path)
        self.result_path = os.path.join(os.path.dirname(base_directory), result_path)
        self.config_path = os.path.join(os.path.dirname(base_directory), config_path)

        # Check if the input folder exists and raise an exception if it doesn't
        if not os.path.exists(self.input_path):
            raise Exception(f"Input folder '{self.input_path}' does not exist.")

        for filename in os.listdir(self.input_path):
            print(f"Processing input file: {filename}")
            print("Loading... Please wait!!")
            if filename.endswith(".xlsx") or filename.endswith(".XLSX") or filename.endswith(".xls"):
                self.fullname = os.path.join(self.input_path, filename)
                flname = os.path.basename(self.fullname)
                self.fname = os.path.splitext(flname)[0]

                # Create a result folder for each input file
                self.result_folder = os.path.join(self.result_path, self.fname)
                self.create_result_folder(self.result_folder)

                self.read_file()
                self.createwindow()
                self.exl_cfg()

                time.sleep(3)  
                #plt.show()
            print(f"Processing completed: {filename}")
            print("\n")
                
val = TorqueAnalysis()
val.main()