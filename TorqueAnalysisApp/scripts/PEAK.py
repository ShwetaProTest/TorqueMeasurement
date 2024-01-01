#!/usr/bin/env python
# coding: utf-8

# import libraries
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import xlsxwriter
import scipy.optimize
import scipy as sy
import math
from scipy.optimize import curve_fit
import os


########################################################################################################
# PEAK file                                                                                            
# created by:      RHM, KZE 
# purpose:         Provide functions for main file:                                                    
#                  Calculate precise max peak TRQ                                                      
#                  Calculate slewrate                                                                  
#                  Calculate rise time                                                                 
#                  Calculate precise holding TRQ                                                       
#                  Calculate settling times                                                            
#                  Calculate measurement time (from apply input until end of holding TRQ)              
#                  Print calculated values                                                             
#                  Create plots (PEAK: 6 plots)                                                        
#                  Create excel file (contains calculated values)                                      
# created:         19.10.2022 - 14.12.2022                                                             
# Modified:        20.01.2023
# Modified by:     RHM
# version:         C                                                 
########################################################################################################

# calculate precise max peak TRQ
def getMorePreciseMaxPeakTRQ(self):
    # set maxPeakTRQ interval with excel config settings
    lowerMaxPeakTorque = self.maxPeakTRQ * self.preciseMaxPeakTRQLLFactor
    upperMaxPeakTorque = self.maxPeakTRQ * self.preciseMaxPeakTRQULFactor
    # get all points inside this interval
    maxPeakTorqueRangeData = self.measurementData[
        (self.measurementData['Torque'] >= lowerMaxPeakTorque) & (self.measurementData['Torque'] <= upperMaxPeakTorque)]
    self.morePreciseMaxPeakTRQ = maxPeakTorqueRangeData['Torque'].mean()
    # override max peak value
    self.maxPeakTRQ = self.morePreciseMaxPeakTRQ
    # get index of the last point
    lastMaxPeakPoint = maxPeakTorqueRangeData.tail(1)
    indexLastMaxPeakPoint = lastMaxPeakPoint['Torque'].idxmax()
    # get values from the last point
    lastMaxPeakPointValues = maxPeakTorqueRangeData.iloc[-1]
    self.timeLastMaxPeakPointValues = lastMaxPeakPointValues['Time']
    torqueLastMaxPeakPointValues = lastMaxPeakPointValues['Torque']
    # max peak calc. with % band
    datasetMaxPeakBandFirst = self.measurementData.iloc[self.indexMaxOvershoot:indexLastMaxPeakPoint, :]
    allBandPoints = datasetMaxPeakBandFirst.loc[(datasetMaxPeakBandFirst['Torque'] >= self.maxPeakTRQ * 1.05) | (
                datasetMaxPeakBandFirst['Torque'] <= self.maxPeakTRQ * 0.95)]
    counterBandValues = len(allBandPoints)
    # if band empty --> 2%
    if counterBandValues == 0:
        self.bandPercentageUL = 1.02
        self.bandPercentageLL = 0.98
    else:
        self.bandPercentageUL = 1.05
        self.bandPercentageLL = 0.95

    allBandPoints = allBandPoints = datasetMaxPeakBandFirst.loc[
        (datasetMaxPeakBandFirst['Torque'] >= self.maxPeakTRQ * self.bandPercentageUL) | (
                    datasetMaxPeakBandFirst['Torque'] <= self.maxPeakTRQ * self.bandPercentageLL)]
    # 1st point
    pointFirstOutsideBand = allBandPoints.tail(1)
    # get index of first point
    indexPointFirstOutsideBand = pointFirstOutsideBand['Torque'].idxmax()
    pointFirstOutsideBandValues = allBandPoints.iloc[-1]
    # get time value of first point
    self.timePointFirstOutsideBandValues = pointFirstOutsideBandValues['Time']
    # 2nd point
    datasetMaxPeakBandSecond = self.measurementData.iloc[indexLastMaxPeakPoint:, :]
    allBandPointsTwo = datasetMaxPeakBandSecond.loc[
        (datasetMaxPeakBandSecond['Torque'] >= self.maxPeakTRQ * self.bandPercentageUL) | (
                    datasetMaxPeakBandSecond['Torque'] <= self.maxPeakTRQ * self.bandPercentageLL)]
    pointFirstOutsideBandTwo = allBandPointsTwo.head(1)
    # get index of last band point
    self.indexPointFirstOutsideBandTwo = pointFirstOutsideBandTwo['Torque'].idxmax()
    pointFirstOutsideBandValuesTwo = allBandPointsTwo.iloc[0]
    # get time value of last band point
    self.timePointFirstOutsideBandValuesTwo = pointFirstOutsideBandValuesTwo['Time']
    # calc max peak torque
    datasetThirdCalc = self.measurementData.iloc[indexPointFirstOutsideBand:self.indexPointFirstOutsideBandTwo, :]
    thirdMaxPeak = datasetThirdCalc['Torque'].mean()
    self.maxPeakTRQ = thirdMaxPeak

# calculate slewrate PEAK
def calculateSlewratePEAK(self):
    self.ninetyPercentTorque = self.maxPeakTRQ * 0.9
    self.tenPercentTorque = self.maxPeakTRQ * 0.1
    # interval
    # for the 90%
    ninetyPercentTorqueLowerLimit = self.ninetyPercentTorque * self.TRQNinetyLLFactor
    ninetyPercentTorqueUpperLimit = self.ninetyPercentTorque * self.TRQNinetyULFactor
    # for the 10%
    tenPercentTorqueLowerLimit = self.tenPercentTorque * self.TRQTenLLFactor
    tenPercentTorqueUpperLimit = self.tenPercentTorque * self.TRQTenULFactor
    # get all points in the 90% intervall
    # this could return more than 1 values --> take first point
    allNinetyPercentPoints = self.dataUntilMaxOvershoot.loc[
        (self.dataUntilMaxOvershoot['Torque'] >= ninetyPercentTorqueLowerLimit) & (
                    self.dataUntilMaxOvershoot['Torque'] <= ninetyPercentTorqueUpperLimit)]
    # get first point of 90% interval
    firstNinetyPercentPoint = allNinetyPercentPoints.iloc[0]
    # get the torque and time value of the 90% point
    self.ninetyPercentPointTRQ = firstNinetyPercentPoint['Torque']
    self.ninetyPercentPointTime = firstNinetyPercentPoint['Time']
    # get all 10% points
    allTenPercentPoints = self.dataUntilMaxOvershoot.loc[(self.dataUntilMaxOvershoot['Torque'] >= tenPercentTorqueLowerLimit) & (
                self.dataUntilMaxOvershoot['Torque'] <= tenPercentTorqueUpperLimit)]
    # get last point of 10% interval
    firstTenPercentPoint = allTenPercentPoints.iloc[-1]
    # get torque and time value of the 10% point
    self.tenPercentPointTRQ = firstTenPercentPoint['Torque']
    self.tenPercentPointTime = firstTenPercentPoint['Time']
    refTimeTen = self.tenPercentPointTime
    refTimeNinety = self.ninetyPercentPointTime
    refTRQTen = self.tenPercentPointTRQ
    refTRQNinety = self.ninetyPercentPointTRQ

    # function: calculate precise time values for slewrate
    def getPreciseTime(TRQ):
        return ((TRQ - refTRQTen) * (refTimeNinety - refTimeTen)) / (refTRQNinety - refTRQTen) + refTimeTen

    # use function for calculation
    self.preciseTimeTen = getPreciseTime(self.tenPercentTorque)
    self.preciseTimeNinety = getPreciseTime(self.ninetyPercentTorque)
    self.preciseTimeZero = getPreciseTime(0)
    self.applyInputTime = self.preciseTimeZero
    self.applyInputTorque = 0
    # 0.000027 seconds is the time difference between two measurement points
    inputTimeCalcData = self.dataUntilMaxOvershoot.loc[(self.dataUntilMaxOvershoot['Time'] >= self.preciseTimeZero - 0.000027) & (
                self.dataUntilMaxOvershoot['Time'] <= self.preciseTimeZero + 0.000027)]
    getInputTimeIndex = inputTimeCalcData.head(1)
    self.indexGetInputTime = getInputTimeIndex['Torque'].idxmax()

    # calculate slew rate
    slewratePeakWithSeconds = (self.ninetyPercentTorque - self.tenPercentTorque) / (self.preciseTimeNinety - self.preciseTimeTen)
    # divide by 1000 --> get Nm/ms (otherwise it would be Nm/s)
    self.slewratePeak = slewratePeakWithSeconds / 1000

    # Effective Slewrate with curve fit (e^x)
    # Define window for e^x curve fit
    intIndexStartPolynomialFit = self.indexGetInputTime
    intIndexEndPolynomialFit = self.indexPointFirstOutsideBandTwo
    curveFitData = self.measurementData.iloc[intIndexStartPolynomialFit:intIndexEndPolynomialFit, :]
    # Get time values --> window of curve fit
    polynomialFitStart = self.measurementData.iloc[intIndexStartPolynomialFit]
    timeStartPolynomialFit = polynomialFitStart['Time']
    polynomialFitEnd = self.measurementData.iloc[intIndexEndPolynomialFit]
    timeEndPolynomialFit = polynomialFitEnd['Time']
    # Shift time values --> Zero time
    dataShiftedTime = curveFitData.copy()
    # Subtract first time value of curve dataset
    dataShiftedTime['Time (Shifted)'] = dataShiftedTime['Time'] - timeStartPolynomialFit
    # assign columns to variables
    self.yData = dataShiftedTime['Torque']
    self.xData = dataShiftedTime['Time (Shifted)']
    self.yDataForFit = self.yData.to_numpy()
    self.xDataForFit = self.xData.to_numpy()

    # define function with e^x
    def curveFitFunction(x, C):
        y = (-self.maxPeakTRQ * math.e ** (-C * x)) + self.maxPeakTRQ
        # C = parameter for optimal fit
        # x = time values
        return y

    # get value of C
    parameters, covariance = curve_fit(curveFitFunction, self.xDataForFit, self.yDataForFit)
    self.fit_C = parameters[0]
    self.fit_C = self.fit_C / 2

    # fit data
    self.fit_yData = curveFitFunction(self.xDataForFit, self.fit_C)
    # Resolve function according to x
    # --> Get x values (time) of curve fit function
    # timeNinetyPercentTRQCurve = -(math.log((self.maxPeakTRQ-self.ninetyPercentTorque)/self.maxPeakTRQ))/(parameters)
    # For curve fit C/2
    timeNinetyPercentTRQCurve = -(math.log((self.maxPeakTRQ - self.ninetyPercentTorque) / self.maxPeakTRQ)) / (parameters / 2)
    self.timeTenPercentTRQCurve = self.tenPercentPointTime - timeStartPolynomialFit
    # get value of array
    self.valueTimeNinetyPercentTRQCurve = timeNinetyPercentTRQCurve[0]
    self.effectiveSlewrate = (self.ninetyPercentTorque - self.tenPercentTorque) / (
                self.valueTimeNinetyPercentTRQCurve - self.timeTenPercentTRQCurve)
    self.effectiveSlewrate = self.effectiveSlewrate / 1000

# calculate rise time
def calculateRisetimePEAK(self):
    self.applyInputTime = self.preciseTimeZero
    self.applyInputTorque = 0
    # calculate rise time
    riseTimeInSeconds = self.preciseTimeNinety - self.preciseTimeTen
    # Multiply by 1000 --> get time in ms (not seconds)
    self.riseTime = riseTimeInSeconds * 1000

# calculate precise holding TRQ
def getMorePreciseHoldingTRQ(self):
    # calculate a small intervall with the initial holding TRQ
    lowerLimitAverageHoldingTorque = self.averageHoldingTorque * self.avgHoldingTRQLLFactor
    upperLimitAverageHoldingTorque = self.averageHoldingTorque * self.avgHoldingTRQULFactor
    # get all measurement points inside the interval
    holdingTorqueIntervall = self.dataForHoldingTorque.loc[
        (self.dataForHoldingTorque['Torque'] >= lowerLimitAverageHoldingTorque) & (
                    self.dataForHoldingTorque['Torque'] <= upperLimitAverageHoldingTorque)]
    # get 1st point of the interval (index)
    startHoldingTorque = holdingTorqueIntervall.head(1)
    indexStartHoldingTorque = startHoldingTorque['Torque'].idxmax()
    # get last point of the interval (index)
    endHoldingTorque = self.dataHoldingTorque.tail(1)
    indexEndHoldingTorque = endHoldingTorque['Torque'].idxmax()
    dataPreciserHoldingTorque = self.measurementData.iloc[indexStartHoldingTorque:indexEndHoldingTorque, :]
    self.averagePreciseHoldingTorque = dataPreciserHoldingTorque['Torque'].mean()
    lowerAveragePreciseHoldingTorque = self.averagePreciseHoldingTorque * self.preciseAverageHoldingTRQLLFactor
    upperAveragePreciseHoldingTorque = self.averagePreciseHoldingTorque * self.preciseAverageHoldingTRQULFactor
    # data set: all data of holding TRQ
    self.preciseHoldingTorqueIntervall = self.dataForHoldingTorque.loc[
        (self.dataForHoldingTorque['Torque'] >= lowerAveragePreciseHoldingTorque) & (
                    self.dataForHoldingTorque['Torque'] <= upperAveragePreciseHoldingTorque)]

# calculate settling times
def calculateSettlingTimes(self):
    # get points
    holdingTorqueStart = self.preciseHoldingTorqueIntervall.iloc[0]
    holdingTorqueEnd = self.preciseHoldingTorqueIntervall.iloc[-1]
    # get their values
    self.holdingTorqueStartTime = holdingTorqueStart['Time']
    self.holdingTorqueEndTime = holdingTorqueEnd['Time']
    # calculations
    self.timeHoldingTorque = self.holdingTorqueEndTime - self.holdingTorqueStartTime
    self.settlingTimeTwo = self.holdingTorqueStartTime - self.applyInputTime
    self.settlingTimeOne = self.timePointFirstOutsideBandValues - self.applyInputTime

# calculate measurement time
def getMeasurementTime(self):
    self.measurementTime = self.holdingTorqueEndTime - self.applyInputTime

# print results
def printResultsPEAK(self):
    # print all values we're interested in:
    print("max overshoot torque:", '\t', self.maxOvershootTorqueValue, "Nm")
    print("max peak torque:", '\t', self.maxPeakTRQ, "Nm")
    # draw line between results
    print("-" * 60)
    print("slewrate:", '\t', '\t', self.slewratePeak, "Nm/ms")
    print("effective slewrate:", '\t', self.effectiveSlewrate, "Nm/ms")
    print("C value:", '\t', '\t', self.fit_C)
    print("rise time:", '\t', '\t', self.riseTime, "ms")
    print("-" * 60)
    print("settling time Max Peak:", '\t', self.settlingTimeOne, "seconds")
    print("settling time Holding:", '\t', self.settlingTimeTwo, "seconds")
    print("-" * 60)
    print("holding torque:", '\t', self.averagePreciseHoldingTorque, "Nm")
    print("holding torque time:", '\t', self.timeHoldingTorque, "seconds")
    print("holding torque start:", '\t', self.holdingTorqueStartTime, "seconds")
    print("holding torque end:", '\t', self.holdingTorqueEndTime, "seconds")
    print("-" * 60)
    print("measurement time:", '\t', self.measurementTime, "seconds")

# create plots
def getPlotsPEAK(self):
    # round values + set variablies for plots (torque: 2 digits; time: 3 digits)
    self.roundedSlewrate = round(self.slewratePeak, 2)
    self.roundedEffectiveSlewrate = round(self.effectiveSlewrate, 2)
    self.roundedMaxOvershootTRQ = round(self.maxOvershootTorqueValue, 2)
    self.roundedMaxPeakTRQ = round(self.maxPeakTRQ, 2)
    self.roundedHoldingTRQ = round(self.averagePreciseHoldingTorque, 2)
    roundedAverageNinetyPercentPointsTRQ = round(self.ninetyPercentTorque, 2)
    roundedAverageTenPercentPointsTRQ = round(self.tenPercentTorque, 2)
    roundedAverageNinetyPercentPointsTime = round(self.preciseTimeNinety, 3)
    roundedAverageTenPercentPointsTime = round(self.preciseTimeTen, 3)
    # set variables for 2% and 5% bands
    # for max peak torque
    maxPeakTRQPlotLLFive = self.morePreciseMaxPeakTRQ * 0.95
    maxPeakTRQPlotULFive = self.morePreciseMaxPeakTRQ * 1.05
    maxPeakTRQPlotLLTwo = self.morePreciseMaxPeakTRQ * 0.98
    maxPeakTRQPlotULTwo = self.morePreciseMaxPeakTRQ * 1.02
    # for holding torque
    self.holdingTRQPlotLLFive = self.averagePreciseHoldingTorque * 0.95
    self.holdingTRQPlotULFive = self.averagePreciseHoldingTorque * 1.05
    self.holdingTRQPlotLLTwo = self.averagePreciseHoldingTorque * 0.98
    self.holdingTRQPlotULTwo = self.averagePreciseHoldingTorque * 1.02
    # concat slewrate; first have to cast slewrate into string
    self.slewrateString = str(self.roundedSlewrate)
    self.slewrateText = "Slewrate: " + self.slewrateString
    # concat effective slewrate
    self.effectiveSlewrateString = str(self.roundedEffectiveSlewrate)
    self.effectiveSlewrateText = "Effective slewrate: " + self.effectiveSlewrateString
    # concat maxOvershootTorque
    maxOvershootString = str(self.roundedMaxOvershootTRQ)
    self.maxOvershootText = "   " + maxOvershootString
    # concat maxPeakTorque
    maxPeakString = str(self.roundedMaxPeakTRQ)
    self.maxPeakText = "   " + maxPeakString
    # concat holdingTorque
    self.holdingTRQString = str(self.roundedHoldingTRQ)
    self.holdingTRQText = "   " + self.holdingTRQString
    # concat 90% peak torque (TRQ)
    ninetyPercentPointsTRQString = str(roundedAverageNinetyPercentPointsTRQ)
    self.ninetyPercentPointsTRQText = "   " + ninetyPercentPointsTRQString
    # concat 10% peak torque (TRQ)
    tenPercentPointsTRQString = str(roundedAverageTenPercentPointsTRQ)
    self.tenPercentPointsTRQText = "   " + tenPercentPointsTRQString
    # concat 90% peak torque (Time)
    ninetyPercentPointsTimeString = str(roundedAverageNinetyPercentPointsTime)
    self.ninetyPercentPointsTimeText = "  " + ninetyPercentPointsTimeString
    # concat 10% peak torque (Time)
    tenPercentPointsTimeString = str(roundedAverageTenPercentPointsTime)
    self.tenPercentPointsTimeText = "  " + tenPercentPointsTimeString
    # height of first plot
    self.heightOfPlot = self.maxOvershootTorqueValue + 1
    # fix slewrate appearance for 2nd plot (in the middle between of 90% (time) and 10% (time))
    self.positionSlewrateTime = (self.preciseTimeNinety + self.preciseTimeTen) / 2
    # in between of 90% (torque) and 10% (torque)
    self.positionSlewrateTorque = (self.ninetyPercentTorque + self.tenPercentTorque) / 2
    # fix effective slewrate appearance for plot
    self.positionEffectiveSlewrateTime = (self.valueTimeNinetyPercentTRQCurve + self.timeTenPercentTRQCurve) / 2
    self.positionEffectiveSlewrateTorque = (self.ninetyPercentTorque + self.tenPercentTorque) / 2
    # 2nd plot: fix start and end of plot
    # fix start
    self.startSecondPlot = self.applyInputTime - 0.004
    # fix end
    self.endSecondPlot = self.preciseTimeNinety + 0.005
    # 5th plot: fix start end of plot
    # fix start
    self.startThirdPlot = self.holdingTorqueStartTime - 0.15
    # fix end
    self.endThirdPlot = self.holdingTorqueStartTime + 0.4
    # 6th plot
    # fix start
    self.startFourthPlot = self.timeLastMaxPeakPointValues - 0.20
    # fix end
    self.endFourthPlot = self.timeLastMaxPeakPointValues + 0.35
    # get variables to adjust the placement of text/values in the plots
    # assign the variables from config excel
    plotHeightPositiveFactor = self.getCorrectSetting['PlotHeightPositiveFactor']
    plotHeightNegativeFactor = self.getCorrectSetting['PlotHeightNegativeFactor']
    # set variables for text positioning
    self.positionMaxOvershootTRQ = self.maxOvershootTorqueValue + plotHeightPositiveFactor
    self.positionMaxPeakTRQ = self.maxPeakTRQ + plotHeightPositiveFactor
    self.positionOneHoldingTRQ = self.averagePreciseHoldingTorque + plotHeightPositiveFactor
    positionTwoHoldingTRQ = self.averagePreciseHoldingTorque - plotHeightNegativeFactor
    self.positionOneNinetyPercentTRQ = self.ninetyPercentTorque + plotHeightPositiveFactor
    self.positionTwoNinetyPercentTRQ = self.ninetyPercentTorque - plotHeightNegativeFactor
    self.positionOneTenPercentTRQ = self.tenPercentTorque + plotHeightPositiveFactor
    self.positionTwoTenPercentTRQ = self.tenPercentTorque - plotHeightNegativeFactor
    # set font size for all plots
    plt.rcParams.update({'font.size': 12})

    #############################################################################################
    # PLOTS
    # 1st plot: Show the whole measurement graph
    # 2nd plot: Show slewrate part
    # 3rd & 4th plot: Effective slewrate plots
    # 5th plot: Show start of holding torque
    # 6th plot: Show area of max peak torque
    ###### plot 1
    fig = plt.figure(figsize=(20, 10))
    ax = fig.add_subplot(111)
    x = list(self.measurementData['Time'])
    y = list(self.measurementData['Torque'])
    line, = ax.plot(x, y)
    ax.set_ylim(0, self.heightOfPlot)
    # Slewrate
    ax.plot(self.slewratePeak, label='Slew rate:                       ' + self.slewrateString + ' Nm/ms', color='white', alpha=0)
    # Effective Slewrate
    ax.plot(self.slewratePeak, label='Effective Slew rate:        ' + self.effectiveSlewrateString + ' Nm/ms', color='white',
            alpha=0)
    # show start of holding torque
    # show holding torque
    plt.axhline(self.averagePreciseHoldingTorque, color='magenta', linestyle='--',
                label="holding torque:            " + self.holdingTRQText + ' Nm')
    plt.text(0, self.positionOneHoldingTRQ, self.roundedHoldingTRQ)
    # show max peak torque
    plt.axhline(self.maxPeakTRQ, color='darkblue', linestyle='--', label="max peak torque:        " + self.maxPeakText + ' Nm')
    plt.text(0, self.positionMaxPeakTRQ, self.roundedMaxPeakTRQ)
    # show max overshoot torque
    plt.axhline(self.maxOvershootTorqueValue, color='black', linestyle='--',
                label='Max overshoot torque:' + self.maxOvershootText + ' Nm')
    plt.text(0, self.positionMaxOvershootTRQ, self.roundedMaxOvershootTRQ)
    plt.axvline(self.holdingTorqueStartTime, color="green", label="Start of holding torque")
    # show 90% max peak point
    # to make the points appear in the foreground, set a positive zorder value
    plt.scatter([self.preciseTimeNinety], [self.ninetyPercentTorque], color="darkred", zorder=5, label="90% of max peak")
    # show 10% max peak point
    plt.scatter([self.preciseTimeTen], [self.tenPercentTorque], color="red", zorder=5, label="10% of max peak")
    # show apply input point
    plt.scatter([self.applyInputTime], [self.applyInputTorque], color="lime", zorder=5, label="Apply input")
    plt.xlabel("Time[Seconds]")
    plt.ylabel("Torque[Nm]")
    # color band
    ax.axhspan(self.holdingTRQPlotLLTwo, self.holdingTRQPlotULTwo, alpha=0.75, color='lightblue',
               label="2% band holding torque (reference)", zorder=1)
    ax.axhspan(self.holdingTRQPlotLLFive, self.holdingTRQPlotULFive, alpha=0.4, color='yellow',
               label="5% band holding torque (reference)", zorder=0)
    # check for clockwise
    if self.antiClockwise == True:
        plt.title("Torque Measurement (Anti clockwise)")
    else:
        plt.title("Torque Measurement (Clockwise)")
    
    # plt.title("Torque Measurement")
    # plt.legend(loc='lower center')
    plt.legend(loc=(1.02, 0))
    plot_filename = 'TorqueMeasurement_' + self.fname + '.png'
    plot_filepath = os.path.join(self.result_folder, plot_filename)
    plt.savefig(plot_filepath, bbox_inches="tight")
    # plt.show()

    ###### plot 2
    fig = plt.figure(figsize=(20, 10))
    ax = fig.add_subplot(111)
    x = list(self.measurementData['Time'])
    y = list(self.measurementData['Torque'])
    line, = ax.plot(x, y)
    ax.set_ylim(0, self.heightOfPlot)
    ax.set_xlim(self.startSecondPlot, self.endSecondPlot)
    # show max peak torque
    plt.axhline(self.maxPeakTRQ, color='darkblue', linestyle='--', label="max peak torque")
    plt.text(self.startSecondPlot, self.positionMaxPeakTRQ, self.maxPeakText)
    # show max overshoot torque
    plt.axhline(self.maxOvershootTorqueValue, color='black', linestyle='--', label="max overshoot torque")
    plt.text(self.startSecondPlot, self.positionMaxOvershootTRQ, self.maxOvershootText)
    # show time of 90% point
    plt.axvline(self.preciseTimeNinety, color="gray", label="Time 90% max peak")
    # show time of 10% point
    plt.axvline(self.preciseTimeTen, color="orange", label="Time 10% max peak")
    # show torque of 90% point
    plt.axhline(self.ninetyPercentTorque, color="gray", linestyle='--', label="Torque 90% max peak")
    plt.text(self.startSecondPlot, self.positionOneNinetyPercentTRQ, self.ninetyPercentPointsTRQText)
    # show torque of 10% point
    plt.axhline(self.tenPercentTorque, color="orange", linestyle='--', label="Torque 10% max peak")
    plt.text(self.startSecondPlot, self.positionOneTenPercentTRQ, self.tenPercentPointsTRQText)
    plt.text(self.positionSlewrateTime, self.positionSlewrateTorque, self.slewrateText)
    point1 = [self.preciseTimeTen, self.tenPercentTorque]
    point2 = [self.preciseTimeNinety, self.ninetyPercentTorque]
    x_values = [point1[0], point2[0]]
    y_values = [point1[1], point2[1]]
    # draw a line between the measurement points for the slewrate
    plt.plot(x_values, y_values, color="turquoise", linestyle="dotted", label="Slewrate")

    # show prcise 10% TRQ Time and precise 90% TRQ Time (all: color: darkblue)
    # 10
    plt.scatter([self.preciseTimeTen], [self.tenPercentTorque], color="darkblue", zorder=5)
    plt.text(self.preciseTimeTen, self.positionTwoTenPercentTRQ, self.tenPercentPointsTimeText)
    # 90
    plt.scatter([self.preciseTimeNinety], [self.ninetyPercentTorque], color="darkblue", zorder=5)
    plt.text(self.preciseTimeNinety, self.positionTwoNinetyPercentTRQ, self.ninetyPercentPointsTimeText)
    # 0
    plt.scatter([self.applyInputTime], [self.applyInputTorque], color="darkblue", zorder=5)
    # display reference points
    plt.scatter([self.tenPercentPointTime], [self.tenPercentPointTRQ], color="lightgray", zorder=4)
    plt.scatter([self.ninetyPercentPointTime], [self.ninetyPercentPointTRQ], color="lightgray", zorder=4)
    # display area of slewrate calculation; uncommented for now
    ax.axvspan(self.preciseTimeTen, self.preciseTimeNinety, alpha=0.03, color='lime', zorder=1)
    plt.xlabel("Time[Seconds]")
    plt.ylabel("Torque[Nm]")
    # check for clockwise
    if self.antiClockwise == True:
        plt.title("Slewrate (Anti clockwise)")
    else:
        plt.title("Slewrate (Clockwise)")
    
    # plt.title("Torque Measurement")
    # plt.legend(loc='lower right')
    plt.legend(loc=(1.02, 0))
    figure_path = os.path.join(self.result_folder, 'Slewrate_' + self.fname + '..png')
    plt.savefig(figure_path, bbox_inches="tight")
    # plt.show()

    ########################################## plot 3: effective slewrate
    fig = plt.figure(figsize=(20, 10))
    ax = fig.add_subplot(111)
    plt.plot(self.xDataForFit, self.yDataForFit)
    plt.plot(self.xDataForFit, self.fit_yData, color="crimson", label="Curve Fit")
    plt.axhline(self.ninetyPercentTorque, color='gray', linestyle='--', label="Torque 90% max peak")
    plt.axhline(self.tenPercentTorque, color="orange", linestyle='--', label="Torque 10% max peak")
    ax.set_ylim(0, self.heightOfPlot)
    # set plot width with time values; starts at 0.00 seconds as this plot uses time shifted data
    ax.set_xlim(-0.0001, 0.025)
    # show effective slewrate
    plt.text(self.positionEffectiveSlewrateTime, self.positionEffectiveSlewrateTorque, self.effectiveSlewrateText)
    point1 = [self.timeTenPercentTRQCurve, self.tenPercentTorque]
    point2 = [self.valueTimeNinetyPercentTRQCurve, self.ninetyPercentTorque]
    x_values = [point1[0], point2[0]]
    y_values = [point1[1], point2[1]]
    plt.scatter([self.timeTenPercentTRQCurve], [self.tenPercentTorque], color="mediumblue", zorder=5)
    plt.scatter([self.valueTimeNinetyPercentTRQCurve], [self.ninetyPercentTorque], color="mediumblue", zorder=5)

    # draw a line between the measurement points for the effecitve slewrate
    plt.plot(x_values, y_values, color="turquoise", linestyle="dotted", label="Slewrate (effective)")
    plt.xlabel("Time shifted[Seconds]")
    plt.ylabel("Torque[Nm]")
    # check for clockwise
    if self.antiClockwise == True:
        plt.title("Effective slewrate [e^x curve fit (Anti clockwise)]")
    else:
        plt.title("Effective slewrate [e^x curve fit (Clockwise)]")
    #
    plt.legend(loc=(1.02, 0))
    figure_path = os.path.join(self.result_folder, 'EffectiveSlewrateZoom_' + self.fname + '..png')
    plt.savefig(figure_path, bbox_inches="tight")
    # plt.show()
    ########################################## plot 3: effective slewrate

    ########################################## plot 4: effective slewrate (bigger)
    fig = plt.figure(figsize=(20, 10))
    ax = fig.add_subplot(111)
    plt.plot(self.xDataForFit, self.yDataForFit)
    plt.plot(self.xDataForFit, self.fit_yData, color="crimson", label="Curve Fit")
    plt.axhline(self.ninetyPercentTorque, color='gray', linestyle='--', label="Torque 90% max peak")
    plt.axhline(self.tenPercentTorque, color="orange", linestyle='--', label="Torque 10% max peak")
    ax.set_ylim(0, self.heightOfPlot)
    ax.set_xlim(-0.005, 0.3)
    # show effective slewrate
    plt.text(self.positionEffectiveSlewrateTime, self.positionEffectiveSlewrateTorque, self.effectiveSlewrateText)
    point1 = [self.timeTenPercentTRQCurve, self.tenPercentTorque]
    point2 = [self.valueTimeNinetyPercentTRQCurve, self.ninetyPercentTorque]
    x_values = [point1[0], point2[0]]
    y_values = [point1[1], point2[1]]
    plt.scatter([self.timeTenPercentTRQCurve], [self.tenPercentTorque], color="mediumblue", zorder=5)
    plt.scatter([self.valueTimeNinetyPercentTRQCurve], [self.ninetyPercentTorque], color="mediumblue", zorder=5)

    # draw a line between the measurement points for the effecitve slewrate
    plt.plot(x_values, y_values, color="turquoise", linestyle="dotted", label="Slewrate (effective)")
    plt.xlabel("Time shifted[Seconds]")
    plt.ylabel("Torque[Nm]")
    # check for clockwise
    if self.antiClockwise == True:
        plt.title("Effective slewrate [e^x curve fit (Anti clockwise)]")
    else:
        plt.title("Effective slewrate [e^x curve fit (Clockwise)]")
    
    plt.legend(loc=(1.02, 0))
    figure_path = os.path.join(self.result_folder, 'EffectiveSlewrate_' + self.fname + '..png')
    plt.savefig(figure_path, bbox_inches="tight")
    # plt.show()
    ########################################## plot 4: effective slewrate (bigger)

    ###### plot 5
    fig = plt.figure(figsize=(20, 10))
    ax = fig.add_subplot(111)
    x = list(self.measurementData['Time'])
    y = list(self.measurementData['Torque'])
    line, = ax.plot(x, y)
    ax.set_ylim(0, self.heightOfPlot)
    ax.set_xlim(self.startThirdPlot, self.endThirdPlot)
    # show holding torque line
    plt.axhline(self.averageHoldingTorque, color='magenta', linestyle='--', label="holding torque" + self.holdingTRQText + " Nm",
                linewidth=2)
    # show start of holding torque
    plt.axvline(self.holdingTorqueStartTime, color="green", label="Start of holding torque")
    # color band of holding TRQ
    ax.axhspan(self.holdingTRQPlotLLTwo, self.holdingTRQPlotULTwo, alpha=0.75, color='lightblue',
               label="2% band holding torque (reference)", zorder=1)
    ax.axhspan(self.holdingTRQPlotLLFive, self.holdingTRQPlotULFive, alpha=0.4, color='yellow',
               label="5% band holding torque (reference)", zorder=0)
    plt.xlabel("Time[Seconds]")
    plt.ylabel("Torque[Nm]")
    # check for clockwise
    if self.antiClockwise == True:
        plt.title("Beginning of holding torque (Anti clockwise)")
    else:
        plt.title("Beginning of holding torque (Clockwise)")
    #
    # plt.title("Torque Measurement")
    plt.legend(loc=(1.02, 0))
    figure_path = os.path.join(self.result_folder, 'BeginningOfHoldingTRQ_' + self.fname + '.png')
    plt.savefig(figure_path, bbox_inches="tight")
    #plt.show()

    ###### plot 6
    self.bandTRQUL = self.morePreciseMaxPeakTRQ * self.bandPercentageUL
    self.bandTRQLL = self.morePreciseMaxPeakTRQ * self.bandPercentageLL
    if self.bandPercentageUL == 1.05:
        labelTextBand = "5% band max peak torque (reference)"
        self.bandNoteExcel = "5% band max peak torque"
    else:
        labelTextBand = "2% band max peak torque (reference)"
        self.bandNoteExcel = "2% band max peak torque"

    dataSetShowMaxPeak = self.measurementData.iloc[self.indexMaxOvershoot:self.indexMaxOvershoot + 20000, :]
    fig = plt.figure(figsize=(20, 10))
    ax = fig.add_subplot(111)
    x = list(self.measurementData['Time'])
    y = list(self.measurementData['Torque'])
    line, = ax.plot(x, y)
    ax.set_ylim(0, self.heightOfPlot)
    ax.set_xlim(self.startFourthPlot, self.endFourthPlot)
    # show settling point (settling time 1)
    plt.axvline(self.timePointFirstOutsideBandValues, color="black", label="settling point (settling time 1)")
    plt.axhline(self.maxPeakTRQ, color='darkblue', linestyle='--', label="max peak torque" + self.maxPeakText + " Nm",
                linewidth=2)
    # show band (5% +- max peak TRQ)
    ax.axhspan(self.bandTRQLL, self.bandTRQUL, alpha=0.4, color='red', label=labelTextBand, zorder=0)
    # display area of slewrate calculation; uncommented for now
    ax.axvspan(self.timePointFirstOutsideBandValues, self.timePointFirstOutsideBandValuesTwo, alpha=0.1,
               label="area of max peak calc.", color='lime', zorder=1)
    plt.xlabel("Time[Seconds]")
    plt.ylabel("Torque[Nm]")
    # check for clockwise
    if self.antiClockwise == True:
        plt.title("Max Peak Torque (Anti clockwise)")
    else:
        plt.title("Max Peak Torque (Clockwise)")
    #
    # plt.title("Torque Measurement")
    plt.legend(loc=(1.02, 0))
    figure_path = os.path.join(self.result_folder, 'MaxPeakTRQ_' + self.fname + '.png')
    plt.savefig(figure_path, bbox_inches="tight")
    # plt.show()

# create excel file with results
def writeExcelResultsPEAK(self):
    workbook_path = os.path.join(self.result_folder, "Result_" + self.fname + ".xlsx")
    workbook = xlsxwriter.Workbook(workbook_path)
    # by default worksheet names in the spreadsheet will be Sheet1, Sheet2, etc.
    worksheet = workbook.add_worksheet("My sheet")
    # create text / concat
    checkAntiClock = 'Clockwise'
    if self.antiClockwise == True:
        checkAntiClock = ' (Anticlockwise)'
    textCycle = 'Cycle' + checkAntiClock
    # data we want to write to the worksheet
    scores = (
        ['Parameters [unit]', textCycle],
        ['Direction [CW/ACW]', checkAntiClock],
        ['max overshoot torque [Nm]', self.maxOvershootTorqueValue],
        ['max peak torque [Nm]', self.maxPeakTRQ],
        ['slewrate [Nm/ms]', self.slewratePeak],
        ['effective slewrate [Nm/ms]', self.effectiveSlewrate],
        ['rise time [ms]', self.riseTime],
        ['Max Peak TRQ settling time [s]', self.settlingTimeOne],
        ['Holding TRQ settling time [s]', self.settlingTimeTwo],
        ['holding torque [Nm]', self.averagePreciseHoldingTorque],
        ['holding torque time [s]', self.timeHoldingTorque],
        ['holding torque start [s]', self.holdingTorqueStartTime],
        ['holding torque end [s]', self.holdingTorqueEndTime],
        ['measurement time [s]', self.measurementTime],
        ['MAX zero error [Nm]', self.maxZeroError],
        ['MIN zero error [Nm]', self.minZeroError],
        ['Average zero error (+-) [Nm]', self.averageZeroError],
        ['Note (TRQ reference band):', self.bandNoteExcel],
        ['Apply input [Index]', self.indexGetInputTime],
        ['Apply input time [s]', self.applyInputTime],
        ['Apply input torque [Nm]', self.applyInputTorque],
        ['Actual 10% time [s]', self.tenPercentPointTime],
        ['Actual 10% torque [Nm]', self.tenPercentPointTRQ],
        ['Corrected 10% time [s]', self.preciseTimeTen],
        ['Corrected 10% torque [Nm]', self.tenPercentTorque],
        ['Actual 90% time [s]', self.ninetyPercentPointTime],
        ['Actual 90% torque [Nm]', self.ninetyPercentPointTRQ],
        ['Corrected 90% time [s]', self.preciseTimeNinety],
        ['Corrected 90% torque [Nm]', self.ninetyPercentTorque],
    )
    # start from the first cell
    # rows and columns are zero indexed
    row = 0
    col = 0
    # iterate over the data and write it out row by row
    for name, score in (scores):
        worksheet.write(row, col, name)
        worksheet.write(row, col + 1, score)
        row += 1
    workbook.close()