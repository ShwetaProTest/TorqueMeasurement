#!/usr/bin/env python
# coding: utf-8

# import libraries
import pandas as pd
import os
import numpy as np
import matplotlib.pyplot as plt
import xlsxwriter
import scipy.optimize
import scipy as sy
import math
from scipy.optimize import curve_fit


########################################################################################################
# LIN file                                                                                             
# created by:      RHM, KZE                                                                            
# purpose:         Provide functions for main file:                                                    

#                  Calculate slewrate
#                  Calculate rise time


#                  Print calculated values
#                  Create plots (LIN: 5 plots)
#                  Create excel file (contains calculated values)
# created:         19.10.2022 - 14.12.2022
# Modified:        20.01.2023
# Modified by:     RHM
# version:         C
########################################################################################################
# calculate more precise max peak TRQ
def calculateMorePreciseMaxPeakTRQLIN(self):
    if self.maxPeakTRQ < 3.5:
        # no band
        # set flags
        self.bandTRQFlag = True
        self.settlingTimeFlag = True
        self.morePreciseMaxPeakTRQ = self.maxPeakTRQ

    elif 3.5 <= self.maxPeakTRQ < 15:
        # band 5%
        # if band empty --> 2%
        # don't set flag
        datasetSettling = self.measurementData.iloc[self.indexMaxOvershoot:self.indexMaxOvershoot + 7500, :]
        allPointsFivePercentBand = datasetSettling.loc[(datasetSettling['Torque'] >= self.maxPeakTRQ * 1.05) | (
                datasetSettling['Torque'] <= self.maxPeakTRQ * 0.95)]
        counterBandValues = len(allPointsFivePercentBand)

        # if band empty --> 2%
        if counterBandValues == 0:
            bandPercentageValueUL = 1.02
            bandPercentageValueLL = 0.98
            self.bandTRQUL = 1.02
            self.bandTRQLL = 0.98
            allPointsFivePercentBand = datasetSettling.loc[(datasetSettling['Torque'] >= self.maxPeakTRQ * 1.02) | (
                    datasetSettling['Torque'] <= self.maxPeakTRQ * 0.98)]
            counterBandValues = len(allPointsFivePercentBand)
            if counterBandValues == 0:
                bandPercentageValueUL = 1.015
                bandPercentageValueLL = 0.985
                self.bandTRQUL = 1.015
                self.bandTRQLL = 0.985
        else:
            bandPercentageValueUL = 1.05
            bandPercentageValueLL = 0.95
            self.bandTRQUL = 1.05
            self.bandTRQLL = 0.95

        allPointsFivePercentBand = datasetSettling.loc[
            (datasetSettling['Torque'] >= self.maxPeakTRQ * bandPercentageValueUL) | (
                    datasetSettling['Torque'] <= self.maxPeakTRQ * bandPercentageValueLL)]
        self.morePreciseMaxPeakTRQ = self.maxPeakTRQ
        # 1st point
        pointFirstOutsideBand = allPointsFivePercentBand.tail(1)
        indexPointFirstOutsideBand = pointFirstOutsideBand['Torque'].idxmax()
        pointFirstOutsideBandValues = allPointsFivePercentBand.iloc[-1]
        self.timePointFirstOutsideBandValues = pointFirstOutsideBandValues['Time']
        # 2nd point

        datasetMaxPeakBandSecond = self.measurementData.iloc[self.indexMaxOvershoot + 7500:, :]
        allBandPointsTwo = datasetMaxPeakBandSecond.loc[
            (datasetMaxPeakBandSecond['Torque'] >= self.maxPeakTRQ * bandPercentageValueUL) | (
                    datasetMaxPeakBandSecond['Torque'] <= self.maxPeakTRQ * bandPercentageValueLL)]
        pointFirstOutsideBandTwo = allBandPointsTwo.head(1)
        indexPointFirstOutsideBandTwo = pointFirstOutsideBandTwo['Torque'].idxmax()
        pointFirstOutsideBandValuesTwo = allBandPointsTwo.iloc[0]
        self.timePointFirstOutsideBandValuesTwo = pointFirstOutsideBandValuesTwo['Time']
        # 3rd calc max peak
        datasetThirdCalc = self.measurementData.iloc[indexPointFirstOutsideBand:indexPointFirstOutsideBandTwo, :]
        thirdMaxPeak = datasetThirdCalc['Torque'].mean()
        self.maxPeakTRQ = thirdMaxPeak
        self.morePreciseMaxPeakTRQ = self.maxPeakTRQ

        # don't set flags
        self.bandTRQFlag = False
        self.settlingTimeFlag = False

    # self.maxPeakTRQ > 15 Nm
    else:
        # band 2%
        # if band empty --> 1%
        # don't set flag
        datasetSettling = self.measurementData.iloc[self.indexMaxOvershoot:self.indexMaxOvershoot + 7500, :]
        allPointsFivePercentBand = datasetSettling.loc[(datasetSettling['Torque'] >= self.maxPeakTRQ * 1.02) | (
                datasetSettling['Torque'] <= self.maxPeakTRQ * 0.98)]
        counterBandValues = len(allPointsFivePercentBand)

        # if band empty --> 1%
        if counterBandValues == 0:
            bandPercentageValueUL = 1.01
            bandPercentageValueLL = 0.99
            self.bandTRQUL = 1.01
            self.bandTRQLL = 0.99
        else:
            bandPercentageValueUL = 1.02
            bandPercentageValueLL = 0.98
            self.bandTRQUL = 1.02
            self.bandTRQLL = 0.98

        allPointsFivePercentBand = datasetSettling.loc[
            (datasetSettling['Torque'] >= self.maxPeakTRQ * bandPercentageValueUL) | (
                    datasetSettling['Torque'] <= self.maxPeakTRQ * bandPercentageValueLL)]
        self.morePreciseMaxPeakTRQ = self.maxPeakTRQ
        # 1st point
        pointFirstOutsideBand = allPointsFivePercentBand.tail(1)

        indexPointFirstOutsideBand = pointFirstOutsideBand['Torque'].idxmax()
        pointFirstOutsideBandValues = allPointsFivePercentBand.iloc[-1]

        self.timePointFirstOutsideBandValues = pointFirstOutsideBandValues['Time']
        # 2nd point
        datasetMaxPeakBandSecond = self.measurementData.iloc[self.indexMaxOvershoot + 7500:, :]
        allBandPointsTwo = datasetMaxPeakBandSecond.loc[
            (datasetMaxPeakBandSecond['Torque'] >= self.maxPeakTRQ * bandPercentageValueUL) | (
                    datasetMaxPeakBandSecond['Torque'] <= self.maxPeakTRQ * bandPercentageValueLL)]
        pointFirstOutsideBandTwo = allBandPointsTwo.head(1)

        indexPointFirstOutsideBandTwo = pointFirstOutsideBandTwo['Torque'].idxmax()
        pointFirstOutsideBandValuesTwo = allBandPointsTwo.iloc[0]

        self.timePointFirstOutsideBandValuesTwo = pointFirstOutsideBandValuesTwo['Time']
        # 3rd calc max peak
        datasetThirdCalc = self.measurementData.iloc[indexPointFirstOutsideBand:indexPointFirstOutsideBandTwo, :]
        thirdMaxPeak = datasetThirdCalc['Torque'].mean()
        self.maxPeakTRQ = thirdMaxPeak
        self.morePreciseMaxPeakTRQ = self.maxPeakTRQ

        # don't set flags
        self.bandTRQFlag = False
        self.settlingTimeFlag = False


# calculate slewrate LIN
def calculateSlewrateLIN(self):
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
    allTenPercentPoints = self.dataUntilMaxOvershoot.loc[
        (self.dataUntilMaxOvershoot['Torque'] >= tenPercentTorqueLowerLimit) & (
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
    inputTimeCalcData = self.dataUntilMaxOvershoot.loc[
        (self.dataUntilMaxOvershoot['Time'] >= self.preciseTimeZero - 0.000027) & (
                self.dataUntilMaxOvershoot['Time'] <= self.preciseTimeZero + 0.000027)]
    getInputTimeIndex = inputTimeCalcData.head(1)
    self.indexGetInputTime = getInputTimeIndex['Torque'].idxmax()

    # calculate slew rate
    slewratePeakWithSeconds = (self.ninetyPercentTorque - self.tenPercentTorque) / (
                self.preciseTimeNinety - self.preciseTimeTen)
    # divide by 1000 --> get Nm/ms (otherwise it would be Nm/s)
    self.slewratePeak = slewratePeakWithSeconds / 1000

    # Effective Slewrate with curve fit (e^x)
    # Define window for e^x curve fit
    intIndexStartPolynomialFit = self.indexGetInputTime
    intIndexEndPolynomialFit = self.indexGetInputTime + 50000
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
    timeNinetyPercentTRQCurve = -(math.log((self.maxPeakTRQ - self.ninetyPercentTorque) / self.maxPeakTRQ)) / (
            parameters / 2)
    self.timeTenPercentTRQCurve = self.preciseTimeTen - timeStartPolynomialFit
    # get value of array
    self.valueTimeNinetyPercentTRQCurve = timeNinetyPercentTRQCurve[0]

    self.effectiveSlewrate = (self.ninetyPercentTorque - self.tenPercentTorque) / (
            self.valueTimeNinetyPercentTRQCurve - self.timeTenPercentTRQCurve)
    self.effectiveSlewrate = self.effectiveSlewrate / 1000


# calculate rise time
def calculateRisetimeLIN(self):
    # calculate rise time
    riseTimeInSeconds = self.preciseTimeNinety - self.preciseTimeTen
    # Multiply by 1000 --> get time in ms (not seconds)
    self.riseTime = riseTimeInSeconds * 1000


# calculate settling time
def calculateSettlingTimeLIN(self):
    if self.settlingTimeFlag == False:
        self.settlingTime = self.timePointFirstOutsideBandValues - self.applyInputTime
    else:
        self.settlingTime = "Not possible to calculate"


# calculate measurement time
def getMeasurementTimeLIN(self):
    holdingTRQEnd = self.tenThousandData.iloc[-1]
    timeHoldingTRQEnd = holdingTRQEnd['Time']
    self.measurementTime = timeHoldingTRQEnd - self.applyInputTime


# calculate holding TRQ with 5 seconds (if possible)
def calculatePreciseHoldingTRQ(self):
    if self.measurementTime >= 7:
        # 5 seconds holding TRQ
        lowerLimitAverageHoldingTorque = self.averageHoldingTorque * self.avgHoldingTRQLLFactor
        upperLimitAverageHoldingTorque = self.averageHoldingTorque * self.avgHoldingTRQULFactor
        # get all measurement points inside the interval
        holdingTRQIntervall = self.dataForHoldingTorque.loc[
            (self.dataForHoldingTorque['Torque'] >= lowerLimitAverageHoldingTorque) & (
                    self.dataForHoldingTorque['Torque'] <= upperLimitAverageHoldingTorque)]
        # get last point of interval (index)
        endHoldingTRQ = holdingTRQIntervall.tail(1)
        indexEndHoldingTRQ = endHoldingTRQ['Torque'].idxmax()
        # get time of last point
        EndHoldingTRQ = holdingTRQIntervall.iloc[-1]
        timeEndHoldingTRQ = EndHoldingTRQ['Time']
        # get all data inside the 5 seconds of holding TRQ
        dataHoldingTRQFiveSeconds = self.measurementData.loc[
            (self.measurementData['Time'] >= timeEndHoldingTRQ - 5) & (
                        self.measurementData['Time'] <= timeEndHoldingTRQ)]
        self.preciseHoldingTRQ = dataHoldingTRQFiveSeconds['Torque'].mean()
    else:
        self.preciseHoldingTRQ = self.averageHoldingTorque


# print results
def printResultsLIN(self):
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
    print("settling time:", '\t', '\t', self.settlingTime, "seconds")

    print("-" * 60)
    print("holding torque:", '\t', self.preciseHoldingTRQ, "Nm")

    print("-" * 60)
    print("measurement time:", '\t', self.measurementTime, "seconds")


# create plots
def getPlotsLIN(self):
    # round values + set variablies for plots (torque: 2 digits; time: 3 digits)
    self.roundedSlewrate = round(self.slewratePeak, 2)
    self.roundedEffectiveSlewrate = round(self.effectiveSlewrate, 2)
    self.roundedMaxOvershootTRQ = round(self.maxOvershootTorqueValue, 2)
    self.roundedMaxPeakTRQ = round(self.maxPeakTRQ, 2)
    self.roundedHoldingTRQ = round(self.preciseHoldingTRQ, 2)
    holdingTRQString = str(self.roundedHoldingTRQ)
    self.roundedHoldingTRQ = holdingTRQString + " *"
    roundedAverageNinetyPercentPointsTRQ = round(self.ninetyPercentTorque, 2)
    roundedAverageTenPercentPointsTRQ = round(self.tenPercentTorque, 2)
    roundedAverageNinetyPercentPointsTime = round(self.preciseTimeNinety, 3)
    roundedAverageTenPercentPointsTime = round(self.preciseTimeTen, 3)
    # set variables for 2% and 5% band of holding torque (1st plot)

    self.maxPeakTRQPlotLLFive = self.morePreciseMaxPeakTRQ * 0.95
    self.maxPeakTRQPlotULFive = self.morePreciseMaxPeakTRQ * 1.05
    self.maxPeakTRQPlotLLTwo = self.morePreciseMaxPeakTRQ * 0.98
    self.maxPeakTRQPlotULTwo = self.morePreciseMaxPeakTRQ * 1.02

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
    # set variables
    # height of first plot
    self.heightOfPlot = self.maxOvershootTorqueValue + 1
    # fix slewrate appearance for 2nd plot
    # in between of 90% (time) and 10% (time)
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
    # get variables to adjust the placement of text/values in the plots
    latestHoldingPoint = self.tenThousandData.iloc[-1]
    timeLatestHoldingPoint = latestHoldingPoint['Time']
    # assign the variables from config excel
    plotHeightPositiveFactor = self.getCorrectSetting['PlotHeightPositiveFactor']
    plotHeightNegativeFactor = self.getCorrectSetting['PlotHeightNegativeFactor']
    # set variables for text positioning
    self.positionMaxOvershootTRQ = self.maxOvershootTorqueValue + plotHeightPositiveFactor
    self.positionMaxPeakTRQ = self.maxPeakTRQ + plotHeightPositiveFactor
    self.positionOneHoldingTRQ = self.averageHoldingTorque + plotHeightPositiveFactor
    positionTwoHoldingTRQ = self.averageHoldingTorque - plotHeightNegativeFactor
    self.positionOneNinetyPercentTRQ = self.ninetyPercentTorque + plotHeightPositiveFactor
    self.positionTwoNinetyPercentTRQ = self.ninetyPercentTorque - plotHeightNegativeFactor
    self.positionOneTenPercentTRQ = self.tenPercentTorque + plotHeightPositiveFactor
    self.positionTwoTenPercentTRQ = self.tenPercentTorque - plotHeightNegativeFactor
    self.positionHoldingTRQ = timeLatestHoldingPoint + 0.05
    # set font size for all plots
    plt.rcParams.update({'font.size': 12})

    #############################################################################################
    # PLOTS
    # 1st plot: Show the whole measurement graph
    # 2nd plot: Show slewrate part
    # 3rd & 4th plot: Effective slewrate plots

    # 5th plot: Show area of max peak torque (if possible)
    ###### plot 1
    fig = plt.figure(figsize=(20, 10))
    ax = fig.add_subplot(111)
    x = list(self.measurementData['Time'])
    y = list(self.measurementData['Torque'])
    line, = ax.plot(x, y)
    ax.set_ylim(0, self.heightOfPlot)
    # Slewrate
    ax.plot(self.slewratePeak, label='Slew rate:                       ' + self.slewrateString + ' Nm/ms',
            color='white')
    # Effective Slewrate
    ax.plot(self.slewratePeak, label='Effective Slew rate:        ' + self.effectiveSlewrateString + ' Nm/ms',
            color='white',
            alpha=0)
    # show holding torque line
    # colors: holding TRQ --> magenta, max peak --> turquoise, max overshoot --> black
    # Note: if only one (turquoise line) is displayed: max peak and holding TRQ are too close
    plt.axhline(self.preciseHoldingTRQ, color='magenta', linestyle='--',
                label="holding torque *:         " + self.holdingTRQText + ' Nm')
    plt.text(self.positionHoldingTRQ, self.positionOneHoldingTRQ, self.roundedHoldingTRQ)
    # show max peak torque line
    plt.axhline(self.maxPeakTRQ, color='darkblue', linestyle='--',
                label="max peak torque:        " + self.maxPeakText + ' Nm')
    plt.text(0, self.positionMaxPeakTRQ, self.roundedMaxPeakTRQ)
    # show max overshoot torque
    plt.axhline(self.maxOvershootTorqueValue, color='black', linestyle='--',
                label='Max overshoot torque:' + self.maxOvershootText + ' Nm')
    plt.text(0, self.positionMaxOvershootTRQ, self.roundedMaxOvershootTRQ)

    # show 90% max peak point
    # to make the points appear in the foreground, set a positive zorder value:
    plt.scatter([self.preciseTimeNinety], [self.ninetyPercentTorque], color="darkred", zorder=5,
                label="90% of max peak")
    # show 10% max peak point
    plt.scatter([self.preciseTimeTen], [self.tenPercentTorque], color="red", zorder=5, label="10% of max peak")
    # show apply input point
    plt.scatter([self.applyInputTime], [self.applyInputTorque], color="lime", zorder=5, label="Apply input")

    # color band of max peak torque
    ax.axhspan(self.maxPeakTRQPlotLLTwo, self.maxPeakTRQPlotULTwo, alpha=0.75, color='lightblue',
               label="2% band holding torque (reference)", zorder=1)
    ax.axhspan(self.maxPeakTRQPlotLLFive, self.maxPeakTRQPlotULFive, alpha=0.4, color='yellow',
               label="5% band holding torque (reference)", zorder=0)
    plt.xlabel("Time[Seconds]")
    plt.ylabel("Torque[Nm]")
    if self.antiClockwise == True:
        plt.title("Torque Measurement (Anti clockwise)")
    else:
        plt.title("Torque Measurement (Clockwise)")
    #
    # plt.title("Torque Measurement")

    plt.legend(loc=(1.02, 0))
    plot_filename = "TorqueMeasurement_" + self.fname + ".png"
    plot_filepath = os.path.join(self.result_folder, plot_filename)
    plt.savefig(plot_filepath, bbox_inches="tight")
    # plt.show()

    #### plot 2
    fig = plt.figure(figsize=(20, 10))
    ax = fig.add_subplot(111)

    line, = ax.plot(x, y)
    ax.set_ylim(0, self.heightOfPlot)
    ax.set_xlim(self.startSecondPlot, self.endSecondPlot)
    # show max peak torque line
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
    # show slewrate
    plt.text(self.positionSlewrateTime, self.positionSlewrateTorque, self.slewrateText)
    ############################################################
    point1 = [self.preciseTimeTen, self.tenPercentTorque]
    point2 = [self.preciseTimeNinety, self.ninetyPercentTorque]
    x_values = [point1[0], point2[0]]
    y_values = [point1[1], point2[1]]
    # draw a line between the measurement points for the slewrate
    plt.plot(x_values, y_values, color="turquoise", linestyle="dotted", label="Slewrate")
    ############################################################
    ax.axvspan(self.preciseTimeTen, self.preciseTimeNinety, alpha=0.03, color='lime', zorder=1)
    # show precise 10% TRQ Time and precise 90% TRQ Time
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

    plt.xlabel("Time[Seconds]")
    plt.ylabel("Torque[Nm]")

    if self.antiClockwise == True:
        plt.title("Slewrate (Anti clockwise)")
    else:
        plt.title("Slewrate (Clockwise)")
    #
    # plt.title("Torque Measurement")
    # here we should set the position of the legend to lower right
    plt.legend(loc=(1.02, 0))
    plot_filename = "Slewrate_" + self.fname + ".png"
    plot_filepath = os.path.join(self.result_folder, plot_filename)
    plt.savefig(plot_filepath, bbox_inches="tight")
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
    ax.set_xlim(-0.001, 0.025)
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
    plot_filename = "EffectiveSlewrateZoom_" + self.fname + ".png"
    plot_filepath = os.path.join(self.result_folder, plot_filename)
    plt.savefig(plot_filepath, bbox_inches="tight")
    # plt.show()
    ########################################## plot 3: effective slewrate

    ########################################## plot 4: effective slewrate (bigger)
    fig = plt.figure(figsize=(20, 10))
    ax = fig.add_subplot(111)
    plt.plot(self.xDataForFit, self.yDataForFit)
    plt.plot(self.xDataForFit, self.fit_yData, color="crimson", label="Curve Fit")
    plt.axhline(self.ninetyPercentTorque, color='gray', linestyle='--', label="Torque 90% max peak")
    plt.axhline(self.tenPercentTorque, color="orange", linestyle='--', label="Torque 10% max peak")
    ax.set_ylim(0,self.heightOfPlot)
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
    #
    plt.legend(loc=(1.02, 0))
    plot_filename = "EffectiveSlewrate_" + self.fname + ".png"
    plot_filepath = os.path.join(self.result_folder, plot_filename)
    plt.savefig(plot_filepath, bbox_inches="tight")
    # plt.show()
    ########################################## plot 4: effective slewrate (bigger)

    ###### plot 5 (only show if band exists)
    if self.bandTRQFlag == False:
        self.startFourthPlot = self.timePointFirstOutsideBandValues - 0.1
        self.endFourthPlot = self.timePointFirstOutsideBandValuesTwo + 0.1
        self.bandTRQULPlot = self.morePreciseMaxPeakTRQ * self.bandTRQUL
        self.bandTRQLLPlot = self.morePreciseMaxPeakTRQ * self.bandTRQLL

        if self.bandTRQUL == 1.05:
            labelTextBand = "5% band max peak torque (reference)"
            self.bandNoteExcel = "5% band max peak torque"
        elif self.bandTRQUL == 1.02:
            labelTextBand = "2% band max peak torque (reference)"
            self.bandNoteExcel = "2% band max peak torque"
        elif self.bandTRQUL == 1.015:
            labelTextBand = "1.5% band max peak torque (reference)"
            self.bandNoteExcel = "1.5% band max peak torque"
        else:
            labelTextBand = "1% band max peak torque (reference)"
            self.bandNoteExcel = "1% band max peak torque"

        dataSetShowMaxPeak = self.measurementData.iloc[self.indexMaxOvershoot:self.indexMaxOvershoot + 26000, :]
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
        linewidth = 2)
        # show band (% +- max peak TRQ)
        ax.axhspan(self.bandTRQLLPlot, self.bandTRQULPlot, alpha=0.4, color='red', label=labelTextBand, zorder=0)
        # display area of slewrate calculation; uncommented for now
        ax.axvspan(self.timePointFirstOutsideBandValues, self.timePointFirstOutsideBandValuesTwo, alpha=0.1,
                   label = "area of max peak calc.", color = 'lime', zorder = 1)
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
        plot_filename = "MaxPeakTRQ_" + self.fname + ".png"
        plot_filepath = os.path.join(self.result_folder, plot_filename)
        plt.savefig(plot_filepath, bbox_inches="tight")
        # plt.show()
    else:
        self.bandNoteExcel = "Too low FFB settings, no settling time calculation possible"

# create excel file with results
def writeExcelResultsLIN(self):
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
        ['Max Peak TRQ settling time [s]', self.settlingTime],
        ['Holding TRQ settling time [s]', 'NA'],
        ['holding torque [Nm]', self.preciseHoldingTRQ],
        ['holding torque time [s]', self.measurementTime - self.settlingTime],
        ['holding torque start [s]', self.timePointFirstOutsideBandValues],
        ['holding torque end [s]', self.timePointFirstOutsideBandValues + self.measurementTime - self.settlingTime],
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
    # rows and columns are zero indexed.
    row = 0
    col = 0
    # iterate over the data and write it out row by row
    for name, score in (scores):
        worksheet.write(row, col, name)
        worksheet.write(row, col + 1, score)
        row += 1
    workbook.close()