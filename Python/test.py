#Vissim-COM programming - Python example code
from __future__ import print_function
import os
# COM-Server
#this needs module install: >> pip install pywin32
import win32com.client as com

# Connecting the COM Server
# Vissim = com.gencache.EnsureDispatch("Vissim.Vissim") #
# once the cache has been generated, its faster to call Dispatch which also creates the connection to Vissim.
vis = com.Dispatch("Vissim.Vissim")

# Loading the traffic network
Filename               = os.path.join(os.path.abspath(os.getcwd()), 'test.inpx')
flag_read_additionally = True # you can read network(elements) additionally, in this case set "flag_read_additionally" to true
vis.LoadNet(Filename, flag_read_additionally)

# Load a Layout:
Filename = os.path.join(os.path.abspath(os.getcwd()), 'test.layx')
vis.LoadLayout(Filename)

## Simulation settings
period_time = 3600 # simulation second [s]
vis.Simulation.SetAttValue('SimPeriod', period_time)
step_time=3
vis.Simulation.SetAttValue('SimRes', step_time)

sim=vis.Simulation # create Simulation COM-interface
vnet=vis.Net # create Net COM-interface

## Setting the traffic demands of the network
vehins=vnet.VehicleInputs
vehin_1=vehins.ItemByKey(1)
vehin_1.SetAttValue('Volume(1)', 1500) # main road
vehin_2=vehins.ItemByKey(2)
vehin_2.SetAttValue('Volume(1)', 100) # side street
#'Volume(1)' means the first defined time interval


## The objects of the traffic signal control
scs=vnet.SignalControllers
sc=scs.ItemByKey(1)# sc = SignalController 1
sgs = sc.SGs # sgs=SignalGroups
sg_1 = sgs.ItemByKey(1)  # SignalGroup 1
sg_2 = sgs.ItemByKey(2)  # SignalGroup 2

dets = sc.Detectors 
det_all = dets.GetAll() #All detectors are queried first
det_1 = det_all[0]    #The first detector of the detectors
# Note that in Python tuple the 1st element is called by 0, although in Vissim GUI its number is 1!

## Measure total travel time in the network 
netperform=vnet.VehicleNetworkPerformanceMeasurement

## Access to DataCollectionPoint object
datapoints = vnet.DataCollectionMeasurements
datapoint1 = datapoints.ItemByKey(1)

## Access to Link object
links = vnet.Links
link_1 = links.ItemByKey(1)

## Running the simulation
# To run the simulation continuous (it stops at breakpoint or end of simulation): sim.RunContinuous()
# Activate QuickMode: vis.Graphics.CurrentNetworkWindow.SetAttValue("QuickMode", 1)
# Set maximum speed: sim.SetAttValue('UseMaxSimSpeed', True)

verify=20; # verifying at every 20 seconds
# Also set it in Vissim GUI! Here: Evaluation\Configuration...\Interval


vis.SaveNetAs(os.path.join(os.path.abspath(os.getcwd()), 'test.inpx'))

for i in range(0,period_time*step_time):
    sim.RunSingleStep()
    if (i / step_time) % verify == 0: # verifying at every 20 seconds
        demand = det_1.AttValue('Presence')  # get detector occupancy:0/1
        print(demand)
        if demand == 1:
            sg_1.SetAttValue('SigState',1) # main road red (1)
            sg_2.SetAttValue('SigState',3) # side street green (3)
        else: # no demand on loop -> main road's signal is green
            sg_1.SetAttValue('SigState',3)
            sg_2.SetAttValue('SigState',1)
            # total travel time at the end of each eval. interval:
            print('Travel time of the network: ') 
            # TravTmTot's subattributes in parentheses:
            # (SimulationRun, TimeInterval, VehicleClass)
            print(netperform.AttValue('TravTmTot(Current,Last,All)'))
            # Query the avg. speed and vehicle number at the end of each eval. interval:
            print(datapoint1.AttValue('Vehs(Current, Last, All)'))
            print(datapoint1.AttValue('Speed(Current, Last, All)'))

# End Vissim and shut down the COM interface   
vis = None

