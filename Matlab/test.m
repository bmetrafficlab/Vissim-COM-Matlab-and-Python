%% Vissim-COM programming - Matlab example code %%
clear all;
close all;
format compact;
clc; % Clears the command window
%% Create Vissim-COM server
vis=actxserver('VISSIM.vissim');
%% Loading the traffic network
access_path=pwd;
vis.LoadNet([access_path '\test.inpx']);
vis.LoadLayout([access_path '\test.layx']);
%% Simulation settings
sim=vis.Simulation;
period_time=3600;
sim.set('AttValue', 'SimPeriod', period_time);
step_time=3;
sim.set('AttValue', 'SimRes', step_time);
%% Define the network object
vnet=vis.Net;
%% Setting the traffic demands of the network
vehins=vnet.VehicleInputs;
vehin_1=vehins.ItemByKey(1);
vehin_1.set('AttValue', 'Volume(1)', 1500); % main road
vehin_2=vehins.ItemByKey(2);
vehin_2.set('AttValue', 'Volume(1)', 100); % side street
%% The objects of the traffic signal control
scs=vnet.SignalControllers;
sc=scs.ItemByKey(1);
sgs=sc.SGs; % SGs=SignalGroups
sg_1=sgs.ItemByKey(1);
sg_2=sgs.ItemByKey(2);
dets=sc.Detectors;
det_all=dets.GetAll;
det_1=det_all{1};
%% to measure total travel time in the network 
netperform=vnet.VehicleNetworkPerformanceMeasurement;
%% Access to DataCollectionPoint object
datapoints=vnet.DataCollectionMeasurements;
datapoint1=datapoints.ItemByKey(1);
%% Access to Link object
links=vnet.Links;
link_1=links.ItemByKey(1);
%% Running the simulation
verify=20; % verifying at every 20 seconds
%Evaluation\Configuration...\Interval in the Vissim GUI
for i=0:(period_time*step_time)
sim.RunSingleStep;
if rem(i/step_time, verify)==0 % verifying at every 20 seconds
demand=det_1.get('AttValue', 'Presence'); %get detector occupancy:0/1
if demand==1 % demand -> demand-actuated stage
sg_1.set('AttValue', 'SigState', 1); % main road red (1)
sg_2.set('AttValue', 'SigState', 3); % side street green (3)
else % no demand on loop -> main road's signal is green
sg_1.set('AttValue', 'SigState', 3);
sg_2.set('AttValue', 'SigState', 1);
%total travel time at the end of each eval. interval:
disp('Travel time of the network: ')
% TravTmTot's subattributes in parentheses:
% (SimulationRun, TimeInterval, VehicleClass)
netperform.get('AttValue','TravTmTot(Current,Last,All)')  
end
% Query the avg. speed and vehicle number at the end of each eval. interval:
datapoint1.get('AttValue', 'Vehs(Current, Last, All)')
datapoint1.get('AttValue', 'Speed(Current, Last, All)')
end
end
%% Delete Vissim-COM server (also closes the Vissim GUI)
vis.release;
disp('The end')