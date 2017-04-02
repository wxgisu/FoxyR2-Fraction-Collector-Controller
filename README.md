This is a colaborative project between two graduate students, Xiaoguang Wang and Jagannathan Alagurajan in Dr. Mark Hargrove's lab in the Roy J. Carver Department of Biochemistry, Biophysics and Molecular Biology of Iowa State University. The goal of the project is to add automatic OD monitoring and sampling functions to a New Brunswick BioFlo110 STR system, to reduce labor intensive steps in bioreactor experiments. The automated system is only tested with *E.coli* cells. The key to this project is to reprogram a FoxyR2 fraction collector to collect discontinuous samples based on time. Its default function is to collect continuous protein elution samples in different fractions.


# Equipments
![Equipment Figure](https://github.com/wxgisu/Robotic-Stirred-Tank-Reactor-System/blob/master/Equipment%20Figure.png)

1. New Brunswick BioFlo110 control system
2. New Brunswick 5L STR glass reactor
3. Varian Cary 50 Spectrometer (controled by  Cary WinUV Software)
4. [Precision Cells 10mm Lightpath Flow Through Cell](http://www.precisioncells.com/products/Spectrophotometer-Cuvettes/Flow-Through-Cells/17/76/Precision-Cells-Type-58-Macro-Flow-Through-Cell-with-Top-Tubes-Lightpath-10mm)
5. Teledyen ISCO FoxyR2 Fraction Collector (reprogramed in house, code is forked from Jagannathan's GitHub site and included in this repository)
6. [Uniclife UL80 Submersible Water Pump](https://www.amazon.com/Uniclife-Submersible-Aquarium-Powerhead-Hydroponic/dp/B00ZW6OHHY/ref=sr_1_1?ie=UTF8&qid=1491107247&sr=8-1-spons&keywords=fish+pump&psc=1)
7. Fisher Scientific ISOTEMP 1006S circulating water bath heater

# Workflow
![Workflow Figure](https://github.com/wxgisu/Robotic-Stirred-Tank-Reactor-System/blob/master/Workflow%20Figure.png)

The strategy to add automated functions to the existing STR system is to setup a continuous circulation system out side of the reactor. As shown in the figure above, steps of workflow is as follow:
1. Culture in reactor (2) is first circulated out through sample line on the head plate via a peristatic pump on BioFlo100 controller (1). flow speed of pump is about 3 ml/min.
2. Culture flow into the flowcell (4) sits in spectrometer (3).
3. OD of flowcell (4) is monitored at 600 nm by Cary WinUV Kinetics program (1 min/read).
4. Culture continue flowing to the fraction collector (5) that is in a refridgerator (4 oC)
5. Fraction collector is reprogramed to withdraw a sample for a certain amount of time at a certain time interval.
6. When Fraction collector is not sampling, culture continue flowing back to reactor through a small feed port. This completes the circulation.

A [powerpoint file](https://github.com/wxgisu/Robotic-Stirred-Tank-Reactor-System/blob/master/Robotic%20STR%20System%20Workflow%20Animation.pptx) is included to show the animation of workflow.

# Setup
To start a bioreactor experiment, the below steps are followed:
1. 


# Reprogram FoxyR2 Fraction Collector Controller

This is VBA-excel code to control Foxy R2 fraction collector. Designed to take periodic sample of cell culture from a in-house fermentor

How to Run:
Download Modules 1,2,3,6,7.

Import the modules into Excel VBA and and run Runmefirst Subroutine to intialize the control page.

Setup Lan Ip address and subnet mask for fraction collector and computer. 

Enter IP address into the sheet and have fun collecting. 
