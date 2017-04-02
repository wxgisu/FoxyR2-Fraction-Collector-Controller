This is a colaborative project between two graduate students, Xiaoguang Wang and Jagannathan Alagurajan in Dr. Mark Hargrove's lab in the Roy J. Carver Department of Biochemistry, Biophysics and Molecular Biology of Iowa State University. The goal of the project is to add automatic OD monitoring and sampling function to a New Brunswick BioFlo110 STR system, to reduce labor intensive steps in bioreactor experiments. 


# Equipments

![Equipment Figure](Robotic-Stirred-Tank-Reactor-System/Equipment Figure.png)

1. New Brunswick BioFlo110 control system
2. New Brunswick 5L STR glass reactor
3. Varian Cary 50 Spectrometer (controled by  Cary WinUV Software)
4. [Precision Cells 10mm Lightpath Flow Through Cell](http://www.precisioncells.com/products/Spectrophotometer-Cuvettes/Flow-Through-Cells/17/76/Precision-Cells-Type-58-Macro-Flow-Through-Cell-with-Top-Tubes-Lightpath-10mm)
5. Teledyen ISCO FoxyR2 Fraction Collector (reprogramed in house, code is forked from Jagannathan's GitHub site and included in this repository)
6. [Uniclife UL80 Submersible Water Pump](https://www.amazon.com/Uniclife-Submersible-Aquarium-Powerhead-Hydroponic/dp/B00ZW6OHHY/ref=sr_1_1?ie=UTF8&qid=1491107247&sr=8-1-spons&keywords=fish+pump&psc=1)
7. Fisher Scientific ISOTEMP 1006S circulating water bath heater

# Work Flow




# Reprogram FoxyR2 Fraction Collector Controller

This is VBA-excel code to control Foxy R2 fraction collector. Designed to take periodic sample of cell culture from a in-house fermentor

How to Run:
Download Modules 1,2,3,6,7.

Import the modules into Excel VBA and and run Runmefirst Subroutine to intialize the control page.

Setup Lan Ip address and subnet mask for fraction collector and computer. 

Enter IP address into the sheet and have fun collecting. 
