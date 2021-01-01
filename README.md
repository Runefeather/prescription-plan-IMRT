# Optimization: Final Project, Fall 2020

## Introduction
This is the repository for the Final project for Optimization Methods (Fall 2020, Cornell Tech). The objective of this project was to create various (3) prescription plans for radiation therapy for a particular patient.

A 20x20 grid of voxels was used to depict the target area for a patient. This grid representation contains the CTV (clinical target volume) and other organs that may be surrounding it. There are 6 beamlets (1 x 10 grids) surrounding this main grid whose intensities determine how much radiation the voxels in the main grid get. 

Also given are the constraints for the amount of radiation that the different parts of the grid can take. Furthermore, an excel sheet, *DosageMatrix.xlsx* describing the radiation amount each voxel gets from each cell in the beamlet is provided.

Using this information, the code outputs prescription plans optimized based on different tradeoffs, which state the units of radiation that every cell in the 6 beamlets outputs. They are also represented by a heatmap that shows radiation on the main grid. An example, Plan 1, is shown below:

![Plan 1](https://github.com/Runefeather/prescription-plan-IMRT/blob/master/images/Plan_1.png)

## Methodology
The first step was to label all the voxels in the main grid into their respective regions. This was done in a separate excel file called *Variables.xlsx*. 

Next, to achieve the goal, an optimization problem was formulated, and different plans were generated using two methods:
* The first was by making decisions about which constraints should be maintained and which ones could be relaxed. This allowed the linear program to find different solutions, some which returned better objective values than with tight constraints. 
* The second method was by changing how the voxels were labeled in the given diagram. By labelling a voxel as a part of one region instead of the other, it was possible to change the limit on the amount of radiation that can pass through the voxel. This was especially effective while labeling voxels in the borders of the regions, and from that, it was possible to obtain a safer set of beamlet intensities. 

## Requirements
There are a few libraries required to run this code. The commands for their installation using the package manager [pip](https://pip.pypa.io/en/stable/) is as follows:

```bash
pip install xlrd
pip install matplotlib
pip install numpy
pip install ortools
pip install seaborn
```

## Running the Code
Ensure that the two excel files - *Variables.xlsx* and *DosageMatrix.xlsx* are in the same directory. Then, to run the code, execute the following: 

```bash
python CreatePlan.py
```

