# constraint-optimization-duty-planner

This repository provides a duty scheduling system built on a point-based framework and developed using the concept of constraint optimization.
The planner ensures fair distribution of duties among staff while respecting constraints such as availability and workload balance.

duty_planner.py - python script of the aforementioned. 

Template.xlsx - example template of the input excel required. 

Try the colab version at: https://colab.research.google.com/drive/14h-zG9mEOG_Wq6zmoUgcTRQkYVIFHJQM?usp=sharing

# Alternatively run on your host machine 
## Clone this repository
```bash
git clone https://github.com/keitheorem/constraint-optimization-duty-planner
```
## Install the necessary libraries
```bash
pip install ortools
pip install holidays
```
## Run the script
```bash
python3 duty_planner.py
```
