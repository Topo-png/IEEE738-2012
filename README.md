# IEEE738-2012
Python implementation of the IEEE 738-2012 standard

The goal of this project was to create a Python class that provides better performance and capability than the original OHT Conductor Ratings Excel document released in 2010.

Make sure you validate your results before using them in the real world. This has not been put through thorough testing and validation.   


**This library is provided AS-IS, Use at your own risk**

****

**How to run**

1. Open your favorite Python editor (PyCharm/Spyder/Idle) and load demo.py
2. Run demo.py
   1. If demo = True, then a predefined conductor/configuration will automatically run and output results to the screen and also to a file labeled" export_test.xlsx"
   2. If demo = False, a command line input is provided to allow the user to select the configuration and conductor


******
**Conductor_Prop-Sample.xlsx**
Holds conductor properties used in calculations. Column names must remain the same, but the order is adjustable. 
Sample data provided, to add new conductors add in proper information with the proper units.

**config-sample.xlsx**
Holds configuration setup. Column names must remain the same, but the order is adjustable.

****

***Allowable Units***
1. Reference UnitConversion.py to see the full list of allowable units
2. Units matter
