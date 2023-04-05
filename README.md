# stock-analysis_Ch2
DU-VIRT-DATA-PT-10-2022-U-B-MW—Module 2 Challenge—Wall Street Stock Analysis—VBA Excel 16.67

# Refactoring VBA Code for Stock Market Analysis
#  Overview of the Project
The purpose of this project was to refactor existing VBA code to analyze the entire stock market over the last few years, which could include thousands of stocks. The original code worked well for a small number of stocks, but it might not work as efficiently for a larger number. Refactoring the code would make it more efficient, take fewer steps, and use less memory. This project aims to analyze the before and after runtime of the VBA script to determine whether the refactored code successfully made it run faster.

## Results
The original VBA code looped through each worksheet for every stock ticker to collect the required information. However, the refactored code only looped through all the data once, using arrays to store the data, and then outputted the results. The code used a tickerIndex variable to access the correct stock ticker and an array to store the required information. 
The code ran significantly faster due to looping only once: #

                       2017 Data           2018 Data
Original Time (sec)    0.3554688           0.3007812
Refactored Time (sec)  0.1171875           0.0937500

The original code took amore than three times longer to run.

While on a small dataset this is negligible lag.
It would be unacceptable on any big dataset.
Additionally, the code would need to be more robust than this code to work on data that might not arrive sorted alphabetically.
Even the refactored code might have errored out with unalphabetized data.


