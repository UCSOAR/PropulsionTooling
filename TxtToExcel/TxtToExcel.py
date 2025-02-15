import csv
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference, Series
import matplotlib.pyplot as plt

#----------------------------------------These variables can have their values changed------------------------------#
# Input file name
input_file_name = 'CS2-10sec.txt'                                                           # EDIT TO CHANGE INPUT FILE NAME

# Output file name
output_file_name = 'CS#4- Temp Vs Time - 10sec Burn.xlsx'                                   # EDIT TO CHANGE OUPUTFILE NAME

# Chart Boundaries
chart_bounds = (0, -1) # EDIT TO CHANGE CHART BOUNDARIES (and calculate the slope)
                      # its of the form (start_row, end_row) like you would see in the excel sheet
                      # LEAVE UNCHANGED (0,0) FOR AUTO TEST DETECTION
                      # Set to (0, -1) to use the whole boundaries

# Chart Title
chart_title = "Temperature vs Time Calibration, Dist 0.75\", Copper Slug #2"                 # EDIT TO CHANGE CHART TITLE
x_axis_title = "Time (s)"                                                                    # EDIT TO CHANGE X AXIS TITLE
y_axis_title = "Temperature (C)"                                                             # EDIT TO CHANGE Y AXIS TITLE

# Slope threshold
slope_threshold = 0.5 # Threshold for slope before program detects the test has started
#--------------------------------------------------------------------------------------------------------------------#



# Create a new Excel workbook
wb = Workbook()
# Select the active worksheet
ws = wb.active

with open(input_file_name, 'r') as file:
    oldfile = file.read()

# Create a list to hold the data
data = []

# keep track of the start and end of the test
test_boundaries = [0, 0]
# Open the text file for reading
with open(input_file_name, 'r') as file:
    # Write column headers
    ws.append(['thermocouple', 'time', 'temperature', 'unit'])

    # remove and make note of start and end tests
    lines = file.readlines()
    for i in range(len(lines)):
        if lines[i] == 'STARTTEST\n':
            test_boundaries[0] = i
        elif lines[i] == 'ENDTEST\n':
            test_boundaries[1] = i - 1  # move it one back because we're removing the previous STARTTEST line
    
    if not (test_boundaries[0] == 0 and test_boundaries[1] == 0):
        lines.pop(test_boundaries[0])
        lines.pop(test_boundaries[1])
    else:
        test_boundaries = [-1, -1]

with open(input_file_name, 'w') as file:
    file.writelines(lines)

with open(input_file_name, 'r') as file:

    # Read data from the text file and write to Excel
    reader = csv.reader(file)
    first_time = None  # Variable to store the first time value
    for idx, row in enumerate(reader):
        # Convert time value to float and divide by 1000
        time = float(row[1]) / 1000
        if idx == 0:
            first_time = time
        # Subtract the first time value from all time values
        row[1] = time - first_time
        row[2] = float(row[2])
        data.append(row)
        ws.append(row)

# If auto detection is enabled
start_idx = 0
end_idx = len(data)+1
if chart_bounds == (0, 0):
    # go through the data until the slope is greater than the threshold
    for i in range(len(data)-1):
        slope = (data[i+1][2] - data[i][2]) / (data[i+1][1] - data[i][1]) 
        # If the slope is greater than the threshold
        if slope > slope_threshold:
            start_idx = i
            break
        else:
            continue

elif chart_bounds == (0, -1):
    start_idx = 1
    end_idx = len(data)+1
else:
    start_idx = min(max(chart_bounds[0]-2, 0), end_idx)
    end_idx = min(max(chart_bounds[1], 0), end_idx)

if test_boundaries == [-1, -1]:
    test_boundaries = [start_idx, end_idx]

# create the chart with some of the data
chart = ScatterChart()
chart.title = chart_title
chart.style = 7#3
chart.height = 18
chart.width = 27

chart.x_axis.title = x_axis_title   
chart.x_axis.number_format = '0'
chart.x_axis.majorTickMark = None
chart.x_axis.tickLblPos = 'low'

chart.y_axis.title = y_axis_title           
chart.y_axis.number_format = '0'
chart.y_axis.majorTickMark = None
chart.y_axis.tickLblPos = 'low'

chart.legend = None

# Set the chart boundaries
chart.x_axis.scaling.min = data[start_idx-2][1]*0.75
chart.y_axis.scaling.min = data[start_idx-2][2]*0.75
chart.x_axis.scaling.max = data[end_idx-2][1]*1.15
chart.y_axis.scaling.max = data[end_idx-2][2]*1.15

# Select the data for the chart
x = Reference(ws, min_col=2, min_row=test_boundaries[0]+2, max_row=test_boundaries[1]+2)
y = Reference(ws, min_col=3, min_row=test_boundaries[0]+2, max_row=test_boundaries[1]+2)
series = Series(values=y, xvalues=x)

chart.series.append(series)

# Add the chart to the sheet
ws.add_chart(chart, "E2")

# Save the workbook to a file
wb.save(output_file_name)

with open(input_file_name, 'w') as file:
    file.write(oldfile)

# Plot the data using matplotlib
plt.plot([data[i][1] for i in range(0, len(data))], [data[i][2] for i in range(0, len(data))])
# plt.plot([data[test_boundaries[0][1]], data[test_boundaries[0][2]]], [data[0][2], data[-1][2]], color='red')
# plt.plot([data[test_boundaries[1][1]], data[test_boundaries[1][2]]], [data[0][2], data[-1][2]], color='red')
plt.xlabel(x_axis_title)
plt.ylabel(y_axis_title)
plt.title(chart_title)
plt.show()