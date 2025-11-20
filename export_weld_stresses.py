from src.pyideastatica.export import export_weld_stress

r"""
This script exports weld stress into Excel.

1. Open your file in Ideastatica. You need valid license.

2. Execute the following exe file, and wait about 15-30 seconds to allow it start:

C:\Program Files\IDEA StatiCa\StatiCa 25.1\IdeaStatiCa.ConnectionRestApi.exe

3. Execute this script.

Some information:
Two types of analysis are supported:
stress-strain - weld stresses are exported into the Excel in one Tab
fatigue - weld stresses are exported to 'Welds' sheet and weld sections are exported to 'Weld Sections' sheet

Info sheet contains information about connection name, time and analysis type. 

Inputs:
project_file_path - path to your Ideastatica file.
connection_name - name of the connection you want to export, thi is shown in left top corner.
output_path - pase any output folder you want Excel file to be saved

"""

# INPUTS:
project_file_path = r"C:\Users\mwo\Downloads\P0160_ICCP_Check.ideaCon"
connection_name = "N6005"  # Connection name
output_path = r"output"  # Path to the output folder

# EXECUTION CODE:
output = export_weld_stress(project_file_path, connection_name, write_json=False, output_path=output_path)
