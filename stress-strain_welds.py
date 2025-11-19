from openpyxl import Workbook
import json
import ideastatica_connection_api
import ideastatica_connection_api.connection_api_service_attacher as connection_api_service_attacher
from ideastatica_connection_api.models.con_load_effect import ConLoadEffect
from ideastatica_connection_api.models.con_calculation_parameter import ConCalculationParameter

"""
Execute the following file before running this script:
C:\Program Files\IDEA StatiCa\StatiCa 25.1\IdeaStatiCa.ConnectionRestApi.exe"""

BASE_URL = "http://localhost:5000"  # Default, do not touch

project_file_path = r"C:\Users\mwo\Downloads\P0160_ICCP_Check.ideaCon"


with connection_api_service_attacher.ConnectionApiServiceAttacher(BASE_URL).create_api_client() as api_client:

    # Open project
    uploadRes = api_client.project.open_project_from_filepath(project_file_path)

    project_id = api_client.project.active_project_id

    # Get the project data
    project_data = api_client.project.get_project_data(project_id)

    # Get list of all connections in the project
    connections_in_project = api_client.connection.get_connections(project_id)

    # Get connection in the project
    connection1 = connections_in_project[1]  # INPUT - WHICH CONNECTION INDEX IN FILE

    # Get all loads
    loads = api_client.load_effect.get_load_effects(project_id, connection1.id)
    no_loads = len(loads)

    loads_copy = loads

    output = []

    i = 0
    for n in range(no_loads):
        loop_loads = api_client.load_effect.get_load_effects(project_id, connection1.id)
        print("load count at the start of loop: " + str(len(loop_loads)))
        load_ef = loop_loads[n]
        current_name = load_ef.name
        print("Processing case:.... " + str(current_name))
        # Remove other loads than one currently considered
        for load in loop_loads:
            if load.name != current_name:
                api_client.load_effect.delete_load_effect(project_id, connection1.id, load.id)
                print("Deleted case: " + str(load.name))

        currrent_loads = api_client.load_effect.get_load_effects(project_id, connection1.id)
        # print(len(currrent_loads))

        calcParams = ideastatica_connection_api.ConCalculationParameter()
        calcParams.connection_ids = [connection1.id]
        # calcParams.analysis_type = "fatigues"

        # Model must be calculated first
        con1_cbfem_results = api_client.calculation.calculate(
            api_client.project.active_project_id, calcParams.connection_ids
        )

        # print(con1_cbfem_results)

        # detailed_results = api_client.calculation.get_results(
        #     api_client.project.active_project_id, calcParams.connection_ids
        # )
        # print(detailed_results)

        results_text = api_client.calculation.get_raw_json_results(api_client.project.active_project_id, calcParams)
        firstConnectionResult = results_text[0]

        raw_results = json.loads(firstConnectionResult)
        # weld_results = json.dumps(raw_results["welds"])
        weld_results = raw_results["welds"]

        # Writing to the .json
        # print(weld_results)
        # with open("data.json", "w") as f:
        #     json.dump(weld_results, f)

        # Stressess are in N/m2 (Pa) so require *1e-6
        for weld_id, weld in weld_results.items():
            lcase = weld.get("loadCase")
            jname = weld.get("joinedItemName")
            name = weld.get("name")
            thickness = weld.get("designedThickness")
            weld_type = weld.get("weldType2")
            weld_length = weld.get("length")
            sigma_per = weld.get("sigmaPerpendicular") * 1e-6
            tau_y = weld.get("tauy") * 1e-6
            tau_x = weld.get("taux") * 1e-6
            output.append([lcase, jname, name, thickness, weld_type, weld_length, sigma_per, tau_y, tau_x])

        # print(output)

        # Add all loads back again
        for copied_load in loads_copy:
            if copied_load.name != current_name:
                api_client.load_effect.add_load_effect(project_id, connection1.id, con_load_effect=copied_load)

    # Preamble for output data
    preamble = [
        "Load Case",
        "joinedItemName",
        "Name",
        "Thickness",
        "Weld type",
        "Weld length",
        "\u03c3_\u27c2",  # sigma_perpendicular
        "\u03c4_\u27c2",  # tau_perpendicular
        "\u03c4_\u2225",  # tau_parallel
    ]
    output.insert(0, preamble)

    file_name = connection1.name

    # # Write to .txt
    # with open("weld_data.txt", "w") as f:
    #     for line in output:
    #         f.write(f"{line}\n")

    # Write to excel
    wb = Workbook()
    ws = wb.active
    # Append rows to the worksheet
    for row in output:
        ws.append(row)
    # Save the file
    wb.save(f"{file_name}_weld_stress.xlsx")
