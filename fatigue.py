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

connection_index = 0  # INPUT - WHICH CONNECTION INDEX IN FILE

with connection_api_service_attacher.ConnectionApiServiceAttacher(BASE_URL).create_api_client() as api_client:

    # Open project
    uploadRes = api_client.project.open_project_from_filepath(project_file_path)

    project_id = api_client.project.active_project_id

    # Get the project data
    project_data = api_client.project.get_project_data(project_id)

    # Get list of all connections in the project
    connections_in_project = api_client.connection.get_connections(project_id)

    # Get connection in the project
    connection1 = connections_in_project[0]  # INPUT - WHICH CONNECTION INDEX IN FILE
    print(connection1.name)
    print(connection1.analysis_type)

    # Get all loads
    loads = api_client.load_effect.get_load_effects(project_id, connection1.id)
    no_loads = len(loads)

    loads_copy = loads

    fatigue_output = []
    weld_output = []

    for n in range(1, no_loads):
        loop_loads = api_client.load_effect.get_load_effects(project_id, connection1.id)
        print("Load cases count at the start of loop: " + str(len(loop_loads)))
        load_ef = loop_loads[n]
        current_name = load_ef.name
        print("Processing case:.... " + str(current_name))
        # Remove other loads than one currently considered
        for load in loop_loads:
            if load.name != current_name and load.name != "Ref_0MPa":
                api_client.load_effect.delete_load_effect(project_id, connection1.id, load.id)
                print("Deleted case: " + str(load.name))

        currrent_loads = api_client.load_effect.get_load_effects(project_id, connection1.id)
        print("Count of cases left in iteration: " + str(len(currrent_loads)))

        calcParams = ideastatica_connection_api.ConCalculationParameter()
        calcParams.connection_ids = [connection1.id]
        calcParams.analysis_type = "fatigues"
        # connection1.analysis_type = calcParams.analysis_type
        # updated_connection1 = api_client.connection.update_connection(
        #     api_client.project.active_project_id, connection1.id, connection1
        # )

        # Model must be calculated first
        con1_fatigue_results = api_client.calculation.calculate(
            api_client.project.active_project_id, calcParams.connection_ids
        )

        # print(con1_fatigue_results)

        # detailed_results = api_client.calculation.get_results(
        #     api_client.project.active_project_id, calcParams.connection_ids
        # )
        # print(detailed_results)

        results_text = api_client.calculation.get_raw_json_results(api_client.project.active_project_id, calcParams)
        firstConnectionResult = results_text[0]
        raw_results = json.loads(firstConnectionResult)

        fatigue_results = raw_results["fatigueChecks"]
        # Writing to the .json
        # print(weld_results)
        with open("data.json", "w") as f:
            json.dump(raw_results, f)

        # Stressess are in N/m2 (Pa) so require *1e-6
        for weld_id, weld in fatigue_results.items():
            lcase = weld.get("loadCase")
            jname = weld.get("joinedItemName")
            thickness = weld.get("designedThickness")
            leg_size = weld.get("legSize")
            name = weld.get("name")
            weld_type = weld.get("weldType2")
            weld_length = weld.get("length")
            sigma_max = weld.get("normalStress") * 1e-6
            tau = weld.get("shearStress") * 1e-6
            sigma = weld.get("normalStress2") * 1e-6
            tau_max = weld.get("shearStress2") * 1e-6
            fatigue_output.append(
                [lcase, jname, name, thickness, leg_size, weld_type, weld_length, sigma_max, tau, tau_max, sigma]
            )

        weld_results = raw_results["fatigueWelds"]

        for weld_id, weld in weld_results.items():
            lcase = weld.get("loadCase")
            jname = weld.get("joinedItemName")
            name = weld.get("name")
            thickness = weld.get("designedThickness")
            leg_size = weld.get("legSize")
            weld_type = weld.get("weldType2")
            weld_length = weld.get("length")
            max_eq_stress = weld.get("maxEquivalentStress") * 1e-6
            tau_y = weld.get("tauy")
            tau_x = weld.get("taux") * 1e-6
            tau_wf_max = weld.get("tauxwf") * 1e-6
            sigma_wf = weld.get("sigmawf") * 1e-6
            weld_output.append(
                [
                    lcase,
                    jname,
                    name,
                    thickness,
                    leg_size,
                    weld_type,
                    weld_length,
                    max_eq_stress,
                    tau_y,
                    tau_x,
                    tau_wf_max,
                    sigma_wf,
                ]
            )

        # print(weld_output)

        # Add all loads back again
        for copied_load in loads_copy:
            if copied_load.name != current_name and copied_load.name != "Ref_0MPa":
                api_client.load_effect.add_load_effect(project_id, connection1.id, con_load_effect=copied_load)

    # Preamble for output data
    preamble_fatigue = [
        "Load Case",
        "Joined Item Name",
        "Name",
        "Designed Thickness",
        "Leg size",
        "Weld type",
        "Weld length",
        "Normal stress",
        "Shear stress",
        "Shear stress 2",
        "Normal stress 2",
    ]
    preamble_weld = [
        "Load Case",
        "Joined Item Name",
        "Name",
        "Designed Thickness",
        "Leg size",
        "Weld type",
        "Weld length",
        "Max \u03c3_eq stress",
        "\u03c4_y",
        "\u03c4_x",
        "\u03c4_wf_max",
        "\u03c3_wf",
    ]
    fatigue_output.insert(0, preamble_fatigue)
    weld_output.insert(0, preamble_weld)

    file_name = connection1.name

    # # # Write to .txt
    # # with open("weld_data.txt", "w") as f:
    # #     for line in output:
    # #         f.write(f"{line}\n")

    # Write to excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Welds"

    ws2 = wb.create_sheet("Weld sections")

    # Append rows to the worksheet
    for row in weld_output:
        ws.append(row)

    for row in fatigue_output:
        ws2.append(row)

    # Save the file
    wb.save(f"{file_name}_fatigue_weld_stress.xlsx")
