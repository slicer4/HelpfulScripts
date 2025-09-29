import xlwings as xw
import pandas as pd

MODEL_FILE = "Model.xlsx"
IO_FILE = "Input-Output.xlsx"
INPUT_SHEET = "Processor"
INPUT_TABLE = "Inputs"
OUTPUT_TABLE = "Outputs"

def clear_inputs(ws):
    cells_to_clear = [
        "D4","D5","E11","V10","V11","V12","V13","X17","X21",
        "E12","W5","V15","V19","K5","K6",
        "AB10","AB11","AB12","AB13","AB14","AB15","AB16",
        "X30","W29","X31","X32","X33",
        "Z30","Y29","Z31","Z32","Z33",
        "AB30","AA29","AB31","AB32","AB33"
    ]
    for c in cells_to_clear:
        ws[c].value = None

def apply_inputs(ws, row):
    ws["D4"].value = row["Pathfinder ID"]
    ws["D5"].value = row["Initiative Name"]

    scope = str(row["Scope"])
    ws["E11"].value = scope
    ws["V10"].value = "x" if scope == "Entire Enterprise" else None
    ws["V11"].value = "x" if scope == "Pharmacy" else None
    ws["V12"].value = "x" if scope == "International" else None
    ws["V13"].value = "x" if scope == "No International" else None
    if "records" in scope:
        ws["X17"].value = int(scope.split()[0])
    if "assets" in scope:
        ws["X21"].value = int(scope.split()[0])

    ws["E12"].value = row["New/Enhancing"]
    ws["W5"].value = row["Directness"]

    cia = str(row["CIA"])
    ws["V15"].value = "x" if "C" in cia else None
    ws["V19"].value = "x" if "A" in cia else None

    ws["K5"].value = row["Initial Cost"]
    ws["K6"].value = row["Subsequent Cost"]

    for col, cell in [
        ("Ransomware","AB10"),("DDoS","AB11"),("Digital Fraud","AB12"),
        ("Hacking","AB13"),("Privilege Abuse","AB14"),
        ("Phishing","AB15"),("Errors","AB16")
    ]:
        ws[cell].value = "Yes" if bool(row[col]) else "No"

    ws["X30"].value = row["Control 1 Function"]
    ws["W29"].value = row["Control 1 Automation"]
    ws["X31"].value = row["Control 1 MITRE"]
    ws["X32"].value = row["Control 1 Location"]
    ws["X33"].value = row["Control 1 Difficulty"]

    ws["Z30"].value = row["Control 2 Function"]
    ws["Y29"].value = row["Control 2 Automation"]
    ws["Z31"].value = row["Control 2 MITRE"]
    ws["Z32"].value = row["Control 2 Location"]
    ws["Z33"].value = row["Control 2 Difficulty"]

    ws["AB30"].value = row["Control 3 Function"]
    ws["AA29"].value = row["Control 3 Automation"]
    ws["AB31"].value = row["Control 3 MITRE"]
    ws["AB32"].value = row["Control 3 Location"]
    ws["AB33"].value = row["Control 3 Difficulty"]

def extract_outputs(metrics_ws):
    return {
        "Likelihood Reduction": metrics_ws["B14"].value,
        "Impact Reduction": metrics_ws["B15"].value,
        "Risk Reduction": metrics_ws["B16"].value,
        "ROI": metrics_ws["B1"].value
    }

def main():
    # Open IO workbook with xlwings to grab named table
    io_wb = xw.Book(IO_FILE)
    inputs_tbl = io_wb.sheets[INPUT_SHEET].api.ListObjects(INPUT_TABLE)
    inputs_rng = inputs_tbl.Range
    df_inputs = pd.DataFrame(inputs_rng.value[1:], columns=inputs_rng.value[0])

    # Open Model workbook
    app = xw.App(visible=False)
    model_wb = xw.Book(MODEL_FILE)
    model_ws = model_wb.sheets["Processor"]
    metrics_ws = model_wb.sheets["Metrics"]

    outputs = []
    for _, row in df_inputs.iterrows():
        clear_inputs(model_ws)
        apply_inputs(model_ws, row)
        model_wb.app.calculate_full()
        result = extract_outputs(metrics_ws)
        result["Pathfinder ID"] = row["Pathfinder ID"]
        outputs.append(result)

    model_wb.close(save=False)
    app.quit()

    # Write back to Outputs table
    outputs_df = pd.DataFrame(outputs)
    outputs_ws = io_wb.sheets[INPUT_SHEET]
    outputs_tbl = outputs_ws.api.ListObjects(OUTPUT_TABLE)
    # Clear old rows (except headers)
    outputs_tbl.DataBodyRange.ClearContents()
    # Write new outputs
    outputs_ws.range(outputs_tbl.DataBodyRange.Cells(1,1)).value = outputs_df.values.tolist()
    io_wb.save()
    io_wb.close()

if __name__ == "__main__":
    main()
