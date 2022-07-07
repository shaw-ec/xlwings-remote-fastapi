import datetime as dt

import xlwings as xw
import yfinance as yf
from fastapi import Body

from app import app


@app.post("/hello")
def hello(data: dict = Body):
    # Instantiate a Book object with the deserialized request body
    book = xw.Book(json=data)

    # Use xlwings as usual
    sheet = book.sheets[0]
    if sheet["A1"].value == "Hello xlwings!":
        sheet["A1"].value = "Bye xlwings!"
    else:
        sheet["A1"].value = "Hello xlwings!"

    # Pass the following back as the response
    return book.json()


@app.post("/yahoo")
def yahoo_finance(data: dict = Body):
    """
    This is a sample function using the yfinance package to query
    Yahoo! Finance. It writes a pandas DataFrame to Excel/Google Sheets.
    """
    book = xw.Book(json=data)

    if "yahoo" not in [sheet.name for sheet in book.sheets]:
        # Insert and prepare the sheet for first use
        sheet = book.sheets.add("yahoo")
        sheet["A1"].value = [
            "Ticker:",
            "MSFT",
            "Start:",
            dt.date.today() - dt.timedelta(days=30),
            "End:",
            dt.date.today(),
        ]
        for address in ["B1", "D1", "F1"]:
            sheet[address].color = "#D9E1F2"
        for address in ["D1", "F1"]:
            sheet[address].columns.autofit()
        sheet[
            "A3"
        ].value = "'=> Adjust the colored parameters and run the script again!"
        sheet.activate()
    else:
        # Query Yahoo! Finance
        sheet = book.sheets["yahoo"]
        target_cell = sheet["A3"]
        target_cell.expand().clear_contents()
        try:
            df = yf.download(
                sheet["B1"].value,
                start=sheet["D1"].value,
                end=sheet["F1"].value,
                progress=False,
            )
            target_cell.value = df
            target_cell.offset(row_offset=1).columns.autofit()
        except Exception as e:
            target_cell.value = repr(e)

    return book.json()


@app.post("/startMC")
def module_count(data: dict = Body):
    """
    This function downloads the data from Mixpanel using its query API by a service account,
    then generate a DataFrame (spreadsheet-like object) of the module counts.
    """
    import requests
    import json
    import numpy as np
    import pandas as pd
    from datetime import datetime as dt
    
    # Download the relevant data from Mixpanel using its query API
    url = "https://mixpanel.com/api/2.0/insights?project_id=2601976&bookmark_id=31169655"
    headers = {
    "Accept": "application/json",
    "Authorization": "Basic YXBpLWFjY2Vzcy1jb2x1bWJpYS40NGY5ZjQubXAtc2VydmljZS1hY2NvdW50OmxBUUtGOFRlUmxQYmRqMXZjMFVxUXJHNnRWaHNORU9p"
    }
    response = requests.get(url, headers=headers)
    
    # Process the data first into JSON, then into DataFrame
    df = pd.DataFrame(json.loads(response.text)['series']['Click a Module - Unique'])
    module_list = df.columns[1:]
    date_list = [idx[:10] for idx in df.index]

    # Record the release date of each module
    release_date = []
    for mol in module_list:
        for idx in range(len(date_list)):
            if df[mol][idx] > 0:
                release_date.append(date_list[idx])
                break
    
    # Count the accumulate number of modules released each date
    accumulate = 0
    module = []
    for d in date_list:
        if d in release_date:
            accumulate += release_date.count(d)
        module += [accumulate]
    
    # Save the result as a DataFrame, which will be returned by the function
    df_new = pd.DataFrame({"Date": date_list, "Module Title": module})
    df_new.set_index("Date", inplace=True)

    # Setting up the google sheets
    book = xw.Book(json=data)
    sheet = book.sheets[0]
    sheet["A1"].value = ["Last Updated:", dt.now().date(), dt.now().time()]
    sheet["A2"].value = ["Start Date:", date_list[0], "End Date:", date_list[-1]]
    for address in ["B2", "D2"]:
        sheet[address].color = "#D9E1F2"
    sheet["A4"].value = df_new

    return book.json()


@app.post("/updateMC")
def module_count_update(data: dict = Body):
    """
    This function updates the module counts for customized dates.
    """
    import requests
    import json
    import numpy as np
    import pandas as pd
    from datetime import datetime as dt
    
    # Download the relevant data from Mixpanel using its query API
    url = "https://mixpanel.com/api/2.0/insights?project_id=2601976&bookmark_id=31169655"
    headers = {
    "Accept": "application/json",
    "Authorization": "Basic YXBpLWFjY2Vzcy1jb2x1bWJpYS40NGY5ZjQubXAtc2VydmljZS1hY2NvdW50OmxBUUtGOFRlUmxQYmRqMXZjMFVxUXJHNnRWaHNORU9p"
    }
    response = requests.get(url, headers=headers)
    
    # Process the data first into JSON, then into DataFrame
    df = pd.DataFrame(json.loads(response.text)['series']['Click a Module - Unique'])
    module_list = df.columns[1:]
    date_list = [idx[:10] for idx in df.index]

    # Record the release date of each module
    release_date = []
    for mol in module_list:
        for idx in range(len(date_list)):
            if df[mol][idx] > 0:
                release_date.append(date_list[idx])
                break
    
    # Count the accumulate number of modules released each date
    accumulate = 0
    module = []
    for d in date_list:
        if d in release_date:
            accumulate += release_date.count(d)
        module += [accumulate]
    
    # Setting up the google sheets
    book = xw.Book(json=data)
    sheet = book.sheets[0]
    sheet["A1"].value = ["Last Updated:", dt.now().date(), dt.now().time()]

    if (sheet["B2"].value == None) and (sheet["D2"].value == None):
        start = date_list[0]
        end = date_list[-1]
        sheet["A2"].value = ["Start Date:", start, "End Date:", end]
        for address in ["B2", "D2"]:
            sheet[address].color = "#D9E1F2"
    elif (sheet["B2"].value == None) or (sheet["D2"].value == None):
        if sheet["B2"].value == None:
            start = date_list[0]
            end = str(sheet["D2"].value)[:10]
            sheet["B2"].value = start
        else:
            start = str(sheet["B2"].value)[:10]
            end = date_list[-1]
            sheet["D2"].value = end
    else:
        start = str(sheet["B2"].value)[:10]
        end = str(sheet["D2"].value)[:10]
    
    # Save the result as a DataFrame, which will be returned by the function
    df_new = pd.DataFrame({"Date": date_list, "Module Title": module})
    df_new.set_index("Date", inplace=True)
    sheet["A4"].expand().clear_contents()
    try:
        sheet["A4"].value = df_new.loc[start: end]
    except Exception as e:
        sheet["A4"].value = repr(e)
    
    return book.json()


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
