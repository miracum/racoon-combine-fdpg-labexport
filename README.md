# combine_labexporter.py

## Overview

`combine_labexporter.py` is a Python script designed to export laboratory data for the RACOON subproject COMBINE. It uses two excel spreadsheets containing patient cohort information and the LOINC labvalues as input and creates JSON-bundles of FHIR data per sheet and patient, as well as csv-files per sheet.

As of now there is no way to import the data via FHIR, so the bundles are merely a feasability feature. To actually import the data two CSVs are needed one containing PID, Date, LOINC, Value (which is created with this script) and one containing the Lower and Upper bounds per LOINC-Code, which has to be supplied additionally to make the import possible.

## Setup

Before using the script, ensure the following adjustments are made:

- **Input File Paths:**
    - Update the `input_file_path` variable to point to the excel-file containing the COMBINE cohort with time periods.
    - Update the `lab_file_path` variable to specify the path to the excel-file of LOINC-codes.

- **FHIR Server URL and Credentials:**
    - Update the `fhir` variable to your DIZ-FHIR-server URL.
    - Create an `.env` file and store the credentials in the variables `FHIR_USERNAME` and `FHIR_PASSWORD`
    - `pat_system`: Set your naming system for patient-ID. Only needed if `REPLACE_PAT_ID` is set to true

- **Additional adjustments:**
    - `MAX_COUNT`: Set the number of resources per GET request to the fhir server
    - `SINGLE_PERIODS`: For debug purposes. If true each time period (preCovid, initial, peak, followups) is exported in a separate json
    - `REPLACE_PAT_ID`: Replaces pseudonym with pat-id in JSON output

## Preparing Patient Spreadsheets

Before running the script, ensure that the patient spreadsheet is properly filled out. Each row should represent a patient's information, with patient IDs filled into column 'A'. There shouldn't be any empty rows between patient entries, as this disrupts data processing.

## Functionality

- **`get_patients()`:**
    - This function crawls through the patient spreadsheet, extracting patient IDs and writing them into a `.csv` file. This CSV file can be utilized to obtain DIZ-pseudonyms from your THS, as the table contains clear patient data. The resulting lookup table path should be set in the `input_patient_map` variable. By default, the CSV is separated by semicolons with the format `PatID;Pseudonym` and includes a header.

- **Data Retrieval and Storage:**
    - The script retrieves laboratory observations containing the specified lab codes within the given time periods of the cohort spreadsheet for each patient. The results are stored in JSON bundles per patient and written out as csv-files per cohort-sheet.

## Usage

To use the script:

1. Make sure the necessary adjustments mentioned above are made.
2. First run get_patients() to obtain the PatIDs
3. Obtain your DIZ pseudonyms from the THS and fill in the path `input_patient_map`
4. Run the rest of the code
5. Output should be a json-bundle of Observations per sheet and patient, as well as a csv per sheet containing PID, Date, LOINC and Value

## Dependencies

The script requires the following dependencies:

- `pandas`: For data manipulation and extraction from Excel spreadsheets.
- `requests`: For making HTTP requests to the FHIR server.
- `python-dotenv`: For loading environment variables from the `.env` file.
