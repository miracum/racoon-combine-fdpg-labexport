# combine_labimporter.py

## Overview

`combine_labimporter.py` is a Python script designed to import laboratory data based on an Excel spreadsheet containing patient information and time periods. The script extracts relevant data from patient spreadsheets, combines it with a list of lab codes, and then queries a FHIR server to retrieve corresponding lab observations within specified time periods for each patient. The results are stored in JSON bundles, organized by patient.

## Setup

Before using the script, ensure the following adjustments are made:

- **Input File Paths:**
    - Update the `input_file_path` variable to point to the excel-file containing the COMBINE cohort with time periods.

- **Lab Code File Path:**
    - Update the `lab_file_path` variable to specify the path to excel-file of LOINC-codes.

- **FHIR Server URL and Credentials:**
    - Update the `fhir` variable to your DIZ-FHIR-server URL.
    - Create an `.env` file and store the credentials in the variables `FHIR_USERNAME` and `FHIR_PASSWORD`

- **Adjust Number of Requested Resources:**
    - Optionally, adjust the number of requested resources per request if needed.

## Preparing Patient Spreadsheets

Before running the script, ensure that the patient spreadsheet is properly filled out. Each row should represent a patient's information, with patient IDs filled into column 'A'. There shouldn't be any empty rows between patient entries, as this might disrupt data processing.

## Functionality

- **`get_patients()`:**
    - This function crawls through the patient spreadsheet, extracting patient IDs and writing them into a `.csv` file. This CSV file can be utilized to obtain DIZ-pseudonyms from your THS, as the table contains clear patient data. The resulting lookup table path should be set in the `input_patient_map` variable. By default, the CSV is separated by semicolons with the format `PatID;Pseudonym` and includes a header.

- **Data Retrieval and Storage:**
    - The script retrieves laboratory observations containing the specified lab codes within the given time periods of the cohort spreadsheet for each patient. The results are stored in JSON bundles per patient.

## Usage

To use the script:

1. Make sure the necessary adjustments mentioned above are made.
2. Run the script using Python.
3. Monitor the output for any errors or warnings.
4. Retrieve the generated JSON bundles containing the lab data for each patient.

## Dependencies

The script requires the following dependencies:

- `pandas`: For data manipulation and extraction from Excel spreadsheets.
- `requests`: For making HTTP requests to the FHIR server.
- `python-dotenv`: For loading environment variables from the `.env` file.

Ensure these dependencies are installed in your Python environment before running the script.