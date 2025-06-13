# Friends and Family Test Response Processing ETL Pipeline Documentation

[![status: experimental](https://github.com/GIScience/badges/raw/master/status/experimental.svg)](https://github.com/GIScience/badges#experimental)

This ELT pipeline currently takes the output from the monthly Inpatient Friends and Family Test (FFT) raw data submissions (stored in a restricted access 
SharePoint folder to ensure Data governance compliance before processing), and following transformation/processing, Loads it into a Macro-enabled Excel 
template for publication at https://www.england.nhs.uk/fft/friends-and-family-test-data/ ensuring compliance with NHS England’s mandate to publish monthly FFT figures. 
The following outlines the full ETL process, technology stack employed, and action required to ensure successful running.  

The ambition is for the pipleine to be expanded to run across all FFT collections. 

_**Note:** No sensitive data is shared in this repository._


ETL Process Overview
The pipeline is designed to fit future NHS England architecture under a Common Data Platform, with a view to being lifted into the Unified Data Access Layer (UDAL) 
once all raw data files are made available in a private data environment, where Databricks will be used to fully automate production.  Until then, the pipeline:
-	validates specified column lengths and column types (for int/float type)
-	generates period strings for current and previous period for use in referencing/labelling
-	removes unnecessary columns and renames remaining columns to conform with stakeholder requirements and to map to expected output
-	with Trust/Organisation level data extracted to a DataFrame, aggregates all Independent Providers (IS1) and NHS providers (NHS) respectively, 
and aggregates is1/NHS lines together, producing national totals, IS1 totals and NHS totals for the period
-	creates percentage positive/negative fields for the totals and adds counts of IS1 and NHS providers to respective rows in the DataFrame
-	extracts Monthly Rolling Totals.xlsx inpatient sheet to a DataFrame and updates it with generated current monthly totals for national, IS1 and NHS. 
-	adds cumulative totals into the monthly rolling totals DataFrame using previous months cumulative values and current months totals
-	extracts previous months values from Rolling Monthly Total.xlsx to produce a national summary DataFrame of current/previous month figures
-	loads updated monthly rolling totals DataFrame back into Rolling Monthly Total.xlsx
-	creates an ICB level DataFrame from a copy of the Trust/Organisation level DataFrame by dropping required fields and aggregating all values by ICB Code/Name
-	creates and implements first level and second level suppression processes to ICB level DataFrame and sorts the DataFrame by ICB Code (descending order) moving 
all IS1 data to the bottom
-	orders and ranks the Organisation/Trust level DataFrame and using the ICB level DataFrame for suppression reference of upper-level suppressions, implement first, 
second and upper-level suppression to Organisations/Trusts within ICBs, then sort the DataFrame by ICB Code (ascending) and Total Responses (descending) and move all 
IS1 data to the bottom
-	extracts Trust Collection Modes data to a new DataFrame and join these by Site Code to the Organisation/Trust level DataFrame.
-	uses a copy of the Collection Mode DataFrame to generate aggregated sum totals including/excluding IS1s
-	extracts site level raw data to a DataFrame and repeat all stages carried out at Trust/Organisation level for suppression, excluding adding collection mode detail. 
Following site level suppression sort the DataFrame by ICB Code (ascending), Trust Code (ascending) and Total Responses (descending) and move all independent provider 
rows to the bottom
-	extracts ward level raw data to a DataFrame and repeat all stages carried out at site level for suppression. Following ward level suppression sort the DataFrame by 
ICB Code (ascending), Trust Code (ascending) Site Code (ascending) and Total Responses (descending) and move all independent provider rows to the bottom
-	removes special characters from ward names to prevent error generation
-	generates Macro Excel back sheet Dropdown lists to ensure all tab filters in the Macro-enabled Excel file contain correct ICB/Trust/Site/Ward name details 
-	opens the Macro-Enabled Excel template to workbook object and all loads in all DataFrames to the correct workbook sheets using a list of tuples stating which 
DataFrame to paste in which sheet starting in which row and column
-	updates period subheadings on the Summary sheet with correct current/previous period labels and formatting
-	creates a Percentage style within the Workbook and converts all percentage columns to correct format with 0 decimal places
-	updates Period in the ‘Note’ sheet title to the current period  and saves the updated workbook


Data source/Linage
NHS/Independent Providers submit monthly FFT returns using NHS England’s Strategic Data Collection Service (SDCS). It is then made available to FFT analysts for download 
via NHS England’s Secure Electronic File transfer (SEFT) system. Data is downloaded into a restricted access SharePoint folder as an Excel (.xlsx) file. Data is extracted 
directly from this folder for transformation/processing using Python in VS Code. 


Scheduling and Automated Run Process
In its current format the Pipeline can be manually triggered once new data is loaded to the restricted SharePoint folder. All data is loaded by 22nd(?) of every month. 
To avoid manually running the process the pipeline can be triggered using Windows Task Scheduler. Given Windows Task Scheduler cannot be triggered by the raw data 
file being loaded to the SharePoint folder, the run process will need to be set to run on the 22nd(?) of each month. Alternatively automation could be run through Airflow.
Once all raw data files are made available in a private data environment on UDAL or FDP, the process can be automated via Databricks or FDP.
It is possible to  


Technology Stack
The pipeline was built in python using the following tools/libraries:
-	pandas – for data manipulation/transformation
-	glob – to find matching file name patterns
-	 openpyxl – to enable interaction (open/ manipulate/save) macro-enabled Excel files while retaining VBA code
-	pytest - for development and testing of all unit tests run to ensure code performs as expected
-	unittest.mock and monkeypatch - enabling mocked filepath/files for running unit tests checking accessing raw data files and saving files
-	pytest-cov - to allow for selecting which folder/file is tests
-	logging – to generate log reports each time the pipeline runs to check pipeline health, all pipeline processes that complete 
    successfully or where in the pipeline an error causes failure


Troubleshooting
If any errors arise from running the pipeline, access the logfiles folder and review messages – (logfiles\inpatient_fft). 
The folder contains datetime suffixed logfiles with the most recent run datetime at the bottom.
Working through the most recent logfile in each subfolder, it will be easy to track at which stage and what type of error caused the pipeline failure 
highlighting what needs to be checked and fixed. Where an unplanned anomaly creates the fault, TTD tests and function updates should be added to the code
to ensure this doesn’t generate the same issue in the future. A full run of 
all tests (pytest -vv) must be run to ensure new code don’t have a negative impact on existing code/modules. All changes must be completed on a new 
branch and pull request submitted detailing changes to ensure version control and change log history.


### Contributing

Contributions are what make the open source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

_See [CONTRIBUTING.md](./CONTRIBUTING.md) for detailed guidance._


### License

Unless stated otherwise, the codebase is released under [the MIT Licence][mit].
This covers both the codebase and any sample code in the documentation.

_See [LICENSE](./LICENSE) for more information._

The documentation is [© Crown copyright][copyright] and available under the terms
of the [Open Government 3.0][ogl] licence.

[mit]: LICENCE
[copyright]: http://www.nationalarchives.gov.uk/information-management/re-using-public-sector-information/uk-government-licensing-framework/crown-copyright/
[ogl]: http://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/