# Pipeline Analyzer

This script fetches and analyzes pipeline data from Harness.io, processes the stages, and generates various summaries and reports. It supports exporting the data to CSV files and updating an Excel spreadsheet.

## Table of Contents

- [Requirements](#requirements)
- [Setup](#setup)
- [Usage](#usage)
- [Functions](#functions)
- [Logging](#logging)
- [Exporting Data](#exporting-data)
- [Updating Spreadsheet](#updating-spreadsheet)

## Requirements

- Python 3.x
- The following Python libraries:
  - requests
  - json
  - yaml
  - collections
  - csv
  - time
  - openpyxl
  - pandas
  - logging
  - dotenv
  - tenacity

## Setup

1. **Install the required packages:**

   ```sh
   pip install requests pyyaml openpyxl pandas python-dotenv tenacity
   ```

2. **Set up environment variables:**

   Create a `.env` file in the project root and add your Harness API key and account ID:

   ```env
   API_KEY=your_harness_api_key
   HARNESS_ACCOUNT_ID=your_harness_account_id
   ```

## Usage

Run the script using:

```sh
python pipeline_analyzer.py
```

This will fetch the organizations, projects, and pipelines, process the data, and export the results to CSV files and an Excel spreadsheet.

## Functions

- **get_orgs()**: Fetches all organizations.
- **get_projects(org_identifier)**: Fetches all projects for a given organization.
- **get_pipelines(org_identifier, project_identifier)**: Fetches all pipelines for a given project.
- **get_pipeline_yaml(org_identifier, project_identifier, pipeline_identifier, store_type, connector_ref=None, repo_name=None)**: Fetches the YAML definition of a pipeline.
- **get_template_yaml(template_ref, version_label='0.0.1', current_level='account', org_identifier=None, project_identifier=None, parent_pipeline_id=None)**: Fetches the YAML definition of a template.
- **process_stages(stages, processed_templates, template_count, current_level='project', org_identifier=None, project_identifier=None, parent_pipeline_id=None)**: Processes the stages of a pipeline or template.
- **analyze_pipelines(pipelines, org_identifier, project_identifier, processed_templates, template_count)**: Analyzes the pipelines to generate various summaries.
- **calculate_build_times(executions)**: Calculates the average and maximum build times for pipeline executions.
- **fetch_pipeline_executions(org_identifier, project_identifier, pipeline_identifier)**: Fetches the execution summaries for a pipeline.
- **export_to_csv(org_summary, account_summary)**: Exports the organization and account summaries to CSV files.
- **export_pipeline_details_to_csv(pipeline_details)**: Exports the detailed pipeline information to a CSV file.
- **export_pipeline_errors_to_csv(pipeline_errors)**: Exports the pipeline errors to a CSV file.
- **export_template_details_to_csv(template_count)**: Exports the template usage details to a CSV file.
- **update_spreadsheet(org_summary, account_summary, pipeline_details, template_count_dict)**: Updates an Excel spreadsheet with the analysis results.

## Logging

Logs are written to `pipeline_analysis.log` and include information about the script's progress and any errors encountered during execution.

## Exporting Data

The script generates the following CSV files:

- `account_summary.csv`: Summary of the account-level analysis.
- `org_summary.csv`: Summary of the organization-level analysis.
- `pipeline_details.csv`: Detailed information about each pipeline.
- `pipeline_errors.csv`: Errors encountered while processing pipelines.
- `template_details.csv`: Information about template usage.

## Updating Spreadsheet

The script updates an Excel spreadsheet located at `/Users/diegopereira/Documents/Development/git/serenity/CI-AdoptionPlan-Hosted_Builds_Migration.xlsx`. If the file does not exist, it creates a new one. The updated spreadsheet is saved as `CI-AdoptionPlan-Hosted_Builds_Migration_Updated.xlsx` in the same directory.

## License

This project is licensed under the MIT License.
