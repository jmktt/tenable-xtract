
# Tenable xtract

Tenable Xtract is a script that allows users to securely and efficiently export asset and agent data from the Tenable.IO platform. The following is a summary of how the program works:

**Client Selection**: The user selects one of the pre-configured. Each client uses its own set of API credentials defined via environment variables.

**Authentication**: The script establishes a secure session with Tenable.IO using the client’s Access Key and Secret Key. If authentication fails, execution is stopped.

**Asset Export**: When asset export is selected, the script collects all available assets from the platform. It then parses and normalizes the data into a structured format. Each field is processed to ensure clean, readable output. The final dataset is exported into an Excel spreadsheet with formatted columns and table styling.

**Agent Export**: When agent export is selected, the script collects all linked agents and provides filtering options:

- **Offline agents only**
- **Agents without any group assigned**
- **All agents**
- **Offline and No Group (Compare)**: This option generates multiple Excel sheets comparing agents that are offline, without group, in both, or in either condition.

> Excel Formatting: All exports are saved as Excel spreadsheets with
> dynamic column sizing and stylized tables for clarity and
> presentation. If Excel export fails, the script automatically falls
> back to a CSV version.
