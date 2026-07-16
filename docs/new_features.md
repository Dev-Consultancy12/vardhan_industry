# Upcoming Features & Requirements

Based on the latest client feedback, the following new features will be implemented in the Quality Inspection Automation system.

## 1. Group-Specific Dynamic Test Readings
**Current Behavior:** 
The system uses a single `TEST READINGS.xlsx` file and pulls baseline readings from the exact same columns for every item, regardless of its classification group. 

**New Requirement:**
* The client will provide specific test readings **per group**.
* Different groups may have different testing parameters (columns).
* The backend parser must become dynamic: instead of hardcoding column lookups for parameters like "Strands dia" or "Insulation dia", the system will map the item to its specific Group, and then dynamically pull the correct parameters and baseline readings specifically defined for that Group's test profile.
* **Impact:** This ensures that distinct wire types/groups get their highly specific inspection criteria automatically applied rather than relying on a generalized template.

## 2. Frontend Test Reading Editor & Tweaker
**Current Behavior:** 
Test readings are generated automatically in the backend during the PDF/Excel generation process. The user only sees the final result on the PDF and cannot modify the baseline data from the web interface.

**New Requirement:**
* **Display Readings:** The frontend will be updated to display the generated test readings (or the base profiles) before or after the file is processed.
* **Interactive Editor:** Users will have an interactive table on the frontend where they can tweak/edit the test readings manually.
* **Save Functionality:** There will be a clear "Edit" and "Save" workflow. 
* **Single Source of Truth:** The `TEST READINGS` Excel file will be treated as the absolute source of truth for the entire system. 
* **Backend Synchronization:** When the user hits "Save", the frontend will send the updated readings to the backend. The backend will then overwrite/update the main master `TEST READINGS` file (updating the actual Excel file locally). Because it is the source of truth, any edits made on the frontend will permanently alter the master file, guaranteeing that all future generations use the exact same tweaked baseline data.
* **Impact:** This gives the quality control team complete power to fine-tune the parameters directly from the browser without needing to manually open and edit the `TEST READINGS.xlsx` file on the server/local filesystem.

## Next Steps for Implementation
1. **Analyze New Data Structure:** Wait for the client to provide the new multi-group test reading Excel sheet to understand the new column mapping and layout.
2. **Backend Architecture Update:** 
   * Modify the Excel generator logic to dynamically read columns based on the group.
   * Create new API endpoints (e.g., `GET /test-readings` and `POST /test-readings`) to allow the frontend to fetch and update the master test readings file.
3. **Frontend UI Update:** Build the interactive data table in React with Edit/Save states and API integration.
