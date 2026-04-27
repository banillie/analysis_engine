# Analysis Engine Project Report

### **Project Overview**
The `analysis_engine` is a specialized Python-based CLI tool designed for the UK Department for Transport (DfT). It automates the ingestion, analysis, and reporting of major project portfolio data, specifically catering to the **IPDC** (Investment Portfolio and Delivery Committee) and **CDG** (Corporate Delivery Group).

### **Core Capabilities**

*   **Data Lifecycle Management:**
    *   **Ingestion:** Extracts data from Excel "master" files using the `datamaps` library.
    *   **Validation:** The `initiate` command performs data integrity checks and converts raw Excel data into JSON for high-performance access in subsequent operations.
    *   **Multi-Quarter Support:** Built-in capability to handle and compare data across different reporting periods for trend analysis.

*   **Automated Visualizations:**
    *   **Dandelion Charts:** Unique portfolio infographics that visualize project scale (cost/resource) and delivery confidence.
    *   **Cost Profiling:** Generates time-series trend graphs and stack plots for portfolio expenditures.
    *   **Milestone Analysis:** Produces schedule charts with progress markers and "blue line" status indicators.
    *   **Speed Dials:** Visual RAG (Red-Amber-Green) status indicators for various confidence metrics.

*   **Reporting & Document Generation:**
    *   **Project Summaries:** Programmatically generates Word document reports for individual projects.
    *   **Dashboards:** Automatically populates complex Excel dashboard templates (e.g., `dashboards_master.xlsx`).
    *   **Risk Analysis:** Extracts risk registers and formats them into Excel or Word, supporting views by project or by risk type.

*   **Data Querying:**
    *   A flexible `query` interface allows users to extract specific "Keys of Interest" (KOIs) across the entire portfolio.

### **Technical Architecture**

*   **Entry Point:** `analysis_engine/main.py` utilizes a dispatch pattern with `argparse` to route CLI commands.
*   **Data Engine:** `analysis_engine/core_data.py` manages data cleaning, normalization, and the `PythonMasterData` abstraction.
*   **Modular Design:** Analysis logic is strictly decoupled into domain-specific modules:
    *   `costs.py`, `milestones.py`, `risks.py`: Data processing for specific domains.
    *   `dandelion.py`, `speed_dials.py`: Visualization logic using `matplotlib`.
    *   `summaries.py`, `render_utils.py`: Document rendering using `python-docx`.
*   **Configuration:** Relies on `config.ini` files and a standardized directory structure (`input/`, `output/`, `core_data/`) to manage environment-specific settings.

### **Tech Stack**
*   **Language:** Python 3
*   **Data Handling:** `datamaps`, `openpyxl`, `xlrd`
*   **Visuals:** `matplotlib`
*   **Office Integration:** `python-docx`
*   **External Dependencies:** Requires **Poppler** for PDF/Image rendering tasks.
