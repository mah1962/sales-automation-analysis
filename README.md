```mermaid
graph TD
    %% Styling
    classDef source fill:#f9f,stroke:#333,stroke-width:2px;
    classDef process fill:#bbf,stroke:#333,stroke-width:2px;
    classDef output fill:#fbf,stroke:#333,stroke-width:2px;
    classDef bi fill:#bfb,stroke:#333,stroke-width:2px;

    %% Nodes
    subgraph Ingestion ["Stage 1: Ingestion"]
        A[Raw CSV Files<br>Sales & Inventory Data]:::source
    end

    subgraph Processing ["Stage 2: Automation & Processing"]
        B[Automation.py<br>Python / Pandas]:::process
    end

    subgraph Storage ["Stage 3: Outputs & Storage"]
        C[Cleaned CSV Data]:::output
        D[Generated Charts & Reports]:::output
    end

    subgraph Analytics ["Stage 4: Business Intelligence"]
        E[Power BI Dashboard<br>Dashboard.pbix]:::bi
    end

    %% Analytics Highlights
    subgraph KPIs ["Dashboard Key KPIs"]
        F[Total Revenue / Profit]
        G[Total Orders / AOV]
        H[Monthly Trends & Slicers]
    end

    %% Connections
    A -->|Place inside /data folder| B
    B -->|Process & Clean| C
    B -->|Generate| D
    C -->|Load into| E
    E --> F
    E --> G
    E --> H
