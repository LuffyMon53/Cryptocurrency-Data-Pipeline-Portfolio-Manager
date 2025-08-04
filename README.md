# Cryptocurrency-Data-Pipeline-Portfolio-Manager
A dual-mode solution that fetches real-time coin data from CoinGecko API and transforms it into actionable portfolio insights through Power BI.


## 🌐 Dual Deployment Options

Option 1: Local Environment Setup

Workflow Path:-

Execute Python script (crypto_fetcher.py) locally

Script fetches data from CoinGecko API

Data gets stored/updated in local Excel file (data/crypto_data.xlsx)

Power BI Desktop connects to this local file

Reports are developed and viewed in Power BI Desktop

```mermaid
graph LR
    %% Local Workflow
    A[Run Python Script] -->|Fetch Data| B[(Local Excel)]
    B -->|Connect| C[PowerBI Desktop]
    C -->|Generate| D[[Local Dashboard]]
    
    %% Styling
    style A fill:#2ecc71,stroke:#27ae60
    style B fill:#3498db,stroke:#2980b9
    style C fill:#9b59b6,stroke:#8e44ad
    style D fill:#f1c40f,stroke:#f39c12
    classDef bg fill:#f0f0f0,stroke:#333,stroke-width:0px;
    class graphArea bg;
```


Option 2: Cloud-Based Workflow (Current Implementation)

Workflow Path:-

Run Python script in Google Colab notebook

Processed data gets saved to Google Drive

Power BI Web imports data directly from Drive link

Reports are built and shared via Power BI online service

```mermaid
graph LR
    %% Cloud Workflow
    E[Colab Notebook] -->|Save to| F[(Google Drive)]
    F -->|Import| G[PowerBI Web]
    G -->|Publish| H[[Online Dashboard]]
    
    %% Styling
    style E fill:#2ecc71,stroke:#27ae60
    style F fill:#3498db,stroke:#2980b9
    style G fill:#9b59b6,stroke:#8e44ad
    style H fill:#f1c40f,stroke:#f39c12
    classDef bg fill:#f0f0f0,stroke:#333,stroke-width:0px;
    class graphArea bg;
```

Key Considerations:-

```mermaid
%% Comparison Table Diagram
classDiagram
    class Local {
        Python Script Run: Shadule
        Refresh: Manual
        PowerBI: Full features
        Security: Self-contained
        Access: Single device
    }
    
    class Cloud {
        Python Script Run: Manual
        Refresh: Manual
        PowerBI: Web limitations
        Security: Cloud-dependent
        Access: Anywhere
    }
    
    Local --|> Comparison
    Cloud --|> Comparison
    
    note for Local "Best for advanced analytics\nand sensitive data"
    note for Cloud "Best for team collaboration\nand automation"
```
