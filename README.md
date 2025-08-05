# Cryptocurrency-Data-Pipeline-Portfolio-Manager
A dual-mode solution that fetches real-time coin data from CoinGecko API and transforms it into actionable portfolio insights through Power BI.

This Power BI dashboard delivers real-time crypto market intelligence and portfolio analytics by integrating with CoinGecko's API. Designed for both active traders and long-term investors, it transforms raw market data into actionable insights through dynamic visualizations and performance metrics.

## Power BI Preview

<img width="1467" height="808" alt="image" src="https://github.com/user-attachments/assets/2fac233e-401f-40dc-a4eb-c3305f88d0a6" />


## 📌 Key Features

### Automated Data Fetching

Fetches real-time data for 50+ (expandable 1500+ with CoinGecko API) cryptocurrencies (price, market cap, volume, dominance).

Tracks Fear & Greed Index and historical trends (BTC/ETH/BNB/SOL/SUI).

### Excel Output (crypto_portfolio.xlsx)

📊 Executive Dashboard: Top market metrics (total cap, dominance, sentiment).

📈 Market Overview: Ranked list of 50 coins with price history.

🌍 Global Metrics: Aggregate crypto market data.

😰 Fear & Greed Index: 30-day sentiment analysis.

🔄 Transactions/Portfolio Sheets: Manual entry for P/L tracking.

## Power BI Integration

Connect Excel to Power BI for interactive charts:

Portfolio performance

Market dominance trends

Buy/sell timing analysis


## ⚙️ How It Works

### 1. Python Script Setup

#### Clone repo
git clone https://github.com/LuffyMon53/Cryptocurrency-Data-Pipeline-Portfolio-Manager

cd Cryptocurrency-Data-Pipeline-Portfolio-Manager

#### Install dependencies
pip install pandas requests openpyxl

#### Run script (fetches fresh data)
Python_Scripts/Data_Fatcher.ipynb

### 2. Excel File Structure

| Sheet Name                  | Description                                                                 | Data Source           | Update Frequency |
|-----------------------------|-----------------------------------------------------------------------------|-----------------------|------------------|
| **📊 Executive Dashboard**   | Key market metrics (total cap, dominance, sentiment)                        | CoinGecko API         | Automatic        |
| **📈 Market Overview**       | Top 50 cryptocurrencies with price, volume, and historical performance      | CoinGecko API         | Automatic        |
| **🌍 Global Metrics**        | Aggregate market statistics (active coins, total markets)                   | CoinGecko API         | Automatic        |
| **😰 Fear & Greed Index**    | Daily sentiment scores with classification                                  | Alternative.me        | Automatic        |
| **📅 BTC History**           | 30-day price, volume, and moving averages for Bitcoin                       | CoinGecko API         | Automatic        |
| **📅 ETH History**           | 30-day price, volume, and moving averages for Ethereum                      | CoinGecko API         | Automatic        |
| **📅 BNB History**           | 30-day price, volume, and moving averages for Binance Coin                  | CoinGecko API         | Automatic        |
| **📅 SOL History**           | 30-day price, volume, and moving averages for Solana                        | CoinGecko API         | Automatic        |
| **📅 SUI History**           | 30-day price, volume, and moving averages for SUI                           | CoinGecko API         | Automatic        |
| **🔄 Transactions**          | Manual trade history (date, coin, type, quantity, price)                    | User Input            | Manual           |
| **💰 Current Portfolio**     | Holdings with auto-calculated P/L (purchase price vs current value)         | User Input            | Manual           |

### 3. Power BI Setup

Generate data: python fetch_crypto_data.py

Open CryptoDashboard.pbix in Power BI

Link to crypto_portfolio.xlsx when prompted

Click Refresh to update visuals

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


