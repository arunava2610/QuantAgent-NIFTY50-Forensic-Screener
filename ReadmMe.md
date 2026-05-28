# QuantAgent: NIFTY 50 Agentic Portfolio & Forensic Screening Engine 🚀📊

QuantAgent is an enterprise-grade, two-pass hybrid AI investment analysis pipeline designed to process, evaluate, and screen equities listed on the NIFTY index. By marrying algorithmic financial models with advanced generative AI orchestration, the system conducts rigorous multi-period equity research, generates automated trend-line visualizations, and delivers professional investment summaries to a distributed email pool.

## 🛠️ System Architecture & Engineering Workflow

The pipeline operates as an automated investment committee executing via a secure, two-pass analytical lifecycle:

1. **Pass 1: Mathematical Extraction & Optimization Engine**
   * Dynamically tracks and downloads 5-year historical income statement, balance sheet, and cash flow data using the Yahoo Finance API.
   * Compiles data into an optimized, structured multi-sheet Excel ledger with professional styling, section merges, and clean data formatting rules.
   * Mathematically evaluates key structural financial health indicators (Profitability, ROI, Solvency, and Liquidity) across all available fiscal periods.
   * Renders high-resolution ratio trend charts natively using Matplotlib.

2. **Pass 2: Context-Enriched Agentic Screening & Fallback Verification**
   * Packages high-density financial data layers into an optimized context schema, bypassing token exhaustion constraints.
   * Feeds metrics directly to a customized, prompt-engineered Gemini 2.5 generative engine to model a Wall Street credit auditor and value investor persona.
   * **Hybrid Reliability System:** Features an automatic programmatic fallback layer. If public API rate limits (`429 Quota Exhausted`) are hit, the system switches to a local rule-based expert system to evaluate margins and debt, ensuring the report compiles successfully without skipping data.
   * Generates a beautifully formatted Word Document report (`.docx`) featuring color-coded trade actions (**Green for BUY**, **Red for SELL**, **Gold for HOLD**), calculated ratio grids, embedded trendline graphics, and dense multi-sentence analytical rationale blocks.

3. **Pass 3: Short-Circuit Orchestration & Automated Mail Gateway**
   * Includes structural checks to prevent redundant API computing outlays: skips the crawler if raw data is present, and skips the entire analysis matrix if the report is already compiled.
   * Securely establishes an encrypted TLS tunnel to log into an SMTP email gateway, distributing the finalized investment portfolio report directly to a multi-recipient mailing list parsed dynamically from system environments.

---

## 📈 Tech Stack & Infrastructure

* **Language:** Python 3.14+
* **Data Core:** Pandas, NumPy, OpenPyXL (Excel Layout Sheet Modeling)
* **Visualization Engine:** Matplotlib (Trend-line Graphics Engine)
* **Generative Orchestration:** Google GenAI SDK (`gemini-2.5-flash`)
* **Document Processing:** Python-Docx (Native MS Word Composition API)
* **Security & Environments:** Python-Dotenv (Credential Isolation Structures)
* **Distribution Protocols:** Smtplib, Email (Secure Network Mail Handlers)

---

## 🔍 Interactive Output Sample Showcase

When the pipeline compiles, it yields a professional Microsoft Word Executive Report. Below is a sample illustration showing exactly how an individual equity is formatted and presented to the investment committee:

### 🏢 InterGlobe Aviation Ltd. (INDIGO.NS) | *Action Recommendation:* **BUY**

#### 🔢 Mathematically Derived Core Ratio History:
* • Fiscal Year 2023 | Profitability (NPM): -1.54% | ROI (ROA Proxy): -0.45% | Solvency (Debt/Assets Ratio): 42.10% | Liquidity (Cash/Assets Ratio): 8.30%
* • Fiscal Year 2024 | Profitability (NPM): 8.90%  | ROI (ROA Proxy): 6.20%  | Solvency (Debt/Assets Ratio): 38.45% | Liquidity (Cash/Assets Ratio): 12.15%
* • Fiscal Year 2025 | Profitability (NPM): 11.25% | ROI (ROA Proxy): 8.40%  | Solvency (Debt/Assets Ratio): 31.20% | Liquidity (Cash/Assets Ratio): 15.60%

#### 🧠 Agentic Committee Forensic Rationale:
> "QuantAgent indicates a strong **BUY** validation for InterGlobe Aviation Ltd. based on a structural operational turnaround completed between FY2023 and FY2025. Net Profit Margins expanded dramatically from a negative 1.54% baseline to a robust 11.25%, backed explicitly by an upward trajectory in core Operating Cash Flows, which signals high earnings quality. Solvency frameworks reflect substantial deleveraging, with the Long-Term Debt-to-Assets ratio contracting from 42.10% to 31.20%, minimizing system default risk. Furthermore, cash-to-asset metrics reveal a fortified liquidity cushion sitting at 15.60%, giving the enterprise significant runway to fund upcoming capital expenditures internally without introducing further debt strain."

#### 📊 Visual Performance Trend Model:
*The system automatically plots and embeds custom multi-period trend visualizations right beneath the rationale text:*

| Metric Scale (%) | Trend Line Chart Preview |
| :--- | :--- |
| 50% | `  /\ `  *-- Solvency (Debt/Assets declining safely)* |
| 25% | ` /--\`  *-- Profitability (NPM surging upwards)* |
| 0%  | `____/`  *-- ROI / ROA (Recovering into strong positive bounds)* |
|     | **2023 -------- 2024 -------- 2025** |

---

## ⚙️ Secure Configuration & Installation

This project utilizes professional decoupling standards. Private API tokens and server login passwords are never hardcoded inside the source file.

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/YOUR_USERNAME/QuantAgent-NIFTY50-Forensic-Screener.git](https://github.com/YOUR_USERNAME/QuantAgent-NIFTY50-Forensic-Screener.git)
   cd QuantAgent-NIFTY50-Forensic-Screener