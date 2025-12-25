# Derivatives-Pricing

---

### 📁 **Excel VBA Tools for Derivatives Pricing and Risk Analytics**

This repository contains a set of **Excel workbooks with embedded VBA macros** designed to assist in derivatives pricing, yield curve construction, and financial instrument valuation. These tools support core concepts used in institutional risk systems (e.g., Bloomberg MARS, Murex) but offer transparency and flexibility suitable for prototyping or independent analysis.

---

### 📄 `Binomial Tree - CRR.xlsm`

**Description:**
An Excel-based **option pricing tool** using the **Cox-Ross-Rubinstein (CRR)** binomial tree method. Suitable for valuing **European and American options**, with flexible model parameters and VBA-driven tree generation.

**Key Features:**

* Inputs for spot price, strike, risk-free rate, volatility, time to maturity
* Select between European and American style
* Dynamic tree generation with adjustable steps
* VBA-powered pricing engine with optional tree visualization

---

### 📄 `IRS.xlsm`

**Description:**
A **Net Present Value (NPV)** calculator for **Interest Rate Swaps (IRS)**. This workbook allows input of market curves and swap deal terms to compute the value of fixed and floating legs using standard industry conventions.

**Key Features:**

* Input modules for discount and forward curves
* Handles various tenors, frequencies, and day count conventions
* Outputs full cashflow schedules and NPV breakdown
* VBA automation for schedule generation and valuation logic

---

### 📄 `YieldCurve.xlsm`

**Description:**
A **bootstrapping engine** built in Excel to derive a **zero-coupon yield curve** from observable market instruments like deposits, futures, and swaps.

**Key Features:**

* Input interface for curve instruments and rates
* Bootstraps zero rates and computes discount/forward factors
* Fully automated via VBA for iterative curve construction
* Useful for pricing fixed income instruments or benchmarking curves


