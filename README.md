# XU100 Money Flow Engine

Quantitative money flow analysis engine for XU100 stocks with sector-level insights and Excel reporting.

## Overview

This project analyzes short-term money flow behavior across XU100 stocks and aggregates results at both stock and sector levels.

Instead of focusing only on price, it attempts to answer a more relevant question:

Where is the money actually going?

## Features

- 10-day signed money flow calculation (volume-weighted)
- Flow ratio and statistical z-score normalization
- Flow strength scoring system
- Regression-based trend classification (89-period)
- Sector-level capital aggregation
- Excel-based reporting with charts and formatting
- Automatic data fetching via Yahoo Finance

## Output

The script generates a structured Excel report including:

- Stock-level money flow analysis
- Top-ranked stocks based on flow strength
- Sector-level capital flow summary
- Embedded mini price charts
- Conditional formatting for quick interpretation

## Methodology

Core logic:

- Typical price × volume → proxy for capital flow  
- Direction determined by price change  
- Aggregation over rolling 10-day window  
- Normalization using z-score  
- Composite scoring using:
  - flow ratio
  - net flow
  - return

This creates a relative strength measure of capital movement.

## Tech Stack

- Python
- pandas, numpy
- yfinance
- matplotlib
- openpyxl

## Usage

```bash
python your_script_name.py
