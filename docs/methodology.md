# Methodology

## Barrier Probability Engine

The engine follows a practical desk-oriented workflow:

1. Read operations from an Excel workbook.
2. Read listed option series data from an exchange series file.
3. Prepare option tickers required for a market data / RTD export.
4. Read quote and implied volatility data from the export sheet.
5. Filter quotes by liquidity, open interest, volume, bid/ask spread and delta.
6. Fit a simple volatility smile/skew approximation around the barrier level.
7. Estimate volatility at the barrier level.
8. Estimate the probability of a downside barrier breach by closing price.
9. Write results back to the workbook.

This is an operational approximation. It is not intended to replace a full pricing/risk library.

## Early Unwind Analysis

The early unwind module compares:

- Entry reference;
- Current reference;
- Bid/offer result components;
- Quantity;
- Remaining term;
- CDI accumulated over the period.

The objective is to identify operations with positive outcome potential and provide a structured report.

## Structured Products Fixing

The fixing module identifies products with fixing date on the current day, calculates outcome fields and prepares advisor-level communication outputs.

## Limitations

- Market data quality directly affects the result.
- The volatility fit is simplified.
- Operational assumptions should be validated before production usage.
- Outputs should be reviewed by a qualified professional.
