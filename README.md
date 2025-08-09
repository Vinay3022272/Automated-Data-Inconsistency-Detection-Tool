# Automated Data Inconsistency Detection Tool

## Overview

The **Automated Data Inconsistency Detection Tool** is a professional-grade solution for identifying and flagging data discrepancies within structured or semi-structured datasets. It is highly effective for auditing financial statements, sales reports, and performance dashboards where data accuracy is critical.

The tool detects:

* **Numerical inconsistencies**: Conflicting numerical values for the same metric across sources.
* **Percentage mismatches**: Incorrect percentage calculations based on given values.
* **Timeline mismatches**: Misaligned or inconsistent dates and periods.
* **Textual contradictions**: Directly conflicting statements.

---

## Key Features

* Automated scanning of structured and semi-structured datasets.
* Cross-page and cross-slide comparison to spot discrepancies.
* Context-rich, detailed reporting of identified issues.
* Supports multiple input formats, including extracted data from PDFs, PPTs, and spreadsheets.

---

## Workflow

1. **Data Input**: Provide structured data as Python dictionaries or objects.
2. **Analysis**: The tool runs detection modules for:

   * Numerical inconsistencies
   * Percentage mismatches
   * Timeline mismatches
   * Textual contradictions
3. **Output**: Generates `Inconsistency` objects with:

   * Type of inconsistency
   * Relevant context (e.g., page/slide number, time period)
   * Detected values

---

## README Summary

This tool streamlines the detection of inconsistencies in datasets, making it invaluable for business, finance, and research data validation.



