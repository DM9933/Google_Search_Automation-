# Google Search Automation

## Overview
This Selenium-based Java automation project automates Google search suggestions extraction based on keywords from an Excel file. The longest and shortest suggestions for each keyword are stored back into the Excel sheet.

## Features
- Reads keywords from an Excel file based on the current day.
- Performs Google search and captures autocomplete suggestions.
- Identifies the longest and shortest suggestions for each keyword.
- Writes the extracted results back into the corresponding Excel sheet.
- Uses TestNG for structured test execution.

## Technologies Used
- **Java** (Programming Language)
- **Selenium WebDriver** (Web Automation)
- **Apache POI** (Excel Handling)
- **TestNG** (Testing Framework)
- **WebDriverManager** (Driver Management)

## Installation & Setup

### Prerequisites
Ensure you have the following installed:
- **Java JDK (17 or later)**
- **Maven**
- **Chrome Browser**
- **Chromedriver** (Managed automatically by WebDriverManager)

### Clone the Repository
```sh
git clone https://github.com/your-username/GoogleSearchAutomation.git
cd GoogleSearchAutomation
