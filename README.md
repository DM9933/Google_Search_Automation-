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

## Dependencies
All dependencies are managed via Maven. Ensure your pom.xml contains:
```xml
<dependencies>
    <dependency>
        <groupId>org.seleniumhq.selenium</groupId>
        <artifactId>selenium-java</artifactId>
        <version>4.9.1</version>
    </dependency>
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi-ooxml</artifactId>
        <version>5.2.3</version>
    </dependency>
    <dependency>
        <groupId>io.github.bonigarcia</groupId>
        <artifactId>webdrivermanager</artifactId>
        <version>5.5.3</version>
    </dependency>
    <dependency>
        <groupId>org.testng</groupId>
        <artifactId>testng</artifactId>
        <version>7.6.1</version>
        <scope>test</scope>
    </dependency>
</dependencies>
```
### Running the Project
1. Open the project in an IDE (Eclipse/IntelliJ).
2. Ensure TestNG is installed and configured.
3. Update excelFilePath in GoogleSearch.java to point to your local Excel file.
4. Run the test class using
```sh
mvn test
```
## File Structure
```
GoogleSearchAutomation/
│-- src/main/java/me/selenium/POM/GoogleSearch.java
│-- pom.xml
│-- Excel.xlsx (Your data file)
```
## How It Works
1. Opens Chrome browser using WebDriverManager.
2. Reads today's sheet from the Excel file.
3. Extracts keywords from column C (Index 2) in the Excel sheet.
4. Searches Google for each keyword.
5. Retrieves autocomplete suggestions and finds:
    - The longest suggestion.
    - The shortest suggestion.
6. Writes results back to column D & E in the Excel sheet.
7. Closes the browser after execution.
## Contributing
1. Fork the repository.
2. Create a feature branch (feature-branch-name).
3. Commit your changes.
4. Open a Pull Request.

## Watch Video
![Video](https://github.com/DM9933/Google_Search_Automation-/blob/main/Assignment1Q1%20(online-video-cutter.com).mp4)
