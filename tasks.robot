*** Settings ***
Documentation     Robot to enter weekly sales data into the RobotSpareBin Industries Intranet.
Library   RPA.Browser.Selenium
Library   RPA.HTTP
Library   RPA.Excel.Files
Library   RPA.PDF

*** Keywords ***
Open The Intranet Website
    Open Available Browser   https://robotsparebinindustries.com/
    Maximize Browser Window    

*** Keywords ***
Login With Valid Credential
    Input Text    username    maria
    Input Password    password    thoushallnotpass
    Submit Form
    Wait Until Page Contains Element    id:sales-form

*** Keywords ***
Download The Excel File
    Download    https://robotsparebinindustries.com/SalesData.xlsx    overwrite=True

*** Keywords ***
Fill The Form Using The Data From The Excel File
    Open Workbook    SalesData.xlsx
    ${excel_data}=    Read Worksheet As Table    header=True
    Close Workbook
    FOR  ${sales_data}  IN  @{excel_data}
        Fill And Submit The Form For One Person    ${sales_data}
    END

*** Keywords ***
Fill And Submit The Form For One Person
    [Arguments]   ${sales_data}
    Input Text    firstname    ${sales_data}[First Name]
    Input Text    lastname    ${sales_data}[Last Name]
    Input Text    salesresult    ${sales_data}[Sales]
    ${target_as_string}=    Convert To String    ${sales_data}[Sales Target]
    Select From List By Value    salestarget    ${target_as_string}
    Click Button    Submit

*** Keywords ***
Collect The Result
    Screenshot    css:div.sales-summary    ${CURDIR}${/}output${/}sales_summary.png

*** Keywords ***
Export The Table As A PDF
    Wait Until Element Is Visible    id:sales-results
    ${sales_results_html}=    Get Element Attribute    id:sales-results    outerHTML
    Html To Pdf    ${sales_results_html}    ${CURDIR}${/}output${/}sales_results.pdf

*** Keywords ***
Log Out And Close The Browser
    Click Button    id:logout
    Close Browser

*** Tasks ***
Insert the sales data for the week and export it as a PDF
    Open The Intranet Website
    Login With Valid Credential
    Download The Excel File
    Fill The Form Using The Data From The Excel File
    Collect The Result
    Export The Table As A PDF
    [Teardown]    Log Out And Close The Browser
