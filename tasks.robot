*** Settings ***
Documentation       Orders robots from RobotSpareBin Industries Inc.
...                 Saves the order HTML receipt as a PDF file.
...                 Saves the screenshot of the orderd robot.
...                 Embeds the screenshot of the robot to the PDF receipt.
...                 Creates ZIP archive of the receipts and the images.

Library             RPA.Browser.Selenium    #auto_close=${False}
Library             RPA.HTTP
Library             RPA.Excel.Files
Library             RPA.Tables
Library             RPA.PDF
Library             RPA.Archive
Library             RPA.Robocorp.Vault.FileSecrets
Library             RPA.Dialogs


*** Variables ***
${ERROR_MESSAGE}                xpath:/html/body/div/div/div[1]/div/div[1]/div
${PDF_TEMP_OUTPUT_DIRECTORY}    ${OUTPUT_DIR}/PDFs


*** Tasks ***
Order robots from RobotSpareBin Industries Inc
    Open Browser robot order
    ${orders}=    Get orders
    FOR    ${row}    IN    @{orders}
        Log    ${row}
        Click yep button
        Fill the form    ${row}
        Preview the robot
        Click Order
        Store receipt as PDF    ${row}
        Take screenshot of robot    ${row}
        Embed the screenshot to the receipt PDF file    ${row}
        Make another order
    END

    Create Zip file of recipts

Minimal task
    Log    Done.


*** Keywords ***
Open Browser robot order
    ${loginwebsite}=    RPA.Robocorp.Vault.FileSecrets.Get Secret    loginwebsite
    Open Available Browser    ${loginwebsite}[website]
    Maximize Browser Window

Click yep button
    Click Button    Yep

Get orders
    Add text input    UrlExcel    label=Insert Excel URL
    ${excelfile}=    Run dialog
    Log    ${excelfile.UrlExcel}
    Download    ${excelfile.UrlExcel}    overwrite=True
    ${orders}=    Read a csv file into Tables
    Close Workbook
    RETURN    ${orders}

Fill the form
    [Arguments]    ${row}
    Select From List By Value    head    ${row}[Head]
    Click Element    id-body-${row}[Body]
    Input Text    xpath:/html/body/div/div/div[1]/div/div[1]/form/div[3]/input    ${row}[Legs]
    Input Text    address    ${row}[Address]

Preview the robot
    Wait And Click Button    id:preview
    Wait Until Element Is Visible    id:robot-preview-image

Click Order
    Wait Until Keyword Succeeds    5x    0.5 sec    Submit Order

Submit Order
    Wait And Click Button    id:order
    Is Element Visible    id:order-completion    missing_ok=False

Read a csv file into Tables
    ${orders}=
    ...    Read table from CSV
    ...    ${CURDIR}${/}orders.csv
    ...    header=True
    Log    ${orders}
    RETURN    ${orders}

Make another order
    Wait And Click Button    id:order-another

Take screenshot of robot
    [Arguments]    ${row}
    Screenshot    id:robot-preview-image    ${OUTPUT_DIR}/Pictures${/}${row}[Order number].png

Store receipt as PDF
    [Arguments]    ${row}
    ${recipt_element}=    Get Element Attribute    id:receipt    outerHTML
    Html To Pdf    ${recipt_element}    ${OUTPUT_DIR}/PDFs${/}${row}[Order number].pdf    overwrite=True

Embed the screenshot to the receipt PDF file
    [Arguments]    ${row}
    Open Pdf    ${OUTPUT_DIR}/PDFs${/}${row}[Order number].pdf
    Add Watermark Image To Pdf
    ...    ${OUTPUT_DIR}/Pictures${/}${row}[Order number].png
    ...    ${OUTPUT_DIR}/PDFs${/}${row}[Order number].pdf

Create Zip file of recipts
    ${zip_file_name}=    Set Variable    ${OUTPUT_DIR}/PDFs.zip
    Archive Folder With Zip
    ...    ${PDF_TEMP_OUTPUT_DIRECTORY}
    ...    ${zip_file_name}
