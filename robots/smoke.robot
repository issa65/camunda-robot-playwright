*** Settings ***
Library    Browser

*** Test Cases ***
Open Example Page
    New Browser    chromium    headless=False
    New Page    https://example.com

    ${title}=    Get Title
    Should Be Equal    ${title}    Example Domain

    ${heading}=    Get Text    h1
    Should Be Equal    ${heading}    Example Domain

    Take Screenshot
    Close Browser
