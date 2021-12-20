
*** Settings ***
Documentation     scraping data from kind
Library           RPA.Browser.Selenium
Library           RPA.Excel.Files



*** Variables ***
${URL}     https://kind.krx.co.kr/listinvstg/listingcompany.do?method=searchListingTypeMain
${SEARCH_TEXT}    %{SEARCH_TEXT}


*** Keywords ***
Search for
    [Arguments]    ${text}
    Wait Until Page Contains Element             xpath://input[@id='AKCKwd']
    Input Text    xpath://input[@id='AKCKwd']    ${text}
    Press Keys    xpath://input[@id='AKCKwd']    ENTER
    Wait Until Page Contains Element    xpath://td[@title='${SEARCH_TEXT}']

*** Keywords ***
재무상태표 스크래핑
    FOR    ${index}    IN RANGE    2  50   
        ${count}=     Get Element Count    xpath:/html[1]/body[1]/form[1]/section[1]/div[1]/table[1]/tbody[1]/tr[${index}]/th[1]
        IF    ${count} > 0
            ${항목}=              Get Text    xpath:/html[1]/body[1]/form[1]/section[1]/div[1]/table[1]/tbody[1]/tr[${index}]/th[1]
            ${2018}=              Get Text    xpath:/html[1]/body[1]/form[1]/section[1]/div[1]/table[1]/tbody[1]/tr[${index}]/td[1]
            ${2019}=              Get Text    xpath:/html[1]/body[1]/form[1]/section[1]/div[1]/table[1]/tbody[1]/tr[${index}]/td[2]
            ${2020}=              Get Text    xpath:/html[1]/body[1]/form[1]/section[1]/div[1]/table[1]/tbody[1]/tr[${index}]/td[3]  

            &{row}=       Create Dictionary
            ...           항목           ${항목}
            ...           2018           ${2018}
            ...           2019           ${2019}
            ...           2020           ${2020}
            Append Rows to Worksheet  ${row}  header=${TRUE}                              
        ELSE
            Log    noelement
        END
    END


*** Keywords ***
포괄손익계산서 스크래핑
    FOR    ${index}    IN RANGE    2  50   
        ${count}=     Get Element Count    xpath:/html[1]/body[1]/form[1]/section[1]/div[1]/table[2]/tbody[1]/tr[${index}]/th[1]
        IF    ${count} > 0
            ${항목}=              Get Text    xpath:/html[1]/body[1]/form[1]/section[1]/div[1]/table[2]/tbody[1]/tr[${index}]/th[1]
            ${2018}=              Get Text    xpath:/html[1]/body[1]/form[1]/section[1]/div[1]/table[2]/tbody[1]/tr[${index}]/td[1]
            ${2019}=              Get Text    xpath:/html[1]/body[1]/form[1]/section[1]/div[1]/table[2]/tbody[1]/tr[${index}]/td[2]
            ${2020}=              Get Text    xpath:/html[1]/body[1]/form[1]/section[1]/div[1]/table[2]/tbody[1]/tr[${index}]/td[3]  

            &{row}=       Create Dictionary
            ...           항목           ${항목}
            ...           2018           ${2018}
            ...           2019           ${2019}
            ...           2020           ${2020}
            Append Rows to Worksheet  ${row}  header=${TRUE}                              
        ELSE
            Log    noelement
        END
    END


*** Keywords ***
현금흐름표 스크래핑
    FOR    ${index}    IN RANGE    2  50   
        ${count}=     Get Element Count    xpath:/html[1]/body[1]/form[1]/section[1]/div[1]/table[3]/tbody[1]/tr[${index}]/th[1]
        IF    ${count} > 0
            ${항목}=              Get Text    xpath:/html[1]/body[1]/form[1]/section[1]/div[1]/table[3]/tbody[1]/tr[${index}]/th[1]
            ${2018}=              Get Text    xpath:/html[1]/body[1]/form[1]/section[1]/div[1]/table[3]/tbody[1]/tr[${index}]/td[1]
            ${2019}=              Get Text    xpath:/html[1]/body[1]/form[1]/section[1]/div[1]/table[3]/tbody[1]/tr[${index}]/td[2]
            ${2020}=              Get Text    xpath:/html[1]/body[1]/form[1]/section[1]/div[1]/table[3]/tbody[1]/tr[${index}]/td[3]  

            &{row}=       Create Dictionary
            ...           항목           ${항목}
            ...           2018           ${2018}
            ...           2019           ${2019}
            ...           2020           ${2020}
            Append Rows to Worksheet  ${row}  header=${TRUE}                              
        ELSE
            Log    noelement
        END
    END



*** Tasks ***
Minimal task
    Open Available Browser    ${URL}
    Search for                ${SEARCH_TEXT}
    Sleep    1.5s
    Wait Until Page Contains Element    xpath://td[@title='${SEARCH_TEXT}']
    Click Element    xpath://td[@title='${SEARCH_TEXT}']
    Switch Window    new
    Wait Until Page Contains Element    xpath://a[@title='재무정보']
    Click Element                       xpath://a[@title='재무정보']
    Sleep    1.5s
    Create Workbook  ${SEARCH_TEXT}.xlsx
    
    Rename worksheet    Sheet    재무상태표
    재무상태표 스크래핑

    Create worksheet    포괄손익계산서
    포괄손익계산서 스크래핑

    Create worksheet    현금흐름표
    현금흐름표 스크래핑

    Save Workbook    .//output//${SEARCH_TEXT}.xlsx
    Close All Browsers