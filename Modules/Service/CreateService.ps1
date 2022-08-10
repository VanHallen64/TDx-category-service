#------ Create Main Service Offering ------#

function New-Service($CategoryId, $UserInput) {
    $ServiceName = "General $($UserInput.CatName) Support"
    Enter-SeUrl ("$Domain"+"TDClient/81/askit/Requests/New?CategoryID=$CategoryId") -Driver $Driver

    #Name
    $CurrentField = Find-SeElement -Wait -Timeout 60 -Driver $Driver -Id "ctl00_ctl00_cpContent_cpContent_txtName"
    Send-SeKeys -Element $CurrentField -Keys $ServiceName
    $CurrentField = Find-SeElement -Driver $Driver -Id "ctl00_ctl00_cpContent_cpContent_txtShortDescription"
    Send-SeKeys -Element $CurrentField -Keys "If you don't find what you are looking for, please submit a $ServiceName ticket"

    # Long description
    $SourceBtn = Find-SeElement -Wait -Timeout 15 -Driver $Driver -Id "cke_16"
    $WebDriverWait = [OpenQA.Selenium.Support.UI.WebDriverWait]::new($Driver, (New-TimeSpan -Seconds 20))
    $Condition = [OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementToBeClickable($SourceBtn)
    $WebDriverWait.Until($Condition) | Out-null
    Invoke-SeClick -Element $SourceBtn
    $CurrentField = Find-SeElement -Wait -Timeout 10 -Driver $Driver -XPath '//div[@id="cke_1_contents"]//textarea'
    Send-SeKeys -Element $CurrentField -Keys "If you don't find what you are looking for, please submit a $ServiceName ticket"

    # Order
    $CurrentField = Find-SeElement -Driver $Driver -Id "ctl00_ctl00_cpContent_cpContent_txtOrder"
    $CurrentField.SendKeys([OpenQA.Selenium.Keys]::Up)

    # Manager
    $CloseBtn = Find-SeElement -Driver $Driver -Id "ctl00_ctl00_cpContent_cpContent_taluManager_btnCleartaluManager"
    Invoke-SeClick -Element $CloseBtn
    $CurrentField = Find-SeElement -Driver $Driver -Id "ctl00_ctl00_cpContent_cpContent_taluManager_txtinput"
    Send-SeKeys -Element $CurrentField -Keys "Client Services - Leadership"
    $CurrentField = Find-SeElement -Driver $Driver -CssSelector "#ctl00_ctl00_cpContent_cpContent_taluManager_txttaluManager_feed > li[rel='216']"
    Invoke-SeClick -Element $CurrentField

    # Request Application Type
    $Option = Find-SeElement -Driver $Driver -Id "ctl00_ctl00_cpContent_cpContent_ddlRequestApplication"
    $SelectElement = [OpenQA.Selenium.Support.UI.SelectElement]::new($Option)
    $SelectElement.SelectByValue("82")

    # Request Type ID
    $CurrentField = Find-SeElement -Driver $Driver -Id "ctl00_ctl00_cpContent_cpContent_taluRequestType_txtinput"
    Send-SeKeys -Element $CurrentField -Keys "IT Service Delivery and Support"
    $CurrentField = Find-SeElement -Driver $Driver -XPath "//ul[@id='ctl00_ctl00_cpContent_cpContent_taluRequestType_txttaluRequestType_feed']//li[@rel='757']"
    Invoke-SeClick -Element $CurrentField

    # Request Service Text
    $CurrentField = Find-SeElement -Driver $Driver -Id "ctl00_ctl00_cpContent_cpContent_txtRequestText"
    Send-SeKeys -Element $CurrentField -Keys "Request Support"

    # Tags
    $GeneralTags = @("general", "technical", "support", "GTS")
    foreach ($tag in $GeneralTags) {
        $CurrentField = Find-SeElement -Driver $Driver -Id "s2id_autogen1"
        Send-SeKeys -Element $CurrentField $tag
        $CurrentField = $Driver.FindElements([OpenQA.Selenium.By]::classname("select2-result-selectable")) | Where-Object {$_.Text -eq $tag}
        Invoke-SeClick -Element $CurrentField
    }

    # Audience
    $Audience = ""
    if ($UserInput.Faculty) {
        $Audience += "Faculty, "
    }
    if ($UserInput.Staff) {
        $Audience += "Staff, "
    }
    if ($UserInput.Students) {
        $Audience += "Students, "
    }
    if ($Audience) {
        $Audience = $Audience.Substring(0, $Audience.Length - 2)
    }
    $CurrentField = Find-SeElement -Driver $Driver -Id "ctl00_ctl00_cpContent_cpContent_attribute4210"
    Send-SeKeys -Element $CurrentField -Keys $Audience
    
    # Requirements
    if ($Audience) {
        if ($Audience.split(',').Length - 1 -gt 0) {
            $Audience = $Audience.Substring(0,$Audience.LastIndexOf(',')) + "or" + $Audience.Substring($Audience.LastIndexOf(',')+1, $Audience.Length - $Audience.LastIndexOf(',')-1)
        }
        $CurrentField = Find-SeElement -Driver $Driver -Id "ctl00_ctl00_cpContent_cpContent_attribute4212"
        Send-SeKeys -Element $CurrentField -Keys "Be $($Audience.ToLower())."
    }
}  