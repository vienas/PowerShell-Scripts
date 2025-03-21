
Import-Module Selenium

$DriverPath = "C:\Program Files\Google\Chrome\Application\" 

$ChromeOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions
$DriverService = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService($DriverPath)
$Driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($DriverService, $ChromeOptions)

if ($Driver -eq $null) {
    Write-Host "Failed to initialize the browser."
    exit
}

if ($MaximoURL) {

$MaximoURL = "../maximo/webclient/login/login.jsp?welcome=true"
$Username = ""
$Password = ""
}


    $Driver.Navigate().GoToUrl($MaximoURL)

    $UsernameField = $Driver.FindElementById("username")
    $UsernameField.SendKeys($Username)
    
    $PasswordField = $Driver.FindElementById("password")
    $PasswordField.SendKeys($Password)

    $LoginButton = $Driver.FindElementById("loginbutton")
    $LoginButton.Click()

    $LoginButSpecjalistaTabton = $Driver.FindElementById("m1e20cba1-sct_anchor_1")
    $LoginButSpecjalistaTabton.Click()

    sleep 5

    $OpenList = $Driver.FindElementByXPath("/html/body/form/table[2]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/div/table/tbody/tr/td/div/table/tbody/tr/td/table/tbody/tr/td/div/table/tbody/tr[3]/td[2]/div[8]/table/tbody/tr[1]/td/table/tbody/tr/td[11]/img")
    $OpenList.Click()
    
    sleep 30
    
    $ChoiseRecords = $Driver.FindElementByXPath("//*[@id='m6a7dfd2f_tbod_ttselrows-ti_label']")
    $ChoiseRecords.Click()
