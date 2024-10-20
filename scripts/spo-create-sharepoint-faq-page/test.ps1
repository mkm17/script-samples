$m365Status = m365 status
if ($m365Status -match "Logged Out") {
    m365 login
}

$csvFilePath = "assets/FAQSetup.csv"
$webUrl = "https://1g044k.sharepoint.com/sites/BATTestSite"
$page = "test2.aspx"

$csvContent = Import-Csv -Path $csvFilePath
$index = 1
foreach ($row in $csvContent) {

    switch($row.BackgroundType) {
        "Gradient" {
            m365 spo page section add --pageName $page --webUrl $webUrl --sectionTemplate OneColumn --zoneEmphasis Gradient --gradientText $row.BackgroundDetails --isCollapsibleSection --order $index
        }
        "Image" {
            m365 spo page section add --pageName $page --webUrl $webUrl --sectionTemplate OneColumn --zoneEmphasis Image --imageUrl $row.BackgroundDetails --fillMode Tile --isCollapsibleSection --order $index
        }
        Default {
            m365 spo page section add --pageName $page --webUrl $webUrl --sectionTemplate OneColumn --zoneEmphasis $row.BackgroundType --isCollapsibleSection --order $index
        }
    }

    m365 spo page clientsidewebpart add --pageName $page --webUrl $webUrl --standardWebPart SiteActivity --section $index
    $index++
}
