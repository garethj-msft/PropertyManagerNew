#region Setup
$graphApiVersion = "test_IntuneApplications_20160621"
$graphUri = "https://graph.microsoft.com/$graphApiVersion/mobileApps"
$authHeader= "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6IlliUkFRUlljRV9tb3RXVkpLSHJ3TEJiZF85cyIsImtpZCI6IlliUkFRUlljRV9tb3RXVkpLSHJ3TEJiZF85cyJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC8wYzNkN2QxNy0wODZkLTRmZDUtOWM3Ny1mMjJhOTFhNDdiMjkvIiwiaWF0IjoxNDcwMDY4MzA0LCJuYmYiOjE0NzAwNjgzMDQsImV4cCI6MTQ3MDA3MjIwNCwiYWNyIjoiMSIsImFtciI6WyJwd2QiXSwiYXBwaWQiOiI4YTNlYjg2Yi04MTQ5LTQyMzEtOWZmMy0zYzUwOTU4ZWEwZmQiLCJhcHBpZGFjciI6IjAiLCJmYW1pbHlfbmFtZSI6Ik1vbmtleSIsImdpdmVuX25hbWUiOiJUZW5hbnQiLCJpcGFkZHIiOiIxMzEuMTA3LjE1OS44IiwibmFtZSI6IlRlbmFudCBNb25rZXkgQWNjb3VudCIsIm9pZCI6ImIzYTA5MDM5LWIwOGItNGJlYS1hMTY5LWY2ZTFhMmQ0ODQyMCIsInBsYXRmIjoiV2luIiwicHVpZCI6IjEwMDMzRkZGOTk0MDQwQjUiLCJzY3AiOiJDYWxlbmRhcnMuUmVhZCBDYWxlbmRhcnMuUmVhZFdyaXRlIENvbnRhY3RzLlJlYWQgQ29udGFjdHMuUmVhZFdyaXRlIERpcmVjdG9yeS5BY2Nlc3NBc1VzZXIuQWxsIERpcmVjdG9yeS5SZWFkLkFsbCBEaXJlY3RvcnkuUmVhZFdyaXRlLkFsbCBlbWFpbCBGaWxlcy5SZWFkIEZpbGVzLlJlYWQuQWxsIEZpbGVzLlJlYWQuU2VsZWN0ZWQgRmlsZXMuUmVhZFdyaXRlIEZpbGVzLlJlYWRXcml0ZS5BbGwgRmlsZXMuUmVhZFdyaXRlLkFwcEZvbGRlciBGaWxlcy5SZWFkV3JpdGUuU2VsZWN0ZWQgR3JvdXAuUmVhZC5BbGwgR3JvdXAuUmVhZFdyaXRlLkFsbCBJZGVudGl0eVJpc2tFdmVudC5SZWFkLkFsbCBNYWlsLlJlYWQgTWFpbC5SZWFkV3JpdGUgTWFpbC5TZW5kIE1haWxib3hTZXR0aW5ncy5SZWFkV3JpdGUgTm90ZXMuQ3JlYXRlIE5vdGVzLlJlYWQgTm90ZXMuUmVhZC5BbGwgTm90ZXMuUmVhZFdyaXRlIE5vdGVzLlJlYWRXcml0ZS5BbGwgTm90ZXMuUmVhZFdyaXRlLkNyZWF0ZWRCeUFwcCBvZmZsaW5lX2FjY2VzcyBvcGVuaWQgUGVvcGxlLlJlYWQgcHJvZmlsZSBTaXRlcy5SZWFkLkFsbCBUYXNrcy5SZWFkIFRhc2tzLlJlYWRXcml0ZSBVc2VyLlJlYWQgVXNlci5SZWFkLkFsbCBVc2VyLlJlYWRCYXNpYy5BbGwgVXNlci5SZWFkV3JpdGUgVXNlci5SZWFkV3JpdGUuQWxsIiwic3ViIjoielRaQ09Kd200eXF4Vm5mYlBqMUhULTdQSzdBejc4MUJPYWQtQ1Y3Wm9lVSIsInRpZCI6IjBjM2Q3ZDE3LTA4NmQtNGZkNS05Yzc3LWYyMmE5MWE0N2IyOSIsInVuaXF1ZV9uYW1lIjoiYWRtaW5AcGV0cmljaHRlY2hyZWFkeS5vbm1pY3Jvc29mdC5jb20iLCJ1cG4iOiJhZG1pbkBwZXRyaWNodGVjaHJlYWR5Lm9ubWljcm9zb2Z0LmNvbSIsInZlciI6IjEuMCJ9.uj3i7xCbfqtFKsAvc0nfRxyRMdztT_l3-gCbaKkQ5UqPI8Hw_i-mYpgPBQB2FBBJdkjgRXw_KCguzw20LEANUSLSX_Hqe0wlEtHNsaSxcBRMe2_2ltRAvCplsRDpLeOLmVPoPvDXF9uJcJeyV9axQBBy63nb5xQrnO2isNx6OHTcDeWFAHguC8sihIqJii8jGp4rtYuFqG9npFYOKgM9zLFtq15BZCQvXCMM2R0NANQRwAk4w7bH0V8C-A5JtJhf3uMNkevo2W0AjJAtZBD6bQn7GM7Ktnet_d0_2bMZHFIkyXI9t_hIwEYmT9TAOVHGfm1m1xcv3_F6lEvgIhZ7rw"
$deleteAll = $false
$publish = $true
#endregion

# Step 0) Delete any existing apps (for demo only)
if ($deleteAll) 
{
    $existingApps = Invoke-RestMethod -Uri $graphUri -Method Get -Headers @{Authorization=$authHeader}

    Write-Host "Found $($existingApps.value.Count) existing apps"
    ForEach ($existingApp in $existingApps.value)
    {
        Write-Host "Deleting $($existingApp.displayName) ($($existingApp.id))"
        $appUri = "$graphUri/$($existingApp.id)"
        Invoke-RestMethod -Uri $appUri -Method Delete -Headers @{Authorization=$authHeader}
    }
}

if ($publish)
{
    # Step 1) Get the list of apps from the iTunes Search API
    $iTunesUrl = "https://itunes.apple.com/search?entity=software&term=microsoft+corporation&attribute=softwareDeveloper&limit=100"
    $apps = Invoke-RestMethod -Uri $iTunesUrl -Method Get

    Write-Host "Found $($apps.resultCount) apps"
    foreach ($app in $apps.results)
    {
        Write-Host "Publishing $($app.trackName)"
        # Step 2) Download the icon for the app
        $iconUrl = $app.artworkUrk60

        if ($iconUrl -eq $null)
        {
            Write-Host "`t60x60 icon not found, using 100x100 icon"
            $iconUrl = $app.artworkUrl100
        }
        if ($iconUrl -eq $null)
        {
            Write-Host "`t60x60 icon not found, using 512x512 icon"
            $iconUrl = $app.artworkUrl512
        }

        $iconResponse = Invoke-WebRequest $iconUrl
        $base64icon = [System.Convert]::ToBase64String($iconResponse.Content)
        $osVersion = [Convert]::ToDouble($app.minimumOsVersion)

        # Step 3) Create the entity object
        $graphApp = @{
            "@odata.type"="microsoft.graph.iosStoreApp";
            displayName=$app.trackName;
            publisher=$app.artistName;
            description=$app.description.Replace('“','"').Replace('”', '"')
            largeIcon= @{
               type=$iconResponse.Headers["Content-Type"];
               value=$base64icon;
            };
            informationUrl=$app.sellerUrl;
            bundleId=$app.bundleId;
            appStoreUrl=$app.trackViewUrl;
            applicableDeviceType=@{
                iPad=$app.supportedDevices -contains "iPadMini";
                iPhoneAndIPod=$app.supportedDevices -contains "iPhone6";
            };
            minimumSupportedOperatingSystem=@{
                v7_1=$osVersion -le 7.1;
                v8_0=$osVersion -eq 8.0
                v9_0=$osVersion -gt 8.0
            };
        }
    
        #Step 4) Publish the application to Graph
        Write-Host "`tCreating application via Graph"
        $createResult = Invoke-RestMethod -Uri $graphUri -Method Post -ContentType "application/json" -Body (ConvertTo-Json $graphApp) -Headers @{Authorization=$authHeader}
        Write-Host "`tApplication created as $graphUri/$($createResult.id)"
    }
}