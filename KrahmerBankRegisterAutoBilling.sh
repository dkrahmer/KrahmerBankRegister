#!/usr/bin/bash

# Configuration
appName="KrahmerBankRegisterAutoBilling"
triggerDirectory="/opt/kerio/mailserver/store/mail/???/???/???/#msgs"
logFile="/home/services/.config/${appName}/${appName}.log"
# Use /dev for development or /exec for production
url_dev="https://script.google.com/macros/s/?????/exec"
url_prod="https://script.google.com/macros/s/?????/exec"
url="$url_prod"
passphrase="This is my super secure passphrase"
# Strip non-alphanumeric characters to match server-side validation
passphrase_clean="${passphrase//[^a-zA-Z0-9]/}"

# Ensure log directory exists
mkdir -p "$(dirname "$logFile")"

echo "$(date '+%Y-%m-%d %H:%M:%S') Starting ${appName}..." | tee -a "$logFile"

while : ; do
{
    echo "$(date '+%Y-%m-%d %H:%M:%S') Waiting for trigger file in ${triggerDirectory}"

    # Monitor for file events continuously
    sudo inotifywait -e create -e moved_to --format '%f' -m "${triggerDirectory}" | while read triggerEvent
    do
        filePath="${triggerDirectory}/${triggerEvent}"

        echo "$(date '+%Y-%m-%d %H:%M:%S') File detected: ${triggerEvent}. Processing..."

        # Read file content
        emailContent=$(sudo cat "$filePath")

        # Send Web Post
        # Note: Don't follow redirects - the script executes on first POST
        # Note: We'll get HTTP 302 but the script runs successfully
        # Note: Query parameters for function and passphrase
        # Note: JSON body contains the email data
        response=$(curl -s \
             -X POST "${url}?function=processEmailBill&passphrase=${passphrase_clean}" \
             -H "Content-Type: application/json" \
             -d "{\"email\": $(echo "$emailContent" | jq -Rs .)}" \
             --write-out "\nHTTP_STATUS:%{http_code}")

        http_status=$(echo "$response" | grep "HTTP_STATUS:" | cut -d: -f2)
        response_body=$(echo "$response" | sed '/HTTP_STATUS:/d')

        if [[ -n "$response_body" && "$response_body" != "HTTP_STATUS:"* ]]; then
            echo "Response: $response_body"
        fi

        # HTTP 302 is expected and means the script executed successfully
        if [[ "$http_status" -eq 302 || ($http_status -ge 200 && $http_status -lt 300) ]]; then
            echo "Success (HTTP $http_status) - Script executed"
        else
            echo "Error: HTTP $http_status"
        fi

        echo -e "\n$(date '+%Y-%m-%d %H:%M:%S') POST completed for ${triggerEvent}"
    done
} 2>&1 | tee -a "$logFile"
done
