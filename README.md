# IdentifyIncomingWebhooks
This code sample identifies teams with incoming webhooks

**Disclaimer: This script is provided as a sample as-is.**

Note: This sample script is using Client ID and Secret and runs in the application context.

# Steps 1
1. Register Entra App ID. While registering the app:
    * Provide Name
    * Selet Accounts in this org directory only (Single Tenant)
    * Rediret URI - empty
2. Gather the Tenant Id, Client Id, Client Secret
3. Ensure below application permissions are added and admin consent granted.
    * Group.Read.All
    * TeamsAppInstallation.Read.All

# Step 2
1. Update the script with the Tenant Id, Client id, and Client Secret.
2. Run the script

# Output
You can see the sample output file in this repository. 