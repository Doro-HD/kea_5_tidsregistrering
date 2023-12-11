# Deployment
A detailed guide on how to deploy this Microsoft Office Outlook add-in for your own organisation.

## Prerequisites
To deploy this setup you are going to need to setup a few Azure resources. This guide assumes you have a microsoft exchange organisation.

**Azure resources**

- Sql database
- App service for the api
- Storage account / optionally use another way to host your files
    - A container that holds the relevant files

### Pipeline
For the pipeline you can use ours as an example, .github/workflows/deploy.yml, but feel free to change it arround if it dosn't entirely suit your needs.
Currently we have only automated the upload process of files to azure but not the deployment of the manifest file.

### Deploy to exchange
You can run this attended script, with your modifications, to add the app for some users

```powershell
$users = @("user1@yourOrg.com", "user2@YourOrg.com")

# Will open a window for you to log in
Connect-ExchangeOnline
foreach ($user in $users) {
    New-App -Mailbox $user -Url "link to your manifest file"
}
```