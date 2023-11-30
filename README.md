# Deployment
A detailed guide on how to deploy this Microsoft Office add-in for your own organisation.

## Prerequisites
To deploy this setup you are going to need to setup a few Azure resources. This guide assumes you have a microsoft exchange organisation.

**Azure resources**

- Sql database
- App service for the api
- Storage account to host the relevant files

### Pipeline
For the pipeline you can use ours as an example, .github/workflows/deploy.yml, but feel free to change it arround if it dosn't entirely suit your needs.