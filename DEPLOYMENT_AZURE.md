# Azure App Service Deployment

## 1. Create GitHub repository
1. Create private repository under organization `Myridius-India`.
2. Push this project to branch `main`.
3. Enable branch protection and required PR review.

## 2. Provision Azure resources
1. Resource Group
2. App Service Plan (Linux)
3. Web App for Containers
4. Application Insights
5. Key Vault

## 3. Configure app settings
Set these application settings in App Service:
- `RAS_DATA_SOURCE=sharepoint`
- `RAS_AUTH_MODE=sso`
- `RAS_AUTH_SESSION_MINUTES=60`
- `RAS_SSO_CLIENT_ID`
- `RAS_SSO_TENANT_ID`
- `RAS_SSO_ALLOWED_DOMAIN`
- `RAS_SSO_SCOPES=User.Read`
- `RAS_SP_TENANT_ID`
- `RAS_SP_CLIENT_ID`
- `RAS_SP_CLIENT_SECRET` (use Key Vault reference)
- `RAS_SP_SITE_ID`
- `RAS_SP_DRIVE_ID`
- `RAS_SP_FOLDER_PATH`
- `RAS_SP_TEMPLATE_FILE`

## 4. Configure GitHub Actions secrets
Repository or environment secrets:
- `AZURE_CREDENTIALS`
- `AZURE_WEBAPP_NAME`
- `AZURE_WEBAPP_RESOURCE_GROUP`

## 5. Deploy
1. Push to `main` for CI validation.
2. Trigger workflow `deploy-azure-appservice` manually for first release.
3. Validate app URL, sign-in, and SharePoint data sync.

## 6. Post-deploy checks
1. Verify only organization users can sign in.
2. Confirm monthly files and template are pulled from SharePoint.
3. Confirm Dashboard and Summary metric pivots for GM%, Revenue, and Cost.
4. Enable App Service logs + alerts.
