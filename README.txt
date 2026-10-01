OPLAN KALULUWA 2026 V14

Team Leader login fixed:
- Select Deployment Unit
- Enter unit password
- No username required

Admin login remains:
- Username + password

Backend fixes:
- Added default Team Leader passwords chounit1 ... chounit14
- Fixed Admin User Access save bug caused by role variable order
- Team Leader accounts created/updated from Admin are automatically tied to unit password

Deployment:
1) Replace Apps Script Code.gs with Code.gs
2) Save
3) Run upgradeOplanDatabaseV12() once
4) Deploy > Manage deployments > Edit > New version > Deploy
5) Replace GitHub/Vercel index.html with index.html

The public dashboard remains unauthenticated and does not expose patient details.
