# Master Sheet Setup

This is for development only. Do not replace the live app sheets yet.

Current dev web app URL:

`https://script.google.com/macros/s/AKfycbwRRSc5J2X3SYjJ_9pFDq7eoVwTICDICtyc8DUC_H-DIsgJHjnzmehbKrk7CCv8vZ0G/exec`

## File

Use:

`apps/scripts/master-data-script.gs`

## Create The Master Sheet

1. Create a new Google Sheet named `Pinaka Master Data DEV`.
2. Open `Extensions > Apps Script`.
3. Paste the full contents of `apps/scripts/master-data-script.gs`.
4. Save.
5. Run `setupMasterSheet`.
6. Deploy as Web App only after setup works.

## Tabs Created

- `Employees`
- `Vehicles`
- `Settings`
- `Audit Log`
- `Portal Apps`
- `Portal Users`
- `User PIN Manager`

## Prepared Imports

Generated from `employee data (1).xlsx`:

`data-import/master-employees-import.csv`

This file contains 197 employee rows mapped to the master `Employees` format, including:

- employee name
- mobile
- father/husband name
- DOB
- DOJ
- designation
- salary/CTC
- bank details
- app access flags where obvious from designation

Generated from `Diesel Report May-2026 .xlsx`:

`data-import/master-vehicles-import.csv`

This file contains 47 cleaned unique vehicle rows using the actual vehicle numbers from the diesel report, including:

- vehicle number
- vehicle type
- driver name
- fuel metric (`KM` or `HRS`)
- app enable flags for Fuel, Weighbridge, MRF, and Maintenance
- employee access flags for Fuel, Weighbridge, MRF, User Charge, Maintenance, and Joining
- ward mapping where already known from the MRF app

Obvious typo/variant duplicates from the diesel report were merged. Review them here:

`data-import/vehicle-number-alias-review.csv`

The two vehicle deployment PDFs do not contain extractable text in this environment. The diesel report is now the primary vehicle import source.

## API Actions

GET:

- `?action=debug`
- `?action=employees`
- `?action=vehicles`
- `?action=config`
- `?action=config&app=fuel`
- `?action=config&app=weighbridge`
- `?action=config&app=mrf`
- `?action=config&app=usercharge`
- `?action=config&app=maintenance`

POST:

- `upsertEmployee`
- `upsertVehicle`
- `bulkUpsertEmployees`
- `bulkUpsertVehicles`

## Migration Rule

Keep Fuel, Weighbridge, User Charge, MRF, Maintenance, and Staff Joining logs in their own app sheets for now.

Only master data moves here:

- employee/user list
- app access
- PINs
- vehicles
- drivers
- wards
- fuel efficiency settings

## Deployment Order

1. Redeploy the master data Apps Script first.
2. Redeploy each dev app script after that: Fuel, Weighbridge, User Charge, MRF, Maintenance, and Staff Joining.
3. Fill missing master login data in the `Portal Users` tab:
   - One row per user.
   - Give each user a different `PIN`.
   - Tick only the apps that user can open.
   - Tick `Admin` only for full access.

The main index now uses Master Sheet users for all tiles. If no active PIN users exist in the master sheet, apps stay locked.

## Portal Apps Tab

Use this tab to decide which apps need a PIN.

Columns:

- `App Key`
- `App Name`
- `PIN Required` checkbox
- `Status`

Tick `PIN Required` when the app should ask for a PIN.

Untick `PIN Required` when the app should open directly without a PIN.

Default:

| App | PIN Required |
| --- | --- |
| Fuel | checked |
| Weighbridge | checked |
| User Charge | checked |
| MRF | checked |
| Staff Joining | unchecked |
| Maintenance | unchecked |

## Portal Users Tab

Use this tab for all portal user/PIN management. Do not use the old `User PIN Manager` tab; it is kept only for backwards compatibility.

Columns:

- `Login ID`
- `User Name`
- `PIN`
- `Active` checkbox
- `Admin` checkbox
- `Fuel` checkbox
- `Weighbridge` checkbox
- `User Charge` checkbox
- `MRF` checkbox
- `Staff Joining` checkbox
- `Maintenance` checkbox
- `Site`
- `Role`

Example:

| Login ID | User Name | PIN | Active | Admin | Fuel | Weighbridge | User Charge | MRF | Staff Joining | Maintenance |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| owner | Owner | 4321 | checked | checked | checked | checked | checked | checked | checked | checked |
| fuel-manager | Fuel Manager | 1111 | checked | unchecked | checked | unchecked | unchecked | unchecked | unchecked | unchecked |
| weighbridge | Weighbridge Incharge | 2222 | checked | unchecked | unchecked | checked | unchecked | unchecked | unchecked | unchecked |

After editing this tab, run:

`Master Data > Sync Portal Users`

Each active user must have a unique PIN. Duplicate PIN rows are ignored until you fix them.

## Add Users Or Change PINs

After redeploying the master data script, reload the master Google Sheet.

Use:

`Master Data > Open Portal Users - PIN & Access`

Then:

1. Add one row per user.
2. Enter `Login ID`, `User Name`, and unique `PIN`.
3. Keep `Active` checked.
4. Tick the app checkboxes this user can open.
5. Tick `Admin` only for full access.
6. Run `Master Data > Sync Portal Users`.

The old `User PIN Manager` sheet is not needed anymore.

## Default App PINs

The Master Data menu also has:

`Master Data > Set Default App PINs`

This creates or updates these rows in `Portal Users`:

| App | Login ID | PIN |
| --- | --- | --- |
| Fuel Manager | `fuel-manager` | `1111` |
| Weighbridge | `weighbridge` | `2222` |
| User Charge Collection | `uc-collector` | `3333` |
| MRF Supervisor | `mrf-supervisor` | `4444` |

To change a PIN later, edit the `PIN` cell in the `Portal Users` tab. Keep every active PIN unique.

After the Master Sheet has an active Portal Users row for an app, the hosted portal uses the Sheet PIN for that app. The built-in `.env` PINs are only first-time fallback defaults.
