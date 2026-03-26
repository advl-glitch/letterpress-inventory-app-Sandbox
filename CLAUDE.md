# CardStockApp — Admin Dashboard

## Owner
Angel Verde — "Prints by Angel" letterpress card business

## Tech Stack
- Vanilla JS single-page app (no framework)
- Google Apps Script (GAS) REST backend with POST-based routing (`action` field)
- Google Sheets as database, Google Drive for photo storage
- Vercel hosting (auto-deploys from GitHub `advl-glitch/cardstock-app-admin`)

## Key Files
- `cardstock-app.js` — entire admin frontend
- `CardStockApp-Script_code.gs` — GAS backend (local copy; paste into GAS editor & redeploy)

## GAS Deployment
- Deployment ID: `AKfycbzM4FR2fzJ9LAE7w1R-J9u0-6V6HSBP5om2NsX59D8-fCH045rVvLvKMZBia8PX5zME`
- To deploy: paste .gs into GAS editor → Manage deployments → new version
- Sheet ID: `1FiDZXPV6aimKpKUvzDCQczq01nCdvMZBia8PX5zME`

## Architecture Patterns
- Header-based column mapping in GAS (not position-based appendRow)
- Google Drive photos use `lh3.googleusercontent.com/d/{fileId}` format
- `fixPhotoUrl()` converts old Drive URLs at display time (3 formats supported)
- `handlePhotoUrlPaste()` auto-converts pasted Drive URLs + shows preview
- All forms hide submit button during save, show status bar
- Dog mascot (`doggycutetonguev2.png`) for loading/empty/error states
- Phone fields use `formatPhoneField()` for auto-formatting

## App Pages
1. Dashboard/Home
2. Main Inventory — card/list view, detail panel, edit modal with status (Open/Limited/Retired)
3. Add New Design — auto-price from ProductType, auto-format tag, "+ Add New Type" option, Open/Limited status
4. Print Runs
5. Retail Partners — add/edit with modals
6. Retail Stock & Sales — hide 0-stock cards, auto-price
7. Orders — pending on top, fulfilled grouped by year (collapsible), undo fulfillment
8. Vending Machines — photo upload, phone formatting
9. Tags — manage tags, assign to designs
10. Inventory Audit — retired designs in collapsible section at bottom

## Sheet Columns
- Items: Status column (Open/Limited/Retired) — NOT Active
- ProductType: TypeID (auto-generated, e.g. "LalaCards"), TypeName, DefaultRetailPrice
- Tags/Partners still use Active column (works fine, just different naming)

## Important Notes
- Format tags are auto-applied when selecting Card/Item Type (hidden from tag pills in Add Design)
- Add Design page checks `sessionStorage('pba_current_page')` to avoid hijacking navigation
- Three separate photo upload handlers (edit design, add design, machines) — intentionally not consolidated
