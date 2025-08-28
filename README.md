# Google Ads Budget Tracker | Monthly Spend (MCC-Level)

## Overview
This Google Ads script creates pacing dashboards for accounts under a Google Ads Manager (MCC).  
For each client account, the script generates a Google Sheet with:

- Monthly budget pacing overview  
- Daily spend tables  
- Line graphs showing forecast vs. actual  
- Recommended new daily budgets to stay on track  

The output includes color-coded metrics (red/yellow/green) and a progress bar for quick triage.  
**Note:** This script focuses only on Account-level budgets (not campaigns or ad groups).

## Setup Instructions
1. In your MCC: **Tools & Settings → Bulk Actions → Scripts → New Script**  
2. Paste this file. Review the **CONFIG** section:
   - Configure per-account monthly budgets in the template sheet.  
   - Ensure email + sheet permissions are enabled for your MCC.  
3. Authorize and **Preview** to verify logs and generated Sheets.  
4. Run the script.  
5. Go into the Google Sheet **CONFIG** and add the monthly budget to the *Monthly Budget* column.  
6. Re-run the script.  
7. Schedule to run **daily** so pacing data stays fresh.  

## Author
[Sam Lalonde](https://www.linkedin.com/in/samlalonde/)  
---

## License
This project is licensed under the [MIT License](https://opensource.org/licenses/MIT).  
Free to use, modify, and distribute.
