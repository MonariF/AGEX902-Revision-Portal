# AGEX 902 Advanced Econometrics Revision Portal

**Lecturer:** Dr. Fred Nyamitago Monari, PhD  
**Institution:** Kisii University, Kenya  
**Hosting:** Google Apps Script + Google Sheets

Interactive revision portal source code for AGEX 902. GitHub stores the project source; Apps Script hosts the running web app.

## Features
- Ten topic modules with revision notes and step-by-step solution guidance.
- Theory multiple-choice questions and numerical practice marked to 2 decimal places.
- Student login using name, registration number and email.
- Lecturer monitoring dashboard with student progress and recent attempts.
- Google Sheets data store with Students, Attempts and Activity tabs.

## Modules
1. Matrix Algebra and OLS
2. GLS and Heteroscedasticity
3. Panel Data — Fixed Effects
4. Instrumental Variables
5. Logit, Probit and Tobit
6. Time Series — Stationarity and Unit Roots
7. Simultaneous Equation Models
8. Difference-in-Differences and Propensity Score Matching
9. Diagnostics
10. Spatial and Quantile Regression

## Deploy to Google Apps Script
1. Create a standalone project at https://script.google.com/.
2. Add a script file named Code.gs and paste the repository Code.gs.
3. Add an HTML file named Index and paste Index.html.
4. Run setupPortal() once and authorize the requested Google permissions. This creates the spreadsheet and stores its ID in Script Properties.
5. In Project Settings, open Script Properties and set AGEX902_LECTURER_KEY to a strong, unique secret. Do not commit this secret to GitHub or share it with students.
6. Deploy as a Web app with access settings appropriate for your class.
7. Test with dummy student accounts before sharing the web app URL.

## Student data and privacy
Student names, registration numbers, email addresses and assessment activity are stored in the Google Sheet owned by the deploying account. Restrict access to authorized staff. Never commit student records, spreadsheet IDs, passwords, tokens or deployment-specific secrets to this public repository.

## Important
This is a starter implementation. Test with dummy records and review all questions and worked examples against the approved course notes before using it for formal assessment.
