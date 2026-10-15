# GSEA English Learning Portfolio

A formative assessment portfolio application designed with the signature GSEA Semester 1 Flight-Deck aesthetic, precision CAD textured header, and clean institutional canvas.

## Project Structure

```
├── src/
│   └── input.css                    # Tailwind input stylesheet with @source directives
├── output.css                       # Compiled & minified Tailwind CSS output
├── index.html                       # Main portfolio application landing page
├── imsmanifest.xml                  # IMS Content Package specification manifest for Moodle
├── GoogleAppsScript_Backend.js      # Backend script for PIN verification & reading logs
├── package.json                     # NPM configuration & build scripts
└── moodle_portfolio_package.zip     # Ready-to-upload Moodle IMS Content Package
```

## Backend Architecture: Zero-Lock Concurrency & PIN Protection

To eliminate Google Sheets write-lock collisions when 20 students submit at the exact same moment:

1. **Submissions (Write Path)**:
   - Submissions route directly to **Google Forms** via its high-throughput `formResponse` endpoint.
   - Google Forms absorbs concurrent traffic smoothly without lock contention and writes directly to your linked Google Sheet.
   - The student's 4-digit PIN is submitted alongside each session entry.

2. **Authentication & Portfolio Logs (Read Path)**:
   - "Load My Portfolio" performs a read-only `GET` query to your deployed Google Apps Script (`SCRIPT_URL`).
   - Apps Script verifies the student's 4-digit PIN against the `Students` sheet.
   - If verified, it fetches and displays the student's logged history from the `Form Responses 1` sheet.
   - Read-only operations never encounter write locks.

---

## Google Forms Setup Instructions

### Step 1: Create the Google Form
Create a new Google Form with the following 11 questions:
1. **Student Name** (Dropdown or Short Answer)
2. **4-Digit Security PIN** (Short Answer)
3. **Module** (Dropdown or Multiple Choice)
4. **Session Number** (Dropdown or Multiple Choice)
5. **Key vocabulary learnt** (Paragraph)
6. **Lesson content review (What I learnt in this session)** (Paragraph)
7. **What I need to work on following this session** (Paragraph)
8. **How will this help me with my formative evaluation?** (Paragraph)
9. **How well prepared do you think you are for your module 1 formative evaluation? (Rate out of 10)** (Linear Scale 1-10 or Short Answer)
10. **What I have done for TOEIC prep this week** (Short Answer)
11. **How well prepared do you think you currently are for your TOEIC test? (Rate out of 10)** (Linear Scale 1-10 or Short Answer)

### Step 2: Link Google Form to Your Spreadsheet
1. In your Google Form, open the **Responses** tab.
2. Click **Link to Sheets** (the green Sheets icon).
3. Select **Select existing spreadsheet** and choose your spreadsheet containing the `Students` sheet (Column A: Student Name, Column B: PIN).
4. Google Forms will create a sheet titled `Form Responses 1`.

### Step 3: Get Form Action URL & Entry IDs
1. In Google Form, click the three vertical dots (⋮) in the top-right corner.
2. Click **Get pre-filled link**.
3. Fill in dummy answers for each field (e.g. `NAME`, `PIN`, `MODULE`, `SESSION`, etc.).
4. Click **Get link**, then **Copy link**.
5. Paste the link into a text editor. It looks like:
   ```
   https://docs.google.com/forms/d/e/1FAIpQLSc.../viewform?usp=pp_url&entry.1234567=NAME&entry.2345678=PIN...
   ```
6. **Action URL**: Change `viewform` at the end to `formResponse`:
   ```
   https://docs.google.com/forms/d/e/1FAIpQLSc.../formResponse
   ```
7. Note down the `entry.XXXXXXX` numbers corresponding to each question.

### Step 4: Configure `index.html`
Open `index.html` and update lines 606–623:
```javascript
const GOOGLE_FORM_ACTION_URL = "https://docs.google.com/forms/d/e/YOUR_FORM_ID/formResponse";

const GOOGLE_FORM_ENTRIES = {
  student: "entry.1234567",
  pin: "entry.2345678",
  module: "entry.3456789",
  session: "entry.4567890",
  vocab: "entry.5678901",
  contentReview: "entry.6789012",
  workOn: "entry.7890123",
  evalHelp: "entry.8901234",
  evalRating: "entry.9012345",
  toeicPrep: "entry.0123456",
  toeicRating: "entry.1122334"
};
```

### Step 5: Deploy the Google Apps Script Reader
1. In your Google Sheet, open **Extensions** > **Apps Script**.
2. Replace `Code.gs` with the content of `GoogleAppsScript_Backend.js`.
3. Click **Deploy** > **Manage deployments** (or **New deployment**).
4. Select **Web app** with:
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Ensure `SCRIPT_URL` in `index.html` points to this Web App URL.

---

## Build & Moodle Deployment

```bash
# Compile CSS
npm run build

# Package for Moodle (creates moodle_portfolio_package.zip)
npm run package:moodle
```

Upload `moodle_portfolio_package.zip` as an **IMS Content Package** activity in your Moodle course.
