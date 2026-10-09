# CareerLift AI — free prototype

This is a **working local web prototype**, not a deployed public website.

## Run
1. Install Python 3.10+.
2. In this folder: `pip install -r requirements.txt`
3. Run: `python app.py`
4. Open http://127.0.0.1:5000

## Current functionality
- Upload .docx resume and paste a job description.
- Show estimated keyword overlap and missing terms.
- No account, payment, or AI API required.
- The resume is read in memory, not persisted by the app.

## Next development milestones
- Human-reviewed AI tailoring grounded only in actual experience.
- Download edited .docx while preserving existing styles and layout.
- Authentication, user data isolation, admin tools, job application tracker.
- Secure public deployment, rate limits, privacy policy, malware scanning, encrypted storage.

**Important:** The percentage is a simple keyword-overlap heuristic, not a verified ATS score. This starter does NOT yet automatically tailor or download resumes. Never run Flask debug mode on a public server.
