# TOW Apartment Stay System

A paperless booking and logging system for TOW's 4 church apartment units. It runs on GitHub Pages, a Google Sheet and one Apps Script.

| Page | Who uses it | What it does |
|---|---|---|
| `index.html` | Church Admin (passcode) | Inbox of new requests, assign units, WhatsApp acknowledgement to the Ministry Head, guest speakers, Send to Accounts, PDFs, Excel export |
| `head.html` | Ministry Heads (no login; personal passcode to approve) | Check availability and approve a stay. Their passcode stamps their name on the member's form |
| `request.html` | Members (opened from the Ministry Head's link) | Fill in, agree to the terms, pledge a love gift and sign. The Ministry Head's approval is stamped on automatically |
| `hosting.html` | PIC / Hosting Pastor (opened from the admin's link) | Guest Speaker Hosting Checklist: apartment, day trip or staying elsewhere |

**What happens automatically**
- Each request is saved to the Google Sheet.
- A filled PDF backup goes to the Drive folder **TOW Apartment Records**.
- An alert goes to **Telegram and email**.
- Confirmed stays **from the day you start using this system** are added to the **Church Space & Events Scheduling Portal** calendar, and updated or removed if they change. Past stays and stays imported from the old log are never added, so nothing already in your calendar is duplicated. The script only edits events it created itself.

**Privacy**
- No guest data or phone numbers are stored in this repository.
- Ministry Heads' WhatsApp numbers live only in the sheet's **Ministries** tab, and only the admin page (after the passcode) can read them.
- Member and hosting links each carry a private token. A member link works once only.

---

## Setup (one time, about 15 minutes)

### 1. Load the data into your Google Sheet
1. Open the Apartment Stay Log Google Sheet and go to **File → Import → Upload** `TOW-Apartment-Stay-Log.xlsx`.
2. Choose **Insert new sheet(s)**. This adds the Stay Log (18 past stays), Summary, Hosting, Approvals, Ministries and Lists tabs.
3. Check the **Ministries** tab. You can edit heads and numbers there any time. The "Accounts" row is used for **Send to Accounts**.
4. **Don't upload the .xlsx file to GitHub.** It contains guest details.

### 2. Add the script
1. In the sheet, go to **Extensions → Apps Script**. Paste in all of `Code.gs` and save.
2. Run `setup` and allow the permissions it asks for (Sheets, Drive, Calendar, Gmail, external requests).
3. Edit `setPasscodes()` with your own admin code and alert email, then run it. The default admin code is `tow-admin`.
4. Each Ministry Head gets their own 6-digit passcode automatically. They appear in the admin page under **⚙ Admin → Ministry Heads & passcodes**, with a button that sends each head the link and their code on WhatsApp.

### 3. Telegram alerts (about 2 minutes)
1. In Telegram, message **@BotFather** → `/newbot` → name it (e.g. *TOW Apartment Alerts*) → copy the **token**.
2. Open your new bot and send it "hi". To use an admin group instead, add the bot to the group and say "hi" there.
3. In Apps Script, paste the token into `setupTelegram()` and run it. You should get "✅ alerts connected" in Telegram.
4. Run `testAlert` to check that both Telegram and email work.

### 3b. Daily jobs: backup and guest speaker reminders (1 minute)
In Apps Script, run `setupDailyJobs` once. It turns on:
- **Backup (about 2am):** a full Excel copy of the sheet goes to **TOW Apartment Records → Backups** in Drive. The latest 30 are kept.
- **Guest speaker reminders (about 8am):**
  - **14 days before:** the PIC gets the hosting link by email, and you get a Telegram + email alert with a WhatsApp message ready to forward.
  - **3 days before:** you get a reminder to prepare the apartment (if the guest is staying), plus a list of missing details. The PIC is emailed what's still missing.

### 4. Deploy the script as a web app
1. Go to **Deploy → New deployment → Web app**. Set **Execute as** to *Me* and **Who has access** to *Anyone*. Copy the URL (it ends in `/exec`).
2. Paste that URL into `config.js`.

### 5. Publish on GitHub Pages
1. Create a repository, e.g. `tow-apartment-log`. Upload everything in this folder **except** the .xlsx file.
2. Go to **Settings → Pages**, choose *Deploy from a branch*, then *main / root*.
3. Your site will be at `https://<username>.github.io/tow-apartment-log/`.
4. In `Code.gs`, set `SITE_URL` to that address and save. Then go to **Deploy → Manage deployments → Edit → New version → Deploy**, so the links in alerts point to your site.

### 6. Share
- **Ministry Heads:** in **⚙ Admin**, tap **💬 Send link & code** next to each head.
- **Admin:** open `…/index.html` with the admin passcode.

---

## Daily workflow

**Paid stay**
1. The Ministry Head checks availability on `head.html`, then taps **Approve & create link** and sends it to the member on WhatsApp.
2. The member fills in the form and signs → you get a Telegram + email alert with the PDF.
3. In **Inbox**, open the request, type the unit (free units are shown) and tap **Save**. The status becomes *Confirmed* and the stay goes on the calendar.
4. Tap **💬 Acknowledge to [Ministry Head]**. WhatsApp opens to his number with the confirmation, check-in/out notes and love gift details. He forwards it to the member.

**FOC guest speaker**
1. When you get the quarter's schedule, tap **+ Guest speaker** for each one: the guest, preaching date, whether they need the apartment, planned stay dates (optional), and the PIC's WhatsApp and email. Use **Save & add another** to go quickly. Each one goes on the calendar straight away.
2. **14 days before**, the reminders described in step 3b go out automatically. You can also send the link any time with **💬 Send link to PIC**.
3. The PIC checks with the guest and fills in the checklist (apartment, day trip or elsewhere). They can save and come back, then submit.
3. You get an alert. If the guest is staying in the apartment, a linked stay request appears in **Inbox** for you to assign a unit.
4. Tap **💼 Send to Accounts** to send the love gift and honorarium bank details to Accounts on WhatsApp.

**Backup and records**
- **🖨 PDF** on any record opens its PDF in Drive.
- **⬇ Excel** downloads the Stay Log and Hosting records.
- **Reports** tab: pick a year, tap 👁 to hide any section you don't need, then **🖨 Print / Save as PDF**. Hidden sections are left out of the printout.
- The **Summary** tab in the sheet shows stays, nights and love gifts by unit, year, category, ministry and month.

## After changing Code.gs
Go to **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy**. The URL stays the same.
