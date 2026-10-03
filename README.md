# In-Out Register

A phone app for recording when employees come **IN** and go **OUT**. Tap a name, tap IN or OUT, and the phone's current time is saved. A reason for going out (Lunch, Bank and so on) is optional.

- Works offline. Entries are saved on the phone first.
- Backs up automatically to a Google Sheet with a dashboard, daily totals, a time-out list and a full log.
- Reports for today, this week, this month or any dates, with Excel download or share (for example, to WhatsApp).
- Add employees one at a time, many at once, or import them from Excel.
- Undo for 10 seconds after every tap, and any entry can be edited or a missed one added.

There is no server to run. The app is a set of static files, and Google runs the backup script for free.

---

## 1. Put the app online (GitHub Pages, free)

1. Create a new repository on GitHub, for example `inout-register`.
2. Upload everything in this folder **except** the `google-apps-script` folder (it is harmless to upload, but it isn't needed).
3. In the repository go to **Settings > Pages**. Under "Build and deployment", choose **Deploy from a branch**, pick `main` and `/ (root)`, then click **Save**.
4. After a minute the app is live at `https://<your-username>.github.io/inout-register/`.

**Is it safe for the repository to be public?** Yes. The repository holds only the app's code. No employee names, times, Google Sheet link or password are stored in it. All of that stays on the phone and in your Google Sheet. The site also tells search engines not to list it.

## 2. Create the Google Sheet backup (about 5 minutes, one time)

**Privacy:** do this while signed in to **your mom's Google account**, so she owns the sheet and nobody else can open it. Never use the Share button on it.

1. Go to <https://sheets.new> and name the sheet, for example **In-Out Register**.
2. Click **Extensions > Apps Script**. Delete everything in the editor, then paste the full contents of `google-apps-script/Code.gs`.
3. On the line `const SECRET = 'change-this-password';`, replace the text inside the quotes with a password of at least 10 characters (letters and numbers, not a name or birthday). Click **Save**.
4. Click **Deploy > New deployment**. Click the gear icon, choose **Web app**, and set:
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Click **Deploy**, then **Authorize access**. Choose your Google account. If you see "Google hasn't verified this app", click **Advanced**, then **Go to ... (unsafe)**. This is normal for your own script.
6. Copy the **Web app URL**. It ends in `/exec`.

> If you change `Code.gs` later, use **Deploy > Manage deployments > Edit (pencil) > Version: New version > Deploy** so the URL stays the same.

## 3. Set up your mom's phone

1. Open the GitHub Pages link in **Chrome**.
2. Tap the Chrome menu (three dots) and choose **Add to Home screen** (or **Install app**). It now opens like a normal app.
3. In the app, go to **Settings > Google Sheet backup**. Paste the Web app URL and the password, then tap **Save and test**.
4. Go to **Employees** and add the staff.
5. Optionally, turn on **Settings > Keep screen on**.

The badge at the top right shows the backup state: **Backed up**, **3 waiting** (offline, will send later) or **Backup failed**.

## Using it

**Every morning (one tap):** the card at the top of Home shows who is Present, on Half day, Absent or not marked yet.
- Tap **Mark everyone present** to mark all remaining staff present, with IN at the opening time (10:00 AM).
- Or, on each person's row, tap **Present** or **Absent**.
- Tap a **name** to choose **Half day**, or to set a different arrival time (for example 10:40 AM) before tapping Present.

**During the day:** each person who is present has an **OUT** and an **IN** button. Tap one and the time is recorded. You can choose a reason before tapping OUT (optional).

**At closing time you do nothing.** If someone went OUT and never came back, their time out is counted until closing (7:00 PM).

**Reports:** choose a period, then tap **Download Excel** or **Share Excel**.

### How the numbers are worked out
- **Time out** = from each OUT to the IN that follows it, cut off at closing time.
- **No IN after an OUT** (for example OUT at 4:45 PM): counted until closing, so **2h 15m**.
- **Half day:** leaving early is **not** counted as time out; the day is counted as a half day instead.
- **Late arrival** is just recorded as the IN time and is not counted as time out.
- **Absent** days are counted separately, and absent staff have no IN or OUT.
- Opening and closing times can be changed in **Settings > Office hours**.

## The Google Sheet

| Tab | What it shows |
|---|---|
| **Dashboard** | A month picker (change it and every number updates), totals, a summary for each employee (present, half days, absent, time out), a chart of hours out, and who is out right now |
| **Daily** | One row per employee per day: attendance (Present, Half day or Absent), first IN, last OUT, times out, time out, reasons |
| **Time Out** | Every OUT and when they came back (or closing time), with the duration and reason |
| **Punch Log** | Every single IN and OUT, newest first |
| **Employees** | The employee list |

The visible tabs are rebuilt after every backup, so **don't type in them**, because your changes would be overwritten. The raw data is in hidden tabs (`_punches`, `_employees`, `_attendance`). To restyle the sheet by hand, open Apps Script, choose `rebuildAll` and click **Run**.

## New phone or reset phone

Install the app as in step 3. Enter the same URL and password, then go to **Settings > Restore from Google Sheet**. Without the sheet, use **Settings > Save backup file** and **Load backup file**.

## Updating the app

Change the files, bump `VERSION` in `sw.js` (for example to `inout-2.0.1`) and upload again. The phone picks up the new version the next time the app is opened twice.

## Files

```
index.html             page shell
styles.css             design (light and dark)
app.js                 all app logic
sw.js                  offline support
manifest.webmanifest   install-as-app settings
icons/                 app icons
vendor/xlsx.full.min.js  Excel library (SheetJS)
google-apps-script/Code.gs  paste into the Google Sheet's Apps Script
```

## Who can see the data

- **The Google Sheet** can only be opened by the Google account that created it. Create it in your mom's account and don't share it.
- **The web app URL** has "Who has access: Anyone" because the phone app calls it without a Google sign-in. It is protected by the password: without the password it returns nothing, opening it in a browser shows nothing, and wrong guesses are slowed down. Keep the URL and password private; only your mom's phone needs them.
- **The app on the phone** stores data only on that phone.
- **GitHub** holds only the app's code, never any data.
