# Tail Trails: Pet Care Website + Automatic Client System

| File | What it is |
|---|---|
| `index.html` | The website: interactive dog/cat, quick price finder, services, sign-up form |
| `images/` | Put your own photos here (optional, see below) |
| `backend/Code.gs` | The automatic back office (runs free in your Google account) |
| `backend/appsscript.json` | Settings file for the back office |

**How it works:** a customer signs up on the site. The sign-up is saved to your Google Sheet, you get an instant email, and the customer gets a confirmation. Every morning at 7 am you get a "Tail Trails today" email listing who to contact.

---

## Part 1: Website on GitHub

1. Go to github.com, click **+**, then **New repository**. Name it `tail-trails`, choose **Public**, and click **Create**.
2. Click **uploading an existing file**.
3. Open the unzipped `tail-trails` folder. **Select the files INSIDE it** (`index.html`, `README.md`, `backend`, `images`) and drag those in. Don't drag the `tail-trails` folder itself. Click **Commit changes**.
4. Open **Settings**, then **Pages**. Set Branch to **main** and the folder to **/ (root)**, then click **Save**.
5. After 2 minutes your site is live at `https://YOUR-USERNAME.github.io/tail-trails/`

## Part 2: Google Sheet back office

6. Go to sheets.google.com and create a blank sheet. Name it **Tail Trails Clients**.
7. Click **Extensions**, then **Apps Script**. Delete the sample code and paste all of `backend/Code.gs`.
8. Change `YOUR_EMAIL@gmail.com` to your email, and add your phone in `BUSINESS_PHONE` if you have one. Click 💾 **Save**.
9. Click ⚙️ **Project Settings** and tick **Show "appsscript.json"**. Back in the editor, open `appsscript.json` and replace its contents with `backend/appsscript.json`. Save.
10. Choose **setup** in the dropdown and click **▶ Run**. Allow the permissions (**Advanced**, then **Go to … (unsafe)**, then **Allow**). It says "unsafe" only because you wrote the script yourself.
11. Refresh the sheet. You'll see a **Clients** tab and a **Tail Trails** menu.

## Part 3: Connect the form

12. In Apps Script, click **Deploy**, then **New deployment**. Click ⚙️ and choose **Web app**.
13. Set **Execute as** to **Me** and **Who has access** to **Anyone**, then click **Deploy**. Copy the URL that ends in `/exec`.
    *"Anyone" only lets people send the form. Nobody can see your sheet.*
14. On GitHub, open `index.html` and click ✏️. Click inside the code, press **Ctrl+F** and search `PASTE_YOUR`. Replace it with your `/exec` link and keep the quotes.
    Optionally fill in `phone`, `phoneDisplay`, `email`, and `overnightPrice` (e.g. `55`).
15. Click **Commit changes**.

## Part 4: Test

16. Open your site and sign up with your own details. Check that a row appears in the sheet and that you got both emails.
17. In the sheet, click **Tail Trails**, then **Send me today's to-do email now**. Then delete the test row.

---

## Daily routine

Read the 7 am email, contact each person, and change the **Status**:

| Status | What happens automatically |
|---|---|
| New | On today's list until you contact them |
| No answer | Shows up again tomorrow |
| Meet & greet booked | Type the date in **Next follow-up**. It shows up that morning |
| Booked | Reminder to check in after 1 week |
| Regular client | Reminder to ask for a review after 3 weeks |
| Paused / Lost | Removed from the lists |

**Follow-up emails:** click a client's row, then **Tail Trails** → *first-week check-in* or *review request*. You see the email first, and it only sends when you click **Yes**.

## Adding your own photos (optional)

Put photos in the `images` folder with these exact names. Until you do, the site shows paw-print art.
- `cat-home-visit.jpg` (a cat at home)
- `winter-walk.jpg` (a dog in a winter coat)

Only use photos you took yourself or have permission to use.

## Changing things

- **Prices and services:** in `index.html`, edit the service cards **and** the `SERVICES` list near the bottom.
- **If you edit `Code.gs` later:** go to **Deploy** → **Manage deployments** → ✏️ → **New version** → **Deploy**.
- **Own domain** (e.g. tailtrails.ca): in GitHub **Settings** → **Pages** → **Custom domain**. The form keeps working.

## Check before going public

The page says caregivers are **background-checked, Pet First Aid trained, insured, GPS-tracked, get 20 hours of training**, and that there is a **backup team**. Only keep each claim if it's true. If it isn't yet, edit or delete those lines in `index.html`.
