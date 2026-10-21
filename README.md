# Trail Camera Reviewer

A Google Apps Script web app for classes that review trail-camera photos stored in Google Drive. Students sign in with their name and the last 4 digits of their student ID, claim a camera, step through its photos and tag the species. Tagged photos are renamed and moved into a "Tagged Photos" folder, and every tag is logged in a Google Sheet.

It runs entirely in **your own Google account**. There is no server and no cost.

## Features

**Students**
- Claim one camera at a time, flip through photos and videos with the keyboard, tag one or several species per photo, undo, and resume later.
- Quick-key species (keys 1–6) with a **Pin** button to lock favourites in place.
- Review or fix tags already made: change or remove a tag from the "Tagged photos" view.
- **View tagged photos** for any completed camera: search, sort, star favourites, click **Nahhh** if an ID looks wrong and suggest another species.

**Instructors** (marked "yes" in the Instructor column of the Student Names tab)
- Officially change a tag from the viewer; it updates the file name, Master Data and everything else.
- **Instructor login** on the first screen opens two extra tabs:
  - **Camera setup** – edit camera data, add cameras, release a reviewer, mark complete/reopen, and choose each camera's (or site's) photo and tagged folders from an in-app folder browser. Nothing is saved until you press *Save changes*; every change is written to a **Setup Change Log** tab.
  - **Review cameras** – open any camera for review, including claimed and completed ones, without claiming it.

## What you need

- A Google account (a school/Workspace account works well).
- Camera photos or videos (png, jpg, mp4, avi, mov) in Google Drive.

## Setup (about 15 minutes)

1. **Make the sheets.** In Google Drive choose *New > File upload* and upload both files from the `template/` folder, then open each and choose *File > Save as Google Sheets*:
   - `Camera Review Sign Ups (template).xlsx` → rename it to exactly **Camera Review Sign Ups**
   - `species (template).xlsx` → rename it to exactly **species** (it stays a separate file; lines starting and ending with `|||` are group headings such as `|||MAMMALS|||`)
2. **Fill in the sheet** (see [Sheet layout](#sheet-layout)): your students in *Student Names*, your cameras in *Data*. Delete the example rows.
3. **Add the script.** In the *Camera Review Sign Ups* sheet choose *Extensions > Apps Script*, then:
   - Replace the contents of `Code.gs` with `src/Code.gs`.
   - Add an HTML file named `Index` and paste `src/Index.html`.
   - Click the gear (*Project Settings*), tick **Show "appsscript.json" manifest file**, and replace its contents with `src/appsscript.json` (change `timeZone` if you like). This also switches on the **Drive** advanced service. If it does not, use *Services (+)* and add **Drive API v3**.
4. **Mark yourself instructor.** In the editor choose `setupInstructorColumn` from the function dropdown and press *Run* (approve the permissions when asked). Then put `yes` in the Instructor column beside every instructor's name.
5. **Deploy.** *Deploy > New deployment > Web app*. Execute as **Me**, Who has access **Anyone** (or *Anyone within your domain*). Copy the web app URL and give it to students. After any later code change use *Deploy > Manage deployments > Edit > New version*.
6. **Point the app at your photos.** Open the web app, press **Instructor login**, go to **Camera setup**, and for each camera (or each site) press **Choose…** to pick the folder of photos, and optionally a folder for the tagged copies. If you leave a camera blank the app falls back to its automatic folder search (see below). Use **Check** on a row to confirm what the app found.
7. Optional: run `setupFeedbackColumns` once from the editor to create the star/Nahhh columns and the Photo Feedback tab (they are also created automatically on first use).

## Sheet layout

**Student Names**

| A | B | C |
|---|---|---|
| `Fall 2026` (cell A1: the **active semester**) | `PIN` | `Instructor` |
| Student name | last 4 of student ID | `yes` for instructors |

Only cameras whose Semester matches A1 appear for students. `Fa26`, `Sp27` style abbreviations are understood.

**Data** (one row per camera deployment)

| Col | Header | Notes |
|---|---|---|
| A | Site | |
| B | Semester | must match Student Names A1 to be visible to students |
| C | Camera | e.g. `Cam 3` |
| D | Deployment | date, `M/D/YYYY` |
| E | Collection | date, `M/D/YYYY` |
| F, G | Latitude, Longitude | |
| H | Brand | |
| I | Reviewer | filled in by the app when a student claims the camera |
| J | How to use | optional, never touched by the app |
| K | Complete | `yes` when finished |
| L | Unique ID override | leave blank |

The camera's identity is `Camera_Deployment_Collection`. Progress, tags and folder choices are keyed to it, so the app locks Site, Camera, Semester and the dates in Camera setup once a camera has a reviewer, is complete or has tagged photos.

**Master Data**, **Photo Feedback**, **Camera Folders** and **Setup Change Log** are written by the app. Do not rename these tabs.

## How the app finds photos

1. A folder you chose in **Camera setup** (camera-level first, then site-level). This is the recommended, portable way.
2. Otherwise an automatic search: the sheet's folder → a folder named for the semester → a site folder → `CAMERA PHOTOS TO REVIEW` → the camera's folder, matched by camera number and dates.

Tagged photos go to `<camera>_<site>_Tagged Photos` in the chosen tagged folder (or beside the photos, as before) and are named `site_species_cameraID_GPS_originalname`.

## Privacy and security

- Students must be signed in (name + PIN) before the app returns any photo or token. PIN guesses are throttled (5 tries per 15 minutes).
- The app runs as the account that deployed it, so that account needs access to the photo folders. Students do not need Drive access.
- Student PINs live in your sheet. Keep the sheet private, and **never commit real rosters to a public repository**.
- Instructor-only functions check the signed-in name on the server, not just in the browser.

## Troubleshooting

- *"Could not locate media folder"* – choose the folder in Camera setup, or check the folder names/dates for the automatic search. Run `clearCaches` after moving or renaming folders.
- *"species sheet not found"* – the species file must be a Google Sheet named exactly `species` owned by or shared with the deploying account.
- Run `diagnoseSite('YOUR_SITE')` in the editor and read *Executions/Logs* to see what the folder search finds.
- Students see an old version – redeploy with *New version*.

## Updating with clasp (optional)

```bash
npm install -g @google/clasp
clasp login
clasp clone <SCRIPT_ID> --rootDir src     # or: clasp create --rootDir src
clasp push
```
`.clasp.json` is git-ignored because it holds your script ID.

## License

MIT – see `LICENSE`.
