# LexCitr setup (Word for Mac)

## 1. Database (Supabase)
1. Create a free account and a new project at supabase.com. Choose an EU region.
2. Open SQL Editor > New query, paste all of `schema.sql`, and click Run.
3. Open Authentication > Sign In / Providers > Email and turn off "Confirm email".
4. Open Project Settings > API and copy the Project URL and the publishable (anon) key.
   Never use the service_role or secret key.
5. Put both values in `config.js`.

## 2. Website (GitHub Pages)
1. Create a free GitHub account and a public repository named `lexcitr`.
2. Choose "uploading an existing file" and upload everything in this folder, keeping the
   `assets` and `icons` folders. Commit.
3. Open Settings > Pages. Set the source to "Deploy from a branch", the branch to `main`, the folder to `/ (root)`, and save.
4. After a minute or two the site is at `https://YOUR-USERNAME.github.io/lexcitr/`.
5. Open it and create your account. Then, in Supabase, turn off new sign-ups
   (Authentication > Sign In / Providers > "Allow new users to sign up").

## 3. Word add-in
1. In `manifest.xml`, replace every `__SITE_URL__` with `https://YOUR-USERNAME.github.io/lexcitr` (no trailing slash).
2. Quit Word.
3. In Finder choose Go > Go to Folder and paste:
   `~/Library/Containers/com.microsoft.Word/Data/Documents/wef`
   If the folder does not exist, create a folder named `wef` inside `Documents` at that location.
4. Copy `manifest.xml` into it.
5. Open Word and a document. Look for LexCitr Footnote on the Home tab. If it is not there,
   check Insert > Add-ins > My Add-ins (the small dropdown arrow).

## Using it
Put your cursor where the footnote belongs, click LexCitr Footnote, choose or create a citation,
add a pinpoint if needed, and click Insert.

## If the add-in does not appear
- University accounts can have add-ins turned off. In Word, check Word > Preferences > Privacy.
  If optional connected experiences are off, turn them on if you are allowed to. If your university blocks them, ask IT.
- Make sure Word is up to date. The add-in needs Word API 1.5.

## Moving existing data
Use Back Up on the old site, then Back Up > Restore on the new one.
