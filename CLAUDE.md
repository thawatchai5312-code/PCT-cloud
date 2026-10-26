# PCT Cloud (Google Apps Script)

- Apps Script project: script ID in `.clasp.json`. Files: `รหัส.js` (server, shows as `รหัส.gs` in the editor), `Index.html` (UI), `appsscript.json`.
- After every code change, deploy immediately: `./deploy.sh "short description of the change"`.
  This runs `clasp push` and redeploys the existing web app deployment, so the `/exec` URL never changes.
- If clasp is not logged in (new session), run `clasp login --no-localhost` and ask the user to open the URL and paste back the redirected `localhost:8888` URL.
