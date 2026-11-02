# Interactive Org Chart for Outlook

An Outlook add-in that lets you pick email recipients from your organization chart. Select someone and their managers can be added to Cc automatically.

There is no backend. The directory, favorites and recent contacts are stored in the browser's `localStorage`, so they stay on the device.

## Features

- Add people by hand, paste CSV text, or upload an Excel/CSV file
- Browse by department or as a reporting tree
- Add selected people to To and Cc while composing an email
- Favorites and recently used contacts

## Import format

Columns: `Name, Title, Email, Department, ReportsToEmail`

Name and Email are required. Rows without a valid email are skipped. `ReportsToEmail` also accepts the manager's name if it matches someone in the directory. See [public/template_excel.csv](public/template_excel.csv) for an example.

## Development

```sh
npm install
npm run dev     # https://localhost:3000
npm run build
npm run lint
```

`xlsx` is installed from the SheetJS CDN because the version on npm is outdated and has known vulnerabilities.

## Deployment

Pushing to `main` deploys to GitHub Pages through [.github/workflows/deploy.yml](.github/workflows/deploy.yml). [manifest.xml](manifest.xml) points at the deployed URLs and is the file to sideload into Outlook or submit to AppSource.
