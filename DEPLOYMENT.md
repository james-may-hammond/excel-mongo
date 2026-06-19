# Simple Deployment: Railway + Vercel

Use:

- **Railway** for the Python/FastAPI backend.
- **Vercel** for the Excel taskpane frontend.
- **Microsoft 365 admin center** to upload the generated Excel manifest.

## 1. Deploy Backend On Railway

1. Push this repo to GitHub.
2. Open Railway.
3. Create a new project.
4. Select **Deploy from GitHub repo**.
5. Choose this repo.
6. Railway will use the included `Dockerfile`.

The backend start command is already handled by Docker:

```bash
uvicorn be.main:app --host 0.0.0.0 --port $PORT
```

After Railway deploys, create a public Railway domain.

Example:

```txt
https://excel-mongo-api.up.railway.app
```

Check that this works:

```txt
https://excel-mongo-api.up.railway.app/health
```

Save this URL. This is your `ADDIN_API_BASE`.

## 2. Deploy Frontend On Vercel

1. Open Vercel.
2. Import the same GitHub repo.
3. Use these settings:

```txt
Framework Preset: Other
Build Command: npm run package
Output Directory: dist
Install Command: npm install
```

The included `vercel.json` already contains these settings.

Add these Vercel environment variables:

```txt
ADDIN_PUBLIC_URL=https://your-vercel-app.vercel.app
ADDIN_API_BASE=https://excel-mongo-api.up.railway.app
```

Replace:

- `https://your-vercel-app.vercel.app` with your real Vercel URL.
- `https://excel-mongo-api.up.railway.app` with your real Railway URL.

Deploy the Vercel project.

Check that this works:

```txt
https://your-vercel-app.vercel.app/fe/taskpane.html
```

## 3. Generate The Excel Manifest

Run this locally with your real URLs:

```bash
ADDIN_PUBLIC_URL=https://your-vercel-app.vercel.app \
ADDIN_API_BASE=https://excel-mongo-api.up.railway.app \
npm run package
```

This creates:

```txt
dist/manifest.xml
```

Validate it:

```bash
ADDIN_PUBLIC_URL=https://your-vercel-app.vercel.app \
ADDIN_API_BASE=https://excel-mongo-api.up.railway.app \
npm run package:validate
```

## 4. Upload To Microsoft 365

For private/internal use:

1. Open Microsoft 365 admin center.
2. Go to **Settings** -> **Integrated apps**.
3. Select **Upload custom apps**.
4. Upload:

```txt
dist/manifest.xml
```

5. Assign the add-in to your users or groups.

After that, users should see the add-in inside Excel.

## Important Notes

- Do not upload `fe/manifest.xml` for production.
- Upload only `dist/manifest.xml`.
- Railway URL goes in `ADDIN_API_BASE`.
- Vercel URL goes in `ADDIN_PUBLIC_URL`.
- Do not use `cert.pem`, `key.pem`, or `mkcert` in production.
- If either deployment URL changes, regenerate `dist/manifest.xml`.
