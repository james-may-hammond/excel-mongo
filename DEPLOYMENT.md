# Deployment Guide

This project has three deployable parts:

- FastAPI backend: the API routes in `be/`.
- Excel taskpane frontend: the static files in `fe/`.
- Office manifest: generated as `dist/manifest.xml` and uploaded to Microsoft 365 or Partner Center.

The manifest does not contain the app code. It only tells Excel where the hosted taskpane lives.

## Option 1: Railway only, recommended first deployment

Use this when you want the simplest production setup. Railway runs the Docker container and serves both the API and the taskpane from one HTTPS URL.

### 1. Push the repo to GitHub

Railway deploys most easily from a GitHub repository.

### 2. Create the Railway service

1. Open Railway.
2. Create a new project.
3. Choose **Deploy from GitHub repo**.
4. Select this repository.
5. Railway should detect `railway.toml` and use the `Dockerfile`.

Railway will build the image and run:

```bash
uvicorn be.main:app --host 0.0.0.0 --port $PORT
```

Do not configure `cert.pem` or `key.pem` in Railway. Railway provides HTTPS at the public domain.

### 3. Generate a Railway public domain

In the Railway service settings, generate a public domain. It will look similar to:

```txt
https://excel-mongo-production.up.railway.app
```

Confirm these URLs open:

```txt
https://excel-mongo-production.up.railway.app/
https://excel-mongo-production.up.railway.app/health
https://excel-mongo-production.up.railway.app/fe/taskpane.html
```

### 4. Generate the production manifest

Run this locally, replacing the URL with your Railway URL:

```bash
ADDIN_PUBLIC_URL=https://excel-mongo-production.up.railway.app npm run package
```

This creates:

```txt
dist/manifest.xml
dist/fe/
```

For Railway-only deployment, the important file is:

```txt
dist/manifest.xml
```

### 5. Validate the manifest

```bash
ADDIN_PUBLIC_URL=https://excel-mongo-production.up.railway.app npm run package:validate
```

### 6. Upload the add-in

For private/internal deployment:

1. Open the Microsoft 365 admin center.
2. Go to **Settings** -> **Integrated apps**.
3. Choose **Upload custom apps**.
4. Upload `dist/manifest.xml`.
5. Assign it to the right users or groups.

For public Marketplace/AppSource deployment, submit `dist/manifest.xml` through Microsoft Partner Center.

## Option 2: Vercel frontend + Railway backend

Use this when you want Vercel to host the static Excel taskpane and Railway to host the Dockerized API.

Important: Vercel is not the right place to run this Dockerized FastAPI backend. In this setup:

- Railway hosts the backend API.
- Vercel hosts the generated `dist` static package.
- The generated manifest points Excel to Vercel for the taskpane and to Railway for API calls.

### 1. Deploy the backend to Railway

Follow Option 1 through the Railway public domain step.

Assume your Railway backend URL is:

```txt
https://excel-mongo-api.up.railway.app
```

Confirm:

```txt
https://excel-mongo-api.up.railway.app/health
```

### 2. Create the Vercel project

1. Open Vercel.
2. Import the same GitHub repository.
3. Set these project settings:

```txt
Framework Preset: Other
Build Command: npm run package
Output Directory: dist
Install Command: npm install
```

The included `vercel.json` already sets these values.

### 3. Configure Vercel environment variables

In Vercel, add these environment variables:

```txt
ADDIN_PUBLIC_URL=https://your-vercel-domain.vercel.app
ADDIN_API_BASE=https://excel-mongo-api.up.railway.app
```

If you use a custom domain on Vercel, use the custom domain for `ADDIN_PUBLIC_URL`.

### 4. Deploy Vercel

After deployment, confirm this opens:

```txt
https://your-vercel-domain.vercel.app/fe/taskpane.html
```

Vercel serves the contents of `dist` as the site root, so the browser URL keeps the `/fe/taskpane.html` path expected by the manifest.

### 5. Generate the manifest locally

Run this locally with the same URLs:

```bash
ADDIN_PUBLIC_URL=https://your-vercel-domain.vercel.app \
ADDIN_API_BASE=https://excel-mongo-api.up.railway.app \
npm run package
```

This creates a manifest at:

```txt
dist/manifest.xml
```

Upload this `dist/manifest.xml` to Microsoft 365 admin center or Partner Center.

Railway-only is still the least confusing first deployment, but the Vercel + Railway split is now supported by the generated package layout.

## Local Docker Test

Build the image:

```bash
docker build -t excel-mongo .
```

Run it:

```bash
docker run --rm -p 8000:8000 excel-mongo
```

Then open:

```txt
http://localhost:8000/health
http://localhost:8000/fe/taskpane.html
```

Local Excel sideloading still uses HTTPS through `run.sh`. The Docker test above is only to verify that the production container starts.

## Production Notes

- Production must use HTTPS.
- Do not use `cert.pem`, `key.pem`, or `mkcert` in production.
- Do not upload `fe/manifest.xml` for production. Upload `dist/manifest.xml`.
- Regenerate `dist/manifest.xml` whenever the production URL changes.
- Increase `package.json` `version` before publishing manifest changes.
- Railway-only is the recommended first deployment because the existing FastAPI app already serves `/fe`.
