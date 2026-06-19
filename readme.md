### Features
- Connect to MongoDB
- Select database
- Select collection
- Build filters from Excel
- Execute filtered queries
- Display query results in Excel
- Show "No Results Found" when empty
- Insert selected Excel rows into MongoDB
- Update existing MongoDB records from Excel
- Refresh query results
- Connection health check
### Architecture Diagram
![Diagram](archdia.png)

---

### Installation

#### Prerequisites

| Requirement | Version | Notes |
|---|---|---|
| Python | 3.10 + | [python.org](https://python.org) |
| Node.js | 18 + | [nodejs.org](https://nodejs.org) |
| MongoDB | Any | Local or Atlas connection string |
| Microsoft Excel | 365 (subscription) | Desktop, signed-in Microsoft account required |
| mkcert | Latest | For trusted local HTTPS cert |

---

#### macOS

**1. Clone the repo**
```bash
git clone https://github.com/james-may-hammond/excel-mongo.git
cd excel-mongo
```

**2. Python virtual environment + dependencies**
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

**3. Node dependencies**
```bash
npm install
```

**4. Create a trusted local SSL certificate**
```bash
brew install mkcert
sudo mkcert -install
mkcert -cert-file cert.pem -key-file key.pem localhost 127.0.0.1 ::1
```

**5. Configure environment**
```bash
cp .env.example .env
# Edit .env — set MONGO_URI and DB_NAME
```

**6. Start the backend**
```bash
./run.sh
```

**7. Load the add-in into Excel**
```bash
# In a second terminal
./node_modules/.bin/office-addin-debugging start fe/manifest.xml --no-debug --dev-server-port 8000
```
Excel opens automatically with the **MongoDB Sync** button in the Home ribbon.

---

### Packaging for production

Office add-ins are published as a hosted web app plus a manifest that points to that hosted app. Localhost manifests are only for development.

For Docker, Railway, and Vercel deployment steps, see `DEPLOYMENT.md`.

#### Build a publishable package

Deploy the FastAPI app and the `fe/` static assets behind trusted HTTPS, then generate the production package:

```bash
ADDIN_PUBLIC_URL=https://your-addin.example.com \
ADDIN_API_BASE=https://your-addin.example.com \
npm run package
```

This creates:

| Path | Purpose |
|---|---|
| `dist/manifest.xml` | Production manifest to upload for admin deployment or Marketplace submission |
| `dist/fe/` | Static taskpane assets to deploy to the public web host |
| `dist/fe/config.js` | Runtime API URL used by the taskpane |

If the frontend and API are hosted separately, point `ADDIN_PUBLIC_URL` at the static frontend host and `ADDIN_API_BASE` at the API host:

```bash
ADDIN_PUBLIC_URL=https://excel-mongo-ui.example.com \
ADDIN_API_BASE=https://excel-mongo-api.example.com \
ADDIN_SUPPORT_URL=https://your-support-page.example.com \
npm run package
```

#### Validate the manifest

```bash
ADDIN_PUBLIC_URL=https://your-addin.example.com npm run package:validate
```

For local development validation:

```bash
npm run validate:dev
```

#### Deployment options

- Internal organization rollout: upload `dist/manifest.xml` through the Microsoft 365 admin center Integrated apps portal.
- Public distribution: submit the add-in through Microsoft Partner Center / Microsoft Marketplace.
- Testing only: sideload `fe/manifest.xml` or run `npm run start:addin`.

Production checklist:

- Use HTTPS URLs in the manifest.
- Host `dist/fe/` at the same path generated in `dist/manifest.xml`.
- Host the FastAPI backend at `ADDIN_API_BASE`.
- Configure CORS for the production Office/add-in domains before locking it down.
- Increment `package.json` `version` before publishing manifest changes.
- Do not publish a manifest that contains `localhost`.

---

#### Windows

**1. Clone the repo**
```cmd
git clone https://github.com/james-may-hammond/excel-mongo.git
cd excel-mongo
```

**2. Python virtual environment + dependencies**
```cmd
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

**3. Node dependencies**
```cmd
npm install
```

**4. Create a trusted local SSL certificate**

Download `mkcert-v*-windows-amd64.exe` from [mkcert releases](https://github.com/FiloSottile/mkcert/releases), rename it to `mkcert.exe`, place it on your PATH, then run in an **Administrator** terminal:
```cmd
mkcert -install
mkcert -cert-file cert.pem -key-file key.pem localhost 127.0.0.1 ::1
```

**5. Configure environment**

Create a `.env` file in the project root:
```
MONGO_URI=mongodb://localhost:27017
DB_NAME=your_database_name
```

**6. Start the backend**
```cmd
.venv\Scripts\activate
python -m uvicorn be.main:app --host localhost --port 8000 --ssl-keyfile key.pem --ssl-certfile cert.pem --reload
```

**7. Load the add-in into Excel**
```cmd
.\node_modules\.bin\office-addin-debugging start fe\manifest.xml --no-debug --dev-server-port 8000
```
Excel opens automatically with the **MongoDB Sync** button in the Home ribbon.
