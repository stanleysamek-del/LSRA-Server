
# LSRA Generator (Logo-Preserving) – FastAPI

This tiny API takes JSON from your Wix page, downloads your Wix-hosted LSRA template (.xlsx with logo),
updates **A15 (merged A15:K19)** with rich text, and returns a downloadable Excel file.
Logos, headers, and formatting are preserved because we only touch `xl/worksheets/sheet1.xml`.

## Endpoints
- `GET /` – health check
- `POST /generate-lsra` – body:
```json
{
  "dateOfInspection": "03/08/2024",
  "address": "3620 Howell Ferry Rd NW, Duluth, GA 30096",
  "inspector": "Stanley Samek"
}
```

## Local run
```bash
pip install -r requirements.txt
uvicorn app:app --host 0.0.0.0 --port 8000
```

## Deploy (two easy options)

### Option A: Render (GitHub-based)
1. Create a new GitHub repo and upload these three files: `app.py`, `requirements.txt`, `README.md`.
2. Go to https://render.com → New → Web Service.
3. Connect your repo → Select **Build Command**: `pip install -r requirements.txt`
4. **Start Command**: `uvicorn app:app --host 0.0.0.0 --port 10000`
5. Set **Environment** = Python 3.11.  
6. Click Deploy. Render will give you a URL like `https://your-app.onrender.com`.

### Option B: PythonAnywhere (upload ZIP; no GitHub needed)
1. Zip these files on your computer (or download the zip I provided).
2. Go to https://www.pythonanywhere.com → create account → **Files** → upload the zip and unzip.
3. **Consoles** → start a Bash console:
   ```bash
   pip3.11 install --user -r requirements.txt
   ```
4. **Web** → Add a new web app → **Manual configuration** (instead of Flask/Django).
5. In **WSGI configuration file**, replace content with:
   ```python
   import sys
   sys.path = ["/home/yourusername/yourfolder"] + sys.path
   from app import app as application  # for ASGI via ASGI wrapper
   ```
   Then click **Reload**.
6. Your site URL will look like: `https://yourusername.pythonanywhere.com/`.

> Tip: If PythonAnywhere requires WSGI (not ASGI), you can instead use `pip install fastapi[all]` and `uvicorn`, then configure via ASGI – see their docs. If this feels heavy, choose Render.

## Wix (frontend) snippet
```html
<script>
async function generateLSRAFromWix() {
  const payload = {
    dateOfInspection: globalHeader?.dateOfInspection || "",
    address: globalHeader?.address || "",
    inspector: globalHeader?.inspector || ""
  };

  const resp = await fetch("https://YOUR_BACKEND_URL/generate-lsra", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });

  if (!resp.ok) { alert("Server error"); return; }

  const blob = await resp.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "LSRA_Report_Generated.xlsx";
  a.click();
  URL.revokeObjectURL(url);
}
</script>
```

Then wire your **Download LSRA** button to `generateLSRAFromWix()`.
