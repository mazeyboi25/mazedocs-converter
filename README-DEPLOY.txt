MazeDocs Railway Converter Backend
==================================

WHAT THIS BACKEND FIXES
-----------------------
- Accepts files up to 200 MB at the MazeDocs application layer.
- Streams uploads to temporary disk in 1 MB chunks instead of reading the
  entire file into Python RAM.
- Streams converted files back with FileResponse instead of loading the
  entire result into RAM.
- Deletes temporary files after the download finishes.
- PDF -> DOCX uses pdf2docx for editable layout reconstruction.
- Docker image installs LibreOffice for legacy DOC/PPT/XLS and higher-fidelity
  Office conversions.
- CORS allows mazedocs.vercel.app, Vercel preview domains, and local testing.

DEPLOY TO RAILWAY
-----------------
1. Put the contents of this ZIP in a new GitHub repository, for example:
   mazedocs-converter

2. In Railway:
   New Project -> Deploy from GitHub repo -> choose mazedocs-converter

3. Railway detects the Dockerfile automatically.

4. After deployment:
   Settings -> Networking -> Generate Domain

5. Open:
   https://YOUR-RAILWAY-DOMAIN/api

   Expected JSON includes:
   "ok": true
   "max_upload_bytes": 209715200

6. Save the Railway domain. The MazeDocs frontend script must be changed so
   Universal Converter requests go directly to:
   https://YOUR-RAILWAY-DOMAIN/api

IMPORTANT
---------
Railway currently documents that request bodies must finish uploading within
5 minutes. There is no Vercel-style 4.5 MB request-body ceiling in the Railway
public-networking limits documentation. Very slow connections may still fail
if a large upload cannot finish within that window.

For production, 200 MB is the application ceiling. Conversion success also
depends on page count, document complexity, CPU, RAM, and conversion time.
