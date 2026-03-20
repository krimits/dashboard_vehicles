# Render deployment

This project is ready to deploy on Render with:

- public dashboard at `/`
- admin upload page at `/admin`
- persistent uploaded workbook storage on a Render disk

## Files

- `render.yaml`: Render Blueprint config
- `requirements.txt`: Python dependencies

## Required environment variable

- `DASHBOARD_ADMIN_SECRET`: secret used by the admin upload UI

## Deploy steps

1. Create a new Git repo for this workspace or copy these files into an existing repo.
2. Push the project to GitHub.
3. In Render, create a new Blueprint and point it to the repo.
4. Keep the disk mount at `/var/data`.
5. Set `DASHBOARD_ADMIN_SECRET` in Render before the first deploy.
6. Deploy the service.

## Runtime behavior

- Render provides the `PORT` variable automatically.
- The app already reads `HOST` and `PORT`.
- Uploaded Excel files and metadata are stored in `/var/data/dashboard`.
- If no uploaded workbook exists yet, the app falls back to the newest matching local Excel file.

## Start command

```bash
python -X utf8 "Qwen_python_20260317_kqvga2wu9.py" --storage-dir /var/data/dashboard
```

## First-time usage

1. Open the public URL and confirm the dashboard loads.
2. Open `/admin`.
3. Enter `DASHBOARD_ADMIN_SECRET`.
4. Upload the latest Excel file.
5. Wait for the next polling interval and confirm the dashboard refreshes.
