# Deployment guide for bank-statement-formatter

This repo is managed separately from `med-acct-system` and uses PM2 on port `3020`.
Public URL target is: `https://www.meditationcenter.net/statement/`

## 1) Server app startup (PM2)

```bash
npm install
PORT=3020 pm2 start ecosystem.config.js --env production
```

Useful commands:

```bash
pm2 status
pm2 logs bank-statement-formatter
pm2 restart bank-statement-formatter
pm2 stop bank-statement-formatter
```

## 2) Reverse proxy (example: nginx)

`www.meditationcenter.net` should route `/statement/` to this app.

```nginx
location /statement/ {
    proxy_pass http://127.0.0.1:3020/;
    proxy_set_header Host $host;
    proxy_set_header X-Real-IP $remote_addr;
    proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
    proxy_set_header X-Forwarded-Proto $scheme;
}
```

> If the UI is used at `/statement`, this preserves existing `fetch('/upload')` calls from the app
> only when your proxy strips `/statement` before forwarding.

If you prefer no path rewrite by proxy, the app also accepts `/statement/upload` directly.

## 3) GitHub Actions auto deploy

Create these repository secrets:

- `AWS_HOST`
- `AWS_USER`
- `AWS_SSH_KEY`
- `AWS_APP_DIR` (for example `/var/www/bank-statement-formatter`)
- optional `AWS_SSH_PORT` (default `22`)

Every push to `master` triggers deployment.
