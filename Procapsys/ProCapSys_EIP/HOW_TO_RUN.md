# How to Run

1. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

2. **`.env` is already included** with a working MongoDB Atlas URI and a
   JWT secret, so you can run it as-is. If you'd rather use your own
   database, just edit `MONGO_URI` / `DB_NAME` in `.env`.

3. **Start the server**
   ```bash
   python app.py
   ```
   The app runs on `http://localhost:5000` (from `FLASK_PORT` in `.env`).

4. **Log in**
   - Super Admin: `login@superadmin.in` / `Superadmin@123`
     (from `SUPER_ADMIN_EMAIL` / `SUPER_ADMIN_PASSWORD` in `.env`)
   - Everyone else: register through the UI, then approve the account
     from the Super Admin / Trainer dashboard.

## Single Active Session — what to test
1. Log in as the same account in two different browsers (or one normal +
   one incognito window).
2. The **first** browser will get logged out automatically on its next
   click/API call, with the message *"Your account has been logged in
   from another device."*, and gets redirected to the login page.
3. Opening a second **tab** in the *same* browser does **not** log you
   out — only a genuinely new login elsewhere does.

See `SINGLE_SESSION_CHANGES.md` for exactly what was changed and why.
