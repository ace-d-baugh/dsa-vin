# VINNY WDW — Microsoft Teams Bot

Posts to any Teams chat/channel whenever a Walt Disney World attraction goes **down (❌ 101)** or comes **back up (✅ 102)**. Polls ThemeParks.wiki every 10 seconds.

---

## Setup

### 1. Azure App Registration

1. Go to [portal.azure.com](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**
2. Name it `DSA VIN`, leave defaults, click **Register**
3. Copy the **Application (client) ID** — this is your `MicrosoftAppId`
4. Go to **Certificates & secrets** → **New client secret** → copy the value — this is your `MicrosoftAppPassword`

### 2. Azure Bot Service

1. Go to **Create a resource** → search **Azure Bot** → Create
2. Set **Microsoft App ID** to the ID from step 1
3. Under **Configuration**, set the **Messaging endpoint** to:
   ```
   https://YOUR-HOST/api/messages
   ```
4. Save

### 3. Run the Bot

```bash
cp .env.example .env
# Fill in MicrosoftAppId and MicrosoftAppPassword

npm install
npm start
```

For local dev, expose port 3978 with [ngrok](https://ngrok.com):
```bash
ngrok http 3978
# Use the https URL as your messaging endpoint in Azure Bot Service
```

### 4. Add to Teams

1. Put `manifest.json` + a `color.png` (192×192) + `outline.png` (32×32) into the `manifest/` folder
2. Replace both `YOUR-APP-REGISTRATION-ID-HERE` placeholders in `manifest.json` with your App ID
3. Zip the three files: `manifest.json`, `color.png`, `outline.png`
4. In Teams: **Apps** → **Manage your apps** → **Upload an app** → upload the zip
5. Add DSA VIN to any chat or channel

---

## How It Works

- On startup, DSA VIN does a silent sync to learn the current state of all attractions
- After that, any transition from `OPERATING → DOWN` posts an ❌ 101 alert
- Any transition from `DOWN → OPERATING` posts a ✅ 102 recovery
- All conversations DSA VIN has been added to receive every alert
- State is persisted in `attractionState.json` so restarts don't cause false alerts
- Conversation references are persisted in `conversationRefs.json`

## Commands

| Message | Response |
|---|---|
| `status` or `dsa vin status` | Confirms the bot is running |

---