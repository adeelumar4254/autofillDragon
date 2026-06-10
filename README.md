# Autofill Dragon — Outlook Add-in

> A powerful React-based Outlook add-in that supercharges your email workflow with intelligent autofill capabilities.


## Table of Contents

- [Prerequisites](#prerequisites)
- [Getting Started](#getting-started)
- [Development Setup](#development-setup)
- [Sideloading the Add-in](#sideloading-the-add-in)
- [Debugging on Mac](#debugging-on-mac)
- [Deployment](#deployment)


## Prerequisites

Before you begin, ensure the following are installed on your machine:

| Requirement | Version | Link |
|---|---|---|
| Node.js | v22+ recommended | [Download](https://nodejs.org/en/download) |
| Visual Studio Code | Latest | [Download](https://code.visualstudio.com) |
| Microsoft Outlook | Desktop or Web | — |

## Getting Started

### 1. Open the Project

Open the project folder in **Visual Studio Code**:

```bash
File → Open Folder → Select project root
```

### 2. Open the Terminal

In VS Code, open an integrated terminal:

```
Terminal → New Terminal  (or Ctrl + ` / Cmd + `)
```

### 3. Navigate to the Project Root

```bash
cd your-project-folder
```

### 4. Install Dependencies

Install all required Node.js packages:

```bash
npm install
```

> This downloads all dependencies listed in `package.json`. Wait for the installation to complete before proceeding.

---

## Development Setup

### 5. Start the Development Server

```bash
npm start
```

The development server will launch and serve the add-in at:

```
https://localhost:3000
```

> ⚠️ Keep this terminal window open while using the add-in — closing it stops the server.

### 6. Trust the Development Certificate

If prompted with a certificate warning, install the development certificate by running:

```bash
npx office-addin-dev-certs install
```

Accept any security prompts that appear in your browser or OS.

---

## Sideloading the Add-in

### Option A — Outlook Desktop

1. Open **Microsoft Outlook**
2. Click **Get Add-ins** from the ribbon
3. Select **My Add-ins**
4. Scroll down to **Custom Add-ins**
5. Click **Add a Custom Add-in → Add from File**
6. Browse to the project root and select `manifest.xml`
7. Click **Install**

The **Autofill Dragon** button will now appear in your Outlook ribbon.


### Option B — Outlook Web

1. Open [Outlook Web](https://outlook.office.com)
2. Click the **Settings** icon (⚙️)
3. Select **Manage Add-ins**
4. Choose **My Add-ins**
5. Click **Add a Custom Add-in → Add from File**
6. Upload `manifest.xml` from the project root
7. Complete the installation prompts


### 7. Verify Installation

1. Open any email message in Outlook
2. Locate the **Autofill Dragon** button in the ribbon
3. Click it — the React application should load in the task pane


## Debugging on Mac

To debug the add-in on macOS, follow the official Microsoft guide:

📖 [Debug Office Add-ins on iPad and Mac](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/debug-office-add-ins-on-ipad-and-mac)



## Deployment

### 1. Build for Production

```bash
npm run build
```

The compiled output will be generated in the `dist/` folder.

### 2. Host the Build Files

Deploy the contents of `dist/` to any static hosting provider:

| Provider | Notes |
|---|---|
| GitHub Pages | Free, great for public repos |
| Azure Static Web Apps | Seamless Microsoft integration |
| AWS S3 + CloudFront | Scalable, enterprise-grade |
| Netlify / Vercel | Fast setup, CI/CD built-in |

After deploying, note your live URL (e.g., `https://your-app.azurestaticapps.net`).

### 3. Update the Manifest

Open `manifest.xml` and replace all occurrences of:

```
https://localhost:3000
```

with your production URL:

```
https://your-live-deployment-url.com
```

> 📖 Full publishing guide: [Publish an Office Add-in using VS Code](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-add-in-vs-code#using-visual-studio-code-to-publish)

---

## Quick Reference

| Command | Description |
|---|---|
| `npm install` | Install dependencies |
| `npm start` | Start development server |
| `npm run build` | Build for production |
| `npx office-addin-dev-certs install` | Trust local SSL certificate |

---

*Built with React · Powered by the Office JavaScript API*