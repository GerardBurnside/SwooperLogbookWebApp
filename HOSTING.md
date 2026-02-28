# Hosting Guide for Skydiving Logbook App

## Free Hosting Options

### 1. GitHub Pages (Recommended)
**Steps:**
1. Create a GitHub repository
2. Upload your files to the repository  
3. Go to repository Settings → Pages
4. Select source branch (usually `main`)
5. Your app will be live at `https://yourusername.github.io/repositoryname`

**Pros:** Easy setup, free custom domain support, automatic deployments
**Cons:** Public repository required for free tier

### 2. Netlify
**Steps:**
1. Sign up at [netlify.com](https://netlify.com)
2. Drag and drop your project folder to Netlify
3. Get instant URL like `https://random-name.netlify.app`
4. Optional: Connect to Git for auto-deployments

**Pros:** Drag-and-drop deployment, form handling, edge functions
**Cons:** Limited build minutes on free plan

### 3. Vercel
**Steps:**
1. Sign up at [vercel.com](https://vercel.com)
2. Import from Git or upload files
3. Get URL like `https://your-project.vercel.app`

**Pros:** Great performance, preview deployments, edge network
**Cons:** Primarily for frameworks (but works with static files)

### 4. Firebase Hosting
**Steps:**
1. Create project at [firebase.google.com](https://firebase.google.com)
2. Install Firebase CLI: `npm install -g firebase-tools`
3. Run `firebase init hosting` in your project
4. Run `firebase deploy`

**Pros:** Google integration (good for Sheets API), fast CDN, SSL
**Cons:** Requires CLI setup

### 5. Surge.sh
**Steps:**
1. Install: `npm install -g surge`
2. Run `surge` in your project directory
3. Follow prompts for custom domain

**Pros:** Simple CLI deployment, custom domains
**Cons:** Limited features on free plan

## Configuration for Hosting

### Environment Variables
Most hosts support environment variables for sensitive data:
- **Netlify:** Use Netlify UI or `netlify.toml`
- **Vercel:** Use Vercel UI or `vercel.json`
- **GitHub Pages:** Store in repository secrets (for actions)

### HTTPS Requirements
Google Sheets API requires HTTPS. All recommended hosts provide free SSL certificates.

### Custom Domains
All hosts support custom domains:
1. Buy domain from registrar (Namecheap, Google Domains, etc.)
2. Update DNS to point to hosting provider
3. Configure custom domain in hosting platform

## Security Considerations
- Never commit your actual `sheets-config.json` file
- Use environment variables for production
- Consider implementing OAuth2 for enhanced security
- Regularly rotate API keys

## Example Deployment Commands

### Netlify CLI:
```bash
npm install -g netlify-cli
netlify deploy
netlify deploy --prod
```

### Surge:
```bash
npm install -g surge
surge
```

### Firebase:
```bash
npm install -g firebase-tools
firebase login
firebase init hosting
firebase deploy
```

Choose the option that best fits your technical comfort level and requirements!