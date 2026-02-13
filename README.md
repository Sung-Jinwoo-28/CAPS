# CAPS Confluence 2026 - Official Website

This repository contains the source code for the official website of **CAPS Confluence 2026**, a student leadership summit organized by the Centre for Academic and Professional Support (CAPS) at Christ University, Bangalore Yeshwanthpur Campus.

🔗 **Live Website**: [https://www.capsypr.in](https://www.capsypr.in)

## 📌 Project Overview
The website serves as the central hub for the event, providing:
- Event details (Seminars, Panels, Workshops, Fireside Chats).
- Registration portal with dynamic pricing (`submission.html`).
- Seat availability tracker (synced with Google Sheets).
- Comprehensive SEO optimization.

## 🚀 Key Features
- **Dynamic Pricing Cards**: Custom views for Capsites, Christites, and Non-Christites.
- **Real-time Availability**: Fetches workshop seat counts from Google Apps Script.
- **SEO Optimized**: Fully implemented meta tags, Open Graph, Twitter Cards, Schema.org (JSON-LD), Sitemap, and Robots.txt.
- **Responsive Design**: Built with Tailwind CSS (CDN) for mobile-first responsiveness.
- **Glassmorphism UI**: Modern aesthetic with glass panels and smooth animations.

## 🛠 Tech Stack
- **Frontend**: HTML5, JavaScript (Vanilla), Tailwind CSS (via CDN).
- **Backend (Logic)**: Google Apps Script (GAS) for form handling and data fetching.
- **Database**: Google Sheets (via GAS).
- **Deployment**: Static Hosting (Vercel/GitHub Pages).

## 📂 Project Structure
```
├── index.html          # Main landing page
├── submission.html     # Registration form logic
├── html.js             # Backend logic (Google Apps Script reference)
├── sitemap.xml         # SEO Sitemap
├── robots.txt          # SEO Crawling rules
└── caps.svg            # Assets
```

## 🔧 Setup & Installation
Since this is a static site using Tailwind via CDN, no build process is required for development.

1.  **Clone the repository**:
    ```bash
    git clone https://github.com/Sung-Jinwoo-28/CAPS.git
    cd CAPS
    ```
2.  **Run locally**:
    Simply open `index.html` in your browser or use a live server extension.

## 🔍 SEO Configuration
The site is optimized for the domain `capsypr.in`.
- **Sitemap**: `https://www.capsypr.in/sitemap.xml`
- **Verification**: Google Search Console tag is implemented in `index.html`.
- **Analytics**: Ready for Google Analytics integration.

## 📝 License
Proprietary software for CAPS Christ University. All rights reserved.
