# SharePoint Hackathon Submission

> **Issue Title:** Project: Synozur Holidays & Birthdays – AI-Assisted SPFx Web Part for SharePoint

---

## Submission Name

**Synozur Holidays & Birthdays – SPFx Web Part & Viva Connections ACE**

---

## Description

### The Problem

Every organization celebrates its people and its culture — but sharing holiday calendars and birthday lists across a SharePoint intranet portal remains a surprisingly unsolved problem. Existing approaches are either too static (a manually maintained image or a pinned document), too fragile (iFrame embeds from external calendars), or too complex for the average SharePoint administrator to install and configure (custom Power Apps, complex Power Automate flows, or enterprise add-ins with heavy licensing requirements). There was clearly a better way.

### The Idea & AI-Assisted Development Journey

I came to this hackathon motivated by a desire to explore **AI-assisted coding** as a true force-multiplier for building production-quality SharePoint solutions. As someone more familiar with business requirements and SharePoint administration than deep front-end development, I leaned heavily on GitHub Copilot and other AI coding tools throughout this project to accelerate scaffold creation, reason through complex recurrence date math, and refine the TypeScript/React component structure.

The result was remarkable: what would previously have taken weeks of iteration and deep SPFx expertise was achievable in a fraction of the time — with AI helping me draft, review, and debug everything from the SharePoint list provisioning service to the adaptive card quick views. This project is as much a story about **what AI-assisted development enables** as it is about the solution itself.

### The Solution

**Synozur Holidays & Birthdays** is a fully self-contained SPFx solution that brings a rich, interactive holiday and birthday experience to any modern SharePoint Online portal page — with zero manual list setup required.

#### Key Capabilities

- 🚀 **Automatic SharePoint List Provisioning**: On first use, the web part silently creates a `HolidaysAndBirthdays` list with all required fields (title, date, type, recurrence rule, image URL, notes, active flag). The provisioning is fully idempotent — safe to run repeatedly, gracefully handles already-existing lists, and respects throttling limits with built-in delays. No administrator setup beyond deploying the `.sppkg` is required.

- 📅 **Intelligent Recurrence Engine**: The solution handles two classes of events:
  - *Fixed-date events*: Birthdays and fixed holidays (e.g., New Year's Day on January 1st) — annual recurrence with the same month and day each year.
  - *Moving holidays*: Complex holidays like Labor Day ("1st Monday of September"), Memorial Day ("last Monday of May"), or Thanksgiving ("4th Thursday of November") — handled through a custom recurrence rule pattern (`MONTHLY_BY_NTH_WEEKDAY:MM:WEEKDAY:NTH`) and a purpose-built `RecurrenceCalculator` service.

- 🌐 **Web Part for Modern SharePoint Pages**: A configurable React-based web part displays upcoming holidays and birthdays in a grouped-by-month layout. Features include:
  - Configurable default view (7–90 days) and expanded view (60–730 days)
  - Optional event images and "Holiday" / "Birthday" type badges
  - Filter by event type (Both / Holidays Only / Birthdays Only)
  - **SharePoint Brand Center fonts support** — automatically inherits custom body and headline fonts from the tenant's Brand Center font packages via CSS variables, no configuration required
  - Fluent UI components for a native SharePoint look and feel
  - Graceful empty states and user-friendly error messages

- 🃏 **Viva Connections Adaptive Card Extensions (ACEs)**: A companion card for Viva Connections dashboards supports:
  - Separate Holiday and Birthday mode per card (allowing multiple cards on the same dashboard)
  - Small and medium card sizes
  - Holiday Quick View: all holidays for the year with previous/next year navigation
  - Birthday Quick View: next N upcoming birthdays (3–10, configurable)
  - Optional "in X days" countdown display
  - Optional branding footer

- 🔒 **Permission-Aware**: Standard SharePoint permissions govern everything — list creation requires "Manage Lists," viewing requires "Read," and editing requires "Contribute." No special app permissions or API grants needed.

### Why This Matters

This solution fills a genuine gap in the SharePoint ecosystem. It is simple enough for a non-technical administrator to deploy in minutes, yet rich enough to delight end users with a polished, interactive experience. It respects the SharePoint data model by storing everything in a standard list that admins can manage directly. And it demonstrates how AI-assisted development can help practitioners who are not full-time developers ship quality SPFx solutions that solve real organizational problems.

---

## Technology & Languages

- **SharePoint Framework (SPFx)** 1.18.2
- **SharePoint Online**
- **Viva Connections** (Adaptive Card Extensions)
- **React** 17 / **TypeScript**
- **Fluent UI** (Office UI Fabric React 7)
- **PnP/sp** (v2) for SharePoint REST operations
- **GitHub Copilot** (AI-assisted development throughout)

---

## Project Video

_[To be recorded and linked prior to final submission]_

### Video Outline

The demo video should be approximately 3–5 minutes and cover the following segments:

**1. Introduction (~ 30 seconds)**
- Brief personal introduction: who I am, my background (more SharePoint administrator / business practitioner than a seasoned front-end developer)
- The motivation: wanting to use AI-assisted coding (GitHub Copilot) to jump-start a real solution to a common organizational problem

**2. The Problem (~ 30 seconds)**
- Show what a typical SharePoint portal looks like today when trying to surface holidays and birthdays: a static image, a pinned document, or nothing at all
- Explain why existing options fall short — too static, too fragile, or far too complex for an average administrator to set up

**3. The AI-Assisted Development Story (~ 45 seconds)**
- Show the GitHub repository and highlight how GitHub Copilot was used throughout: scaffolding the SPFx project, writing the recurrence date-calculation logic, building the React components, and drafting the provisioning service
- Emphasize that this was a project made possible *because* of AI assistance — the kind of solution that would have taken weeks without it came together much faster

**4. Live Demo — Web Part (~ 90 seconds)**
- Starting from a fresh SharePoint site with no list in place, add the "Holidays & Birthdays" web part to a page
- Show the automatic list provisioning in action: the list and all required fields are created silently on first load
- Walk through the web part UI: grouped-by-month layout, event images, Holiday/Birthday type badges, the default vs. expanded day-range views
- Open the property pane and show the key configuration options (day range sliders, display mode filter, list name, show/hide toggles)
- Add a new event directly in the SharePoint list and refresh the page to show it appearing immediately

**5. Live Demo — Recurrence Engine (~ 30 seconds)**
- Show a moving holiday entry (e.g., Labor Day with `MONTHLY_BY_NTH_WEEKDAY:09:MONDAY:1`) and demonstrate that the correct date is calculated for the current year automatically
- Contrast with a fixed birthday entry to show both recurrence types working side by side

**6. Closing — Sharing Freely & What's Next (~ 30 seconds)**
- Acknowledge that this solution is being shared freely and openly — the code is public on GitHub for anyone in the community to use, fork, or build upon
- Mention the next steps: submitting to the Microsoft AppSource marketplace and completing the Viva Connections Adaptive Card Extension (ACE) support that didn't make it into this submission
- Invite the community to try it, file issues, and contribute

---

## Project Repository URL

https://github.com/chris-mcnulty/synozur-comet

---

## Team Members

- https://github.com/chris-mcnulty

---

## Sharing Philosophy

This solution is being released **freely and openly** to the SharePoint community. No licensing fees, no paywalls — just a working web part that any organization can deploy to their tenant in minutes.

Getting here was genuinely exciting. Starting from a standing start and arriving at a polished, production-ready SPFx solution — one that provisions its own list, handles complex date recurrence, integrates with SharePoint Brand Center fonts, and works across both modern pages and Viva Connections dashboards — felt like a real proof point for what AI-assisted development can unlock. We're proud of what came together and happy to pay it forward to the community that made SharePoint a platform worth building on.

---

## Next Steps

While this submission represents a fully working and deployable solution, two items remain on the roadmap for the next release:

1. **Microsoft AppSource Publishing**: The solution is being prepared for submission to the Microsoft AppSource marketplace, so organizations can discover and install it directly from the SharePoint App Catalog without needing to build from source.

2. **Viva Connections Adaptive Card Extension (ACE) — Full Completion**: The ACE component (companion cards for the Viva Connections dashboard) was scoped and partially built during the hackathon period but did not reach a state ready for this submission. The card view, holiday quick view, and birthday quick view are functional; remaining work includes final polish, property pane refinements, and end-to-end testing before the ACE is shipped alongside the web part as an integrated package.

Both of these will be tracked in the public GitHub repository at https://github.com/chris-mcnulty/synozur-comet.

---

## Badge Validation

- [ ] I verify that all of my team members have completed the badge validation form at aka.ms/SharePointHackathon/Badges.
