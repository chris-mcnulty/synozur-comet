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

---

## Project Repository URL

https://github.com/chris-mcnulty/synozur-comet

---

## Team Members

- https://github.com/chris-mcnulty

---

## Badge Validation

- [ ] I verify that all of my team members have completed the badge validation form at aka.ms/SharePointHackathon/Badges.
