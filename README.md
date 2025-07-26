# monarch-ai-search

## Summary

Short summary on functionality and used technologies.

[picture of the solution in action, if possible]

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.21.1-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

> Any special pre-requisites?

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| folder name | Author details (name, company, twitter alias with link) |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.1     | March 10, 2021   | Update comment  |
| 1.0     | January 29, 2021 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

# Monarch AI Search Web Part

## Project Overview
This project is a SharePoint Framework (SPFx) web part using React, designed to provide a banner-style search experience for SharePoint Online. It integrates and extends the PnP Modern Search solution for advanced search functionality and extensibility.

## Key Features
- Integrates PnP Modern Search web parts (Search Results, Filters, Verticals, Box)
- Extensible architecture for custom layouts, UI components, and search features
- Easy update workflow for PnP Modern Search source code
- Ready for collaborative development and future customization

## Setup Instructions

### 1. Install Dependencies
```sh
npm install
```

### 2. Fetch/Update PnP Modern Search Source
```sh
npm run get-pnp-modern-search
```
This will download the latest `search-parts` and `search-extensibility` source code from the official PnP Modern Search repository into `src/pnp-modern-search-parts/` and `src/pnp-modern-search-extensibility/` respectively.
Both directories are ignored in `.gitignore` and are not tracked by git for repository hygiene and best practices.

### 3. Build the Solution
```sh
gulp bundle
gulp package-solution
```

### 4. Integrate PnP Modern Search Components
- Import and use components from `src/pnp-modern-search-parts` in your custom web part code.
- Wire up connections between Search Box, Results, Filters, and Verticals as needed.
- Extend or customize layouts and features using the extensibility model.

### 5. Push Changes to Remote
```sh
git add .
git commit -m "Update web part and search integration"
git push
```

## Update Workflow
- To update PnP Modern Search, run `npm run get-pnp-modern-search` and rebuild the solution.
- Review release notes for breaking changes and test your integration.
- If you accidentally commit these source directories, update `.gitignore`, remove them from git tracking with `git rm -r --cached src/pnp-modern-search-parts src/pnp-modern-search-extensibility`, then commit and push the cleanup.

## Best Practices
- Keep third-party source code in `src/pnp-modern-search-parts` and `src/pnp-modern-search-extensibility` (not `node_modules`).
- Document any customizations or integration patterns clearly.
- Use modular components for easy future extension.

## For Models and Agents
- Always use the provided scripts for setup and updates.
- Follow the integration pattern in this README for consistent development.
- Do not modify code in `node_modules`.
- Communicate any changes or issues in the project documentation and commit messages.

## References
- [PnP Modern Search Documentation](https://microsoft-search.github.io/pnp-modern-search/)
- [PnP Modern Search GitHub](https://github.com/microsoft-search/pnp-modern-search)

---
For questions or issues, contact the project maintainer or open an issue in the repository.
