# EWS Tracker — Early Warning Signs

Employee attrition-risk tracker for team leads and admins: warning indicators, a computed risk score/status, monthly headcount tracking, and inter-team transfer requests. Currently a single-file HTML/JS app (`index.html`) backed by Supabase.

## Migration to SharePoint & Power Platform

Supabase isn't allowed at the org this is deployed to, and neither are custom Azure AD app registrations or custom-script hosting in SharePoint. The rebuild target is **SharePoint Lists + Power Apps (canvas app) + Power Automate** — using only what's already licensed, no new IT approvals needed.

The full build spec — SharePoint List schemas, Power Fx formulas, screen map, Power Automate flows, and a Power Apps delegation/scale warning specific to this app — is in [`ews-rebuild-blueprint.md`](ews-rebuild-blueprint.md).

### Setting up on a work PC with GitHub Copilot (Sonnet/Opus)

1. **Get the file locally** — clone or pull this repo, or download `ews-rebuild-blueprint.md` directly from GitHub (`Raw` → Save As, or the file view's Download button).
2. **Open Copilot Chat and pick the model** — in VS Code: the Copilot Chat icon in the sidebar (or `Ctrl+Alt+I` / `Cmd+Ctrl+I`). Click the model picker at the bottom of the chat input and switch to **Claude Sonnet** — it's fast and covers most of the formula debugging and back-and-forth. Switch to **Opus** only if Sonnet gets stuck on something gnarly (a delegation issue, a formula erroring for a non-obvious reason).
3. **Attach the file as context** — use the 📎 attach icon in Copilot Chat, or drag `ews-rebuild-blueprint.md` straight into the chat panel. Open with something like: *"This is the spec for an app I'm rebuilding on SharePoint Lists + Power Apps + Power Automate. I'll be working in Power Apps Studio and pasting formulas/errors as I go — use this as the source of truth for schema and logic."*
4. **Set up the SharePoint side** — in your SharePoint site: **+ New → List → Blank list** for each of the 5 lists in the spec, then add columns per that list's table (List settings → Create column, matching the type shown). Build `Employees` first — the other lists reference it.
5. **Create the Canvas app** — go to `make.powerapps.com` → **Create → Blank app → Canvas**. Add all 5 SharePoint lists as data sources in the Data panel.
6. **Build in the spec's order** — `App.OnStart` formula first (paste into App properties → OnStart), then the Employee Form + scoring formulas, then the rest of the screens. When a formula throws a red error in the formula bar, copy the exact error text plus the formula into Copilot Chat.
7. **Power Automate flows** — separate tool, `make.powerautomate.com` → **Create → Automated cloud flow**: trigger "When an item is created" (SharePoint) on `TransferRequests` for the approval flow, and a Recurrence trigger for the daily snapshot flow.
8. **Before rollout** — raise the row limit (App settings → Advanced settings → Data row limit → 2000) and pilot against one team lead's real roster before opening it org-wide.
