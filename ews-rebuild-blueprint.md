# EWS Tracker — Power Platform Build Spec

Context for whoever's helping build this: EWS Tracker is an employee attrition-risk tracker currently built as a single static HTML/JS file using Supabase (Postgres + Auth) as its backend. Supabase is not allowed at this org, and neither are custom Azure AD app registrations or custom-script hosting in SharePoint. Power Apps is confirmed available. This doc is the spec for rebuilding it on SharePoint Lists + Power Apps (canvas app) + Power Automate.

## Concept mapping

| Today — Supabase | Rebuilt — Power Platform |
|---|---|
| Postgres tables | SharePoint Lists |
| Supabase Auth (username + password) | Microsoft 365 sign-in — already theirs |
| Row-level security policies | Person-column filters in Power Fx |
| Client-side JS (index.html) | Power Fx formulas in a Canvas app |
| Snapshot captured on page view | Power Automate — daily scheduled flow |
| SheetJS bulk import / CSV export | Native Excel + SharePoint export |

## Architecture

Team leads & admin sign in with their M365 account (no signup form) → Canvas app (Power Apps, screens + Power Fx) → reads/writes 5 SharePoint Lists → Power Automate flows handle approvals and the daily snapshot.

The old `profiles` table did two jobs: held role/approval state, and was half of Supabase's auth system. Those split apart here. Identity comes free from `User()` the moment someone opens the app with their work account. The `TeamLeads` list only tracks: is this person a recognized team lead or admin, and are they currently active. `App.OnStart` looks that up once per session.

## SharePoint Lists

### 1. TeamLeads — replaces `profiles` + Supabase Auth
Who's a recognized team lead or admin. An admin adds people directly by picking their M365 identity — no self-service signup to approve.

| Column | Type | Notes |
|---|---|---|
| Person | Person or Group | Key column. Replaces profiles.id entirely. |
| Role | Choice | Admin · Team Lead — was profiles.role |
| Approved | Yes/No | Now means "currently active." Flip to No to revoke without deleting history. |
| Notes | Single line text | Optional |

### 2. Employees — replaces `employees`
The core roster. RiskScore and RiskStatus are written on save, not computed on the fly (see Delegation note).

| Column | Type | Notes |
|---|---|---|
| Title | Single line text | Employee name — was employees.name |
| EmpID | Single line text | Text, not Number — preserves leading zeros |
| Owner | Person | The team lead this employee belongs to |
| Position | Choice | OCN, Outreach, B vs D, EDITS, C&S, Appeals Verification, Clinical Appeals, Adherence & Monitoring, Outbound Denial, Part_D Tech, Gen_PA Tech, GLP_1, UHC_west, Misroutes, CE, SME, Escalation, WFM/RTA |
| DateHired | Date only | |
| Ind_Tardy, Ind_Absent, Ind_LowProd, Ind_Diseng, Ind_JobHunt, Ind_Conflict, Ind_QADrop, Ind_NoInit, Ind_LOA, Ind_Withdraw | Yes/No ×10 | Flattened from the old `ind{}` JSON object |
| ActiveCAP | Yes/No | |
| Rating | Single line text | e.g. "4.2 / Meets" |
| ConfirmedAttrition | Choice | (blank) · BLACK · ABSCONDING · LOA · MATERNITY |
| AttritionDate, LeaveStart, LeaveReturn | Date only ×3 | |
| Notes | Multiple lines, plain text | TL notes |
| RiskScore | Number | Written on save |
| RiskStatus | Choice | GREEN · YELLOW · RED · BLACK — written on save |
| Modified | built-in | Replaces updated_at |

### 3. TransferRequests — replaces `transfer_requests`
One item per transfer ask. Exists mainly to trigger the approval flow.

| Column | Type | Notes |
|---|---|---|
| Employee | Lookup → Employees | |
| FromOwner, ToOwner | Person ×2 | |
| Status | Choice | Pending · Approved · Declined · Cancelled |
| Note | Multiple lines text | Optional |
| Created | built-in | Replaces requested_at |

### 4. Headcount — replaces `headcount`
One item per team lead per month. Supabase enforced "one row per (owner, year, month)" with a DB constraint — SharePoint has none, so the app must look before it writes (see upsert formula).

| Column | Type | Notes |
|---|---|---|
| Owner | Person | |
| Year | Number | |
| Month | Number (1–12) | |
| OpeningHC, NewHires, TransferIn, TransferOut, VoluntaryAttrition, InvoluntaryAttrition, ProjectedAttrition | Number ×7 | |
| ClosingHC | Number | Computed and written on save |

### 5. RiskSnapshots — replaces `risk_snapshots`
Daily point-in-time counts for the Trends charts. Power Automate makes this a real scheduled job instead of "capture whenever someone opens the page."

| Column | Type | Notes |
|---|---|---|
| Owner | Person | |
| SnapDate | Date only | |
| GreenCount, YellowCount, RedCount, BlackCount | Number ×4 | |
| AvgScore | Number | |

## Power Fx formulas

### Role & identity lookup — App.OnStart
```
// runs once when the app opens — replaces the entire Supabase login screen
Set(varMe, LookUp(TeamLeads, Person.Email = User().Email));
Set(varIsAdmin, !IsBlank(varMe) && varMe.Role.Value = "Admin");
Set(varApproved, !IsBlank(varMe) && varMe.Approved);
```

### RiskScore — Employee form, live preview + before Patch
```
Set(varRiskScore,
  CountIf(
    Table(
      {Yes: swTardy.Value}, {Yes: swAbsent.Value},  {Yes: swLowProd.Value},
      {Yes: swDiseng.Value}, {Yes: swJobHunt.Value}, {Yes: swConflict.Value},
      {Yes: swQADrop.Value}, {Yes: swNoInit.Value},  {Yes: swLOA.Value}, {Yes: swWithdraw.Value}
    ),
    Yes = true
  ) + If(swActiveCAP.Value, 1, 0)
)
```

### RiskStatus — Employee form, same logic as calcStatus() in index.html
```
Set(varRiskStatus,
  If(
    ddAttrition.Selected.Value = "BLACK", "BLACK",
    Or(ddAttrition.Selected.Value = "ABSCONDING",
       ddAttrition.Selected.Value = "LOA",
       ddAttrition.Selected.Value = "MATERNITY"), "RED",
    varRiskScore = 0, "GREEN",
    varRiskScore <= 3, "YELLOW",
    "RED"
  )
)
```

### Active roster filter — Permanent Attrition screen, Items property
```
// a future-dated BLACK stays on the active roster until that date arrives
Filter(Employees,
  ConfirmedAttrition.Value = "BLACK",
  Or(IsBlank(AttritionDate), AttritionDate <= Today())
)
```

### Opening HC auto-carry — Headcount form, defaults OpeningHC from last month's close
```
Set(varAutoOpening,
  With({prevMonth: If(varMonth = 1, 12, varMonth - 1),
        prevYear:  If(varMonth = 1, varYear - 1, varYear)},
    LookUp(Headcount, Owner.Email = varMe.Person.Email
      && Year = prevYear && Month = prevMonth, ClosingHC)
  )
)
```

### Headcount save — upsert — Headcount form, Save button
```
With({existing: LookUp(Headcount, Owner.Email = varMe.Person.Email
                     && Year = varYear && Month = varMonth)},
  Patch(Headcount, If(IsBlank(existing), Defaults(Headcount), existing),
    {
      Owner: varMe.Person, Year: varYear, Month: varMonth,
      OpeningHC: varOpening, NewHires: varHires, TransferIn: varTIn, TransferOut: varTOut,
      VoluntaryAttrition: varVol, InvoluntaryAttrition: varInvol, ProjectedAttrition: varProj,
      ClosingHC: varOpening + varHires + varTIn - varTOut - varVol - varInvol
    }
  )
)
```

## Screen map

| Screen | Was | Visible to | What it does |
|---|---|---|---|
| Home | My Team / Dashboard | Everyone | Role-conditional: team leads get their own roster gallery, admins get the cross-team summary. Same screen, different Items formula. |
| Employee Form | Add/Edit modal | Team leads | 10 toggles + live RiskScore/RiskStatus preview |
| Headcount | Headcount | Everyone | Monthly entries + rollup formulas |
| Permanent Attrition | Permanent Attrition | Everyone | Active-roster filter, inverted |
| Trends | Trends | Everyone | Line chart bound to RiskSnapshots, filtered by owner for team leads |
| Performer | Performer | Admin only | Flags team leads whose avg days-since-Modified falls outside 5–12 days |
| Team Leads | Users | Admin only | Add a person to TeamLeads, set Role, toggle Approved — no signups to review |
| Settings | Settings | Everyone | Display name, etc. |

## Power Automate flows

### 1. Transfer Approval
**Trigger:** item created in TransferRequests
1. Start and wait for an approval, assigned to ToOwner, with employee name + FromOwner's note.
2. On approve — Patch the Employees item: Owner → ToOwner. Set Status → Approved.
3. On reject — Set Status → Declined. Employee stays put.

Bonus over the original: a Teams/Outlook step can notify FromOwner of the outcome — the old app only had an in-session toast.

### 2. Daily Risk Snapshot
**Trigger:** Recurrence, once daily
1. Get every approved TeamLeads Person.
2. For each — get their active Employees, count by RiskStatus, average RiskScore.
3. Upsert one RiskSnapshots row per owner for today's date.

Strictly better than the source: a real scheduled job instead of a snapshot that only fired when someone opened the Trends tab.

### 3. New Team Lead Review (optional)
**Trigger:** item added to TeamLeads with Approved = No
1. Notify the admin group via Teams that a new roster entry needs review.
2. Admin flips Approved to Yes once confirmed.

Only worth building if admins add people speculatively. If admins only ever add someone once confirmed, skip this — set Approved = Yes on entry.

## Delegation & scale — read before building

Power Apps only pushes some filters/formulas down to SharePoint to run there. Past the row limit (500 default, 2000 max), anything non-delegable runs locally on just that truncated slice, **silently, no error**. The admin Home screen sums RiskStatus counts across every team lead's roster at once — exactly the full-table aggregate at risk once total headcount grows past a few hundred.

Mitigations:
- Write RiskScore/RiskStatus to real columns on save — don't recompute live inside a Filter/CountRows over the whole table.
- Add indexed columns on Owner, RiskStatus, ConfirmedAttrition (Employees list settings) — indexing is what makes a filter delegable.
- Raise the app's data row limit from 500 to the 2000 max (App settings → Advanced settings).
- 2000 is still a hard ceiling. If total active headcount is heading toward that, Dataverse becomes worth a separate licensing conversation — not a blocker today, just a number to watch.

## Build order

1. Create the 5 SharePoint Lists — get Employees right first, everything else references it.
2. Add indexed columns to Employees (Owner, RiskStatus, ConfirmedAttrition) — cheap now, painful to retrofit later.
3. Build the Canvas app shell — App.OnStart identity lookup, role variables, navigation frame.
4. Build the Employee Form + scoring formulas — get this solid before anything downstream depends on it.
5. Build the remaining screens — Home, Headcount, Attrition, Trends, Performer, Team Leads, Settings.
6. Build the two Power Automate flows.
7. Raise the row limit to 2000 before real data volume makes the default matter.
8. Pilot with one team lead's real roster before rolling out org-wide.
