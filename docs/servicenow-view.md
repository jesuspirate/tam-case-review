# Set up your ServiceNow export view

The review works with **any** case export, even one with only the Number column.
Every column you add unlocks a bit more. Set up the view below once and every review
gets the full picture. Columns your instance doesn't fill in just come out blank, which
is fine.

The names below are the standard ServiceNow Customer Service Management labels. If your
instance renames a field or adds its own (for example a request-type field per program),
put those names in your profile (see "Your organization's profile" below).

## Recommended columns

| Column | What it adds |
|---|---|
| **Number** | the case (required) |
| **Short description** | the issue |
| **Highest Priority** (or **Priority**) | P1–P4 tags and sort order |
| **State** | whose move it is: "Action: customer" or your team |
| **Substate** | Awaiting Info, Pending RCA… (prep view) |
| **Opened** | case age, the aging chart, 30+/90+ day warnings |
| **Updated** | which cases changed since your last update |
| **Assigned to** | your team's case owner |
| **Customer Contact** | customer contact; also spots "customer replied" |
| **Assignment group** | prep view and CSV (not on the page) |
| **Comments and Work notes** | latest public comment and internal note, the Jira ticket |
| your **request-type** column(s) | page sections (blank = Support) |
| **Component**, **Sub-Component** | product area under each issue, "Focus area" callouts |
| **Product Version** | version under each issue, for self-hosted products |
| **Has Jira?** (or a bug-number field) | "Engineering" tag, linked to the ticket |
| **Assignee Updated** | "no update from the owner in N days" |
| **Non-Assignee Updated**, **Non-Assignee Updated By** | "customer replied, no owner update since" |
| **Follow up** | "follow-up due" / overdue reminders |
| **Action status** | "needs attention" flags |

**Nice to have** (prep view and the app): *Customer Account* (page subtitle),
*Description* (full text in the app), *Tenant URL* (which tenant), *Impact*, *Urgency*,
*Reassignment count*, *Escalation*, *SLA due*, *Project Code*.

**Usually leave out** *Additional comments* and *Work notes*: they repeat *Comments and
Work notes*, and the app merges them so nothing shows twice. One exception: ServiceNow
cuts *Comments and Work notes* off at about 32,000 characters, so on very long cases the
oldest comments are missing. Adding the two separate columns brings those back.

## What goes where

- **On the customer page:** the issue, priority, whose move it is, case age, your latest
  update, the product area or version, and an "Engineering" tag on cases with a Jira ticket.
- **Only for you** (Terminal, the app, "Before the call"): comments and work notes,
  substate, group, tenant, follow-up dates, needs-attention and escalation flags,
  "customer replied" alerts, owner idle time, reassignments, SLA.

## Setting it up in ServiceNow

1. Press **Create a new list**.
2. Select a list you already have for your customer. If you don't have one, create it
   from scratch.
3. Add a condition: **Customer account** is **Acme** (your customer).
4. In **Select columns**, add the columns from the table above. Type a name to find it.
5. Press **Create**.
6. In the list shown, press **Export**.
7. Drop the file into **TAM Case Review** (or put it in `exports/<account>/` for the
   command line).

The list is saved, so next time just open it and press **Export**. Menus differ a little
between ServiceNow versions and workspaces. If yours has no **Create a new list**, use the
list's column settings and its **Export → Excel** menu instead.

The app tells you if your export is missing a recommended column ("Get more out of this
review" at the top of a review, then **Which columns?**). The command line prints the same
tip under the case counts.

## Counting contract usage

If your profile limits some case types per contract year, the review counts them since
each account's contract start. Your list must then also include **closed** cases, or the
counts are only minimums ("at least"). Add a condition group: *Active is true* **OR**
*Opened on or after* the contract start date. Closed cases are only used for the count.

## Your organization's profile

Names that differ between companies live in a small settings file, `profile.json`, not
in the code. Start from [profiles/example.json](../profiles/example.json):

| Key | What it sets |
|---|---|
| `vendor` | your company name as customers see it ("Action: Acme" / "Waiting on Acme") |
| `jira_browse_url` | where ticket links go, e.g. `https://acme.atlassian.net/browse/%s` |
| `types`, `default_type`, `type_aliases` | page sections, their order and colors |
| `type_columns`, `type_column_label` | which export columns hold the case type |
| `topics` | keyword focus areas, used when there's no Component column |
| `self_hosted_products` | products whose cases show Product Version |
| `contract_allowances` | case types limited per contract year, e.g. `{"Ask an Expert": 80, "Go Live Support": 2}` |

- **Command line:** put `profile.json` next to `callprep.py`, or point `TAM_PROFILE` at it.
- **Desktop app:** the build scripts bundle `profile.json` from the toolkit folder. GitHub
  builds read it from the `TAM_PROFILE` repository secret (`gh secret set TAM_PROFILE < profile.json`).
  An app built without one uses neutral defaults (the vendor shows as "Vendor").

`profile.json` is git-ignored, so your organization's details stay out of the repository.
