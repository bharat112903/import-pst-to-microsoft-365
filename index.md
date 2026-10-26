---
layout: default
title: "Import PST to Microsoft 365: Methods, Limits & Real Costs (2026)"
description: "How to import PST files to Microsoft 365 in 2026: Purview network upload, AzCopy, PowerShell, throttling (24 GB/day per mailbox), failures, and plan upgrade costs."
keywords: "import PST to Microsoft 365, PST import Office 365, Purview network upload, AzCopy PST import, New-MailboxImportRequest Exchange Online, PST import slow, Exchange Online archive limits"
author: "Bharat Bhushan, Product Manager"
author_url: "https://www.linkedin.com/in/bharat24/"
reviewer: "Shelly Bhardwaj, Exchange Online Migration Expert"
reviewer_url: "https://www.linkedin.com/in/shelly-bhardwaj/"
date: 2026-10-23
last_modified_at: 2026-10-23
canonical_url: "https://[YOUR-DOMAIN]/import-pst-to-microsoft-365/"
---

# Import PST to Microsoft 365: Methods, Limits & What the Manual Route Really Costs (2026)

> **Short answer:** In Exchange Online, the supported admin method for importing PST files is **Microsoft Purview network upload** (AzCopy + a mapping CSV + the Import page). `New-MailboxImportRequest` works only for **on-premises Exchange**. Microsoft cites roughly **24 GB per day per mailbox** for import speed (not guaranteed), and we found **no documented paid add-on** that raises it.

Need the product-led walkthrough instead? See Stellar's guide to [import PST to Office 365](https://www.stellarinfo.com/import-pst-to-office-365.php).

**By** [Bharat Bhushan](https://www.linkedin.com/in/bharat24/), Product Manager · **Technically reviewed by** [Shelly Bhardwaj](https://www.linkedin.com/in/shelly-bhardwaj/), Exchange Online Migration Expert · **Published:** 23 October 2026 · **Last reviewed:** 23 October 2026

> **Disclosure:** This guide is published by Stellar Information Technology, the maker of Stellar Migrator for Outlook. Microsoft figures come from Microsoft documentation and are linked. Where this guide discusses Stellar, it lists the questions we are testing rather than claims, and we will add measured results as they are completed. See [How we researched this](#how-we-researched-this).

---

## Table of contents

1. [Key takeaways](#key-takeaways)
2. [PST import methods compared](#pst-import-methods-compared)
3. [How to import PST files with Purview network upload](#how-to-import-pst-files-with-purview-network-upload)
4. [Can you still use PowerShell to import PSTs?](#can-you-still-use-powershell-to-import-psts)
5. [Where PST imports fail](#where-pst-imports-fail)
6. [Import speed and throttling](#import-speed-and-throttling)
7. [Mailbox limits and upgrade costs](#mailbox-limits-and-upgrade-costs)
8. [Where a dedicated migration tool fits](#where-a-dedicated-migration-tool-fits)
9. [How to choose a method](#how-to-choose-a-method)
10. [FAQ](#faq)
11. [How we researched this](#how-we-researched-this)
12. [Sources](#sources)

---

## Key takeaways

- **No single "PowerShell way" exists for Exchange Online.** Microsoft says `New-MailboxImportRequest` is not supported in Exchange Online and points to network upload instead.
- **Network upload is the supported admin route.** PowerShell supports it (finding mailbox GUIDs, managing retention hold) but does not run the import.
- **Speed is per mailbox, not per tenant.** About **24 GB per day per mailbox**, not guaranteed. PSTs for the same mailbox import one after another.
- **Bigger plans add storage, not documented import speed.** Drive shipping is the only paid route Microsoft documents, and it is limited to Enterprise Agreement customers.
- **Failures are common enough to plan for.** Microsoft Q&A threads report jobs stuck at "Analysis in progress" for over a day and jobs stuck at 0%.
- **Upgrading for capacity is a recurring cost.** Example: 100 users moving from Exchange Online Plan 1 to Plan 2 costs about **$4,800 a year** at the list prices we found.

---

## PST import methods compared

| Method | Best for | PowerShell role | Main limits |
|---|---|---|---|
| Outlook Import/Export or drag and drop | One user, a few small files | None | Manual, per user, depends on the desktop client |
| Purview **network upload** (AzCopy + mapping CSV) | Admins moving many PSTs | Supporting (GUIDs, retention hold) | 20 GB per PST recommended, 500 rows per CSV, per-mailbox speed limit |
| Purview **drive shipping** | Very large volumes | Supporting | Enterprise Agreement only; 7 to 10 business days after Microsoft receives the drive; paid |
| `New-MailboxImportRequest` | **On-premises Exchange only** | Primary | Not supported in Exchange Online |

---

## How to import PST files with Purview network upload

### Prerequisites

- An admin account that can open the **Import** page in the Microsoft Purview portal.
- PST files on a file share, file server, or Azure Storage location you manage.
- The AzCopy version provided on the Import page. Other versions are not supported, and Azure Storage Explorer cannot be used to upload.

### Step 1: Get the SAS URL and AzCopy

1. In the Purview portal go to **Data lifecycle management > Microsoft 365 > Import**.
2. Select **New import job**, name it (lowercase letters, numbers, hyphens, underscores), and choose **Upload your data**.
3. Select **Show network upload SAS URL**, copy it, and store it like a password.
4. Download AzCopy from the same page.

### Step 2: Upload the PST files with AzCopy

```powershell
azcopy.exe copy "<Source location>" "<SAS URL>"
```

- You cannot point at a single file; everything in the source folder uploads.
- In scripts, escape `%` as `%%` and `&` as `^&`.
- Add `--s2s-preserve-access-tier=false` if the source is a general-purpose v2 storage account.

### Step 3: Build the mapping CSV

The header must stay exactly as supplied:

```csv
Workload,FilePath,Name,Mailbox,IsArchive,TargetRootFolder,ContentCodePage,SPFileContainer,SPManifestContainer,SPSiteUrl
```

Key rules:

- One row per PST; `Workload` is `Exchange`.
- `FilePath` and `Name` are **case-sensitive** and must match the upload exactly.
- Names must be unique within a job.
- Maximum **500 rows** per CSV.
- `IsArchive` set to TRUE requires an enabled archive.

### Step 4: Create the job, analyze, and import

1. After the upload finishes, select the CSV, choose **Validate** (it must pass), accept the terms, and **Save**.
2. Wait for the status to move from **Analysis in progress** to **Analysis completed**.
3. Choose **Import to Microsoft 365**, optionally filter by age or item type, then **Import data**.
4. Refresh until the status reads **Complete**.

### What Microsoft documents about timing

| Stage | Documented timing |
|---|---|
| Upload | "Several hours per TB," depending on your network |
| Analysis | **Not documented** |
| Import | About **24 GB per day per mailbox**, not guaranteed, varies with load |
| Uploaded files | Deleted 30 days after the last job was created, if no jobs are in progress |

---

## Can you still use PowerShell to import PSTs?

**For Exchange Online: no, not as the main method.** PowerShell supports network upload with these commands Microsoft references:

```powershell
# Mailbox GUID (needed when addresses are shared with inactive or soft-deleted mailboxes)
Get-Mailbox <identity> | FL ExchangeGuid
Get-Mailbox <identity> -InactiveMailboxOnly | FL ExchangeGuid
Get-Mailbox <identity> -SoftDeletedMailbox | FL ExchangeGuid

# Imports place the mailbox on retention hold indefinitely. Release it or set an end date:
Set-Mailbox <identity> -RetentionHoldEnabled $false
Set-Mailbox <identity> -EndDateForRetentionHold <date>
```

**For on-premises Exchange: yes.** `New-MailboxImportRequest` works there:

```powershell
New-MailboxImportRequest -Mailbox Ayla -FilePath \\SERVER01\PSTFiles\Recovered.pst -TargetRootFolder "RecoveredFiles" -IncludeFolders "#Inbox#"
```

- It needs the **Mailbox Import Export** role (not assigned by default) and read permission for the Exchange Trusted Subsystem on the share.
- `BadItemLimit` defaults to 0; Microsoft recommends 10 or lower; 51 or higher needs `AcceptLargeDataLoss`.
- In Exchange Online, Microsoft directs you to network upload instead.

---


## Where PST imports fail

### Documented limits that cause failures or skipped items

- PSTs over **20 GB** can slow the import.
- Items over **150 MB** are skipped.
- PSTs with **300 or more nested folder levels** cannot be imported.
- Archive imports cannot exceed **100 GB** per import.
- CSV validation is strict and case-sensitive; a wrong path or name fails the row.
- Public folders and unified groups are not supported targets.
- Uploaded files **cannot be deleted** from the storage container.

### Failures reported by administrators

These are community reports, not Microsoft-confirmed bugs.

| Report | What happened | Outcome on the thread |
|---|---|---|
| [PST import very slow, stuck on analysis (Microsoft Q&A, 2025)](https://learn.microsoft.com/en-us/answers/questions/2244008/purview-pst-import-very-slow-seeminly-stuck-on-ana) | Nearly all jobs sat at Analysis in progress for over 24 hours; one failed with "Failed; Error: Unknown" and no useful log | Advice: cancel and recreate, confirm uploads finished, split PSTs, check network and PST integrity. No confirmed fix from the poster |
| [Import stuck at 0% for half a day (Microsoft Q&A, 2023)](https://learn.microsoft.com/en-us/answers/questions/1397044/how-to-stop-an-m365-pst-import-job-that-is-stuck-o) | A 300 MB PST stayed at 0%; the job could not be cancelled and did not show in `Get-MailboxImportRequest` | Advice: check PowerShell first, try forced removal. No accepted fix |
| [Jobs stuck at Analysis in progress (Microsoft Q&A)](https://learn.microsoft.com/en-us/answers/questions/1132972/import-jobs-to-office-365-mailbox-stuck-at-analysi) | Similar stalls | See thread |

Even tiny test files stalled in one report, so file size alone does not explain every delay.

### Gaps in the manual process

| Gap | Why it matters |
|---|---|
| No documented analysis time or completion estimate | You cannot promise users a cut-over date |
| Opaque failures ("Unknown" errors, thin logs) | Troubleshooting becomes support tickets |
| No PST health check before upload | Corrupt PSTs hang or fail after hours of upload |
| Per-mailbox serial queue | Multiple PSTs per user stack up in time |
| Strict CSV with case-sensitive paths | Typos cost whole job cycles |
| Retention hold left on indefinitely | Easy to forget; affects retention behavior later |
| Fixed AzCopy version, no single-file upload | Less flexible for small fixes and re-runs |
| 500 rows per CSV | Large estates need many jobs |
| No documented way to buy speed | Deadlines cannot be bought out of |

---

## Import speed and throttling

### What Microsoft publishes

- About **24 GB per day per mailbox** (not guaranteed).
- PSTs for **different** mailboxes import in parallel; PSTs for the **same** mailbox import sequentially.
- Recommended PST size: **20 GB or less**.
- Archive import: **100 GB per import**; split larger data across imports.
- Network devices can also slow Exchange Online migration traffic. A [2015 Microsoft post](https://learn.microsoft.com/en-us/archive/blogs/onthewire/http-flood-mitigation-causing-slow-exchange-mailbox-migrations-to-office-365) describes proxies and firewalls with HTTP flood or per-IP connection limits slowing migrations. It is archived and TMG-specific, so check current guidance before applying it.

### How long will a PST import take?

Estimates calculated from Microsoft's 24 GB per day figure. They are arithmetic, not promises.

| Data going into one mailbox | Estimated import time |
|---|---|
| 50 GB | about 2 days |
| 100 GB | about 4 days |
| 200 GB | about 8 days |
| 500 GB | about 21 days |
| 1 TB | about 42 days |

Splitting a user's data across many PSTs does not help if they all target the same mailbox, because they queue. Scaling comes from **more mailboxes in parallel**, not from bigger files or more PSTs.

### Can a subscription upgrade speed up PST import?

Based on Microsoft's PST import documentation, **no paid add-on is described that raises import speed.** Upgrades change **how much a mailbox can hold**, not how fast imports run. The only paid route documented is **drive shipping**: Enterprise Agreement customers only, 7 to 10 business days after Microsoft receives the drive, up to 10 TB per drive and 10 drives per job.

If anyone tells you a license removes the throttle, ask for the exact SKU and written confirmation.

---

## Mailbox limits and upgrade costs

### Storage by plan

| Plan | User mailbox | Archive |
|---|---|---|
| Business Basic / Standard | 100 GB | 50 GB |
| Business Premium, Microsoft 365 E3 / E5 | 100 GB | 1.5 TB (auto-expanding) |
| Exchange Online Plan 1 | 50 GB | 50 GB |
| Exchange Online Plan 2 | 100 GB | 1.5 TB (auto-expanding) |
| Exchange Online Kiosk / F3 | 2 GB | None |

Source: [Microsoft Exchange Online limits](https://learn.microsoft.com/ro-ro/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits). Older third-party articles list different defaults, so **re-check Microsoft's page before acting**; plan limits change. Auto-expanding archives start at 100 GB and grow in increments up to 1.5 TB.

### List prices found (USD per user per month, annual term)

| Plan | Price | Source |
|---|---|---|
| Exchange Online Plan 1 | $4.00 | [o365hq](https://o365hq.com/blog/exchange-online-plan-1-vs-plan-2-vs-kiosk-vs-business-basic-mailbox-licensing) |
| Exchange Online Plan 2 | $8.00 | same |
| Exchange Online Archiving add-on (to Plan 1) | $3.00 | same |
| Business Basic (with Teams) | $7.00 | [AlphaBold](https://www.alphabold.com/microsoft-365-licensing-guide/) |
| Business Standard (with Teams) | $14.00 | same |
| Business Premium (with Teams) | $22.00 | same |
| Microsoft 365 E3 | $39.00 | same |

These are third-party summaries of list prices as of October 2026. Prices vary by region, currency, billing term and agreement. **Confirm in your Microsoft admin center or with your reseller** before quoting.

### Worked example: 100 users who need the larger archive

| Upgrade path | Extra cost per user per month | Extra cost for 100 users per year |
|---|---|---|
| Exchange Online Plan 1 to Plan 2 | $4 | **$4,800** |
| Plan 1 plus Archiving add-on | $3 | **$3,600** |
| Business Basic to Business Premium | $15 | **$18,000** |

These figures buy **capacity**. They are a recurring cost and, per Microsoft's documentation, they do not raise the per-mailbox import rate.

### Total cost of the manual route (worked example)

Example scenario: **100 mailboxes of 50 GB each**, moved by network upload, with admins who need the larger archive. Time comes from Microsoft's 24 GB per day figure; hours and delay cost are **planning assumptions** you should replace with your own numbers.

| Cost item | How to estimate | Worked example (100 mailboxes, 50 GB each) |
|---|---|---|
| Import time | Data per mailbox / 24 GB per day; the longest mailbox sets the finish date | 50 / 24 = about 2 days per mailbox; different mailboxes import in parallel, so about 2 days overall (calculated) |
| Admin hours | Upload, CSV prep, job monitoring, failed-job reruns | 40 hours (assumption): upload, 100-row CSV prep, validation, monitoring |
| Failure allowance | Reruns x hours per rerun | 10 reruns x 2 hours = 20 hours (assumption: about 1 in 10 jobs rerun) |
| License upgrade (if needed) | Section above: capacity cost per year | $4,800 a year for Exchange Online Plan 1 to Plan 2 at list prices found (100 users); $0 if mailboxes already fit your plan |
| Delay cost | Days of delay x business impact | 3 days of stall x $500 a day = $1,500 (illustrative assumption; use your own daily impact) |
| **Total** | Sum of the rows above | **60 admin hours + $4,800 a year in licenses + about $1,500 delay risk, about 2 days import time** |

Swap in your own mailbox count, sizes, hourly rate and delay impact to get a number you can take to a budget meeting. These figures are an illustration, not a measured result.

---

## Where a dedicated migration tool fits

> **Status: testing in progress.** The rows below are questions we are testing, not claims, and results will be published only once measured. Any tool that writes into Exchange Online remains subject to the service limits of the connection method it uses; we will test and state this plainly.

For a step-by-step guide to the Stellar approach, see how to [Import PST to Office 365 with Stellar Migrator for Outlook](https://www.stellarinfo.com/import-pst-to-office-365.php).

### Questions we are testing, one per gap

| Gap in the manual route | Question we are testing with Stellar Migrator for Outlook | Evidence needed |
|---|---|---|
| No PST health check | Does it detect or handle corrupt PSTs before migration? | Test with deliberately damaged PSTs |
| PSTs over 20 GB slow the import | Does it handle large PSTs without manual splitting? | Time and result at 20, 50, 100 GB |
| Opaque failures | Does it give clear progress, logs and error reasons? | Logs from failed and successful runs |
| Strict CSV and AzCopy workflow | Is the setup simpler than CSV plus AzCopy? | Step count and time to start |
| Same-mailbox queue | How does it behave with multiple PSTs per mailbox? | Time for 1, 3, 5 PSTs per user |
| Scaling to many users | Can it migrate many mailboxes in parallel? | Time for 10 and 50 mailboxes |
| Retention hold side effect | Does the tool set or avoid the hold? | Mailbox properties after migration |
| Re-runs after failure | Can it resume or retry without starting over? | Interrupt a run and resume |
| Skipped items | How does it handle items over 150 MB and deep folder trees? | Item counts before and after |

### Test plan (same method for both approaches)

| Variable | Values |
|---|---|
| PST size | 5, 20, 50, 100 GB |
| PSTs per mailbox | 1, 3, 5 |
| Target mailboxes | 1, 10, 50 |
| Approaches | A: Purview network upload / B: Stellar Migrator for Outlook |
| Runs | 3 each; publish average and range |
| Metrics | Total time, setup time, items in vs items migrated, skipped and failed items with reasons, errors, retries, manual steps |

Rules: isolated test tenant, copies only, same network and client for both approaches, fresh PST for every run, tenant and plan type recorded, reviewer sign-off. Raw data will be published in this repository under `/results`.

### Results

**How to read this table.** Network upload times are **calculated from Microsoft's documented rate of about 24 GB per day per mailbox** (not guaranteed), not measured by us. Stellar Migrator for Outlook rows will be filled only from our own measured runs, with logs and screenshots, using the test plan above. Until then they are marked "Test in progress".

| Scenario | Approach | Time | Items migrated | Skipped / failed | Notes |
|---|---|---|---|---|---|
| 20 GB, 1 mailbox | Network upload | About 1 day (20 / 24 GB per day, calculated) | Not measured | Items over 150 MB skipped; PSTs at the 20 GB recommended maximum | Analysis time before import is not documented by Microsoft |
| 20 GB, 1 mailbox | Stellar Migrator | Test in progress | Test in progress | Test in progress | Run 3 times, average and range published |
| 100 GB, 1 mailbox | Network upload | About 4 days (100 / 24 GB per day, calculated) | Not measured | PSTs over 20 GB can slow the import and Microsoft recommends splitting | Archive imports are capped at 100 GB per import |
| 100 GB, 1 mailbox | Stellar Migrator | Test in progress | Test in progress | Test in progress | Includes large-PST handling without manual splitting to be confirmed |
| 5 PSTs, 1 mailbox | Network upload | Sequential: total GB / 24 GB per day (for example 5 x 20 GB = 100 GB, about 4 days, calculated) | Not measured | Strict, case-sensitive CSV rows can fail validation | PSTs for the same mailbox import one after another |
| 5 PSTs, 1 mailbox | Stellar Migrator | Test in progress | Test in progress | Test in progress | Behavior with several PSTs per mailbox to be confirmed |
| 50 mailboxes | Network upload | Parallel across mailboxes: the largest mailbox sets the finish date (for example 100 GB, about 4 days, calculated) | Not measured | 500 rows per CSV; jobs stuck at Analysis in progress reported by administrators | Needs one CSV row per PST and retention hold cleanup afterward |
| 50 mailboxes | Stellar Migrator | Test in progress | Test in progress | Test in progress | Parallel migration of many mailboxes to be confirmed |

---

## How to choose a method

1. **Count and size your PSTs.** Over 20 GB each, or many per user? Expect queueing and splitting.
2. **Find your longest mailbox.** Data divided by 24 GB per day sets your earliest finish date for the manual route.
3. **Check archive targets.** More than 100 GB per archive import needs splitting; storage above plan limits needs an upgrade.
4. **Price the upgrade honestly.** It buys capacity, not speed.
5. **Price your own time.** Include reruns and troubleshooting, not just the license line.
6. **Run a pilot** of both approaches on a representative mailbox before committing.

---

## FAQ

### How do I import a PST file into Microsoft 365?

For admin-led imports, use Microsoft Purview network upload: get the SAS URL from the Import page, upload PSTs with AzCopy, build a mapping CSV, validate it, then run the import job. A single user can instead use Outlook's Import/Export or drag and drop.

### Does New-MailboxImportRequest work in Exchange Online?

No. Microsoft says it is not supported in Exchange Online and directs you to the network upload method. It works for on-premises Exchange.

### How long does a PST import to Microsoft 365 take?

Microsoft cites about 24 GB per day per mailbox, not guaranteed. A 100 GB import into one mailbox is roughly 4 days by that figure. Analysis time is not documented.

### Why is my PST import stuck on "Analysis in progress"?

Administrators report stalls of 24 hours or more on Microsoft Q&A. Community advice includes confirming uploads finished, splitting large PSTs, checking PST integrity and the network, and recreating the job. These are not Microsoft-confirmed fixes.

### What is the maximum PST size for Purview import?

Microsoft recommends 20 GB or less per PST. Items over 150 MB are skipped, and PSTs with 300 or more nested folder levels cannot be imported.

### Can I pay Microsoft to make PST import faster?

We found no documented paid add-on that raises import speed. Drive shipping is paid, but it is limited to Enterprise Agreement customers and takes 7 to 10 business days after Microsoft receives the drive.

### Does upgrading from Exchange Online Plan 1 to Plan 2 help with PST import?

It raises mailbox and archive capacity, not documented import speed. At the list prices we found, 100 users cost about $4,800 more per year.

### Does a PST import leave the mailbox on retention hold?

Per Microsoft, imports place the mailbox on retention hold indefinitely. Release it with `Set-Mailbox <identity> -RetentionHoldEnabled $false` or set an end date with `-EndDateForRetentionHold`.

---

## How we researched this

- **Method:** We read Microsoft Learn documentation for Purview PST import, network upload, `New-MailboxImportRequest`, and Exchange Online limits, plus public Microsoft Q&A threads and third-party price summaries, on **23 October 2026**.
- **What is Microsoft-sourced:** limits, speeds, steps and cmdlet behavior. **What is third-party:** list prices and the administrator failure reports. Both are labeled where they appear.
- **What is our arithmetic:** the import-time table, derived from Microsoft's 24 GB per day figure.
- **What is not yet first-hand:** Stellar Migrator for Outlook performance. Test results will be added only from measured runs, with raw data in this repository.
- **Limits:** Microsoft limits, prices and behavior change. The 24 GB per day figure is not a guarantee. Administrator reports are anecdotal and may reflect tenant-specific conditions.
- **Corrections:** Found an error or an outdated figure? [Open an issue](../../issues) and we will review it, correct the page, and update the "Last reviewed" date.

**About the author:** [Bharat Bhushan](https://www.linkedin.com/in/bharat24/) is a Product Manager at Stellar Information Technology with 15+ years of B2B SaaS marketing experience. [LinkedIn profile](https://www.linkedin.com/in/bharat24/)

**About the reviewer:** [Shelly Bhardwaj](https://www.linkedin.com/in/shelly-bhardwaj/) is an Exchange Online Migration Expert at Stellar Information Technology. [LinkedIn profile](https://www.linkedin.com/in/shelly-bhardwaj/)

---

## Sources

- [Microsoft Purview: PST import overview](https://learn.microsoft.com/en-us/purview/pst-import-overview)
- [Microsoft Purview: Use network upload to import PST files](https://learn.microsoft.com/purview/pst-import-network-upload)
- [Microsoft: New-MailboxImportRequest](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/new-mailboximportrequest)
- [Microsoft: Exchange Online limits](https://learn.microsoft.com/ro-ro/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits)
- [Microsoft Q&A: Purview PST import very slow, stuck on analysis](https://learn.microsoft.com/en-us/answers/questions/2244008/purview-pst-import-very-slow-seeminly-stuck-on-ana)
- [Microsoft Q&A: Stuck at 0% for half a day](https://learn.microsoft.com/en-us/answers/questions/1397044/how-to-stop-an-m365-pst-import-job-that-is-stuck-o)
- [Microsoft Q&A: Jobs stuck at Analysis in progress](https://learn.microsoft.com/en-us/answers/questions/1132972/import-jobs-to-office-365-mailbox-stuck-at-analysi)
- [Microsoft blog (2015, archived): HTTP flood mitigation and slow migrations](https://learn.microsoft.com/en-us/archive/blogs/onthewire/http-flood-mitigation-causing-slow-exchange-mailbox-migrations-to-office-365)
- [CodeTwo: Mailbox size in Microsoft 365](https://www.codetwo.com/admins-blog/mailbox-size-microsoft-365/)
- [o365hq: Exchange Online Plan 1 vs Plan 2 prices and sizes](https://o365hq.com/blog/exchange-online-plan-1-vs-plan-2-vs-kiosk-vs-business-basic-mailbox-licensing)
- [AlphaBold: Microsoft 365 licensing guide 2026](https://www.alphabold.com/microsoft-365-licensing-guide/)

*Microsoft, Microsoft 365, Outlook and Exchange Online are trademarks of Microsoft Corporation. This article is not affiliated with or endorsed by Microsoft.*
