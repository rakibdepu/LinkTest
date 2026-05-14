# classify_profile_sites.py Explanation Log

This document explains the purpose, features, workflow, logic, and code structure of `classify_profile_sites.py`.

## 1. Purpose

The script reads website domains from column `A` of:

- `Large Free Profile Submission Site List.xlsx`

Then it checks each site and writes:

- column `B`: registration/access result
- column `C`: website type guess

The current goal is to identify whether a site:

- supports user registration for public profiles
- only shows login without confirmed registration
- redirects to another domain
- is expired or parked
- is unreachable
- is only an admin login page and should be ignored

## 2. Main Features

The script currently supports all of these features:

- Excel workbook reading and writing with `openpyxl`
- visible-text based auth detection
- auth-path probing such as `/login`, `/register`, `/signup`, `/join`
- admin-login filtering
- redirect detection
- expired/parked domain detection
- unreachable-domain detection
- website type classification
- `requests` mode
- headless `selenium` mode
- `hybrid` mode
- dry-run mode
- interactive menu mode
- random sample testing
- last-N incomplete row processing
- exact row range processing
- forced row range reruns even when rows are already filled
- stop immediately if workbook is locked
- periodic save checkpoints

## 3. High-Level Workflow

The script works in this order:

1. Parse command-line options
2. Open the workbook
3. Build the list of rows to process
4. Select fetch mode: `requests`, `selenium`, or `hybrid`
5. Load each site
6. Extract visible text
7. Search for login/register signals
8. Probe common auth URLs if enabled
9. Detect redirect/admin-login/expired/unreachable conditions
10. Build the final result text for column `B`
11. Guess site type for column `C`
12. Save progress every 25 rows
13. Save again at the end

## 4. Core Configuration

Important constants near the top of the script:

- `WORKBOOK_PATH`
  - Excel file path
- `USER_AGENT`
  - browser-like user agent string
- `KEYWORDS`
  - `Login`, `Sign In`, `Register`, `Sign Up`
- `ADMIN_HINTS`
  - known admin-only phrases and admin-style login indicators
- `AUTH_PATHS`
  - common user auth endpoints
- `TYPE_PATTERNS`
  - regex rules for website category guessing
- `REQUEST_TIMEOUT`
  - request/browser timeout
- `SAVE_EVERY`
  - checkpoint frequency

## 5. FetchResult Object

The `FetchResult` dataclass stores one page-load result:

- `final_url`
  - URL after redirects
- `html`
  - raw page HTML
- `visible_text`
  - stripped or rendered visible content
- `status`
  - `OK`, `HTTP 403`, `Timeout`, etc.
- `engine`
  - `requests` or `selenium`

This object is passed through most detection logic.

## 6. URL Preparation

`normalize_url()` cleans the original value from Excel and builds candidate URLs.

Example:

- `example.com`
  - becomes:
  - `https://example.com`
  - `http://example.com`

This improves success rate when the sheet contains plain domains only.

## 7. Requests Mode

`fetch_requests()` uses standard HTTP requests.

It:

- sends browser-like headers
- downloads page content
- decodes response text
- strips HTML into visible text
- returns a `FetchResult`

Pros:

- faster
- lightweight

Cons:

- weaker on JavaScript-rendered sites

## 8. Selenium Mode

`create_webdriver()` creates a headless Chrome browser.

`fetch_selenium()` then:

- opens the page in the browser
- waits for JS rendering if needed
- reads `page_source`
- grabs visible body text
- returns a `FetchResult`

Pros:

- better for JS-heavy sites
- better for rendered visible text

Cons:

- slower

The browser is headless, so no visible UI is shown.

## 9. Hybrid Mode

Hybrid mode combines speed and accuracy.

How it works:

1. first tries `requests`
2. if the page is sparse or looks JS-heavy
3. and no useful auth signal is found
4. it retries with Selenium

When hybrid escalates to Selenium, the output in column `B` includes:

- `fetched with Selenium`

This helps show which sites needed the browser fallback.

## 10. Visible Text Extraction

`strip_html()` removes:

- scripts
- styles
- tags
- extra whitespace

This is used mainly for `requests` mode, where raw HTML must be converted into a visible-text approximation.

In Selenium mode, visible text is taken directly from the rendered page body.

## 11. Keyword Detection

`detect_keywords_from_text()` searches visible text for:

- `login`
- `sign in`
- `register`
- `sign up`

It also checks the final page path for common auth URLs, such as:

- `/login`
- `/signin`
- `/register`
- `/signup`
- `/join`

This returns a list of matched signals, such as:

- `["Login", "Register"]`
- `["Sign In", "Sign Up"]`

## 12. Admin Login Detection

`looks_like_admin_login()` filters out admin-only pages.

It checks:

- visible text
- raw HTML
- final URL path

Examples of admin-only hints:

- `admin login`
- `wp-admin`
- `wp-login`
- `/administrator`
- `cpanel`
- `backend login`

If detected, the site is classified as:

- `Admin login only, ignored`

This prevents false positives where the script sees login UI that is not meant for public user profile creation.

## 13. Base Domain and Redirect Logic

`get_registered_domain()` extracts a simplified domain name from a URL.

`classify_registration_status()` compares:

- original source domain
- final destination domain

If they differ, the script labels the result as:

- `Redirects to ...`

This helps identify sites that no longer resolve to their original target.

## 14. Auth Path Probing

`probe_auth_paths()` is used when `--scan-auth-paths` is enabled.

It checks likely auth routes such as:

- `/login`
- `/signin`
- `/register`
- `/signup`
- `/join`

Why this matters:

- some homepages do not show registration text
- but the registration page still exists

If this probing finds new auth signals, the script adds a note like:

- `auth-path match: Register, Sign Up`

## 15. JS-Heavy Detection

`should_flag_js_heavy()` checks if the page looks like a JavaScript app with little visible text.

Markers include:

- `__next`
- `id="app"`
- `id="root"`
- `ng-app`
- `data-reactroot`

This is especially important for hybrid mode, because it helps decide when Selenium should retry the page.

## 16. Selenium Retry Logic

`should_retry_with_selenium()` controls hybrid fallback.

Hybrid retries with Selenium when:

- requests mode returned `OK`
- no useful auth signal was found
- and the page looks JS-heavy or nearly empty

This saves time by avoiding Selenium on every site.

## 17. Expired or Parked Domain Detection

`looks_expired_or_parked()` checks for domain parking and expired-domain text such as:

- `domain for sale`
- `buy this domain`
- `this domain is for sale`
- `expired domain`
- `domain parked`
- `hugedomains`
- `sedo domain parking`

If matched, the script labels the site:

- `Domain expired/parked`

## 18. Registration Status Logic

The most important final decision happens in `classify_registration_status()`.

It applies logic in this order:

1. redirect check
2. unreachable-domain check
3. expired/parked check
4. admin-login check
5. public registration check
6. login-only check
7. page-not-found check
8. blocked/restricted check
9. other non-OK status
10. no-registration-found fallback

Possible output labels include:

- `User registration available`
- `Login found, registration not confirmed`
- `Redirects to example.com`
- `Domain expired/parked`
- `Domain unreachable`
- `Page not found`
- `Blocked or restricted`
- `Admin login only, ignored`
- `No registration found`

## 19. Column B Format

Column `B` is not just one label. It is a compact audit string.

Typical structure:

1. main outcome
2. optional raw signal details
3. optional HTTP/error status
4. optional auth-path note
5. optional Selenium fallback note

Examples:

- `User registration available | Signals: Register, Sign Up`
- `Redirects to example.com | fetched with Selenium`
- `Admin login only, ignored | Signals: Login, Sign In`
- `Domain unreachable | URL error: ...`

## 20. Website Type Classification

`classify_site()` writes column `C`.

It uses regex-based heuristics to guess categories such as:

- `Forum`
- `Social Network`
- `Blog`
- `Business Directory`
- `Ecommerce`
- `News/Media`
- `Education`
- `Portfolio/Creative`
- `Wiki/Knowledge`
- `Job Board`
- `Real Estate`
- `General Website`

It works by checking the URL and page content for matching keywords.

This is helpful, but it is only a heuristic guess.

## 21. Interactive Menu

When the script is run with:

```powershell
python classify_profile_sites.py --menu
```

it shows prompts for:

- fetch mode
- auth-path scanning
- stop-if-locked
- JS-heavy handling
- sample size
- random seed
- Selenium render wait

This is useful when you do not want to remember long commands.

## 22. CLI Options

Current command-line options:

- `--stop-if-locked`
  - stop immediately if workbook save fails

- `--sample-size`
  - process only a random sample of rows

- `--seed`
  - control repeatable random sampling

- `--scan-auth-paths`
  - check common auth routes

- `--js-mode`
  - JS handling behavior

- `--fetch-mode`
  - `requests`, `selenium`, or `hybrid`

- `--render-wait`
  - Selenium wait time after page load

- `--menu`
  - open the interactive menu

- `--dry-run`
  - do not save results to Excel

- `--last-n`
  - process last N incomplete rows

- `--row-start`
  - start row for range targeting

- `--row-end`
  - end row for range targeting

- `--force-row-range`
  - rerun the exact specified row range even if column `B` and `C` already contain values

## 23. Row Selection Logic

The script normally skips rows where both:

- column `B` has a value
- column `C` has a value

Then it optionally filters that row list using:

- sample mode
- last-N mode
- row range mode

Important behavior:

- `--row-start` and `--row-end` normally target only incomplete rows
- `--force-row-range` overrides that and forces exact rerun of those rows

## 24. Resume Behavior

The script is resumable because:

- rows already filled in both `B` and `C` are skipped
- progress is saved every 25 rows

If a run stops midway, running it again continues from the next unfinished row, unless forced row rerun options are used.

## 25. Lock Handling

If the workbook is open in Excel, saving may fail.

When run with:

- `--stop-if-locked`

the script stops immediately if saving the workbook fails because the file is locked.

This avoids partial continuation into alternate files.

## 26. Dry Run Mode

`--dry-run` performs all checking logic but does not write anything to the workbook.

This is useful for:

- testing speed
- testing new logic
- comparing `requests`, `selenium`, and `hybrid`
- inspecting printed output safely

## 27. Typical Commands

### Full hybrid run

```powershell
python classify_profile_sites.py --fetch-mode hybrid --scan-auth-paths --stop-if-locked
```

### Interactive mode

```powershell
python classify_profile_sites.py --menu
```

### Safe sample run

```powershell
python classify_profile_sites.py --sample-size 20 --seed 42 --fetch-mode hybrid --scan-auth-paths --dry-run
```

### Exact row rerun

```powershell
python classify_profile_sites.py --fetch-mode hybrid --scan-auth-paths --stop-if-locked --row-start 4414 --row-end 4513 --force-row-range
```

### Requests-only run

```powershell
python classify_profile_sites.py --fetch-mode requests --scan-auth-paths --stop-if-locked
```

### Selenium-only run

```powershell
python classify_profile_sites.py --fetch-mode selenium --scan-auth-paths --stop-if-locked
```

## 28. Strengths

The current script is strong in these areas:

- practical workbook automation
- resume-friendly execution
- visible-text oriented detection
- admin-login filtering
- redirect recognition
- parked/unreachable detection
- JS-capable scanning
- speed/accuracy balance through hybrid mode
- safe testing through dry-run mode

## 29. Known Limitations

Current limitations to keep in mind:

- domain extraction is simplified and may be imperfect for some country-code domains
- website type detection is heuristic only
- auth-page existence does not always guarantee public signup is truly open
- Selenium is slower
- JS detection is heuristic, not perfect

## 30. Best Practical Use

For most real runs, the best balance is:

```powershell
python classify_profile_sites.py --fetch-mode hybrid --scan-auth-paths --stop-if-locked
```

For rerunning a known block:

```powershell
python classify_profile_sites.py --fetch-mode hybrid --scan-auth-paths --stop-if-locked --row-start 4414 --row-end 4513 --force-row-range
```

For testing changes first:

```powershell
python classify_profile_sites.py --sample-size 10 --fetch-mode hybrid --scan-auth-paths --dry-run
```

## 31. Summary

This script has evolved from a simple login-text checker into a more useful profile-submission analyzer.

Its real job now is:

- identify public registration opportunities
- ignore admin-only login screens
- detect redirections
- detect dead/parked/unreachable domains
- classify site type
- write clear results back into Excel

That makes it much closer to the actual decision you need when evaluating profile-submission opportunities.
