# yummeroo
Find saved recipes easily with the yummeroo app

## How the website updates

`.github/workflows/recipe-sync.yml` runs every 6 hours. It downloads the
"Yummeroo Recipes" spreadsheet, runs `scripts/yummeroo_sync.py`, and if the
recipes changed, updates `index.html` on the `gh-pages` branch. To run it now,
open the Actions tab, choose "Recipe sync", then "Run workflow".

© 2026 Lavanya Krishna. All rights reserved.
