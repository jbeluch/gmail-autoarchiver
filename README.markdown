GMail Auto-Archiver
===================

This script will archive emails in your inbox that are older than a
specified number of days.

Usage
-----

1. Set up filters in gmail matching the pattern `aa:\d+` where the
   `\d+` is the age limit in days, e.g. `aa:3`.
2. *Optional*: Enter you email address below for the variable
   `EMAIL_ADDRESS`. If you skip this step, you will be prompted to
   enter your email address interactively when the script runs.
3. Execute the script. If you are running it for the first time, you
   will be required to authorize the script's oauth access in a web
   browser.

The script will print the subject line for any emails that are
auto-archived.

*Note*: Once you authorize the oauth token/secret, they are saved to
disk at `OAUTH_PATH`. If the token/secret no longer work, simply remove
the file at `OAUTH_PATH`. The next time the script is run it will set
up a new token/secret.
