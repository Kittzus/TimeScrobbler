# TimeScrobbler
=============

## What?
TimeScrobbler is a simple PowerShell tool designed to help answer that age-old question of "What the heck was I actually doing on that arbitrary day three weeks ago?"

When provided with a date-range, Timescrobbler will pull information from common & user specified folders, Outlook and Slack together into an HTML report for each day within that range, summarising exactly who you spoke to, what was said, what meetings you attended, and what files you worked on or downloaded that day.


## Why?
TimeScrobbler was primarily created to help me handle filling out back-dated timesheets. By having a single sheet summary of each day's activity, I can more easily & accurately attribute my time, long after I've forgotten where I even was five weeks ago. Hopefully others might find it useful too!


## Instructions
* Download the repository
* Unblock the zip & extract
* Edit the User.conf file to reflect your report output directory, common working directories, and Slack details (if desired)
* Run .\New-TimeScrobble.ps1 as follows:

    ```powershell.exe -File <path>\New-TimeScrobble.ps1 -startDay <datestring> -endDay <datestring>```
    **IMPORTANT:** ```<datestring>``` must be in sortable format yyyy-MM-dd e.g. 2016-08-01 for 8th August 2016


### Prerequisites
* PowerShell 3 or later
* For optional Slack functions:
  * The excellent [PSSlack](https://github.com/RamblingCookieMonster/PSSlack) module by [RamblingCookieMonster](https://github.com/RamblingCookieMonster) must be installed
  * A [valid test token](https://api.slack.com/docs/oauth-test-tokens) from the Slack you wish to scrobble
  
  
## ToDo
* Better input validation
* Improve output report formatting with CSS wizardry
* Build XAML interface
* Add additional data sources