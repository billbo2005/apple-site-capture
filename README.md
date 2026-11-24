# Project Apple · Site Capture

Offline-first web app for logging progress room by room on UCB Project Apple: each activity in a room has a 0 to 100% slider, and 100% closes it.
Address: https://sitecapture.projectcentral.org/
Data lives in Project Central's SharePoint (site *Project Apple Site Capture*: lists **Rooms**, **WorksToGo**, library **Photos**).
Sign-in is Project Central Microsoft 365 (Entra app *Apple Site Capture*). No data is stored in this repository.

Prepared by Project Central Assurance.


**v1.2.1 (24 Nov 2026)** – changes waiting on a phone are sent first, before the list is refreshed, so a slow or stalled refresh cannot hold them up; a request that gets no answer gives up after 90 seconds and is retried; the chip shows 'Updating list…' with a count; the room list is read on every sync so new rooms reach every phone; an app update no longer forces the whole list to be read again.

**v1.2.0 (24 Nov 2026)** – activity list brought into line with the tracker: BMS is now BMS 1st Fix and BMS 2nd Fix per room (Containment, Cabling, Panel Terminations, Field Terminations and IO Test no longer per room), and Walls 1st Fix and Walls 2nd Fix are added. The app and the tracker work side by side: an app figure is used by the tracker only if it was logged in the week.
