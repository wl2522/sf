# sf

This is a Python 3 program that takes a Salesforce media plan in .csv format and converts the data into an invoice summary, which is the standard document format used by the Media Delivery Ops team at WebMD. The Finance department also uses these documents to generate the invoices that are sent to clients.

It uses the OpenPyXL package to interact with Microsoft Excel files. A standalone executable program that doesn't require an existing installation of Python was also created using the PyInstaller package. The original version of this program is included due to later versions of OpenPyXL being incompatible with PyInstaller.

Media plans come in various formats, but are meant to show which ad products are being purchased, how much is being spent on each ad placement, what rate is being charged for each type of ad, and the date range during which each ad will be displayed.

Ad placements on the invoice summary are grouped such that ads of the same type, date range, and reservation note belong in the same group. Reservation notes help to distinguish different ad placements that are of the same type and are running within the same date range. For example, this often occurs in campaigns that target different states with the same ads.

Ad categories are automatically assigned to each ad by searching for identifying terms in the ad name.

Warnings are displayed upon encountering placements with flat fee pricing or placeholder placements with no dollar value.

An example media plan is included to demonstrate how this program works. A media plan of this size and complexity could take 20 or more minutes to convert into the invoice summary format due to the many combinations of date ranges, placement types, and reservation notes that are included.
