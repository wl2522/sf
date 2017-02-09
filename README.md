# sf

A Python 3 program that takes a Salesforce media plan in .csv format and converts the data into an invoice summary, which is the standard document format used by the Media Delivery Ops team at WebMD.

Categories are automatically assigned to each ad by searching for identifying terms in the ad name.

Warnings are displayed upon encountering placements with flat fee pricing or placeholder placements with no dollar value.

A standalone executable program that doesn't require an existing installation of Python is also included.
