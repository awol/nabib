This project extracts information from the Internet Banking site for the National Australia Bank
and produces a QIF file for each of the accounts in the users profile. It is written using the
Python programming language and scrapes the data from the website by parsing the HTML it produces.

This means that the tools are tightly coupled to the structure of the website and changes to the website can cause the utilities to break suddenly.

We use this tool to create QIF files that we subsequently import into GnuCash to manage our financial accounts.
