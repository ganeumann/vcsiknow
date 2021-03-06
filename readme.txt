VCs I Know

vcsiknow.py
Jerry Neumann, June 2014

Pulls all email addresses from your IMAP email account and creates a spreadsheet of all
the people whose emails you have that work for a VC firm. Can take a while to run if you
have a lot of email.

Usage:

python vcsiknow.py userid password imapserver outputFileName

Right now it just searches the inbox. If you have a lot of email not in the inbox, you
will have to modify line 89 to search a different mail folder. I'm hoping to add that
functionality to the program itself, but I need to investigate how different systems
specify folder names first.

The output file will be vcsiknow.csv in your current working directory if not specified.

The output file is a tab-delimited csv, suitable for importing into a spreadsheet. The
format is: name, firm, email address, date last corresponded with, number of times
corresponded with.
 
This program uses a list of VC:domain mappings kept in a Google spreadsheet. This means
that anyone whose email has the VC domain will be surfaced, including non-VCs, and even
non-people. The name field os the name as it appears in the email address, so some emails
won't have a name associated with the address, or the address may also appear as the name.

I excluded domains whose primary business is not VC, even if there are some VCs who work
there (most of the banks, for instance.)

In sum, the data needs to be cleaned even after running the program. I hope it's useful
anyway.

If you find a VC domain that is missing, enter it into this form: http://goo.gl/aTwlNY
I will put it in the spreadsheet.


License:

You have my permission to use this code however you want, for free, forever. Be warned,
however, that I am not a professional programmer. In fact, I'm not even a very good
programmer. Actually, I suck. So if the code doesn't work or if it works wrong or even
if it causes your computer to be invaded by thieves and used to create Skynet, or
anything else bad that happens whatsoever, you attest that that's exactly what you
expected to happen, because you took the responsibility to look at the code yourself
and see what it does, so none of those bad things is any surprise at all to you.

If you don't feel entirely confident using this program, don't.