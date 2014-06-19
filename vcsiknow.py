# VCsIKnow
#
# Pulls all email addresses from your IMAP email account and creates a spreadsheet of all
# the people whose emails you have that work for a VC firm. Can take a while to run if you
# have a lot of email.
# 
# This program uses a list of VC:domain mappings kept in a Google spreadsheet. If you find
# one is missing, enter it into this form: http://goo.gl/aTwlNY
# I will put it in the spreadsheet.
#
# License:
#
# You have my permission to use this code however you want, for free, forever. Be warned,
# however, that I am not a professional programmer. In fact, I'm not even a very good
# programmer. Actually, I suck. So if the code doesn't work or if it works wrong or even
# if it causes your computer to be invaded by thieves and used to create Skynet, or
# anything else bad that happens whatsoever, you attest that that's exactly what you
# expected to happen, because you took the responsibility to look at the code yourself
# and see what it does, so none of those bad things is any surprise at all to you.
#
# If you don't feel entirely confident using this program after reading that, don't.
#

import csv, imaplib, email, StringIO, urllib2, sys, os

outfile = os.getcwd() + "/vcsiknow.csv"
if (len(sys.argv) == 4):
	userid = sys.argv[1]
	password = sys.argv[2]
	imapserver = sys.argv[3]
elif (len(sys.argv) == 5):
	outfile = sys.argv[4]
else:
	print """
vcsiknow:

vcsiknow gets all addresses from your IMAP email and compares their domains to a list
of venture capital firm domains. All matches are put in a .csv file for easy manipulation
in a spreadsheet.

Usage:
python vcsiknow.py userid password imap outfile

userid is your email userID. password is your email password. If you have two-factor authentication
turned on in Gmail (which you should), you'll need to get an 'app password.'

imap is your imap server. Gmail's is 'imap.gmail.com'

outfile is where the results are written. If no outfile is specified, output is put into vcsiknow.csv.
"""
	sys.exit(0)

def split_addrs(s):
    #split an address list into list of tuples of (name,address)
    if not(s): return []
    outQ = True
    cut = -1
    res = []
    for i in range(len(s)):
        if s[i]=='"': outQ = not(outQ)
        if outQ and s[i]==',':
            res.append(email.utils.parseaddr(s[cut+1:i]))
            cut=i
    res.append(email.utils.parseaddr(s[cut+1:i+1]))
    return [(nm,addr.lower()) for nm, addr in res]

# get the list of vc domains from Google spreadsheets, put it into doms={domain: name, ...}

# change to standard libraries only: requests->urllib2
# res=requests.get("http://docs.google.com/feeds/download/spreadsheets/Export?key=11URwkQ3cXgOvPjb23DhmWFWKYMx6ggKztV7w-Bc_mVw&exportFormat=csv&gid=1634616604").content
res = urllib2.urlopen("http://docs.google.com/feeds/download/spreadsheets/Export?key=11URwkQ3cXgOvPjb23DhmWFWKYMx6ggKztV7w-Bc_mVw&exportFormat=csv&gid=1634616604").read()
cr = csv.DictReader(StringIO.StringIO(res))

domains = {i['Domain']:i['Name'] for i in cr}   # doms={vc domain: firm name,...}

# get all the emails
mail=imaplib.IMAP4_SSL(imapserver)
mail.login(userid,password)

# folders=mail.list()					# need to parse folders before this will work
# show all folders, have the user select one
# print "Available folders:"
# for i in range(len(folders)):
# 	print i+1,". ",folders[i]

# folder = raw_input("Which folder (1 to "+len(folders)+")? ")
# if folder<1 or folder>len(folders): raise Exception("No such folder.")

mail.select("INBOX")
# mail.select(folders[folder-1])
result,data=mail.search(None,"ALL")
ids=data[0].split()
print "Fetching "+str(len(ids))+" emails."
msgs = mail.fetch(','.join(ids),'(BODY.PEEK[HEADER])')[1][0::2]
print "Fetched. Parsing emails."

# get all the tos and froms and ccs
addrlist=[]
for x,msg in msgs:
    msgobj = email.message_from_string(msg)
    tmpaddrlist = []
    emlrawdate=email.utils.parsedate(msgobj['date'])
    emldate = "%d/%d/%d" % (emlrawdate[1],emlrawdate[2],emlrawdate[0])
    
    # parse all to, from and cc emails addresses into list of (name, address) tuples
    tmpaddrlist.extend(split_addrs(msgobj['to']))
    tmpaddrlist.extend(split_addrs(msgobj['from']))
    tmpaddrlist.extend(split_addrs(msgobj['cc']))

    # add date to each tuple
    addrlist.extend([(n,e,emldate) for n,e in tmpaddrlist])

print "Processing "+str(len(addrlist))+" email addresses."

vcs={}			# vcs={address: name,...}
vcemlcnt={}		# vcemlcnt={address: # of times found,...}
vclasteml={}	# vclsteml={address: last time emailed,...}

for address in addrlist:
	name,emailaddr,emldate = address
	if '@' not in emailaddr: continue
	domain = emailaddr.split('@')[1]
	if domain in domains:
		if vcs.has_key(emailaddr) and vcs[emailaddr]:
			vcemlcnt[emailaddr] +=1
		else:
			vcs[emailaddr] = name.strip("'")
			vcemlcnt[emailaddr] = 0
		vclasteml[emailaddr] = emldate
		
with open(outfile,'w') as f:
	f.write("Name\tFirm\tEmail\tLast Email\tTimes Found\n")
	i = 0
	for eml in vcs:
		# to file: person name, firm name, person email address, last emailed, # of emails
		name = vcs[eml]
		domain = eml.split('@')[1]
		firm = domains[domain]
		l = name+"\t"+firm+"\t"+eml+"\t"+vclasteml[eml]+"\t"+str(vcemlcnt[eml])+"\n"
		i += 1
		f.write(l)
		
print "Finished. "+str(i)+" records written to "+outfile
	
	